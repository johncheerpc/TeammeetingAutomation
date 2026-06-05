"""
Resolve the Teams meeting joinWebUrl from an encrypted Microsoft Graph
recording/transcript change notification delivered via an Azure Storage Queue.

Flow
----
1. Read the queue message (a Graph "rich" notification with encryptedContent).
2. Decrypt the resource data (Microsoft Graph scheme: RSA-OAEP -> HMAC verify -> AES-CBC).
3. Pull meetingId + organizer id out of the decrypted callTranscript/callRecording.
4. GET /users/{organizerId}/onlineMeetings/{meetingId} to read joinWebUrl.

Dependencies
------------
    pip install cryptography requests msal azure-functions

Required environment variables / app settings
---------------------------------------------
    TENANT_ID            Entra tenant id
    CLIENT_ID            App registration (client) id
    CLIENT_SECRET        App registration client secret
    GRAPH_PRIVATE_KEY    PEM private key matching the encryptionCertificate you
                         supplied when creating the subscription. Either the PEM
                         text itself, or a path to a .pem file.
    GRAPH_PRIVATE_KEY_PASSWORD   (optional) passphrase for the PEM key
    EXPECTED_CLIENT_STATE        (optional) the clientState you set on the subscription

Permissions
-----------
App-only token needs OnlineMeetings.Read.All (to read the meeting) and
OnlineMeetingRecording.Read.All / OnlineMeetingTranscript.Read.All for the
notification scope. Reading another user's onlineMeeting in app-only context
also requires an application access policy granting your app rights over that
organizer (Set-CsApplicationAccessPolicy).
"""

import base64
import hashlib
import hmac
import json
import logging
import os
from functools import lru_cache
from typing import Optional

import requests
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import padding as asym_padding
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.primitives.hashes import SHA1
from cryptography.hazmat.primitives.padding import PKCS7

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# --------------------------------------------------------------------------- #
# 1. Private key loading
# --------------------------------------------------------------------------- #
@lru_cache(maxsize=1)
def _load_private_key():
    raw = os.environ["GRAPH_PRIVATE_KEY"]
    if not raw.lstrip().startswith("-----BEGIN"):
        # treat as a path to a .pem file
        with open(raw, "rb") as fh:
            pem_bytes = fh.read()
    else:
        pem_bytes = raw.encode("utf-8")

    password = os.environ.get("GRAPH_PRIVATE_KEY_PASSWORD")
    return serialization.load_pem_private_key(
        pem_bytes,
        password=password.encode("utf-8") if password else None,
    )


# --------------------------------------------------------------------------- #
# 2. Decryption of the rich-notification resource data
# --------------------------------------------------------------------------- #
def decrypt_resource_data(encrypted_content: dict) -> dict:
    """
    Decrypt one notification item's `encryptedContent` block into the resource
    JSON (a callTranscript or callRecording object).

    encrypted_content has: data, dataKey, dataSignature (all base64 strings).
    """
    private_key = _load_private_key()

    encrypted_sym_key = base64.b64decode(encrypted_content["dataKey"])
    encrypted_payload = base64.b64decode(encrypted_content["data"])
    expected_signature = base64.b64decode(encrypted_content["dataSignature"])

    # 2a. Unwrap the AES symmetric key with the RSA private key (OAEP / SHA-1).
    # If decryption fails with a key created for SHA-256 OAEP, swap SHA1() -> SHA256().
    symmetric_key = private_key.decrypt(
        encrypted_sym_key,
        asym_padding.OAEP(
            mgf=asym_padding.MGF1(algorithm=SHA1()),
            algorithm=SHA1(),
            label=None,
        ),
    )

    # 2b. Verify integrity: HMAC-SHA256 over the *encrypted* bytes, keyed by the
    #     symmetric key. Reject tampered payloads.
    actual_signature = hmac.new(
        symmetric_key, encrypted_payload, hashlib.sha256
    ).digest()
    if not hmac.compare_digest(actual_signature, expected_signature):
        raise ValueError("dataSignature mismatch - notification may be tampered with")

    # 2c. AES-CBC decrypt. IV = first 16 bytes of the symmetric key. PKCS7 padding.
    iv = symmetric_key[:16]
    cipher = Cipher(algorithms.AES(symmetric_key), modes.CBC(iv))
    decryptor = cipher.decryptor()
    padded = decryptor.update(encrypted_payload) + decryptor.finalize()

    unpadder = PKCS7(algorithms.AES.block_size).unpadder()
    plaintext = unpadder.update(padded) + unpadder.finalize()

    return json.loads(plaintext.decode("utf-8"))


# --------------------------------------------------------------------------- #
# 3. Graph app-only token + meeting lookup
# --------------------------------------------------------------------------- #
@lru_cache(maxsize=1)
def _msal_app():
    import msal

    return msal.ConfidentialClientApplication(
        client_id=os.environ["CLIENT_ID"],
        client_credential=os.environ["CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{os.environ['TENANT_ID']}",
    )


def _get_token() -> str:
    result = _msal_app().acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Token request failed: {result.get('error')} - "
            f"{result.get('error_description')}"
        )
    return result["access_token"]


def get_join_web_url(organizer_id: str, meeting_id: str) -> Optional[str]:
    """GET the online meeting and return its joinWebUrl."""
    token = _get_token()
    url = f"{GRAPH_BASE}/users/{organizer_id}/onlineMeetings/{meeting_id}"
    resp = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "id,subject,joinWebUrl"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json().get("joinWebUrl")


# --------------------------------------------------------------------------- #
# 4. Orchestration: queue message -> joinWebUrl
# --------------------------------------------------------------------------- #
def _iter_notifications(payload: dict):
    """A queue message may be a single notification or a {"value": [...]} batch."""
    if isinstance(payload, dict) and "value" in payload:
        yield from payload["value"]
    else:
        yield payload


def handle_message(message_body: str) -> list[dict]:
    """
    Parse one Azure Queue message, decrypt each notification, and resolve the
    joinWebUrl. Returns a list of result dicts (one per notification item).
    """
    payload = json.loads(message_body)
    expected_state = os.environ.get("EXPECTED_CLIENT_STATE")
    results = []

    for note in _iter_notifications(payload):
        # Defence-in-depth: validate clientState if you set one on the subscription.
        if expected_state and note.get("clientState") != expected_state:
            logger.warning("Skipping notification with unexpected clientState")
            continue

        encrypted_content = note.get("encryptedContent")
        if not encrypted_content:
            logger.warning("Notification has no encryptedContent; is it a rich sub?")
            continue

        resource = decrypt_resource_data(encrypted_content)

        meeting_id = resource.get("meetingId")
        organizer_id = (
            resource.get("meetingOrganizer", {}).get("user", {}).get("id")
        )

        # meetingId is null for ad hoc (non-scheduled) calls - no onlineMeeting,
        # hence no joinWebUrl. Fall back to callId for those.
        if not meeting_id:
            logger.info(
                "Ad hoc call (meetingId is null); callId=%s, no joinWebUrl.",
                resource.get("callId"),
            )
            results.append(
                {
                    "callId": resource.get("callId"),
                    "meetingId": None,
                    "organizerId": organizer_id,
                    "joinWebUrl": None,
                }
            )
            continue

        if not organizer_id:
            raise ValueError(
                "No meetingOrganizer.user.id in payload - cannot address the "
                "meeting in app-only context."
            )

        join_web_url = get_join_web_url(organizer_id, meeting_id)
        logger.info("Resolved joinWebUrl for meeting %s", meeting_id)

        results.append(
            {
                "recordingOrTranscriptId": resource.get("id"),
                "meetingId": meeting_id,
                "organizerId": organizer_id,
                "callId": resource.get("callId"),
                "contentCorrelationId": resource.get("contentCorrelationId"),
                "joinWebUrl": join_web_url,
            }
        )

    return results


# --------------------------------------------------------------------------- #
# 5. Azure Functions queue trigger (Python v2 programming model)
# --------------------------------------------------------------------------- #
# If you run this as an Azure Function, uncomment the block below. Otherwise
# call handle_message(...) from your own queue consumer.
#
# import azure.functions as func
#
# app = func.FunctionApp()
#
# @app.queue_trigger(
#     arg_name="msg",
#     queue_name="recording-notifications",
#     connection="AzureWebJobsStorage",
# )
# def process_notification(msg: func.QueueMessage) -> None:
#     body = msg.get_body().decode("utf-8")
#     for r in handle_message(body):
#         logging.info("joinWebUrl=%s meetingId=%s", r["joinWebUrl"], r["meetingId"])


if __name__ == "__main__":
    # Local test: pass the queue message JSON as a string.
    import sys

    logging.basicConfig(level=logging.INFO)
    body = sys.stdin.read() if not sys.stdin.isatty() else "{}"
    print(json.dumps(handle_message(body), indent=2))
