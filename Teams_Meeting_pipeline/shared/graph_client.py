# =============================================================================
# FILE: shared/graph_client.py
#
# PURPOSE:
#   This file contains all the functions that talk to the Microsoft Graph API.
#
# WHAT IS MICROSOFT GRAPH API?
#   It is Microsoft's single gateway to access all Microsoft 365 data —
#   Teams meetings, transcripts, recordings, users, calendars, etc.
#   Every request goes to: https://graph.microsoft.com/
#
# TWO VERSIONS OF THE API:
#   - v1.0  (GRAPH_BASE) → stable, production-ready features
#   - beta  (GRAPH_BETA) → newer features (transcripts/recordings are still beta)
#   We use beta for transcripts and recordings because they are not yet in v1.0.
#
# HOW API CALLS WORK:
#   1. We get an access token from auth.py (Managed Identity)
#   2. We put the token in the HTTP "Authorization" header
#   3. We make an HTTP request (GET, POST, PATCH, DELETE) to a Graph URL
#   4. Graph returns a JSON response
#   5. We return that JSON as a Python dictionary
#
# IMPORTANT — WHY WE USE /users/{organiser_id}/ INSTEAD OF /me/:
#   /me/ in a Graph URL means "the currently signed-in user".
#   With Managed Identity there is NO signed-in user — the Function App
#   authenticates as itself (an application identity, not a person).
#   So /me/ returns 404 or 403 because there is no "me".
#
#   Instead we use /users/{organiser_id}/ where organiser_id is the
#   Entra Object ID of the person who organised the meeting.
#   We extract this ID from the meeting details returned by Graph.
#
#   FLOW:
#     1. get_meeting_details_by_id()  → uses /communications/ endpoint
#        (does NOT need /me/ — works with app-only Managed Identity)
#        → returns full meeting object including organiser's user ID
#     2. get_transcript_content()    → uses /users/{organiser_id}/
#        → downloads the actual VTT transcript text
#     3. get_transcript_metadata()   → uses /users/{organiser_id}/
#        → gets transcript info
#     4. get_recording_*()           → uses /users/{organiser_id}/
#        → gets recording info and download URL
#
# USED BY: function_app.py
# =============================================================================

import requests       # The library we use to make HTTP calls (like a browser)
import logging        # For printing log messages (visible in Azure Portal logs)
import datetime       # For working with dates and times
from .auth import get_access_token  # Our Managed Identity auth function from auth.py

# Base URLs for the two versions of Graph API
GRAPH_BASE = "https://graph.microsoft.com/v1.0"   # Stable version
GRAPH_BETA = "https://graph.microsoft.com/beta"   # Beta version (needed for transcripts)


# =============================================================================
# PRIVATE HELPER FUNCTIONS
# =============================================================================

def _auth_headers() -> dict:
    """
    Builds the HTTP headers needed for every Graph API request.

    WHAT ARE HTTP HEADERS?
        Extra information sent with every web request, like a cover letter
        attached to a document. Graph needs two things:
        1. "Authorization" → proves we are allowed to make this request
        2. "Content-Type"  → tells Graph we are sending JSON data

    RETURNS a dictionary like:
        {
            "Authorization": "Bearer eyJ0eXAiOiJKV1Q...(long token string)",
            "Content-Type": "application/json"
        }
    """
    token = get_access_token()  # Get Managed Identity token from auth.py
    return {
        "Authorization": f"Bearer {token}",  # "Bearer" is required before the token
        "Content-Type": "application/json"   # We're sending/receiving JSON
    }


# =============================================================================
# GENERIC HTTP METHODS
# =============================================================================

def graph_get(url: str) -> dict:
    """
    Makes an HTTP GET request to Microsoft Graph.
    Use GET when you want to READ/FETCH data (not change anything).
    """
    response = requests.get(url, headers=_auth_headers())
    response.raise_for_status()
    return response.json()


def graph_post(url: str, payload: dict) -> dict:
    """
    Makes an HTTP POST request to Microsoft Graph.
    Use POST when you want to CREATE something new.
    """
    response = requests.post(url, headers=_auth_headers(), json=payload)
    response.raise_for_status()
    return response.json()


def graph_patch(url: str, payload: dict) -> dict:
    """
    Makes an HTTP PATCH request to Microsoft Graph.
    Use PATCH when you want to UPDATE/MODIFY an existing resource.
    """
    response = requests.patch(url, headers=_auth_headers(), json=payload)
    response.raise_for_status()
    return response.json()


def graph_delete(url: str) -> None:
    """
    Makes an HTTP DELETE request to Microsoft Graph.
    Use DELETE when you want to remove a resource.
    """
    token = get_access_token()
    response = requests.delete(url, headers={"Authorization": f"Bearer {token}"})
    response.raise_for_status()


# =============================================================================
# SUBSCRIPTION MANAGEMENT
#
# A "subscription" tells Microsoft Graph:
#   "Please send me a notification whenever [something] happens."
# In our case: "Notify me whenever a new transcript is created in Teams."
#
# HOW SUBSCRIPTIONS WORK:
#   1. We call create_subscription() with our notification URL
#   2. Graph immediately sends a validation request to that URL (to verify it works)
#   3. Our fn_receive_notification function responds to that validation
#   4. Graph confirms the subscription — now active
#   5. Whenever a transcript is created, Graph POSTs to our notification URL
#   6. Subscriptions expire after 60 minutes — fn_renew_subscription handles renewal
#
# NOTE ON MANAGED IDENTITY + SUBSCRIPTIONS:
#   Creating subscriptions via /subscriptions does work with Managed Identity
#   as long as the correct Graph application permissions have been granted
#   to the Managed Identity's service principal via PowerShell (see README).
# =============================================================================

def create_subscription(notification_url: str, client_state: str) -> dict:
    """
    Creates a new Graph change notification subscription for Teams transcripts.

    WHAT THIS DOES:
        Tells Microsoft Graph: "Watch for new transcripts in any Teams meeting
        and send a notification to our Azure Function URL when one is created."

    PARAMETERS:
        notification_url → the public HTTPS URL of our fn_receive_notification function
                           Example: "https://myapp.azurewebsites.net/api/receive_notification?code=ABC123"

        client_state     → a secret string WE choose (stored in NOTIFICATION_CLIENT_STATE app setting)
                           Graph includes this in every notification it sends us.
                           We check it matches to verify the notification is genuine.
                           Example: "mySecretState123"

    RETURNS a dict like:
        {
            "id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
            "changeType": "created",
            "notificationUrl": "https://...",
            "resource": "/communications/onlineMeetings/getAllTranscripts",
            "expirationDateTime": "2024-01-15T10:30:00Z",
            "clientState": "mySecretState123"
        }
    """
    expiry = (
        datetime.datetime.utcnow() + datetime.timedelta(minutes=60)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    payload = {
        "changeType": "created",
        "notificationUrl": notification_url,
        "resource": "/communications/onlineMeetings/getAllTranscripts",
        "expirationDateTime": expiry,
        "clientState": client_state
    }

    return graph_post(f"{GRAPH_BETA}/subscriptions", payload)


def renew_subscription(subscription_id: str) -> dict:
    """
    Extends an existing subscription's expiry time by another 60 minutes.
    Called automatically by the fn_renew_subscription timer every 45 minutes.

    PARAMETERS:
        subscription_id → the "id" from when the subscription was created
    """
    expiry = (
        datetime.datetime.utcnow() + datetime.timedelta(minutes=60)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    return graph_patch(
        f"{GRAPH_BETA}/subscriptions/{subscription_id}",
        {"expirationDateTime": expiry}
    )


def delete_subscription(subscription_id: str) -> None:
    """Deletes/cancels a subscription so Graph stops sending notifications."""
    graph_delete(f"{GRAPH_BETA}/subscriptions/{subscription_id}")


def list_subscriptions() -> list:
    """
    Returns a list of ALL active subscriptions for this app.

    RETURNS a list of subscription dicts.
    Returns empty list [] if no subscriptions exist.
    """
    result = graph_get(f"{GRAPH_BETA}/subscriptions")
    return result.get("value", [])


# =============================================================================
# ONLINE MEETING FUNCTIONS
#
# IMPORTANT — APP-ONLY ENDPOINT FOR MEETING DETAILS:
#   With Managed Identity (app-only), we CANNOT use:
#     /me/onlineMeetings/{meetingId}               ← requires signed-in user
#
#   We MUST use the communications endpoint instead:
#     /communications/onlineMeetings/{meetingId}   ← works with app-only
#
#   This endpoint returns the full meeting object including the organiser's
#   user ID, which we then use for fetching transcripts and recordings.
# =============================================================================

def get_meeting_details(meeting_id: str) -> dict:
    """
    Fetches all details about a specific Teams online meeting.

    USES: /communications/onlineMeetings/{meetingId}
    This endpoint works with app-only Managed Identity (no signed-in user needed).

    WHAT WE GET FROM THIS:
        - meeting subject/title      → used as SharePoint folder name
        - start/end times            → informational
        - organiser's user ID        → CRITICAL: needed for transcript/recording calls
          Found at: result["participants"]["organizer"]["identity"]["user"]["id"]

    PARAMETERS:
        meeting_id → the Teams meeting ID from the Graph notification resource path
                     Example: "MSoxNzhhZT..."

    RETURNS a dict like:
        {
            "id": "MSoxNzhhZT...",
            "subject": "Weekly Team Standup",
            "startDateTime": "2024-01-15T09:00:00Z",
            "endDateTime": "2024-01-15T09:30:00Z",
            "participants": {
                "organizer": {
                    "identity": {
                        "user": {
                            "id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",  ← organiser's Entra Object ID
                            "displayName": "Jane Doe"
                        }
                    }
                }
            }
        }
    """
    # NOTE: Using /communications/ endpoint — NOT /me/
    # /communications/ is the correct app-only path for meeting details
    return graph_get(f"{GRAPH_BETA}/communications/onlineMeetings/{meeting_id}")


def get_organiser_id(meeting_details: dict) -> str:
    """
    Extracts the organiser's Entra user ID from the meeting details dict.

    WHY THIS IS NEEDED:
        Transcript and recording endpoints require the organiser's user ID.
        The URL pattern is: /users/{organiser_id}/onlineMeetings/{meetingId}/transcripts/...
        We get the organiser_id from the meeting details returned by get_meeting_details().

    PARAMETERS:
        meeting_details → the dict returned by get_meeting_details()

    RETURNS:
        The organiser's Entra Object ID string
        Example: "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"

    RAISES:
        KeyError if the meeting details don't contain organiser info (shouldn't happen)
    """
    try:
        organiser_id = (
            meeting_details["participants"]["organizer"]["identity"]["user"]["id"]
        )
        logging.info(f"Extracted organiser ID: {organiser_id}")
        return organiser_id
    except KeyError as e:
        raise Exception(
            f"Could not extract organiser ID from meeting details. "
            f"Missing key: {e}. "
            f"Meeting details received: {meeting_details}"
        )


# =============================================================================
# TRANSCRIPT FUNCTIONS
#
# IMPORTANT — WHY WE USE /users/{organiser_id}/ HERE:
#   These endpoints are under a specific user's context.
#   With app-only Managed Identity, we must specify WHOSE meetings we are
#   looking at by using /users/{organiser_id}/ in the URL.
#
#   The organiser_id is obtained from get_meeting_details() → get_organiser_id().
#   In function_app.py, get_meeting_details() is always called FIRST,
#   then organiser_id is extracted and passed to these functions.
# =============================================================================

def get_transcript_metadata(meeting_id: str, transcript_id: str, organiser_id: str) -> dict:
    """
    Fetches metadata about a transcript (created time, content URL, etc.)
    Does NOT return the actual transcript text — use get_transcript_content() for that.

    PARAMETERS:
        meeting_id    → the Teams meeting ID
        transcript_id → the specific transcript ID
        organiser_id  → the meeting organiser's Entra user ID
                        (from get_organiser_id(get_meeting_details(meeting_id)))

    RETURNS a dict like:
        {
            "id": "MSt2AbCdEf...",
            "createdDateTime": "2024-01-15T09:35:00Z",
            "transcriptContentUrl": "https://graph.microsoft.com/beta/..."
        }
    """
    # /users/{organiser_id}/ is required for app-only access (no /me/)
    return graph_get(
        f"{GRAPH_BETA}/users/{organiser_id}/onlineMeetings/{meeting_id}"
        f"/transcripts/{transcript_id}"
    )


def get_transcript_content(meeting_id: str, transcript_id: str, organiser_id: str) -> str:
    """
    Downloads the ACTUAL TEXT of a Teams transcript in VTT format.

    WHAT IS VTT FORMAT?
        WebVTT is a subtitle/caption format. It looks like:
            WEBVTT

            00:00:05.000 --> 00:00:08.000
            <v John Smith>Hello everyone, welcome to the meeting.

            00:00:09.000 --> 00:00:12.000
            <v Jane Doe>Thanks for joining. Let's get started.

        Each block has: timestamp --> timestamp, then speaker name and what they said.
        This text is what we send to Azure OpenAI for analysis.

    PARAMETERS:
        meeting_id    → the Teams meeting ID
        transcript_id → the specific transcript ID
        organiser_id  → the meeting organiser's Entra user ID
                        REQUIRED for app-only Managed Identity access.
                        Get this via: get_organiser_id(get_meeting_details(meeting_id))

    RETURNS:
        A string containing the full VTT transcript text.

    URL USED:
        GET /beta/users/{organiser_id}/onlineMeetings/{meetingId}/transcripts/{transcriptId}/content?$format=text/vtt
        NOT: /me/onlineMeetings/... (that requires a signed-in user)
    """
    token = get_access_token()

    # $format=text/vtt tells Graph to return the transcript as plain VTT text
    # (Graph can also return it as DOCX — we use VTT because it's plain text)
    url = (
        f"{GRAPH_BETA}/users/{organiser_id}/onlineMeetings/{meeting_id}"
        f"/transcripts/{transcript_id}/content?$format=text/vtt"
    )

    # Only Authorization header needed — this is a file download, not JSON
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    response.raise_for_status()

    # response.text is the raw VTT string (not JSON)
    return response.text


# =============================================================================
# RECORDING FUNCTIONS
# Same pattern as transcripts — must use /users/{organiser_id}/ not /me/
# =============================================================================

def get_recording_metadata(meeting_id: str, recording_id: str, organiser_id: str) -> dict:
    """
    Fetches metadata about a Teams meeting recording.

    PARAMETERS:
        meeting_id   → the Teams meeting ID
        recording_id → the specific recording ID
        organiser_id → the meeting organiser's Entra user ID

    RETURNS a dict with recording details (created time, duration, etc.)
    """
    return graph_get(
        f"{GRAPH_BETA}/users/{organiser_id}/onlineMeetings/{meeting_id}"
        f"/recordings/{recording_id}"
    )


def get_recording_download_url(meeting_id: str, recording_id: str, organiser_id: str) -> str:
    """
    Gets a temporary direct download URL for a meeting recording video.

    HOW IT WORKS:
        Graph doesn't stream the video directly. It returns an HTTP 302 redirect
        pointing to a temporary Azure Blob Storage URL where the video lives.
        We capture that redirect URL and return it.

    PARAMETERS:
        meeting_id   → the Teams meeting ID
        recording_id → the specific recording ID
        organiser_id → the meeting organiser's Entra user ID

    RETURNS:
        A temporary HTTPS download URL for the MP4 video.
        WARNING: This URL expires after ~1 hour. Download promptly.
    """
    url = (
        f"{GRAPH_BETA}/users/{organiser_id}/onlineMeetings/{meeting_id}"
        f"/recordings/{recording_id}/content"
    )
    token = get_access_token()

    # allow_redirects=False → capture the redirect URL instead of following it
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=False
    )

    # 302/303 = redirect — the Location header has the actual download URL
    if response.status_code in (302, 303):
        return response.headers["Location"]

    response.raise_for_status()
    return url


def download_recording_bytes(meeting_id: str, recording_id: str, organiser_id: str) -> bytes:
    """
    Downloads the entire recording video as bytes.

    WARNING: Recordings can be very large (hundreds of MB to GB).
             This loads the entire file into memory.
             For very large recordings use upload_large_file() in sharepoint_client.py
             which streams in chunks.

    PARAMETERS:
        meeting_id   → the Teams meeting ID
        recording_id → the specific recording ID
        organiser_id → the meeting organiser's Entra user ID

    RETURNS:
        bytes of the full MP4 video file
    """
    download_url = get_recording_download_url(meeting_id, recording_id, organiser_id)
    response = requests.get(download_url, stream=True)
    response.raise_for_status()
    return response.content
