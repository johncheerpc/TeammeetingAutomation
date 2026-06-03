# =============================================================================
# FILE: shared/queue_client.py
#
# PURPOSE:
#   Handles sending messages to the Azure Storage Queue.
#
# WHAT IS AZURE STORAGE QUEUE?
#   A queue is like a to-do list that multiple systems can share.
#   - fn_receive_notification ADDS items to the queue (one per Teams notification)
#   - fn_process_queue READS items from the queue (one at a time, automatically)
#
# WHY DO WE USE A QUEUE?
#   When Graph sends us a notification, we must respond within a few seconds
#   or Graph will think our endpoint is broken and retry.
#   We can't do all the processing (download transcript, call OpenAI, etc.)
#   in that short time — it takes 10-30 seconds.
#
#   So instead:
#     1. fn_receive_notification gets the notification → QUICKLY adds it to queue → returns 202
#     2. fn_process_queue picks it up from the queue → takes its time processing
#
#   This decouples fast reception from slow processing.
#
# MESSAGE FORMAT:
#   Azure Storage Queue messages must be plain text.
#   We convert our Python dict to a JSON string, then encode it as base64.
#   Base64 encoding ensures special characters don't corrupt the message.
#
# USED BY: function_app.py (fn_receive_notification adds messages)
#          function_app.py (fn_process_queue reads messages — handled automatically by Azure)
# =============================================================================

import os      # To read STORAGE_CONNECTION_STRING and QUEUE_NAME
import json    # To convert Python dict → JSON string
import base64  # To encode the message as base64 text
import logging # For log messages
from azure.storage.queue import QueueClient  # Azure Storage Queue SDK


def get_queue_client() -> QueueClient:
    """
    Creates and returns an Azure Storage Queue client.

    WHAT IS A QUEUE CLIENT?
        It's an object that knows how to talk to a specific Azure Storage Queue.
        Think of it like a connection to the queue — you create it once
        and then use it to send or receive messages.

    REQUIRED APP SETTINGS:
        STORAGE_CONNECTION_STRING → the connection string for your Azure Storage Account
                                    Example: "DefaultEndpointsProtocol=https;AccountName=..."
                                    (Found in Azure Portal → Storage Account → Access keys)

        QUEUE_NAME → the name of the queue to use
                     Example: "transcripts-queue"
                     (The queue must already exist — created in Azure Portal)

    RETURNS:
        A QueueClient object connected to the specified queue
    """
    # Read settings from environment variables (set in Azure Portal → App settings)
    conn_str   = os.environ["STORAGE_CONNECTION_STRING"]
    queue_name = os.environ["QUEUE_NAME"]

    # Create a QueueClient using the connection string + queue name
    # from_connection_string() handles all the authentication automatically
    return QueueClient.from_connection_string(conn_str, queue_name)


def enqueue_notification(data: dict) -> None:
    """
    Converts a notification payload dict into a queue message and sends it.

    THIS IS CALLED BY fn_receive_notification in function_app.py
    every time Microsoft Graph sends us a Teams transcript notification.

    HOW THE MESSAGE IS PREPARED:
        Step 1: Convert the Python dict to a JSON string
                {"resource": "/communications/...", "changeType": "created"}
                becomes → '{"resource": "/communications/...", "changeType": "created"}'

        Step 2: Encode the JSON string as UTF-8 bytes
                (computers store text as bytes internally)

        Step 3: Base64-encode the bytes → safe ASCII text
                Azure Storage Queue requires messages to be plain text,
                and base64 ensures special characters (like < > & ") don't cause issues.

        Step 4: Send the encoded message to the queue

    PARAMETERS:
        data → a Python dictionary with the notification details
               Example:
               {
                   "resource": "/communications/onlineMeetings/MSp.../transcripts/MSt...",
                   "changeType": "created",
                   "subscriptionId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                   "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                   "resourceData": {"id": "MSt..."},
                   "receivedAt": "2024-01-15T09:35:00"
               }

    RETURNS:
        Nothing (None). Just puts the message in the queue.

    AFTER THIS RUNS:
        Azure automatically triggers fn_process_queue with this message.
        fn_process_queue then reads and decodes the message to get the dict back.
    """
    # Get the queue client (connects to our Azure Storage Queue)
    client = get_queue_client()

    # Step 1: Convert Python dict → JSON string
    # json.dumps() = "JSON dump to string"
    # Result looks like: '{"resource": "/communications/...", ...}'
    message = json.dumps(data)

    # Step 2 + 3: UTF-8 encode → base64 encode → decode back to string
    # .encode("utf-8")  → converts string to bytes: b'{"resource":...'
    # base64.b64encode() → encodes bytes to base64 bytes: b'eyJyZXNvdXJjZSI6...'
    # .decode("utf-8")  → converts base64 bytes back to a string: 'eyJyZXNvdXJjZSI6...'
    encoded = base64.b64encode(message.encode("utf-8")).decode("utf-8")

    # Step 4: Send the encoded message to the Azure Storage Queue
    # Azure will hold this message until fn_process_queue picks it up
    client.send_message(encoded)

    logging.info(f"Enqueued notification for resource: {data.get('resource', 'unknown')}")
