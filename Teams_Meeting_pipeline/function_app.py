# =============================================================================
# FILE: function_app.py
#
# PURPOSE:
#   This is the MAIN FILE of the Azure Function App.
#   All four Azure Functions are defined here.
#
# THE FOUR FUNCTIONS:
#   1. fn_subscribe           → You call this MANUALLY to start listening for Teams transcripts
#   2. fn_receive_notification → Microsoft Graph calls this automatically when a transcript is created
#   3. fn_process_queue       → Azure calls this automatically for each queue message
#   4. fn_renew_subscription  → Azure calls this automatically every 45 minutes (timer)
#
# =============================================================================
#
# ┌─────────────────────────────────────────────────────────────────────────┐
# │                    COMPLETE FLOW DIAGRAM                                │
# │                                                                         │
# │  YOU (manual)                                                           │
# │     │                                                                   │
# │     │  POST /api/subscribe                                              │
# │     │  Body: {"notification_url": "https://..."}                        │
# │     ▼                                                                   │
# │  fn_subscribe  ──────────────────────────────────────────────────────   │
# │     │  Calls Microsoft Graph: "watch for new transcripts"              │
# │     │  Graph validates our URL (calls fn_receive_notification once)    │
# │     │  Returns: {"subscription_id": "...", "expires": "..."}           │
# │     │                                                                   │
# │  [Teams meeting happens, recording & transcript created]                │
# │     │                                                                   │
# │  Microsoft Graph (automatic)                                            │
# │     │                                                                   │
# │     │  POST /api/receive_notification                                   │
# │     │  Body: {"value": [{"resource": "...", "changeType": "created"}]} │
# │     ▼                                                                   │
# │  fn_receive_notification                                                │
# │     │  Validates the notification (checks clientState)                 │
# │     │  Puts a message on the Azure Storage Queue                       │
# │     │  Returns 202 immediately (must be fast)                          │
# │     │                                                                   │
# │  Azure Storage Queue (automatic trigger)                               │
# │     │                                                                   │
# │     ▼                                                                   │
# │  fn_process_queue                                                       │
# │     │  1. Parses the queue message                                     │
# │     │  2. Fetches transcript from Graph API                            │
# │     │  3. Downloads template from SharePoint                           │
# │     │  4. Sends to Azure OpenAI → gets JSON analysis                  │
# │     │  5. Fills Word template with analysis                            │
# │     │  6. Creates output folder in SharePoint                          │
# │     │  7. Uploads: report.docx, transcript.vtt, analysis.json         │
# │     │                                                                   │
# │  Azure Timer (automatic, every 45 minutes)                             │
# │     │                                                                   │
# │     ▼                                                                   │
# │  fn_renew_subscription                                                  │
# │     │  Extends subscription expiry (Graph subscriptions last 60 mins) │
# └─────────────────────────────────────────────────────────────────────────┘
#
# =============================================================================

import azure.functions as func  # Azure Functions SDK — provides HttpRequest, HttpResponse, etc.
import json                      # For converting dicts to/from JSON strings
import logging                   # For writing log messages (visible in Azure Portal)
import os                        # For reading environment variables (app settings)
import base64                    # For decoding base64-encoded queue messages
import io                        # For working with files in memory

# Import all our shared module functions
from shared.graph_client import (
    create_subscription,       # Creates a Graph change notification subscription
    list_subscriptions,        # Lists all active subscriptions
    renew_subscription,        # Extends a subscription's expiry time
    get_transcript_content,    # Downloads the VTT transcript text
    get_transcript_metadata,   # Gets metadata about a transcript
    get_meeting_details,       # Gets Teams meeting details via /communications/ endpoint
    get_organiser_id,          # Extracts organiser's Entra user ID from meeting details
    download_recording_bytes,  # Downloads recording video as bytes
)
from shared.sharepoint_client import (
    download_template,   # Downloads the Word template from SharePoint
    upload_file,         # Uploads a file to SharePoint (up to ~4MB)
    upload_large_file,   # Uploads a large file to SharePoint (chunked, for videos)
    create_folder,       # Creates a new folder in SharePoint
)
from shared.openai_client import analyse_transcript    # Sends transcript to OpenAI, gets analysis
from shared.document_filler import fill_template       # Fills the Word template with analysis data
from shared.queue_client import enqueue_notification   # Puts a message on the Azure Storage Queue

from datetime import datetime  # For generating timestamps (used in folder names)

# =============================================================================
# Create the Function App
#
# http_auth_level=func.AuthLevel.FUNCTION means:
#   All HTTP functions require an API key (?code=...) in the URL.
#   This key is automatically generated by Azure and shown in the Portal.
#   It prevents random people on the internet from calling your functions.
# =============================================================================
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


# =============================================================================
# FUNCTION 1: fn_subscribe
#
# TRIGGER: Manual HTTP POST (you call this yourself)
#
# PURPOSE:
#   Registers our Function App with Microsoft Graph to start receiving
#   notifications whenever a Teams transcript is created.
#   You only need to call this ONCE to set up (and again if the subscription
#   is ever fully deleted — the renewal timer handles normal expiry).
#
# HOW TO CALL THIS FUNCTION:
#
#   URL (from Azure Portal → your Function App → Functions → subscribe → Get Function URL):
#     https://YOUR-APP.azurewebsites.net/api/subscribe?code=YOUR_FUNCTION_KEY
#
#   METHOD: POST
#
#   HEADERS:
#     Content-Type: application/json
#
#   BODY (JSON — you must provide this):
#     {
#         "notification_url": "https://YOUR-APP.azurewebsites.net/api/receive_notification?code=YOUR_KEY"
#     }
#
#   HOW TO GET THE notification_url:
#     Go to Azure Portal → your Function App → Functions → receive_notification → Get Function URL
#     Copy that URL and paste it as the notification_url value.
#
#   TOOLS TO CALL THIS:
#     Option A — Using curl (command line):
#       curl -X POST \
#         "https://YOUR-APP.azurewebsites.net/api/subscribe?code=YOUR_KEY" \
#         -H "Content-Type: application/json" \
#         -d '{"notification_url": "https://YOUR-APP.azurewebsites.net/api/receive_notification?code=OTHER_KEY"}'
#
#     Option B — Using Postman:
#       1. Open Postman
#       2. New Request → POST
#       3. URL: https://YOUR-APP.azurewebsites.net/api/subscribe?code=YOUR_KEY
#       4. Body tab → raw → JSON
#       5. Paste: {"notification_url": "https://..."}
#       6. Send
#
#     Option C — Using Python requests:
#       import requests
#       response = requests.post(
#           "https://YOUR-APP.azurewebsites.net/api/subscribe?code=YOUR_KEY",
#           json={"notification_url": "https://YOUR-APP.azurewebsites.net/api/receive_notification?code=OTHER_KEY"}
#       )
#       print(response.json())
#
#   SUCCESS RESPONSE (HTTP 200):
#     {
#         "status": "success",
#         "subscription_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
#         "expires": "2024-01-15T10:30:00Z"
#     }
#     → Save the subscription_id (useful for debugging, but not required)
#     → "expires" is 60 minutes from now — fn_renew_subscription handles renewal
#
#   ERROR RESPONSES:
#     HTTP 400 → "Request body must be JSON" or "Missing 'notification_url'"
#                Fix: check your request body format
#     HTTP 500 → Error from Microsoft Graph
#                Common cause: admin consent not yet granted for API permissions
# =============================================================================

@app.route(route="subscribe", methods=["POST"])
def fn_subscribe(req: func.HttpRequest) -> func.HttpResponse:
    """Registers a Microsoft Graph subscription to watch for new Teams transcripts."""

    logging.info("fn_subscribe was triggered")

    # ── Parse the request body ─────────────────────────────────────────────────
    # req.get_json() reads the HTTP request body and parses it as JSON
    # If the body is not valid JSON, it raises ValueError
    try:
        body = req.get_json()
    except ValueError:
        # Return a 400 Bad Request response with an error message
        return func.HttpResponse(
            "Request body must be JSON. Example: {\"notification_url\": \"https://...\"}",
            status_code=400
        )

    # Extract the notification_url from the parsed body
    # .get() returns None if the key doesn't exist (instead of raising an error)
    notification_url = body.get("notification_url")
    if not notification_url:
        return func.HttpResponse(
            "Missing 'notification_url' in request body. "
            "This should be the URL of the receive_notification function.",
            status_code=400
        )

    # Read the client state secret from our app settings
    # This is a string we chose (e.g. "mySecretState123") stored in Azure App Settings
    # Graph will include this in every notification it sends us, so we can verify it's genuine
    client_state = os.environ["NOTIFICATION_CLIENT_STATE"]

    try:
        # Check existing subscriptions (for informational/debugging purposes)
        existing = list_subscriptions()
        logging.info(f"Currently have {len(existing)} active subscription(s)")

        # Create the subscription — this calls Microsoft Graph
        # Graph will immediately call our notification_url to validate it
        # (fn_receive_notification handles that validation automatically)
        result = create_subscription(notification_url, client_state)
        logging.info(f"Subscription created successfully. ID: {result['id']}")

        # Return success response with subscription details
        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "subscription_id": result["id"],
                "expires": result["expirationDateTime"],
                "message": "Subscription active. fn_renew_subscription will renew it automatically."
            }),
            status_code=200,
            mimetype="application/json"  # Tell the caller this is JSON
        )

    except Exception as e:
        logging.error(f"Failed to create subscription: {e}")
        return func.HttpResponse(
            f"Error creating subscription: {str(e)}",
            status_code=500
        )


# =============================================================================
# FUNCTION 2: fn_receive_notification
#
# TRIGGER: HTTP (called automatically by Microsoft Graph — NOT by you)
#
# PURPOSE:
#   Microsoft Graph calls this URL whenever:
#     a) It needs to VALIDATE our notification URL (first time, during subscription creation)
#     b) A Teams transcript is CREATED (the actual notifications we want)
#
# YOU DO NOT CALL THIS YOURSELF.
#   Microsoft Graph calls it automatically. However, you need to provide its URL
#   when calling fn_subscribe (as the "notification_url" parameter).
#
# SCENARIO A — VALIDATION HANDSHAKE (automatic, happens once):
#
#   When you call fn_subscribe, Graph immediately sends a GET request to this URL:
#     GET /api/receive_notification?validationToken=SOME_TOKEN_STRING
#
#   We MUST respond with that exact token as plain text within 10 seconds.
#   If we don't, Graph rejects our subscription and it won't work.
#   This code handles it automatically — no action needed from you.
#
# SCENARIO B — ACTUAL NOTIFICATION (automatic, happens when transcript is created):
#
#   When a Teams meeting transcript is created, Graph sends:
#     POST /api/receive_notification
#     Body:
#     {
#         "value": [
#             {
#                 "changeType": "created",
#                 "clientState": "mySecretState123",    ← we verify this matches our setting
#                 "resource": "/communications/onlineMeetings/MSp.../transcripts/MSt...",
#                 "subscriptionId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
#                 "tenantId": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
#                 "resourceData": {
#                     "id": "MSt...",
#                     "@odata.type": "#microsoft.graph.callTranscript"
#                 }
#             }
#         ]
#     }
#
#   We MUST respond with HTTP 202 quickly (within a few seconds).
#   If we take too long, Graph will retry and we'll process the same transcript multiple times.
#   So we just put it on the queue and return immediately.
#
# WHAT HAPPENS AFTER:
#   The notification is put on the Azure Storage Queue.
#   fn_process_queue is automatically triggered to do the actual work.
# =============================================================================

@app.route(route="receive_notification", methods=["GET", "POST"])
def fn_receive_notification(req: func.HttpRequest) -> func.HttpResponse:
    """Receives change notifications from Microsoft Graph and queues them for processing."""

    logging.info("fn_receive_notification was triggered")

    # ── SCENARIO A: Validation Handshake ──────────────────────────────────────
    # When creating a subscription, Graph sends a GET with ?validationToken=...
    # We must echo it back as plain text immediately.
    # req.params is a dictionary of URL query parameters
    # e.g. for URL "...?validationToken=abc123", req.params.get("validationToken") = "abc123"
    validation_token = req.params.get("validationToken")
    if validation_token:
        logging.info("Graph validation handshake received — echoing token back")
        # Return the token exactly as received, as plain text
        # This proves to Graph that this URL is under our control
        return func.HttpResponse(
            validation_token,
            status_code=200,
            mimetype="text/plain"  # Must be text/plain, NOT application/json
        )

    # ── SCENARIO B: Actual Change Notification ─────────────────────────────────
    # Graph sends a POST with the notification details in the body
    try:
        body = req.get_json()
    except ValueError:
        logging.error("Received non-JSON body from Graph")
        return func.HttpResponse("Invalid JSON body", status_code=400)

    # Read our secret verification string from app settings
    expected_state = os.environ["NOTIFICATION_CLIENT_STATE"]

    # "value" is a list because Graph can batch multiple notifications together
    notifications = body.get("value", [])
    processed = 0

    for notification in notifications:
        # ── Security check ──────────────────────────────────────────────────────
        # Verify the clientState matches what we set when creating the subscription.
        # This prevents malicious actors from sending fake notifications to our URL.
        if notification.get("clientState") != expected_state:
            logging.warning(
                f"Received notification with wrong clientState: "
                f"{notification.get('clientState')} — ignoring"
            )
            continue  # Skip this notification, process the next one

        # Extract key fields from the notification
        resource    = notification.get("resource", "")    # e.g. "/communications/onlineMeetings/MSp.../transcripts/MSt..."
        change_type = notification.get("changeType", "")  # e.g. "created"
        logging.info(f"Valid notification: {change_type} on {resource}")

        # Build the payload we'll store in the queue
        # We include all the info fn_process_queue will need to do its job
        payload = {
            "resource":       resource,                              # Path to the transcript in Graph
            "changeType":     change_type,                          # "created"
            "subscriptionId": notification.get("subscriptionId"),   # Which subscription triggered this
            "tenantId":       notification.get("tenantId"),         # The Microsoft 365 tenant ID
            "resourceData":   notification.get("resourceData", {}), # Extra data (includes transcript ID)
            "receivedAt":     datetime.utcnow().isoformat()         # Timestamp for debugging
        }

        # Put this payload on the Azure Storage Queue
        # fn_process_queue will automatically pick it up
        enqueue_notification(payload)
        processed += 1

    logging.info(f"Processed {processed} notification(s), enqueued for processing")

    # IMPORTANT: Return 202 Accepted immediately.
    # Graph requires a fast response. If we return 4xx or 5xx, Graph retries.
    # 202 means "I received it and will process it asynchronously"
    return func.HttpResponse("", status_code=202)


# =============================================================================
# FUNCTION 3: fn_process_queue
#
# TRIGGER: Azure Storage Queue (automatic — triggered by new queue messages)
#
# PURPOSE:
#   This is the main processing function. It does all the heavy work:
#   - Fetches the transcript from Microsoft Graph
#   - Downloads the template from SharePoint
#   - Sends to Azure OpenAI for analysis
#   - Fills the Word template
#   - Uploads everything to SharePoint
#
# YOU DO NOT CALL THIS YOURSELF.
#   Azure automatically calls this when a message arrives in the queue.
#   The "msg" parameter contains the queue message body.
#
# WHAT TRIGGERS IT:
#   When fn_receive_notification calls enqueue_notification(),
#   Azure detects the new queue message and immediately calls this function.
#
# IF THIS FUNCTION FAILS:
#   Azure will retry up to 5 times (configurable in host.json).
#   After max retries, the message goes to a "poison queue" for investigation.
#
# LOGS TO WATCH (in Azure Portal → your Function App → Monitor → Logs):
#   "fn_process_queue triggered"             → function started
#   "Processing resource: /communications/..." → which transcript is being processed
#   "Fetched transcript (XXXX chars)"        → transcript downloaded from Graph
#   "Downloaded template (XXXX bytes)"       → template downloaded from SharePoint
#   "OpenAI analysis keys: [...]"            → OpenAI analysis complete
#   "Template filled successfully"           → Word document ready
#   "Created SharePoint folder: ..."         → output folder created
#   "✅ Uploaded Meeting_Report.docx"        → report uploaded
#   "✅ Uploaded transcript.vtt"             → transcript uploaded
#   "✅ Uploaded analysis.json"              → analysis JSON uploaded
#   "🎉 Pipeline complete! Files at: ..."    → everything done!
# =============================================================================

@app.queue_trigger(
    arg_name="msg",                        # Python variable name for the queue message
    queue_name="%QUEUE_NAME%",             # %QUEUE_NAME% reads from app settings (= "transcripts-queue")
    connection="STORAGE_CONNECTION_STRING" # App setting name for the storage connection string
)
def fn_process_queue(msg: func.QueueMessage) -> None:
    """
    Processes a Teams transcript notification from the queue.
    Full pipeline: Graph → OpenAI → SharePoint.
    """

    logging.info("fn_process_queue triggered — starting pipeline")

    # ── STEP 1: Parse the queue message ───────────────────────────────────────
    # The queue message is the base64-encoded JSON we put in queue_client.py
    # We need to reverse that encoding to get the original dict back.
    try:
        # msg.get_body() returns the raw message as bytes
        # .decode("utf-8") converts bytes to a string
        raw = msg.get_body().decode("utf-8")

        # Try to base64-decode (we encoded it in queue_client.py)
        try:
            decoded = base64.b64decode(raw).decode("utf-8")
        except Exception:
            # If base64 decoding fails, use the raw string directly
            decoded = raw

        # Parse the JSON string back into a Python dictionary
        payload = json.loads(decoded)

    except Exception as e:
        logging.error(f"Failed to parse queue message: {e}")
        return  # Stop processing — Azure will retry automatically

    # Extract the fields from the payload
    resource      = payload.get("resource", "")        # Graph resource path
    resource_data = payload.get("resourceData", {})    # Extra data from Graph
    logging.info(f"Processing notification for resource: {resource}")

    # ── STEP 2: Extract Meeting ID and Transcript ID ───────────────────────────
    # The resource path looks like:
    # "/communications/onlineMeetings/MSp1NzhhZT.../transcripts/MSt2AbCdEf..."
    # We need to split this path to get the meeting ID and transcript ID.
    meeting_id    = None
    transcript_id = None

    # Split by "/" and look for the parts we need
    # e.g. ["", "communications", "onlineMeetings", "MSp1NzhhZT...", "transcripts", "MSt2AbCdEf..."]
    parts = resource.split("/")
    for i, part in enumerate(parts):
        # The meeting ID is the item RIGHT AFTER "onlineMeetings"
        if part == "onlineMeetings" and i + 1 < len(parts):
            meeting_id = parts[i + 1]
        # The transcript ID is the item RIGHT AFTER "transcripts"
        if part == "transcripts" and i + 1 < len(parts):
            transcript_id = parts[i + 1]

    # Fallback: sometimes IDs are in the resourceData dict instead
    if not meeting_id:
        meeting_id = resource_data.get("meetingId", None)
    if not transcript_id:
        transcript_id = resource_data.get("id", None)

    # If we still don't have both IDs, we can't proceed
    if not meeting_id or not transcript_id:
        logging.error(f"Could not extract meeting/transcript IDs from resource: {resource}")
        return

    logging.info(f"Meeting ID: {meeting_id}")
    logging.info(f"Transcript ID: {transcript_id}")

    # ── STEP 3: Fetch Meeting Details then Transcript from Microsoft Graph ──────
    try:
        # FIRST: get meeting details using the /communications/ endpoint.
        # This works with app-only Managed Identity — no signed-in user needed.
        # We call this BEFORE the transcript because we need the organiser's
        # Entra user ID to construct the correct transcript URL.
        meeting_details = get_meeting_details(meeting_id)
        logging.info(f"Meeting subject: {meeting_details.get('subject', 'Unknown')}")

        # SECOND: extract the organiser's Entra user ID from the meeting details.
        # This is required because with Managed Identity (app-only), Graph transcript
        # endpoints need /users/{organiser_id}/ in the URL.
        # Using /me/ would fail — there is no signed-in user with Managed Identity.
        organiser_id = get_organiser_id(meeting_details)

        # THIRD: fetch the actual VTT transcript text.
        # URL used: /beta/users/{organiser_id}/onlineMeetings/{meetingId}/transcripts/{id}/content
        transcript_text = get_transcript_content(meeting_id, transcript_id, organiser_id)
        logging.info(f"Fetched transcript ({len(transcript_text)} characters)")

    except Exception as e:
        logging.error(f"Failed to fetch meeting details or transcript from Graph: {e}")
        return

    # ── STEP 4: Download SharePoint Template ──────────────────────────────────
    # Read SharePoint settings from Azure App Settings
    site_id          = os.environ["SHAREPOINT_SITE_ID"]          # e.g. "contoso.sharepoint.com,abc123,def456"
    template_file_id = os.environ["SHAREPOINT_TEMPLATE_FILE_ID"] # e.g. "01ABCDE..."

    try:
        # Download the .docx template as bytes from SharePoint
        template_bytes = download_template(site_id, template_file_id)
        logging.info(f"Downloaded template ({len(template_bytes)} bytes)")

        # We also extract text from the template to send to OpenAI as reference context.
        # This helps OpenAI understand what the meeting was about in the context of the document.
        # python-docx can open a .docx from bytes using io.BytesIO
        from docx import Document as DocxDocument
        ref_doc = DocxDocument(io.BytesIO(template_bytes))

        # Extract all non-empty paragraph texts and join them with newlines
        reference_text = "\n".join(
            p.text for p in ref_doc.paragraphs if p.text.strip()
        )
        logging.info(f"Extracted {len(reference_text)} characters from template for reference")

    except Exception as e:
        logging.error(f"Failed to download or read SharePoint template: {e}")
        return

    # ── STEP 5: Analyse with Azure OpenAI ─────────────────────────────────────
    # Send the transcript text + reference doc text to GPT-4o
    # Get back a structured JSON analysis (meeting summary, action items, etc.)
    try:
        analysis = analyse_transcript(transcript_text, reference_text)
        logging.info(f"OpenAI analysis complete. Keys returned: {list(analysis.keys())}")

    except Exception as e:
        logging.error(f"Azure OpenAI analysis failed: {e}")
        return

    # ── STEP 6: Fill the Word Template with Analysis Data ─────────────────────
    # Replace {{MEETING_TITLE}}, {{EXECUTIVE_SUMMARY}}, etc. in the template
    # and append sections like Action Items, Key Decisions, etc.
    try:
        filled_doc_bytes = fill_template(template_bytes, analysis)
        logging.info("Word template filled successfully")

    except Exception as e:
        logging.error(f"Failed to fill Word template: {e}")
        return

    # ── STEP 7: Create Output Folder in SharePoint ────────────────────────────
    # We create a unique folder for each meeting using the subject + timestamp
    # Example: "ProcessedMeetings/Weekly_Standup_20240115_0930"

    output_folder_root = os.environ["SHAREPOINT_OUTPUT_FOLDER"]  # e.g. "ProcessedMeetings"
    meeting_subject    = meeting_details.get("subject", "Meeting")  # e.g. "Weekly Team Standup"

    # Clean the subject for use as a folder name:
    # Keep only letters, numbers, spaces, underscores, hyphens
    # Replace anything else with underscore
    # Limit to 50 characters to keep folder names manageable
    safe_subject = "".join(
        c if c.isalnum() or c in " _-" else "_"
        for c in meeting_subject
    )[:50].strip()

    # Add timestamp to make folder name unique even if subject is the same
    # strftime('%Y%m%d_%H%M') → "20240115_0930"
    folder_name = f"{safe_subject}_{datetime.utcnow().strftime('%Y%m%d_%H%M')}"

    try:
        # Create the folder in SharePoint
        create_folder(site_id, output_folder_root, folder_name)
        # The full path to the folder (used for uploading files into it)
        folder_path = f"{output_folder_root}/{folder_name}"
        logging.info(f"Created SharePoint output folder: {folder_path}")

    except Exception as e:
        logging.error(f"Failed to create SharePoint folder: {e}")
        return

    # ── STEP 8: Upload All Output Files to SharePoint ─────────────────────────
    # Upload three files into the output folder:
    #   a) Meeting_Report.docx → the filled Word document
    #   b) transcript.vtt      → the raw transcript text
    #   c) analysis.json       → the OpenAI JSON output (for debugging/audit trail)
    try:

        # 8a. Upload the filled Word document
        upload_file(
            site_id,
            folder_path,
            "Meeting_Report.docx",
            filled_doc_bytes,
            # This long string is the official MIME type for .docx files
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        logging.info("✅ Uploaded Meeting_Report.docx")

        # 8b. Upload the raw VTT transcript text
        upload_file(
            site_id,
            folder_path,
            "transcript.vtt",
            transcript_text.encode("utf-8"),  # Convert string to bytes for upload
            "text/vtt"
        )
        logging.info("✅ Uploaded transcript.vtt")

        # 8c. Upload the OpenAI analysis as a JSON file
        # json.dumps(analysis, indent=2) → nicely formatted JSON string
        # .encode("utf-8") → converts to bytes for upload
        upload_file(
            site_id,
            folder_path,
            "analysis.json",
            json.dumps(analysis, indent=2).encode("utf-8"),
            "application/json"
        )
        logging.info("✅ Uploaded analysis.json")

        # ── OPTIONAL: Upload the Recording Video ──────────────────────────────
        # Recording files can be very large (hundreds of MB to GB).
        # Uncomment the block below to enable recording upload.
        # You will also need to extract the recording_id from the notification,
        # which may come via a separate Graph notification for recordings.
        #
        # try:
        #     recording_bytes = download_recording_bytes(meeting_id, recording_id)
        #     upload_large_file(
        #         site_id,
        #         folder_path,
        #         "recording.mp4",
        #         recording_bytes,
        #         "video/mp4"
        #     )
        #     logging.info("✅ Uploaded recording.mp4")
        # except Exception as rec_err:
        #     logging.warning(f"Recording upload skipped: {rec_err}")

        logging.info(f"🎉 Pipeline complete! All files uploaded to: {folder_path}")

    except Exception as e:
        logging.error(f"Failed to upload files to SharePoint: {e}")
        return


# =============================================================================
# FUNCTION 4: fn_renew_subscription
#
# TRIGGER: Azure Timer (runs automatically every 45 minutes)
#
# PURPOSE:
#   Microsoft Graph subscriptions for onlineMeetings expire after 60 minutes.
#   If we don't renew them, Graph stops sending notifications and we miss transcripts.
#   This function renews ALL active subscriptions every 45 minutes as a safety margin.
#
# YOU DO NOT CALL THIS YOURSELF.
#   Azure runs it on the schedule defined in the @app.timer_trigger decorator.
#
# SCHEDULE FORMAT (cron expression): "0 */45 * * * *"
#   Reading left to right: seconds minutes hours day-of-month month day-of-week
#   "0 */45 * * * *" = "at second 0, every 45 minutes, every hour, every day"
#   Other examples:
#     "0 0 * * * *"   = every hour
#     "0 */30 * * * *" = every 30 minutes
#     "0 0 9 * * 1"   = every Monday at 9:00 AM
#
# WHAT HAPPENS IF THIS FUNCTION FAILS:
#   The subscription will expire 60 minutes after the last successful renewal.
#   If you notice you're missing notifications, check this function's logs.
#   You can force a re-subscription by calling fn_subscribe again.
#
# LOGS TO WATCH:
#   "fn_renew_subscription triggered"   → timer fired
#   "Renewed subscription: xxxx"        → each subscription successfully renewed
#   "No active subscriptions found"     → need to call fn_subscribe first
# =============================================================================

@app.timer_trigger(
    schedule="0 */45 * * * *",  # Run at 0 seconds, every 45 minutes
    arg_name="timer"             # Python variable name for the timer info object
)
def fn_renew_subscription(timer: func.TimerRequest) -> None:
    """Automatically renews all active Graph subscriptions every 45 minutes."""

    logging.info("fn_renew_subscription triggered by timer")

    try:
        # Get all currently active subscriptions
        subs = list_subscriptions()

        if not subs:
            # No subscriptions means either:
            # a) fn_subscribe hasn't been called yet
            # b) The subscription was somehow deleted
            logging.warning(
                "No active subscriptions found to renew. "
                "Call the fn_subscribe endpoint to create one."
            )
            return

        # Renew each subscription
        for sub in subs:
            renew_subscription(sub["id"])
            logging.info(f"Successfully renewed subscription: {sub['id']}")

        logging.info(f"Renewed {len(subs)} subscription(s)")

    except Exception as e:
        logging.error(f"Failed to renew subscriptions: {e}")
        # Don't re-raise — let the next timer run try again
