# Teams Recording & Transcript Pipeline — Complete Implementation Guide

## Architecture Overview

```
[Microsoft Teams]
      │  Recording/Transcript created
      ▼
[Microsoft Graph Change Notification]
      │  HTTP POST
      ▼
[Azure Function: fn_receive_notification]   ◄── fn_subscribe (manual trigger)
      │  Enqueues message
      ▼
[Azure Storage Queue]
      │  Queue trigger
      ▼
[Azure Function: fn_process_queue]
      │  Calls Graph API for transcript + recording details
      │  Calls SharePoint for template document
      │  Calls Azure OpenAI
      ▼
[Azure OpenAI (GPT-4)]
      │  Returns JSON output
      ▼
[Azure Function: fn_process_queue (continued)]
      │  Fills template document
      │  Uploads filled doc + recording + transcript
      ▼
[SharePoint Output Folder]
```

---

## Prerequisites

Before writing any code, you need:

| Item | Where to get it |
|---|---|
| Azure Subscription | portal.azure.com |
| Microsoft 365 tenant (Teams) | admin.microsoft.com |
| Azure Function App (Python 3.11) | portal.azure.com |
| Azure Storage Account | portal.azure.com |
| Azure OpenAI resource | portal.azure.com → Azure OpenAI |
| App Registration in Entra ID | portal.azure.com → Entra ID |

---

## PHASE 1 — Azure & Entra ID Setup

### Step 1.1 — Register an App in Microsoft Entra ID

This gives your Function App an identity to call Microsoft Graph.

1. Go to **portal.azure.com** → search **"App registrations"** → **New registration**
2. Fill in:
   - Name: `TeamsTranscriptApp`
   - Supported account types: **Single tenant**
   - Redirect URI: leave blank
3. Click **Register**
4. Note down:
   - **Application (client) ID** → save as `CLIENT_ID`
   - **Directory (tenant) ID** → save as `TENANT_ID`

5. Go to **Certificates & secrets** → **New client secret**
   - Description: `TeamsAppSecret`
   - Expiry: 24 months
   - Click **Add** → copy the **Value** immediately → save as `CLIENT_SECRET`

6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
   Add ALL of these:
   ```
   OnlineMeetings.Read.All
   OnlineMeetingTranscript.Read.All
   OnlineMeetingRecording.Read.All
   CallRecords.Read.All
   Sites.ReadWrite.All
   Files.ReadWrite.All
   ```
7. Click **Grant admin consent for [your tenant]** → confirm

---

### Step 1.2 — Create Azure Storage Account

1. Go to **portal.azure.com** → **Storage accounts** → **Create**
2. Fill in:
   - Resource group: create new `rg-teams-pipeline`
   - Storage account name: `teamspipelinestorage` (must be globally unique)
   - Region: choose nearest
   - Performance: Standard
   - Redundancy: LRS
3. Click **Review + Create** → **Create**
4. Once created, go to the storage account → **Access keys** → copy **Connection string** → save as `STORAGE_CONNECTION_STRING`
5. Go to **Queues** → **+ Queue** → name it `transcripts-queue` → **OK**

---

### Step 1.3 — Create Azure Function App

1. Go to **portal.azure.com** → **Function App** → **Create**
2. Fill in:
   - Resource group: `rg-teams-pipeline`
   - Function App name: `teams-transcript-pipeline` (globally unique)
   - Runtime stack: **Python**
   - Version: **3.11**
   - Region: same as storage account
   - Hosting: **Consumption (Serverless)**
3. Click **Review + Create** → **Create**
4. Once created, go to the Function App → **Configuration** → **Application settings**
   Add all these settings (click **+ New application setting** for each):

   | Name | Value |
   |---|---|
   | `TENANT_ID` | your tenant ID |
   | `CLIENT_ID` | your client ID |
   | `CLIENT_SECRET` | your client secret |
   | `STORAGE_CONNECTION_STRING` | your storage connection string |
   | `QUEUE_NAME` | `transcripts-queue` |
   | `NOTIFICATION_CLIENT_STATE` | any random string e.g. `mySecretState123` |
   | `AZURE_OPENAI_ENDPOINT` | your OpenAI endpoint URL |
   | `AZURE_OPENAI_KEY` | your OpenAI API key |
   | `AZURE_OPENAI_DEPLOYMENT` | your deployment name e.g. `gpt-4o` |
   | `SHAREPOINT_SITE_ID` | your SharePoint site ID (see Phase 3) |
   | `SHAREPOINT_TEMPLATE_FILE_ID` | your template file ID (see Phase 3) |
   | `SHAREPOINT_OUTPUT_FOLDER` | e.g. `ProcessedMeetings` |

5. Click **Save**

---

### Step 1.4 — Create Azure OpenAI Resource (if not done)

1. Go to **portal.azure.com** → search **"Azure OpenAI"** → **Create**
2. Fill in resource group, name, region, pricing tier → **Create**
3. Once created, go to **Azure OpenAI Studio** → **Deployments** → **Deploy model**
   - Model: `gpt-4o`
   - Deployment name: `gpt-4o` (or your choice)
4. Go to the resource → **Keys and Endpoint** → copy **KEY 1** and **Endpoint**

---

## PHASE 2 — Local Development Setup

### Step 2.1 — Install Tools

```bash
# 1. Install Azure Functions Core Tools
npm install -g azure-functions-core-tools@4

# 2. Install Azure CLI
# Windows: https://aka.ms/installazurecliwindows
# macOS:
brew install azure-cli

# 3. Login to Azure
az login
```

### Step 2.2 — Create the Project

```bash
# Create project folder
mkdir teams-transcript-pipeline
cd teams-transcript-pipeline

# Initialise Function App (Python v2 model)
func init --python -m V2

# Install dependencies
pip install -r requirements.txt
```

---

## PHASE 3 — Project File Structure

Create this exact folder structure:

```
teams-transcript-pipeline/
├── function_app.py              # All function definitions
├── host.json
├── local.settings.json          # Local env vars (never commit)
├── requirements.txt
└── shared/
    ├── __init__.py
    ├── auth.py                  # Microsoft Graph token
    ├── graph_client.py          # Graph API calls
    ├── openai_client.py         # Azure OpenAI calls
    ├── sharepoint_client.py     # SharePoint operations
    └── queue_client.py          # Storage queue operations
```

---

## PHASE 4 — Code Implementation

### `requirements.txt`

```
azure-functions
azure-storage-queue
azure-identity
msal
requests
python-docx
openai
```

---

### `local.settings.json`  ⚠️ Never commit this file

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "YOUR_STORAGE_CONNECTION_STRING",
    "FUNCTIONS_WORKER_RUNTIME": "python",
    "TENANT_ID": "YOUR_TENANT_ID",
    "CLIENT_ID": "YOUR_CLIENT_ID",
    "CLIENT_SECRET": "YOUR_CLIENT_SECRET",
    "STORAGE_CONNECTION_STRING": "YOUR_STORAGE_CONNECTION_STRING",
    "QUEUE_NAME": "transcripts-queue",
    "NOTIFICATION_CLIENT_STATE": "mySecretState123",
    "AZURE_OPENAI_ENDPOINT": "https://YOUR_RESOURCE.openai.azure.com/",
    "AZURE_OPENAI_KEY": "YOUR_KEY",
    "AZURE_OPENAI_DEPLOYMENT": "gpt-4o",
    "SHAREPOINT_SITE_ID": "YOUR_SITE_ID",
    "SHAREPOINT_TEMPLATE_FILE_ID": "YOUR_TEMPLATE_FILE_ID",
    "SHAREPOINT_OUTPUT_FOLDER": "ProcessedMeetings"
  }
}
```

---

### `shared/auth.py` — Get Graph Access Token

```python
import os
import msal

def get_access_token() -> str:
    """
    Gets an OAuth2 access token from Microsoft Identity Platform
    using Client Credentials flow (app-only, no user login needed).
    """
    tenant_id    = os.environ["TENANT_ID"]
    client_id    = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    scope = ["https://graph.microsoft.com/.default"]

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret
    )

    result = app.acquire_token_for_client(scopes=scope)

    if "access_token" not in result:
        raise Exception(
            f"Could not get access token: {result.get('error_description', result)}"
        )

    return result["access_token"]
```

---

### `shared/graph_client.py` — Microsoft Graph API Calls

```python
import requests
import logging
import os
from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"

def graph_get(url: str) -> dict:
    """Make an authenticated GET request to Microsoft Graph."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

def graph_post(url: str, payload: dict) -> dict:
    """Make an authenticated POST request to Microsoft Graph."""
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()
    return response.json()

def graph_delete(url: str) -> None:
    """Make an authenticated DELETE request to Microsoft Graph."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.delete(url, headers=headers)
    response.raise_for_status()

# ── Subscriptions ──────────────────────────────────────────────────────────────

def create_subscription(notification_url: str, client_state: str) -> dict:
    """
    Subscribe to Microsoft Graph change notifications.
    Subscribes to onlineMeetings transcripts and recordings.
    """
    import datetime

    # Subscription expires in 60 minutes (max allowed for online meetings)
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
    """Renew a subscription before it expires."""
    import datetime
    expiry = (
        datetime.datetime.utcnow() + datetime.timedelta(minutes=60)
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    response = requests.patch(
        f"{GRAPH_BETA}/subscriptions/{subscription_id}",
        headers=headers,
        json={"expirationDateTime": expiry}
    )
    response.raise_for_status()
    return response.json()

def delete_subscription(subscription_id: str) -> None:
    graph_delete(f"{GRAPH_BETA}/subscriptions/{subscription_id}")

def list_subscriptions() -> list:
    result = graph_get(f"{GRAPH_BETA}/subscriptions")
    return result.get("value", [])

# ── Transcript & Recording ─────────────────────────────────────────────────────

def get_transcript_content(meeting_id: str, transcript_id: str) -> str:
    """
    Downloads the VTT/text content of a transcript.
    Returns raw text string.
    """
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}

    # Get transcript metadata first
    url = (
        f"{GRAPH_BETA}/me/onlineMeetings/{meeting_id}"
        f"/transcripts/{transcript_id}/content?$format=text/vtt"
    )
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.text

def get_meeting_details(meeting_id: str) -> dict:
    """Get full details of an online meeting."""
    return graph_get(f"{GRAPH_BETA}/me/onlineMeetings/{meeting_id}")

def get_transcript_metadata(meeting_id: str, transcript_id: str) -> dict:
    """Get metadata for a specific transcript."""
    return graph_get(
        f"{GRAPH_BETA}/me/onlineMeetings/{meeting_id}/transcripts/{transcript_id}"
    )

def get_recording_metadata(meeting_id: str, recording_id: str) -> dict:
    """Get metadata for a specific recording."""
    return graph_get(
        f"{GRAPH_BETA}/me/onlineMeetings/{meeting_id}/recordings/{recording_id}"
    )

def get_recording_content_url(meeting_id: str, recording_id: str) -> str:
    """
    Get the download URL for a meeting recording.
    Returns a temporary download URL (valid ~1 hour).
    """
    url = (
        f"{GRAPH_BETA}/me/onlineMeetings/{meeting_id}"
        f"/recordings/{recording_id}/content"
    )
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    # Follow redirect to get the actual download URL
    response = requests.get(url, headers=headers, allow_redirects=False)
    if response.status_code in (302, 303):
        return response.headers.get("Location")
    response.raise_for_status()
    return url
```

---

### `shared/sharepoint_client.py` — SharePoint Operations

```python
import requests
import os
import io
import logging
from .auth import get_access_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

def _headers():
    token = get_access_token()
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

def get_sharepoint_site_id(hostname: str, site_path: str) -> str:
    """
    Helper to get a SharePoint site ID.
    Example: hostname='mycompany.sharepoint.com', site_path='/sites/MySite'
    Run this once to get your SHAREPOINT_SITE_ID setting.
    """
    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    response = requests.get(url, headers=_headers())
    response.raise_for_status()
    return response.json()["id"]

def download_template(site_id: str, file_id: str) -> bytes:
    """Download a template .docx file from SharePoint by file ID."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/items/{file_id}/content"
    token = get_access_token()
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=True
    )
    response.raise_for_status()
    return response.content

def upload_file(
    site_id: str,
    folder_path: str,
    filename: str,
    content: bytes,
    content_type: str = "application/octet-stream"
) -> dict:
    """
    Upload a file to a SharePoint folder.
    folder_path example: 'ProcessedMeetings/Meeting_2024_01_15'
    """
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/"
        f"{folder_path}/{filename}:/content"
    )
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type
    }
    response = requests.put(url, headers=headers, data=content)
    response.raise_for_status()
    return response.json()

def create_folder(site_id: str, parent_path: str, folder_name: str) -> dict:
    """Create a new folder in SharePoint."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{parent_path}:/children"
    payload = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }
    response = requests.post(url, headers=_headers(), json=payload)
    response.raise_for_status()
    return response.json()

def get_file_id_by_path(site_id: str, file_path: str) -> str:
    """Get a file's ID by its path. Useful for finding template file IDs."""
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{file_path}"
    response = requests.get(url, headers=_headers())
    response.raise_for_status()
    return response.json()["id"]
```

---

### `shared/openai_client.py` — Azure OpenAI Integration

```python
import os
import json
import logging
from openai import AzureOpenAI

def get_openai_client() -> AzureOpenAI:
    return AzureOpenAI(
        azure_endpoint=os.environ["AZURE_OPENAI_ENDPOINT"],
        api_key=os.environ["AZURE_OPENAI_KEY"],
        api_version="2024-02-01"
    )

def analyse_transcript(transcript_text: str, reference_doc_text: str) -> dict:
    """
    Send transcript + reference document to Azure OpenAI.
    Returns structured JSON with meeting analysis.
    """
    client = get_openai_client()
    deployment = os.environ["AZURE_OPENAI_DEPLOYMENT"]

    system_prompt = """
You are an expert meeting analyst. You will be given:
1. A meeting transcript (in VTT format)
2. A reference document from SharePoint

Your job is to produce a structured JSON summary that can be used to populate a report template.

Return ONLY valid JSON with this exact structure:
{
  "meeting_title": "string",
  "meeting_date": "YYYY-MM-DD",
  "attendees": ["name1", "name2"],
  "executive_summary": "2-3 sentence summary",
  "key_decisions": ["decision 1", "decision 2"],
  "action_items": [
    {"owner": "name", "task": "description", "due_date": "YYYY-MM-DD or TBD"}
  ],
  "discussion_topics": [
    {"topic": "string", "summary": "string"}
  ],
  "reference_doc_relevance": "How the reference doc relates to what was discussed",
  "risks_and_issues": ["risk 1", "risk 2"],
  "next_steps": ["step 1", "step 2"]
}
"""

    user_message = f"""
## TRANSCRIPT:
{transcript_text[:8000]}

## REFERENCE DOCUMENT FROM SHAREPOINT:
{reference_doc_text[:3000]}

Analyse the above and return the JSON structure.
"""

    response = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ],
        temperature=0.1,
        max_tokens=2000,
        response_format={"type": "json_object"}
    )

    raw = response.choices[0].message.content
    return json.loads(raw)
```

---

### `shared/queue_client.py` — Azure Storage Queue

```python
import os
import json
import base64
import logging
from azure.storage.queue import QueueClient

def get_queue_client() -> QueueClient:
    conn_str = os.environ["STORAGE_CONNECTION_STRING"]
    queue_name = os.environ["QUEUE_NAME"]
    return QueueClient.from_connection_string(conn_str, queue_name)

def enqueue_notification(data: dict) -> None:
    """Serialise and enqueue a notification payload."""
    client = get_queue_client()
    message = json.dumps(data)
    # Azure Storage Queue requires base64-encoded messages
    encoded = base64.b64encode(message.encode("utf-8")).decode("utf-8")
    client.send_message(encoded)
    logging.info(f"Enqueued message for resource: {data.get('resource', 'unknown')}")
```

---

### `shared/document_filler.py` — Fill Word Template with OpenAI Output

```python
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

PLACEHOLDER_MAP = {
    "{{MEETING_TITLE}}":       lambda d: d.get("meeting_title", ""),
    "{{MEETING_DATE}}":        lambda d: d.get("meeting_date", ""),
    "{{EXECUTIVE_SUMMARY}}":   lambda d: d.get("executive_summary", ""),
    "{{REFERENCE_RELEVANCE}}": lambda d: d.get("reference_doc_relevance", ""),
}

def fill_template(template_bytes: bytes, analysis: dict) -> bytes:
    """
    Takes a .docx template as bytes, replaces placeholders,
    appends tables for action items and attendees,
    returns filled .docx as bytes.

    In your SharePoint template, add these text placeholders:
        {{MEETING_TITLE}}, {{MEETING_DATE}}, {{EXECUTIVE_SUMMARY}},
        {{REFERENCE_RELEVANCE}}, {{ACTION_ITEMS_TABLE}}, {{ATTENDEES_LIST}}
    """
    doc = Document(io.BytesIO(template_bytes))

    # ── Replace simple text placeholders ───────────────────────────────────────
    for paragraph in doc.paragraphs:
        for placeholder, getter in PLACEHOLDER_MAP.items():
            if placeholder in paragraph.text:
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, getter(analysis))

    # Also replace in tables that already exist in the template
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for placeholder, getter in PLACEHOLDER_MAP.items():
                        if placeholder in para.text:
                            for run in para.runs:
                                if placeholder in run.text:
                                    run.text = run.text.replace(
                                        placeholder, getter(analysis)
                                    )

    # ── Append key decisions section ───────────────────────────────────────────
    doc.add_heading("Key Decisions", level=2)
    for decision in analysis.get("key_decisions", []):
        para = doc.add_paragraph(decision, style="List Bullet")

    # ── Append action items as a table ─────────────────────────────────────────
    doc.add_heading("Action Items", level=2)
    action_items = analysis.get("action_items", [])
    if action_items:
        table = doc.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Owner"
        hdr[1].text = "Task"
        hdr[2].text = "Due Date"
        for item in action_items:
            row = table.add_row().cells
            row[0].text = item.get("owner", "")
            row[1].text = item.get("task", "")
            row[2].text = item.get("due_date", "TBD")

    # ── Append next steps ──────────────────────────────────────────────────────
    doc.add_heading("Next Steps", level=2)
    for step in analysis.get("next_steps", []):
        doc.add_paragraph(step, style="List Number")

    # ── Append attendees ───────────────────────────────────────────────────────
    doc.add_heading("Attendees", level=2)
    for attendee in analysis.get("attendees", []):
        doc.add_paragraph(attendee, style="List Bullet")

    # ── Return as bytes ────────────────────────────────────────────────────────
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()
```

---

### `shared/__init__.py`

```python
# Shared module package
```

---

### `function_app.py` — All Azure Functions

```python
import azure.functions as func
import json
import logging
import os
import base64
import requests
from datetime import datetime

from shared.auth import get_access_token
from shared.graph_client import (
    create_subscription,
    list_subscriptions,
    renew_subscription,
    delete_subscription,
    get_transcript_content,
    get_meeting_details,
    get_transcript_metadata,
    get_recording_content_url,
)
from shared.sharepoint_client import (
    download_template,
    upload_file,
    create_folder,
)
from shared.openai_client import analyse_transcript
from shared.document_filler import fill_template
from shared.queue_client import enqueue_notification

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


# ══════════════════════════════════════════════════════════════════════════════
# FUNCTION 1 — Manual trigger: Subscribe to Graph change notifications
# ══════════════════════════════════════════════════════════════════════════════

@app.route(route="subscribe", methods=["POST"])
def fn_subscribe(req: func.HttpRequest) -> func.HttpResponse:
    """
    Call this manually (via HTTP POST) to create a Graph subscription.
    The notificationUrl must be the URL of fn_receive_notification below.

    Call body (JSON):
    {
        "notification_url": "https://YOUR_FUNCTION_APP.azurewebsites.net/api/receive_notification?code=YOUR_KEY"
    }
    """
    logging.info("fn_subscribe triggered")

    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse("Request body must be JSON", status_code=400)

    notification_url = body.get("notification_url")
    if not notification_url:
        return func.HttpResponse(
            "Missing 'notification_url' in request body", status_code=400
        )

    client_state = os.environ["NOTIFICATION_CLIENT_STATE"]

    try:
        # List existing subscriptions to avoid duplicates
        existing = list_subscriptions()
        logging.info(f"Found {len(existing)} existing subscriptions")

        result = create_subscription(notification_url, client_state)
        logging.info(f"Created subscription: {result['id']}")

        return func.HttpResponse(
            json.dumps({
                "status": "success",
                "subscription_id": result["id"],
                "expires": result["expirationDateTime"]
            }),
            status_code=200,
            mimetype="application/json"
        )

    except Exception as e:
        logging.error(f"Failed to create subscription: {e}")
        return func.HttpResponse(f"Error: {str(e)}", status_code=500)


# ══════════════════════════════════════════════════════════════════════════════
# FUNCTION 2 — HTTP trigger: Receive Graph change notifications
# ══════════════════════════════════════════════════════════════════════════════

@app.route(route="receive_notification", methods=["POST", "GET"])
def fn_receive_notification(req: func.HttpRequest) -> func.HttpResponse:
    """
    Microsoft Graph sends notifications to this endpoint.

    Two scenarios:
    1. Validation handshake (GET with validationToken param): must echo back token
    2. Actual notification (POST): validate clientState and enqueue the payload
    """
    logging.info("fn_receive_notification triggered")

    # ── Scenario 1: Graph validation handshake ─────────────────────────────────
    # When you first subscribe, Graph sends a GET with a validationToken.
    # You MUST return it as plain text within 10 seconds or subscription fails.
    validation_token = req.params.get("validationToken")
    if validation_token:
        logging.info("Responding to Graph validation handshake")
        return func.HttpResponse(
            validation_token,
            status_code=200,
            mimetype="text/plain"
        )

    # ── Scenario 2: Actual change notification ─────────────────────────────────
    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse("Invalid JSON", status_code=400)

    expected_state = os.environ["NOTIFICATION_CLIENT_STATE"]
    notifications = body.get("value", [])

    for notification in notifications:
        # Security check: verify clientState matches what we set
        if notification.get("clientState") != expected_state:
            logging.warning("clientState mismatch — ignoring notification")
            continue

        resource = notification.get("resource", "")
        change_type = notification.get("changeType", "")
        logging.info(f"Received notification: {change_type} on {resource}")

        # Build a payload to enqueue for async processing
        payload = {
            "resource": resource,
            "changeType": change_type,
            "subscriptionId": notification.get("subscriptionId"),
            "tenantId": notification.get("tenantId"),
            "resourceData": notification.get("resourceData", {}),
            "receivedAt": datetime.utcnow().isoformat()
        }

        enqueue_notification(payload)

    # IMPORTANT: Must return 202 Accepted quickly.
    # Graph will retry if you don't respond fast enough.
    return func.HttpResponse("", status_code=202)


# ══════════════════════════════════════════════════════════════════════════════
# FUNCTION 3 — Queue trigger: Process each notification from the queue
# ══════════════════════════════════════════════════════════════════════════════

@app.queue_trigger(
    arg_name="msg",
    queue_name="%QUEUE_NAME%",
    connection="STORAGE_CONNECTION_STRING"
)
def fn_process_queue(msg: func.QueueMessage) -> None:
    """
    Triggered automatically whenever a message appears in the queue.

    Steps:
    1. Parse the queue message
    2. Extract meeting/transcript IDs from the Graph resource path
    3. Fetch transcript content from Graph
    4. Download reference document from SharePoint
    5. Analyse with Azure OpenAI
    6. Fill the Word template
    7. Upload everything to SharePoint output folder
    """
    logging.info("fn_process_queue triggered")

    # ── Step 1: Parse message ──────────────────────────────────────────────────
    try:
        raw = msg.get_body().decode("utf-8")
        # Azure Storage Queue messages may be base64-encoded
        try:
            decoded = base64.b64decode(raw).decode("utf-8")
        except Exception:
            decoded = raw
        payload = json.loads(decoded)
    except Exception as e:
        logging.error(f"Failed to parse queue message: {e}")
        return

    logging.info(f"Processing notification for resource: {payload.get('resource')}")

    # ── Step 2: Extract IDs from Graph resource path ───────────────────────────
    # Resource path example:
    # /communications/onlineMeetings/MSp.../transcripts/MSt...
    resource = payload.get("resource", "")
    resource_data = payload.get("resourceData", {})

    meeting_id = None
    transcript_id = None

    # Parse IDs from the resource URL
    parts = resource.split("/")
    for i, part in enumerate(parts):
        if part == "onlineMeetings" and i + 1 < len(parts):
            meeting_id = parts[i + 1]
        if part == "transcripts" and i + 1 < len(parts):
            transcript_id = parts[i + 1]

    # Also try resourceData
    if not meeting_id:
        meeting_id = resource_data.get("id")
    if not transcript_id:
        transcript_id = resource_data.get("id")

    if not meeting_id or not transcript_id:
        logging.error(f"Could not extract IDs from resource: {resource}")
        return

    # ── Step 3: Fetch transcript from Graph ────────────────────────────────────
    try:
        transcript_text = get_transcript_content(meeting_id, transcript_id)
        transcript_meta = get_transcript_metadata(meeting_id, transcript_id)
        meeting_details = get_meeting_details(meeting_id)
        logging.info(f"Fetched transcript ({len(transcript_text)} chars)")
    except Exception as e:
        logging.error(f"Failed to fetch transcript: {e}")
        return

    # ── Step 4: Download SharePoint reference document ─────────────────────────
    site_id = os.environ["SHAREPOINT_SITE_ID"]
    template_file_id = os.environ["SHAREPOINT_TEMPLATE_FILE_ID"]

    try:
        template_bytes = download_template(site_id, template_file_id)
        logging.info("Downloaded SharePoint template")

        # Extract text from the reference doc for OpenAI context
        import io
        from docx import Document as DocxDocument
        ref_doc = DocxDocument(io.BytesIO(template_bytes))
        reference_text = "\n".join([p.text for p in ref_doc.paragraphs if p.text.strip()])
    except Exception as e:
        logging.error(f"Failed to download SharePoint template: {e}")
        return

    # ── Step 5: Analyse with Azure OpenAI ─────────────────────────────────────
    try:
        analysis = analyse_transcript(transcript_text, reference_text)
        logging.info(f"OpenAI analysis complete: {list(analysis.keys())}")
    except Exception as e:
        logging.error(f"OpenAI analysis failed: {e}")
        return

    # ── Step 6: Fill the Word template ────────────────────────────────────────
    try:
        filled_doc_bytes = fill_template(template_bytes, analysis)
        logging.info("Template filled successfully")
    except Exception as e:
        logging.error(f"Failed to fill template: {e}")
        return

    # ── Step 7: Create output folder and upload everything ────────────────────
    output_folder_root = os.environ["SHAREPOINT_OUTPUT_FOLDER"]
    meeting_subject = meeting_details.get("subject", "Meeting")
    # Sanitise folder name
    safe_subject = "".join(
        c if c.isalnum() or c in " _-" else "_"
        for c in meeting_subject
    )[:50]
    folder_name = f"{safe_subject}_{datetime.utcnow().strftime('%Y%m%d_%H%M')}"

    try:
        create_folder(site_id, output_folder_root, folder_name)
        folder_path = f"{output_folder_root}/{folder_name}"

        # Upload filled report document
        upload_file(
            site_id,
            folder_path,
            "Meeting_Report.docx",
            filled_doc_bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        logging.info("Uploaded filled report document")

        # Upload raw transcript as .vtt file
        upload_file(
            site_id,
            folder_path,
            "transcript.vtt",
            transcript_text.encode("utf-8"),
            "text/vtt"
        )
        logging.info("Uploaded transcript")

        # Upload analysis JSON (useful for debugging / downstream processing)
        upload_file(
            site_id,
            folder_path,
            "analysis.json",
            json.dumps(analysis, indent=2).encode("utf-8"),
            "application/json"
        )
        logging.info("Uploaded analysis JSON")

        # Note: Recording download requires a separate step — recordings are
        # served via a temporary streaming URL. To download and re-upload the
        # recording video, use get_recording_content_url() and then stream
        # the bytes to upload_file(). Recordings can be large (100s of MB).
        # Consider using Azure Blob Storage for large files.

        logging.info(
            f"✅ Pipeline complete! Files uploaded to: {folder_path}"
        )

    except Exception as e:
        logging.error(f"Failed to upload files to SharePoint: {e}")
        return


# ══════════════════════════════════════════════════════════════════════════════
# FUNCTION 4 — Timer trigger: Renew Graph subscription every 45 minutes
# ══════════════════════════════════════════════════════════════════════════════

@app.timer_trigger(
    schedule="0 */45 * * * *",   # every 45 minutes
    arg_name="timer"
)
def fn_renew_subscription(timer: func.TimerRequest) -> None:
    """
    Graph subscriptions for onlineMeetings expire after 60 minutes.
    This timer function renews them automatically.
    """
    logging.info("fn_renew_subscription triggered")
    try:
        subs = list_subscriptions()
        for sub in subs:
            renew_subscription(sub["id"])
            logging.info(f"Renewed subscription: {sub['id']}")
        if not subs:
            logging.warning("No subscriptions found to renew")
    except Exception as e:
        logging.error(f"Failed to renew subscriptions: {e}")
```

---

## PHASE 5 — Deploy to Azure

### Step 5.1 — Deploy the Function App

```bash
# From your project folder:
func azure functionapp publish teams-transcript-pipeline --python
```

After deployment, get your function URLs from:
**portal.azure.com** → your Function App → **Functions** → click each function → **Get Function URL**

---

### Step 5.2 — Get the `fn_receive_notification` URL

From the Azure Portal, copy the URL for `receive_notification`. It looks like:
```
https://teams-transcript-pipeline.azurewebsites.net/api/receive_notification?code=AbCdEf...
```
Save this — you'll use it in the next step.

---

### Step 5.3 — Subscribe to Graph Notifications

Call the `fn_subscribe` function with a POST request:

```bash
curl -X POST \
  "https://teams-transcript-pipeline.azurewebsites.net/api/subscribe?code=YOUR_SUBSCRIBE_KEY" \
  -H "Content-Type: application/json" \
  -d '{
    "notification_url": "https://teams-transcript-pipeline.azurewebsites.net/api/receive_notification?code=YOUR_RECEIVE_KEY"
  }'
```

A successful response looks like:
```json
{
  "status": "success",
  "subscription_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "expires": "2024-01-15T10:30:00Z"
}
```

---

### Step 5.4 — Get Your SharePoint Site ID and Template File ID

Run this once to find your SharePoint IDs (put your values in):

```python
# run_once_get_ids.py — run locally, then delete
import os
from shared.sharepoint_client import get_sharepoint_site_id, get_file_id_by_path
from shared.auth import get_access_token

os.environ["TENANT_ID"] = "YOUR_TENANT_ID"
os.environ["CLIENT_ID"] = "YOUR_CLIENT_ID"
os.environ["CLIENT_SECRET"] = "YOUR_CLIENT_SECRET"

# Get site ID
site_id = get_sharepoint_site_id(
    "yourcompany.sharepoint.com",   # e.g. contoso.sharepoint.com
    "/sites/YourSiteName"
)
print(f"SHAREPOINT_SITE_ID = {site_id}")

# Get template file ID (path relative to the document library root)
file_id = get_file_id_by_path(site_id, "Templates/MeetingReportTemplate.docx")
print(f"SHAREPOINT_TEMPLATE_FILE_ID = {file_id}")
```

Add both values to your Function App **Application settings**.

---

## PHASE 6 — Prepare Your SharePoint Template

Your Word template (`.docx`) should contain these text placeholders that the code will replace:

```
{{MEETING_TITLE}}
{{MEETING_DATE}}
{{EXECUTIVE_SUMMARY}}
{{REFERENCE_RELEVANCE}}
```

The code automatically appends these sections after the placeholders:
- Key Decisions (bullet list)
- Action Items (table with Owner / Task / Due Date)
- Next Steps (numbered list)
- Attendees (bullet list)

---

## PHASE 7 — Testing End to End

### Local Testing

```bash
# Start the function app locally
func start

# Test the subscribe function (triggers validation + subscription)
curl -X POST http://localhost:7071/api/subscribe \
  -H "Content-Type: application/json" \
  -d '{"notification_url": "https://YOUR_NGROK_OR_PUBLIC_URL/api/receive_notification"}'

# Simulate a Graph notification manually
curl -X POST http://localhost:7071/api/receive_notification \
  -H "Content-Type: application/json" \
  -d '{
    "value": [{
      "changeType": "created",
      "clientState": "mySecretState123",
      "resource": "/communications/onlineMeetings/MEETING_ID/transcripts/TRANSCRIPT_ID",
      "resourceData": {"id": "TRANSCRIPT_ID"},
      "subscriptionId": "test-sub-id",
      "tenantId": "YOUR_TENANT_ID"
    }]
  }'
```

> **Tip for local testing:** Use [ngrok](https://ngrok.com) to expose your local port publicly so Graph can send notifications to it:
> ```bash
> ngrok http 7071
> # Use the https URL ngrok gives you as your notification_url
> ```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Subscription creation fails with 403 | Admin consent not granted for API permissions |
| Subscription creation fails with 400 | Notification URL not publicly accessible or not HTTPS |
| `clientState` mismatch in logs | `NOTIFICATION_CLIENT_STATE` env var doesn't match what was used at subscription time |
| Queue messages not processed | Check `STORAGE_CONNECTION_STRING` and `QUEUE_NAME` app settings match |
| OpenAI returns invalid JSON | Ensure `response_format: json_object` is set; check deployment name |
| SharePoint upload 403 | Ensure `Sites.ReadWrite.All` and `Files.ReadWrite.All` permissions are granted |
| Subscription expires / no notifications | Confirm timer function `fn_renew_subscription` is deployed and running |

---

## Security Checklist

- [ ] `local.settings.json` is in `.gitignore`
- [ ] All secrets are in Azure Application Settings, not in code
- [ ] `clientState` is validated on every incoming notification
- [ ] Function URLs use function-level auth keys (`?code=...`)
- [ ] Client secret has an expiry date — calendar a reminder to rotate it
