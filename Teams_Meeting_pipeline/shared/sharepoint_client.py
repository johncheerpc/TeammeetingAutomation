# =============================================================================
# FILE: shared/sharepoint_client.py
#
# PURPOSE:
#   All functions that interact with Microsoft SharePoint.
#   We use SharePoint for two things:
#     1. DOWNLOAD the Word template document (the blank report template)
#     2. UPLOAD the finished files (filled report, transcript, analysis JSON)
#
# HOW SHAREPOINT IS ACCESSED:
#   SharePoint is accessed through the same Microsoft Graph API as Teams.
#   The URL pattern for SharePoint files is:
#     https://graph.microsoft.com/v1.0/sites/{siteId}/drive/...
#
# WHAT IS A SITE ID?
#   Every SharePoint site has a unique ID. It looks like:
#     "yourcompany.sharepoint.com,abc123...,def456..."
#   Run the helper script run_once_get_sharepoint_ids.py to find yours.
#
# USED BY: function_app.py (fn_process_queue)
# =============================================================================

import requests   # For making HTTP calls
import logging    # For log messages
from .auth import get_access_token  # Our Managed Identity auth

# Base URL for Microsoft Graph v1.0 (stable)
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    """
    Builds the standard HTTP headers for SharePoint API requests.
    Same pattern as graph_client.py — Authorization + Content-Type.

    RETURNS:
        {"Authorization": "Bearer <token>", "Content-Type": "application/json"}
    """
    token = get_access_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }


# =============================================================================
# DISCOVERY HELPERS
# Run these ONCE locally to find your SharePoint IDs, then put them
# in your Azure Function App settings. You won't call these at runtime.
# =============================================================================

def get_sharepoint_site_id(hostname: str, site_path: str) -> str:
    """
    Looks up the unique ID of a SharePoint site.

    YOU ONLY NEED TO CALL THIS ONCE to find your site ID.
    Then save the result as the SHAREPOINT_SITE_ID app setting.

    PARAMETERS:
        hostname  → your SharePoint domain
                    Example: "contoso.sharepoint.com"

        site_path → the path to your specific site
                    Example: "/sites/TeamSite"
                    (found in your SharePoint URL:
                     https://contoso.sharepoint.com/sites/TeamSite)

    RETURNS:
        A site ID string like: "contoso.sharepoint.com,abc123,def456"

    EXAMPLE CALL:
        site_id = get_sharepoint_site_id(
            "contoso.sharepoint.com",
            "/sites/TeamSite"
        )
        print(site_id)  # Copy this into your app settings
    """
    # Graph URL format for looking up a site: /sites/{hostname}:{path}
    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    response = requests.get(url, headers=_headers())
    response.raise_for_status()
    # The response JSON has an "id" field with the site ID
    return response.json()["id"]


def get_file_id_by_path(site_id: str, file_path: str) -> str:
    """
    Looks up a file's unique Graph ID using its path in SharePoint.

    YOU ONLY NEED TO CALL THIS ONCE to find your template file's ID.
    Then save the result as the SHAREPOINT_TEMPLATE_FILE_ID app setting.

    PARAMETERS:
        site_id   → your SharePoint site ID (from get_sharepoint_site_id)

        file_path → path to the file, relative to the document library root
                    Example: "Templates/MeetingReportTemplate.docx"
                    This means the file is at:
                    SharePoint Site → Documents → Templates → MeetingReportTemplate.docx

    RETURNS:
        A file ID string like: "01ABCDE..."

    EXAMPLE CALL:
        file_id = get_file_id_by_path(site_id, "Templates/MeetingReportTemplate.docx")
        print(file_id)  # Copy this into your app settings
    """
    # Graph URL format for looking up a file by path: /sites/{siteId}/drive/root:/{path}
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{file_path}"
    response = requests.get(url, headers=_headers())
    response.raise_for_status()
    return response.json()["id"]


# =============================================================================
# RUNTIME FUNCTIONS
# These are called automatically during the pipeline.
# =============================================================================

def download_template(site_id: str, file_id: str) -> bytes:
    """
    Downloads a file from SharePoint and returns it as bytes (raw binary data).

    THIS IS CALLED AUTOMATICALLY during fn_process_queue to download
    the Word template before filling it with meeting data.

    PARAMETERS:
        site_id → your SharePoint site ID (from SHAREPOINT_SITE_ID app setting)
        file_id → the file's unique ID (from SHAREPOINT_TEMPLATE_FILE_ID app setting)

    RETURNS:
        bytes — the raw binary content of the .docx file
        These bytes are then passed to document_filler.py to fill in the content.

    EXAMPLE OF WHAT HAPPENS NEXT:
        template_bytes = download_template(site_id, file_id)
        # template_bytes is now the .docx file in memory as binary data
        # We pass it to fill_template() which opens it as a Word document
    """
    # Graph URL for downloading file content by item ID
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/items/{file_id}/content"

    token = get_access_token()

    # allow_redirects=True → SharePoint sends a redirect to the actual file location
    # We want requests to automatically follow that redirect
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        allow_redirects=True
    )
    response.raise_for_status()

    # .content returns the raw bytes of the file
    return response.content


def create_folder(site_id: str, parent_path: str, folder_name: str) -> dict:
    """
    Creates a new folder inside SharePoint.

    THIS IS CALLED AUTOMATICALLY during fn_process_queue to create
    a unique output folder for each processed meeting.

    EXAMPLE FOLDER STRUCTURE CREATED:
        ProcessedMeetings/
            Weekly_Standup_20240115_0930/
                Meeting_Report.docx
                transcript.vtt
                analysis.json

    PARAMETERS:
        site_id     → your SharePoint site ID

        parent_path → the folder where the new folder will be created
                      Example: "ProcessedMeetings"
                      This folder must already exist in SharePoint.

        folder_name → the name of the new folder to create
                      Example: "Weekly_Standup_20240115_0930"

    RETURNS:
        dict with details of the created folder including its ID and URL

    CONFLICT BEHAVIOR:
        If a folder with the same name already exists, SharePoint will
        automatically rename the new one (e.g. "Meeting_1", "Meeting_2").
        This is controlled by "@microsoft.graph.conflictBehavior": "rename"
    """
    # URL to list/create children of a specific folder path
    url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{parent_path}:/children"

    payload = {
        "name": folder_name,  # The name of the folder to create
        "folder": {},         # Empty "folder" object tells Graph this is a folder (not a file)
        # If a folder with this name exists, rename the new one instead of failing
        "@microsoft.graph.conflictBehavior": "rename"
    }

    response = requests.post(url, headers=_headers(), json=payload)
    response.raise_for_status()
    return response.json()


def upload_file(
    site_id: str,
    folder_path: str,
    filename: str,
    content: bytes,
    content_type: str = "application/octet-stream"
) -> dict:
    """
    Uploads a file to a SharePoint folder. Best for files up to ~4 MB.

    For larger files (like video recordings), use upload_large_file() instead.

    THIS IS CALLED AUTOMATICALLY during fn_process_queue to upload:
        - Meeting_Report.docx (the filled Word template)
        - transcript.vtt      (the raw transcript text)
        - analysis.json       (the OpenAI JSON output)

    PARAMETERS:
        site_id      → your SharePoint site ID

        folder_path  → path to the folder where the file will be uploaded
                       Example: "ProcessedMeetings/Weekly_Standup_20240115_0930"

        filename     → what to name the uploaded file
                       Example: "Meeting_Report.docx"

        content      → the file content as bytes
                       Example: b'\x50\x4B\x03\x04...' (binary file data)

        content_type → the MIME type of the file (tells SharePoint what kind of file it is)
                       Common values:
                         "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                         → for .docx Word files
                         "text/vtt"         → for transcript .vtt files
                         "application/json" → for .json files
                         "video/mp4"        → for .mp4 video recordings

    RETURNS:
        dict with details of the uploaded file including its ID, URL, and name

    HOW THE URL IS BUILT:
        We use "path-based" upload: /sites/{siteId}/drive/root:/{folder}/{filename}:/content
        The ":/content" at the end means "upload the content of this file path"
    """
    # Build the upload URL using the folder path + filename
    url = (
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/"
        f"{folder_path}/{filename}:/content"
    )

    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": content_type  # Must match the type of file being uploaded
    }

    # PUT request with the file bytes as the request body
    response = requests.put(url, headers=headers, data=content)
    response.raise_for_status()
    return response.json()


def upload_large_file(
    site_id: str,
    folder_path: str,
    filename: str,
    content: bytes,
    content_type: str = "video/mp4"
) -> dict:
    """
    Uploads a large file (>4 MB) to SharePoint using a resumable upload session.

    WHY NEEDED:
        The simple upload_file() function has a ~4 MB limit.
        Meeting recordings are typically 100 MB to several GB.
        For large files we use "upload sessions" which split the file
        into chunks and upload them one at a time.

    HOW IT WORKS:
        Step 1: Ask SharePoint for an "upload session" URL
                (a temporary URL that accepts chunks)
        Step 2: Upload the file in 5 MB chunks, one by one
                Each chunk tells SharePoint where it fits:
                "bytes 0-5242879/1073741824" (bytes start-end/total)
        Step 3: After the last chunk, SharePoint assembles the file

    PARAMETERS:
        site_id      → your SharePoint site ID
        folder_path  → path to the destination folder
        filename     → name for the uploaded file (e.g. "recording.mp4")
        content      → the full file as bytes
        content_type → MIME type (default "video/mp4" for recordings)

    RETURNS:
        dict with the completed file's details
    """

    # ── Step 1: Create an upload session ──────────────────────────────────────
    # This asks SharePoint: "I want to upload a large file, give me a temporary URL"
    session_url = (
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/"
        f"{folder_path}/{filename}:/createUploadSession"
    )
    session_payload = {
        "item": {
            "@microsoft.graph.conflictBehavior": "rename",  # Rename if file exists
            "name": filename
        }
    }
    session_resp = requests.post(session_url, headers=_headers(), json=session_payload)
    session_resp.raise_for_status()

    # SharePoint gives us a temporary "uploadUrl" — valid for a few hours
    upload_url = session_resp.json()["uploadUrl"]
    logging.info(f"Upload session created for {filename}")

    # ── Step 2: Upload in 5 MB chunks ─────────────────────────────────────────
    chunk_size = 5 * 1024 * 1024  # 5 MB in bytes (5 × 1024 × 1024)
    total_size = len(content)     # Total file size in bytes
    offset = 0                    # Where we are in the file (starts at byte 0)
    resp = None

    while offset < total_size:
        # Slice out the next chunk from the content bytes
        chunk = content[offset: offset + chunk_size]
        end = offset + len(chunk) - 1  # Last byte index of this chunk

        # Content-Range header tells SharePoint where this chunk fits
        # Format: "bytes {start}-{end}/{total}"
        # Example: "bytes 0-5242879/52428800" means "bytes 0 to 5MB of a 50MB file"
        headers = {
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {offset}-{end}/{total_size}",
            "Content-Type": content_type
        }

        resp = requests.put(upload_url, headers=headers, data=chunk)
        resp.raise_for_status()

        offset += chunk_size  # Move to the next chunk
        logging.info(f"Uploaded {min(offset, total_size)}/{total_size} bytes of {filename}")

    # After the last chunk, SharePoint returns details of the completed file
    return resp.json()
