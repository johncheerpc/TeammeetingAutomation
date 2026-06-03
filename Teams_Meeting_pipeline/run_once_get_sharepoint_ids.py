"""
run_once_get_sharepoint_ids.py

Run this script ONCE locally to discover your SharePoint site ID and template file ID.
Copy the printed values into your Azure Function App Application Settings.

Usage:
    pip install msal requests python-docx
    python run_once_get_sharepoint_ids.py
"""

import os
import sys

# ── Fill these in before running ──────────────────────────────────────────────
TENANT_ID       = "YOUR_TENANT_ID"
CLIENT_ID       = "YOUR_CLIENT_ID"
CLIENT_SECRET   = "YOUR_CLIENT_SECRET"
SHAREPOINT_HOST = "yourcompany.sharepoint.com"   # e.g. contoso.sharepoint.com
SITE_PATH       = "/sites/YourSiteName"           # e.g. /sites/TeamSite
TEMPLATE_PATH   = "Templates/MeetingReportTemplate.docx"  # relative to document library root
# ─────────────────────────────────────────────────────────────────────────────

os.environ["TENANT_ID"]     = TENANT_ID
os.environ["CLIENT_ID"]     = CLIENT_ID
os.environ["CLIENT_SECRET"] = CLIENT_SECRET

sys.path.insert(0, os.path.dirname(__file__))
from shared.sharepoint_client import get_sharepoint_site_id, get_file_id_by_path

print("\n🔍 Fetching SharePoint IDs...\n")

site_id = get_sharepoint_site_id(SHAREPOINT_HOST, SITE_PATH)
print(f"SHAREPOINT_SITE_ID       = {site_id}")

file_id = get_file_id_by_path(site_id, TEMPLATE_PATH)
print(f"SHAREPOINT_TEMPLATE_FILE_ID = {file_id}")

print("\n✅ Copy both values into your Azure Function App → Configuration → Application Settings")
