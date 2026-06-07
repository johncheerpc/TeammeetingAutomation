"""
================================================================================
 CONFIGURATION  ->  shared/config.py   (LOCAL version)
================================================================================

WHERE DO THESE VALUES COME FROM NOW?
------------------------------------
In the Azure Functions version these came from "Application settings". Running
locally there is no Azure app to hold them, so instead we read them from
ENVIRONMENT VARIABLES on your machine. The easiest way to set those is a small
text file named `.env` placed next to main.py - main.py loads it for you.

So the flow is:
    .env file  ->  loaded into environment variables  ->  read here with os.environ

WHY KEEP A DEDICATED CONFIG MODULE?
-----------------------------------
Same reasons as before: one place that names every setting, plus a `validate()`
helper that fails immediately with a clear message if you forgot one - instead
of a confusing crash deep inside the SharePoint code.

NEVER put real secrets directly in source files. Keep them in `.env`, and make
sure `.env` is git-ignored so it is not committed.
================================================================================
"""

import os
from pathlib import Path

# Load the project's `.env` file (if present) into environment variables, so the
# settings below can be read. python-dotenv is optional: if it isn't installed,
# we just rely on environment variables that are already set.
try:
    from dotenv import load_dotenv

    # The .env file lives in the project root - one level up from this `shared`
    # folder. Pointing at it explicitly means it loads no matter where you run
    # the script from.
    load_dotenv(Path(__file__).resolve().parent.parent / ".env")
except ImportError:
    pass


def _get(name: str, default: str | None = None) -> str | None:
    """Read one environment variable by name.

    Args:
        name:    the variable name, e.g. "SharePointSiteUrl".
        default: value to use if the variable is not present.
    """
    return os.environ.get(name, default)


class Config:
    """A typed, central view of every setting this project needs."""

    # --- SharePoint site / list ------------------------------------------ #
    SHAREPOINT_SITE_URL: str | None = _get("SharePointSiteUrl")    # e.g. https://contoso.sharepoint.com/sites/Meetings
    SHAREPOINT_LIST_NAME: str | None = _get("SharePointListName")  # the custom list's title

    # --- SharePoint authentication (Azure AD "app-only" login) ----------- #
    # From an App Registration in Microsoft Entra ID (Azure AD). "App-only"
    # means the script signs in as itself using a client id + secret - ideal
    # for an automation script with no interactive user.
    SHAREPOINT_TENANT_ID: str | None = _get("SharePointTenantId")
    SHAREPOINT_CLIENT_ID: str | None = _get("SharePointClientId")
    SHAREPOINT_CLIENT_SECRET: str | None = _get("SharePointClientSecret")

    # --- AI processing (optional) ---------------------------------------- #
    AI_ENDPOINT: str | None = _get("AIEndpoint")      # e.g. https://<res>.openai.azure.com/
    AI_API_KEY: str | None = _get("AIApiKey")
    AI_DEPLOYMENT: str | None = _get("AIDeployment")  # the model deployment name

    @classmethod
    def validate(cls) -> None:
        """Stop early with a clear error if a REQUIRED setting is missing.

        Turns a confusing "NoneType has no attribute..." crash into one obvious
        message naming exactly which variable you still need to set in `.env`.
        """
        required = {
            "SharePointSiteUrl": cls.SHAREPOINT_SITE_URL,
            "SharePointListName": cls.SHAREPOINT_LIST_NAME,
            "SharePointTenantId": cls.SHAREPOINT_TENANT_ID,
            "SharePointClientId": cls.SHAREPOINT_CLIENT_ID,
            "SharePointClientSecret": cls.SHAREPOINT_CLIENT_SECRET,
        }
        missing = [name for name, value in required.items() if not value]
        if missing:
            raise EnvironmentError(
                "Missing required setting(s): " + ", ".join(missing)
                + ". Add them to your .env file (see .env.example)."
            )
