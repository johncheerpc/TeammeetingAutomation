"""
================================================================================
 CONFIGURATION  ->  shared/config.py
================================================================================

WHAT ARE "APP SETTINGS"?
------------------------
Azure Functions stores configuration as "Application settings" - these are just
environment variables. In the cloud you set them in:
    Function App  ->  Settings  ->  Environment variables / Configuration
When running locally, the file `local.settings.json` provides the same values.

Either way, your Python code reads them with `os.environ`. This is the secure,
recommended place for things like connection strings, site URLs and secrets -
NEVER hard-code those in your source files.

WHY A DEDICATED CONFIG MODULE?
------------------------------
Reading `os.environ[...]` scattered all over the codebase is error-prone. We
centralise every setting here, give each a clear Python name, and add a
`validate()` helper that fails fast with a readable message if something a
function needs was not configured.
================================================================================
"""

import os


def _get(name: str, default: str | None = None) -> str | None:
    """Read one app setting (environment variable) by name.

    Args:
        name:    the exact app-setting key, e.g. "SharePointSiteUrl".
        default: value to use if the setting is not present.

    Returns:
        The setting value, or `default` if it was not set.
    """
    return os.environ.get(name, default)


class Config:
    """A typed, central view of every app setting this project uses.

    Each class attribute below is read ONCE when the module is first imported.
    That is fine for Functions because the worker process keeps these in memory.
    """

    # --- Queue ------------------------------------------------------------ #
    # The name of the storage queue the production trigger listens to.
    QUEUE_NAME: str = _get("QueueName", "transcript-notifications")
    # The CONNECTION to that queue is referenced by the *name* "QueueConnection"
    # inside function_app.py. The actual connection string lives in app settings
    # under that key - we deliberately do not load the secret into code here.

    # --- SharePoint site / list ------------------------------------------ #
    SHAREPOINT_SITE_URL: str | None = _get("SharePointSiteUrl")    # e.g. https://contoso.sharepoint.com/sites/Meetings
    SHAREPOINT_LIST_NAME: str | None = _get("SharePointListName")  # the custom list's title

    # --- SharePoint authentication (Azure AD "app-only" login) ----------- #
    # These come from an App Registration in Microsoft Entra ID (Azure AD).
    # "App-only" means the Function signs in as itself (no user), using a
    # client id + secret, which is ideal for background/automation jobs.
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

        Calling this at the start of the pipeline turns a confusing
        "NoneType has no attribute..." crash deep in the code into a single,
        obvious message naming exactly which app setting you forgot to add.
        """
        required = {
            "SharePointSiteUrl": cls.SHAREPOINT_SITE_URL,
            "SharePointListName": cls.SHAREPOINT_LIST_NAME,
            "SharePointTenantId": cls.SHAREPOINT_TENANT_ID,
            "SharePointClientId": cls.SHAREPOINT_CLIENT_ID,
            "SharePointClientSecret": cls.SHAREPOINT_CLIENT_SECRET,
        }
        # Collect the names of any settings that are empty/missing.
        missing = [name for name, value in required.items() if not value]
        if missing:
            raise EnvironmentError(
                "Missing required application setting(s): " + ", ".join(missing)
            )
