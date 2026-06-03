# =============================================================================
# FILE: shared/auth.py
#
# PURPOSE:
#   This file handles authentication (proving who we are) to Microsoft.
#   Instead of using a username/password or secret keys, we use
#   "Managed Identity" — Azure automatically gives our Function App
#   a trusted identity, so we never have to store any passwords.
#
# HOW MANAGED IDENTITY WORKS (simple explanation):
#   Think of it like a staff ID badge. When you work at a company,
#   you don't need to prove who you are with a password every time
#   you enter a room — your badge does it. Managed Identity works
#   the same way: Azure gives your Function App a "badge" (identity)
#   and Microsoft services (Graph, OpenAI, SharePoint) trust it.
#
# USED BY:
#   - shared/graph_client.py   (to call Microsoft Graph API)
#   - shared/sharepoint_client.py (to access SharePoint)
#   - shared/openai_client.py  (to call Azure OpenAI)
#
# NO SECRETS NEEDED — nothing to configure here.
# =============================================================================

# DefaultAzureCredential is a smart credential that automatically picks
# the right authentication method depending on where the code is running:
#   - Running in Azure Function App → uses Managed Identity (automatic)
#   - Running on your local laptop  → uses your "az login" session
# You do NOT need to change anything for it to work in both places.
from azure.identity import DefaultAzureCredential


def get_access_token() -> str:
    """
    Gets a short-lived access token (like a temporary pass) that proves
    our Function App is allowed to call Microsoft Graph API.

    WHAT IS AN ACCESS TOKEN?
        A long string of characters (looks like random gibberish) that
        Microsoft gives us after we authenticate. We attach this string
        to every API request we make, and Microsoft knows we are authorised.
        Tokens expire after ~1 hour, but DefaultAzureCredential handles
        renewal automatically — we don't need to worry about it.

    RETURNS:
        A string like "eyJ0eXAiOiJKV1QiLCJub....(very long string)"
        This gets put in the HTTP header: "Authorization: Bearer <token>"

    EXAMPLE OF HOW IT IS USED (you don't call this directly):
        token = get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get("https://graph.microsoft.com/...", headers=headers)
    """

    # "scope" tells Microsoft WHAT service we want access to.
    # "https://graph.microsoft.com/.default" means:
    #   "Give me access to Microsoft Graph with all the permissions
    #    that have been granted to this app in the Azure Portal."
    scope = "https://graph.microsoft.com/.default"

    # Create a credential object — this does NOT make any network call yet.
    # It just sets up the authentication strategy.
    credential = DefaultAzureCredential()

    # NOW make the network call to get the actual token.
    # get_token() contacts Microsoft's login service and returns a token object.
    # token.token  → the actual token string
    # token.expires_on → when it expires (handled automatically)
    token = credential.get_token(scope)

    # Return just the token string (not the whole object)
    return token.token
