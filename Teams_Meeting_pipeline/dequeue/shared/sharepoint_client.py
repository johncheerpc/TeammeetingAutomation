"""
================================================================================
 STEPS 2, 5 & 6 - SHAREPOINT OPERATIONS  ->  shared/sharepoint_client.py
================================================================================

WHAT THIS DOES
--------------
All the talking-to-SharePoint happens here, behind a small, friendly class so
the rest of the code never deals with raw REST calls:
  - Step 2: read metadata columns from the custom list.
  - (Step 4 input): download the source document for the AI.
  - Step 5: upload the AI output document.
  - Step 6: move the recording files.

LIBRARY
-------
We use `office365-rest-python-client` (installed via requirements.txt). It is a
community library that wraps the SharePoint REST API in Python objects.

KEY IDEA: "execute_query()"
---------------------------
This library is LAZY. When you write `list.items.filter(...).get()` it only
builds up the request - nothing has been sent yet. You must call
`.execute_query()` to actually send it to SharePoint and get the data back.
If you forget it, you will get empty/unexecuted objects. Watch for that line.

AUTHENTICATION
--------------
We sign in "app-only" with a client id + secret from an Entra ID (Azure AD) App
Registration. That app needs permission on the site (typically
Sites.ReadWrite.All, granted by an admin, or a site-scoped grant).

SERVER-RELATIVE URLs
--------------------
SharePoint identifies files by a "server-relative URL" - the path AFTER the
domain, e.g. "/sites/Meetings/Shared Documents/file.docx". Users often paste
FULL URLs (https://contoso.sharepoint.com/sites/...). The `_to_server_relative`
helper converts either form into what the API expects.
================================================================================
"""

import logging
import posixpath                      # posixpath = forward-slash path math (right for URLs)
from urllib.parse import urlparse     # to split a full URL into parts

from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext

from shared.config import Config
from shared.models import SharePointMetadata


class SharePointClient:
    """A small, reusable wrapper around a SharePoint connection (ClientContext)."""

    def __init__(self) -> None:
        """Create and authenticate the SharePoint connection.

        Called once per pipeline run. `Config.validate()` makes sure all the
        required settings exist before we try to connect.
        """
        Config.validate()

        # Bundle the app-registration client id + secret into a credential.
        credentials = ClientCredential(
            Config.SHAREPOINT_CLIENT_ID,
            Config.SHAREPOINT_CLIENT_SECRET,
        )

        # `ClientContext` represents your connection/session to ONE site.
        self.ctx = ClientContext(Config.SHAREPOINT_SITE_URL).with_credentials(credentials)

        # Remember just the path part of the site URL (e.g. "/sites/Meetings").
        # We need it when converting relative paths to server-relative ones.
        self._site_path = urlparse(Config.SHAREPOINT_SITE_URL).path.rstrip("/")

    # ------------------------------------------------------------------ #
    # STEP 2 - read the metadata row from the custom list
    # ------------------------------------------------------------------ #
    def get_metadata(self, join_web_url: str, meeting_id: str) -> SharePointMetadata:
        """Find the list item matching JoinWebUrl + MeetingId and return its columns.

        Args:
            join_web_url: value to match against the list's `JoinWebUrl` column.
            meeting_id:   value to match against the list's `MeetingId` column.

        Returns:
            A `SharePointMetadata` with the four columns we care about.

        Raises:
            LookupError: if no matching list item is found.
        """
        # Get the list by its display title.
        target_list = self.ctx.web.lists.get_by_title(Config.SHAREPOINT_LIST_NAME)

        # Build an OData filter (SharePoint's query language for the REST API).
        # OData escapes a single quote by DOUBLING it (' -> ''), so we do that to
        # avoid breaking the query if a value contains a quote.
        safe_url = (join_web_url or "").replace("'", "''")
        safe_id = (meeting_id or "").replace("'", "''")
        query = f"JoinWebUrl eq '{safe_url}' and MeetingId eq '{safe_id}'"

        # .filter(query) -> WHERE clause; .top(1) -> only need the first match;
        # .get() -> build a read request; .execute_query() -> actually send it.
        items = target_list.items.filter(query).top(1).get().execute_query()

        if len(items) == 0:
            raise LookupError(
                f"No SharePoint list item found for JoinWebUrl='{join_web_url}', "
                f"MeetingId='{meeting_id}'."
            )

        # `.properties` is a dict of all the item's column values. We read the
        # four columns by their INTERNAL names (which may differ from display
        # names - verify these in your list if a value comes back as None).
        props = items[0].properties
        return SharePointMetadata(
            admin_updated_folder_link=props.get("AdminUpdatedFolderLink"),
            file_to_be_saved=props.get("FileToBeSaved"),
            recordings_link=props.get("RecordingsLink"),
            transcript_link=props.get("TranscriptLink"),
        )

    # ------------------------------------------------------------------ #
    # STEP 4 INPUT - download a document's bytes
    # ------------------------------------------------------------------ #
    def read_file(self, server_relative_or_full_url: str) -> bytes:
        """Download a file from SharePoint and return its raw bytes.

        Args:
            server_relative_or_full_url: either a full https URL or a
                server-relative path to the file.

        Returns:
            The file content as bytes (ready to pass to processAI).
        """
        server_relative_url = self._to_server_relative(server_relative_or_full_url)
        file_obj = self.ctx.web.get_file_by_server_relative_url(server_relative_url)
        # get_content() reads the bytes; remember execute_query() to run it.
        content = file_obj.get_content().execute_query()
        return content.value  # `.value` holds the actual bytes

    # ------------------------------------------------------------------ #
    # STEP 5 - upload the AI output document
    # ------------------------------------------------------------------ #
    def save_file(self, destination: str, content: bytes) -> str:
        """Upload `content` to `destination` (a folder path + file name).

        Args:
            destination: where to save, e.g.
                "/sites/Meetings/Shared Documents/Output/summary.docx".
            content: the bytes to write (the document returned by processAI).

        Returns:
            The server-relative path the file was saved to.
        """
        server_relative = self._to_server_relative(destination)

        # Split "/folder/path/file.ext" into the folder and the file name.
        folder_url = posixpath.dirname(server_relative)   # "/folder/path"
        file_name = posixpath.basename(server_relative)   # "file.ext"

        folder = self.ctx.web.get_folder_by_server_relative_url(folder_url)
        # upload_file creates (or overwrites) the file inside that folder.
        folder.upload_file(file_name, content).execute_query()

        logging.info("Saved output document to %s", server_relative)
        return server_relative

    # ------------------------------------------------------------------ #
    # STEP 6 - move recording files
    # ------------------------------------------------------------------ #
    def move_file(self, source: str, destination_folder: str) -> str:
        """Move ONE file into `destination_folder`.

        Args:
            source: the file to move (full or server-relative URL).
            destination_folder: the folder to move it into.

        Returns:
            The new server-relative path of the moved file.
        """
        src_rel = self._to_server_relative(source)
        dst_folder_rel = self._to_server_relative(destination_folder)

        file_name = posixpath.basename(src_rel)
        dst_rel = posixpath.join(dst_folder_rel, file_name)  # folder + same file name

        source_file = self.ctx.web.get_file_by_server_relative_url(src_rel)
        # moveto(target, flag): flag=1 means "overwrite if it already exists".
        source_file.moveto(dst_rel, 1).execute_query()

        logging.info("Moved recording %s -> %s", src_rel, dst_rel)
        return dst_rel

    def move_folder_contents(self, source_folder: str, destination_folder: str) -> list[str]:
        """Move EVERY file in `source_folder` into `destination_folder`.

        Returns:
            A list of the new paths for all files that were moved.
        """
        src_rel = self._to_server_relative(source_folder)
        folder = self.ctx.web.get_folder_by_server_relative_url(src_rel)

        # List the files in the folder (remember execute_query()).
        files = folder.files.get().execute_query()

        moved: list[str] = []
        for f in files:
            # Reuse move_file for each one. `serverRelativeUrl` is already the
            # correct relative path, so move_file's normaliser leaves it as-is.
            moved.append(self.move_file(f.serverRelativeUrl, destination_folder))
        return moved

    # ------------------------------------------------------------------ #
    # HELPER - normalise any URL/path to "server-relative"
    # ------------------------------------------------------------------ #
    def _to_server_relative(self, url: str) -> str:
        """Convert a full URL, an absolute path, or a relative path to the
        server-relative form SharePoint expects.

        Examples:
            "https://contoso.sharepoint.com/sites/Meetings/Docs/a.txt"
                                            -> "/sites/Meetings/Docs/a.txt"
            "/sites/Meetings/Docs/a.txt"    -> unchanged
            "Docs/a.txt"                    -> "/sites/Meetings/Docs/a.txt"
        """
        if not url:
            raise ValueError("A SharePoint URL/path is required but was empty.")
        if url.lower().startswith("http"):
            # Full URL -> keep only the path part (drop scheme + domain).
            return urlparse(url).path
        if url.startswith("/"):
            # Already an absolute server-relative path.
            return url
        # Otherwise treat it as relative to the configured site.
        return posixpath.join(self._site_path, url)
