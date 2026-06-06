"""
================================================================================
 THE SHARED PIPELINE  ->  shared/processor.py
================================================================================

THE MOST IMPORTANT FILE
-----------------------
`process_message` is the single function that does ALL the real work. Both
triggers (the HTTP one and the Queue one in function_app.py) simply call it.
That is how we follow the requirement "share the same processing logic to avoid
code duplication": the triggers are thin shells, this is the brain.

IT RUNS THE SIX STEPS IN ORDER
------------------------------
  Step 1: extract meeting info from the message       (message_parser)
  Step 2: read metadata from the SharePoint list      (sharepoint_client)
  Step 3: log an audit snapshot                        (logging)
  Step 4: run the document through AI                  (ai_processor.processAI)
  Step 5: save the AI output back to SharePoint        (sharepoint_client)
  Step 6: move the recording files                     (sharepoint_client)

ERROR PHILOSOPHY
----------------
For fatal problems (e.g. message can't be parsed, required metadata missing) we
RAISE. The HTTP trigger turns that into a 500 response; the Queue trigger lets
Azure retry and eventually poison-queue the message - both good for auditing.
Step 6 (moving recordings) is treated as non-fatal: if it fails we log it and
still report success for the parts that worked, because the document was already
produced and saved.
================================================================================
"""

import logging
from typing import Any

from shared.ai_processor import processAI
from shared.config import Config
from shared.message_parser import parse_message
from shared.sharepoint_client import SharePointClient


def process_message(message: str) -> dict[str, Any]:
    """Run the full transcript-processing pipeline for ONE queue message.

    Args:
        message: the raw queue message text (JSON).

    Returns:
        A small dict summarising what happened (ids, where the output was saved,
        which recordings were moved). The HTTP trigger returns this to the
        caller; the Queue trigger ignores the return value.
    """
    logging.info("=== Transcript pipeline started ===")

    # Fail fast if any required app setting is missing.
    Config.validate()

    # --- Step 1: Extract meeting information -------------------------------- #
    meeting = parse_message(message)
    logging.info("Step 1 - meeting info extracted: %s", meeting.to_log_dict())

    # We cannot look anything up in SharePoint without these two keys.
    if not meeting.join_web_url or not meeting.meeting_id:
        raise ValueError(
            "Cannot continue: JoinWebUrl and MeetingId are both required to "
            "look up SharePoint metadata."
        )

    # Open ONE SharePoint connection and reuse it for steps 2, 4-input, 5 and 6.
    sp = SharePointClient()

    # --- Step 2: Retrieve additional metadata from SharePoint --------------- #
    metadata = sp.get_metadata(meeting.join_web_url, meeting.meeting_id)
    logging.info("Step 2 - SharePoint metadata retrieved: %s", metadata.to_log_dict())

    # --- Step 3: Logging (a single audit snapshot of everything so far) ----- #
    logging.info(
        "Step 3 - audit snapshot: meeting=%s metadata=%s",
        meeting.to_log_dict(),
        metadata.to_log_dict(),
    )

    # --- Step 4: Process transcript using AI -------------------------------- #
    if not metadata.admin_updated_folder_link:
        raise ValueError("AdminUpdatedFolderLink is empty; nothing to feed to AI.")

    # Read the source document, then hand its bytes to processAI.
    source_content = sp.read_file(metadata.admin_updated_folder_link)
    logging.info(
        "Step 4 - read %d bytes from AdminUpdatedFolderLink; invoking processAI.",
        len(source_content),
    )
    output_document = processAI(source_content)
    logging.info("Step 4 - processAI returned %d bytes.", len(output_document))

    # --- Step 5: Save output document --------------------------------------- #
    if not metadata.file_to_be_saved:
        raise ValueError("FileToBeSaved is empty; cannot save AI output.")
    saved_path = sp.save_file(metadata.file_to_be_saved, output_document)
    logging.info("Step 5 - output document saved to %s", saved_path)

    # --- Step 6: Move recordings (non-fatal if it fails) -------------------- #
    moved: list[str] = []
    if metadata.recordings_link:
        try:
            moved = sp.move_folder_contents(
                source_folder=metadata.admin_updated_folder_link,
                destination_folder=metadata.recordings_link,
            )
            logging.info("Step 6 - moved %d recording file(s).", len(moved))
        except Exception:  # noqa: BLE001 - log and keep the overall run successful
            logging.exception("Step 6 - failed to move recordings (continuing).")
    else:
        logging.info("Step 6 - no RecordingsLink provided; skipping move.")

    # Build and log the result summary.
    result = {
        "meeting_id": meeting.meeting_id,
        "join_web_url": meeting.join_web_url,
        "saved_document": saved_path,
        "recordings_moved": moved,
    }
    logging.info("=== Transcript pipeline completed: %s ===", result)
    return result
