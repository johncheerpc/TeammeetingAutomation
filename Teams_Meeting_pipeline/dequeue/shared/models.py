"""
================================================================================
 DATA MODELS  ->  shared/models.py
================================================================================

WHAT IS THIS?
-------------
Simple "data container" classes that carry information between the pipeline
steps. They are plain Python `@dataclass` objects - think of them as labelled
boxes. Using them (instead of passing loose dictionaries around) means:
  - your editor can autocomplete `.meeting_id` etc.,
  - a typo like `.meting_id` is caught instead of silently returning None,
  - the SHAPE of the data is documented in one place.

There is nothing Azure-specific here - it is ordinary Python.
================================================================================
"""

from dataclasses import dataclass, field, asdict
from typing import Any


@dataclass
class MeetingInfo:
    """Information extracted from the incoming queue message (Step 1).

    Each field maps to one piece of meeting data we need downstream. The two
    most important are `join_web_url` and `meeting_id`, because we use them to
    look the meeting up in SharePoint.
    """

    join_web_url: str | None = None          # the Teams "join" URL of the meeting
    meeting_id: str | None = None            # the Teams/Graph meeting id
    transcript_info: dict[str, Any] = field(default_factory=dict)  # transcript ids/urls
    meeting_duration: str | None = None      # e.g. "PT45M" (ISO 8601 duration)
    meeting_date: str | None = None          # e.g. "2026-06-05T14:00:00Z"
    raw: dict[str, Any] = field(default_factory=dict)  # the full original payload, kept for reference

    # NOTE on `field(default_factory=dict)`: for mutable defaults like dict/list
    # you must use default_factory, NOT `= {}`. Otherwise every instance would
    # accidentally share the SAME dictionary. This is a common Python gotcha.

    def to_log_dict(self) -> dict[str, Any]:
        """Return a compact view suitable for logging.

        We deliberately leave out `raw` because it can be large and noisy in the
        logs. This keeps the audit log readable.
        """
        return {
            "join_web_url": self.join_web_url,
            "meeting_id": self.meeting_id,
            "transcript_info": self.transcript_info,
            "meeting_duration": self.meeting_duration,
            "meeting_date": self.meeting_date,
        }


@dataclass
class SharePointMetadata:
    """The columns we read back from the SharePoint custom list (Step 2).

    These four columns tell the pipeline WHERE to read the source document,
    WHERE to save the AI output, and WHERE to move the recordings.
    """

    admin_updated_folder_link: str | None = None  # source document -> input to processAI
    file_to_be_saved: str | None = None           # where to save the AI output
    recordings_link: str | None = None            # where to move the recordings
    transcript_link: str | None = None            # link to the transcript (for reference/logging)

    def to_log_dict(self) -> dict[str, Any]:
        """Convert this dataclass to a plain dict for logging.

        `asdict` (from the dataclasses module) turns the object into a regular
        dictionary automatically.
        """
        return asdict(self)
