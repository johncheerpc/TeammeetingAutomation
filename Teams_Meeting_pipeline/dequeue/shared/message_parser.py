"""
================================================================================
 STEP 1 - PARSE THE QUEUE MESSAGE  ->  shared/message_parser.py
================================================================================

GOAL
----
Turn the raw text of the queue message into a tidy `MeetingInfo` object that the
rest of the pipeline can use.

BACKGROUND (why this is defensive)
----------------------------------
The message comes from a Microsoft Graph "change notification" that Microsoft
sends when a Teams meeting transcript is created. Depending on how the
notification was configured (and on whatever Logic App / Power Automate flow
pushed it onto the queue), the JSON can look different:

  - A Graph "changeNotificationCollection":   {"value": [ {notification}, ... ]}
  - A single notification object:             { ...fields... }
  - A flattened custom message someone built: { "JoinWebUrl": "...", ... }

Field names may also differ in casing (joinWebUrl vs JoinWebUrl). So instead of
assuming one exact shape, this parser looks in the likely places and is
case-insensitive. Being forgiving here saves you from brittle failures later.
================================================================================
"""

import json
import logging
import re
from typing import Any

from shared.models import MeetingInfo


def _first(d: dict[str, Any], *keys: str) -> Any:
    """Return the value of the first matching key, ignoring upper/lower case.

    Example:
        _first(data, "joinWebUrl", "JoinWebUrl") will return whichever of those
        two keys exists in `data`, or None if neither does.
    """
    # Build a copy of the dict with all keys lower-cased so we can match
    # regardless of how the source capitalised them.
    lowered = {k.lower(): v for k, v in d.items()}
    for key in keys:
        if key.lower() in lowered:
            return lowered[key.lower()]
    return None


def _extract_meeting_id_from_resource(resource: str | None) -> str | None:
    """Recover the meeting id from a Graph "resource" path, if present.

    Graph notifications include a resource string that looks like:
        communications/onlineMeetings('MSox...')/transcripts('abc-123')
    The bit inside onlineMeetings('...') is the meeting id. We use a regular
    expression to pull it out as a fallback when the id is not given directly.
    """
    if not resource:
        return None
    match = re.search(r"onlineMeetings\('([^']+)'\)", resource)
    return match.group(1) if match else None


def parse_message(message: str) -> MeetingInfo:
    """Parse the raw queue message string into a `MeetingInfo`.

    Args:
        message: the raw text of the queue message (JSON).

    Returns:
        A populated `MeetingInfo`.

    Raises:
        ValueError: if the message is not valid JSON or has an unexpected shape.
    """
    # 1) Turn the JSON text into a Python object (dict/list).
    try:
        payload = json.loads(message)
    except (json.JSONDecodeError, TypeError) as exc:
        # Re-raise with a clearer message; `from exc` keeps the original cause.
        raise ValueError(f"Queue message is not valid JSON: {exc}") from exc

    # 2) Figure out where the actual notification lives.
    notification: dict[str, Any]
    if isinstance(payload, dict) and isinstance(payload.get("value"), list) and payload["value"]:
        # Graph collection form: take the first notification in the array.
        notification = payload["value"][0]
    elif isinstance(payload, dict):
        # Already a single notification object.
        notification = payload
    else:
        raise ValueError("Unexpected message shape; expected a JSON object.")

    # Graph puts some details under "resourceData"; default to {} if absent.
    resource_data = notification.get("resourceData") or {}
    if not isinstance(resource_data, dict):
        resource_data = {}

    # 3) Start building our tidy object. Keep the original payload in `raw`.
    info = MeetingInfo(raw=notification)

    # JoinWebUrl: check the top level first, then resourceData.
    info.join_web_url = _first(notification, "joinWebUrl", "JoinWebUrl") or _first(
        resource_data, "joinWebUrl", "JoinWebUrl"
    )

    # MeetingId: top level -> resourceData (id/meetingId) -> parsed from resource.
    info.meeting_id = (
        _first(notification, "meetingId", "MeetingId")
        or _first(resource_data, "meetingId", "MeetingId", "id")
        or _extract_meeting_id_from_resource(notification.get("resource"))
    )

    # Transcript-related details, grouped together for logging/traceability.
    info.transcript_info = {
        "resource": notification.get("resource"),
        "transcript_id": _first(resource_data, "transcriptId", "id"),
        "content_url": _first(notification, "transcriptContentUrl", "contentUrl"),
        "change_type": notification.get("changeType"),  # e.g. "created"
    }

    # Duration and date - try a few common key names for each.
    info.meeting_duration = _first(notification, "meetingDuration", "duration", "MeetingDuration")
    info.meeting_date = _first(
        notification, "meetingDate", "startDateTime", "MeetingDate", "createdDateTime"
    )

    # 4) Warn (don't crash yet) if the two critical lookup keys are missing.
    #    The orchestrator decides whether to stop; here we just flag it.
    if not info.join_web_url or not info.meeting_id:
        logging.warning(
            "Message parsed but JoinWebUrl/MeetingId could not be fully resolved. "
            "join_web_url=%s meeting_id=%s",
            info.join_web_url,
            info.meeting_id,
        )

    return info
