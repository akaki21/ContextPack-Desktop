"""Parse structured progress events emitted by the PowerShell runner."""

from __future__ import annotations

import json
from typing import Any


def parse_runner_event(line: str) -> dict[str, Any] | None:
    """Parse a ContextPack event and ignore ordinary tool output safely."""

    try:
        payload = json.loads(line)
    except json.JSONDecodeError:
        return None
    if not isinstance(payload, dict) or payload.get("contextpack_event") != 1:
        return None
    return payload

