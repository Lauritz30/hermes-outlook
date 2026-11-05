"""Shared helpers for the outlook tool handlers."""

from __future__ import annotations

import functools
import json
from typing import Any


def dumps(payload: Any) -> str:
    return json.dumps(payload, ensure_ascii=False)


def ok_result(summary: str, data: Any = None) -> str:
    """Successful tool result: a model-facing ``summary`` plus the raw result payload."""
    return dumps({"ok": True, "summary": summary, "data": {} if data is None else data})


def error_result(message: str) -> str:
    """Error tool result. Handlers never raise — the model always gets JSON back."""
    return dumps({"ok": False, "error": message})


def guarded(handler):
    """Wrap a handler so any failure becomes an error JSON string instead of an exception."""

    @functools.wraps(handler)
    def wrapper(*args: Any, **kwargs: Any) -> str:
        try:
            return handler(*args, **kwargs)
        except Exception as error:  # noqa: BLE001 - the model needs the message, not a traceback
            return error_result(str(error) or error.__class__.__name__)

    return wrapper


__all__ = ["dumps", "error_result", "guarded", "ok_result"]
