"""Tool handlers — what runs when the model calls an outlook_* tool.

Each handler is ``guarded``: failures become ``{"ok": false, "error": ...}`` JSON
instead of an exception. Handlers receive the tool's ``args`` dict plus Hermes's
signature-inspected context kwargs.
"""

from __future__ import annotations

from typing import Any, Dict, Mapping

from .. import client
from .shared import error_result, guarded, ok_result


def _run(action: str, args: Mapping[str, Any]) -> str:
    res = client.run(action, dict(args))
    if res.get("ok") is False:
        return error_result(res.get("error") or "unknown error")
    data = res.get("data") if isinstance(res.get("data"), dict) else res
    return ok_result(_summary(action, data), data)


def _summary(action: str, data: Dict[str, Any]) -> str:
    if action == "status":
        return (
            f"Outlook {data.get('outlook_version', '?')} — {data.get('store_count', 0)} mailbox(es) "
            f"(default: {data.get('default_store', '?')})."
        )
    if action == "list_mailboxes":
        return f"{data.get('mailbox_count', 0)} mailbox(es)."
    if action == "list_folders":
        return f"{data.get('folder_count', 0)} folder(s) in {data.get('store', '?')}."
    if action == "search_emails":
        n = data.get("result_count", 0)
        trunc = " (truncated at max_results)" if data.get("truncated") else ""
        return f"Found {n} email(s){trunc}."
    if action == "get_email":
        email = data.get("email") or {}
        return (
            f"Email: {email.get('subject', '(no subject)')} "
            f"from {email.get('sender_name') or email.get('sender_email') or '(unknown)'} "
            f"— {email.get('attachment_count', 0)} attachment(s)."
        )
    return "Done."


@guarded
def outlook_status(args: Mapping[str, Any], task_id: str = "default", **kwargs: Any) -> str:
    del task_id, kwargs
    return _run("status", args)


@guarded
def outlook_list_mailboxes(args: Mapping[str, Any], task_id: str = "default", **kwargs: Any) -> str:
    del task_id, kwargs
    return _run("list_mailboxes", args)


@guarded
def outlook_list_folders(args: Mapping[str, Any], task_id: str = "default", **kwargs: Any) -> str:
    del task_id, kwargs
    return _run("list_folders", args)


@guarded
def outlook_search_emails(args: Mapping[str, Any], task_id: str = "default", **kwargs: Any) -> str:
    del task_id, kwargs
    return _run("search_emails", args)


@guarded
def outlook_get_email(args: Mapping[str, Any], task_id: str = "default", **kwargs: Any) -> str:
    del task_id, kwargs
    return _run("get_email", args)


HANDLERS: Dict[str, Any] = {
    "outlook_status": outlook_status,
    "outlook_list_mailboxes": outlook_list_mailboxes,
    "outlook_list_folders": outlook_list_folders,
    "outlook_search_emails": outlook_search_emails,
    "outlook_get_email": outlook_get_email,
}

__all__ = ["HANDLERS"]
