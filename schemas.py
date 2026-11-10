"""Tool schemas — what the LLM sees."""

from __future__ import annotations

from typing import Any, Dict

STORE: Dict[str, Any] = {
    "type": "string",
    "description": (
        "Mailbox/store to target: a display name (or unique substring of it) or a 0-based index "
        "into the mailbox list from outlook_list_mailboxes. Omit to use the default mailbox (for "
        "folder/email tools) or ALL mailboxes (for search). Shared mailboxes, PST archives and "
        "online archives each appear as their own store."
    ),
}

EMAIL_SUMMARY: Dict[str, Any] = {
    "type": "object",
    "properties": {
        "entry_id": {"type": "string", "description": "Stable Outlook EntryID — pass this to outlook_get_email."},
        "subject": {"type": "string"},
        "sender": {"type": "string", "description": "Sender display name."},
        "sender_email": {"type": "string"},
        "to": {"type": "string", "description": "Recipients (display) as Outlook returns them."},
        "received_time": {"type": "string", "description": "ISO-8601 local time."},
        "unread": {"type": "boolean"},
        "has_attachments": {"type": "boolean"},
        "attachment_count": {"type": "integer"},
        "size": {"type": "integer", "description": "Item size in bytes."},
        "categories": {"type": "array", "items": {"type": "string"}},
        "store": {"type": "string", "description": "Which mailbox this email came from."},
        "folder": {"type": "string", "description": "Folder path within the mailbox."},
    },
}

STATUS: Dict[str, Any] = {
    "name": "outlook_status",
    "description": (
        "Check that classic Outlook is reachable via COM: report the Outlook version, whether it was "
        "already running, the default mailbox, and the list of all connected mailboxes/stores "
        "(accounts, shared mailboxes, PST archives). Use this first to confirm the bridge works and "
        "to see which mailboxes are available before searching."
    ),
    "parameters": {"type": "object", "properties": {}},
}

LIST_MAILBOXES: Dict[str, Any] = {
    "name": "outlook_list_mailboxes",
    "description": (
        "List every mailbox/store the current Outlook profile can see — each email account, shared "
        "mailbox, PST file and online archive — with its kind and top-level folder count. Use the "
        "returned name (or index) as the `store` argument for the other tools."
    ),
    "parameters": {"type": "object", "properties": {}},
}

LIST_FOLDERS: Dict[str, Any] = {
    "name": "outlook_list_folders",
    "description": (
        "List folders in a mailbox, optionally recursively. Returns each folder's name, path, id and "
        "item count. Use the path (e.g. 'Inbox/Projects') as the `folder` argument for search."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "store": STORE,
            "recursive": {
                "type": "boolean",
                "description": "Also list subfolders (default false).",
            },
        },
    },
}

SEARCH_EMAILS: Dict[str, Any] = {
    "name": "outlook_search_emails",
    "description": (
        "Scan emails across one mailbox or ALL mailboxes and return matching messages (newest first). "
        "Filters combine with AND. Date and unread filters are pushed into Outlook's fast index; "
        "subject/sender/recipient/body are matched client-side. Omit `store` to scan every mailbox. "
        "Omit `folder` to scan all mail folders (system folders like Deleted Items/Junk are skipped "
        "unless you name a folder explicitly or set skip_system_folders=false). Results are capped at "
        "max_results (default 50); use entry_id from a result to fetch the full body via "
        "outlook_get_email."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "store": STORE,
            "folder": {
                "type": "string",
                "description": "Folder path to limit the search to, e.g. 'Inbox' or 'Inbox/Projects'. Omit to scan all mail folders.",
            },
            "subject": {"type": "string", "description": "Substring match on subject (case-insensitive)."},
            "sender": {"type": "string", "description": "Substring match on sender name or address."},
            "recipient": {"type": "string", "description": "Substring match on the To line."},
            "body": {"type": "string", "description": "Substring match on message body (slow — only set when needed)."},
            "unread_only": {"type": "boolean", "description": "Only unread messages."},
            "has_attachments": {"type": "boolean", "description": "Only messages with attachments."},
            "received_after": {"type": "string", "description": "ISO date/datetime, e.g. '2024-01-01' or '2024-01-01T09:00:00'."},
            "received_before": {"type": "string", "description": "ISO date/datetime."},
            "include_subfolders": {"type": "boolean", "description": "Descend into subfolders (default true)."},
            "skip_system_folders": {"type": "boolean", "description": "Skip Deleted Items/Junk/Outbox/etc. (default true)."},
            "max_results": {"type": "integer", "description": "Max results (1-500, default 50)."},
        },
    },
}

GET_EMAIL: Dict[str, Any] = {
    "name": "outlook_get_email",
    "description": (
        "Fetch one email by its EntryID and return full details: subject, sender, all recipients "
        "(to/cc/bcc), times, text body (truncated to max_body_chars), attachment list, categories, "
        "importance and unread flag. Pass the entry_id you got from outlook_search_emails together "
        "with the same store name."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "entry_id": {"type": "string", "description": "Outlook EntryID from a search result."},
            "store": STORE,
            "max_body_chars": {"type": "integer", "description": "Truncate body to N chars (default 20000, 0 = no body)."},
            "include_html": {"type": "boolean", "description": "Also return the HTML body (can be large; default false)."},
        },
        "required": ["entry_id"],
    },
}

SEND_EMAIL: Dict[str, Any] = {
    "name": "outlook_send_email",
    "description": (
        "Compose and send (or draft) an email through classic Outlook. By default (send=false) "
        "the message is opened as a draft in Outlook so the user can review and click Send; set "
        "send=true only when the email should actually be sent immediately. Attachments are "
        "absolute paths to local files."
    ),
    "parameters": {
        "type": "object",
        "properties": {
            "to": {"type": "string", "description": "Recipients (To), semicolon-separated email addresses or names."},
            "account": {"type": "string", "description": "Which of your email accounts sends the message — an SMTP address or display name (e.g. 'lauritz.loy@ifco.com'). Valid values are in outlook_status `accounts`. Omit to use the default account."},
            "cc": {"type": "string", "description": "CC recipients, semicolon-separated."},
            "bcc": {"type": "string", "description": "BCC recipients, semicolon-separated."},
            "subject": {"type": "string", "description": "Subject line."},
            "body": {"type": "string", "description": "Email body (plain text, or HTML when body_is_html=true)."},
            "body_is_html": {"type": "boolean", "description": "Treat body as HTML (default false)."},
            "attachments": {"type": "array", "items": {"type": "string"}, "description": "Absolute paths to local files to attach."},
            "importance": {"type": "integer", "description": "0=low, 1=normal, 2=high (default normal)."},
            "send": {"type": "boolean", "description": "true = send now; false (default) = open as a draft for review."},
        },
        "required": ["to"],
    },
}

TOOL_SCHEMAS = (
    STATUS,
    LIST_MAILBOXES,
    LIST_FOLDERS,
    SEARCH_EMAILS,
    GET_EMAIL,
    SEND_EMAIL,
)

__all__ = [
    "EMAIL_SUMMARY",
    "GET_EMAIL",
    "LIST_FOLDERS",
    "LIST_MAILBOXES",
    "SEARCH_EMAILS",
    "SEND_EMAIL",
    "STATUS",
    "STORE",
    "TOOL_SCHEMAS",
]
