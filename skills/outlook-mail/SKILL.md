---
name: outlook-mail
description: Read email from classic Outlook (desktop) with the hermes-outlook extension. Use when asked to check, read, search, scan or summarize email, mailboxes, inboxes, or folders in Microsoft Outlook — across all accounts, shared mailboxes and PST archives.
version: 1.0.0
author: Lauritz Loy, Hermes Agent
license: MIT
platforms: [windows]
metadata:
  hermes:
    tags: [outlook, email, mail, exchange, office, com]
---

# Reading email from classic Outlook

This plugin reads mail from the local classic Outlook (desktop) through a Windows PowerShell
COM bridge. It can see **every mailbox the current Outlook profile has open** — each email
account, auto-mapped shared mailboxes, and PST archives — with no Graph API or M365 admin
consent.

Use the `outlook_*` tools (they start with `outlook_`). This skill covers *how* to use them well.

## The core loop (do this every time)

1. **Check reachability** with `outlook_status` — confirms the COM bridge works and lists every
   mailbox/store that is available. Note the mailbox names (and the default).
2. **Pick the mailbox(es).** Pass a mailbox name (or a unique substring, or its index from
   `outlook_list_mailboxes`) as `store`. Omit `store` in `outlook_search_emails` to scan **all**
   mailboxes.
3. **Search** with `outlook_search_emails`, then **read** promising hits with `outlook_get_email`
   (pass the returned `entry_id` *and* the same `store`).

## Which tool

| Goal | Tool |
|---|---|
| Is Outlook reachable? Which mailboxes exist? | `outlook_status` |
| List mailboxes with kind + folder counts | `outlook_list_mailboxes` |
| List folders in one mailbox (optionally recursive) | `outlook_list_folders` |
| Scan/search messages across one or all mailboxes | `outlook_search_emails` |
| Read one full email (body, recipients, attachments) | `outlook_get_email` |

## Searching well

- **"Scan all my mailboxes"** → `outlook_search_emails` with no `store`. It walks every store's
  mail folders. System folders (Deleted Items, Junk, Outbox, Drafts, Sync Issues, RSS Feeds) are
  skipped by default — set `skip_system_folders: false` to include them, or target a folder
  explicitly with `folder`.
- **Filters combine with AND.** `subject`, `sender`, `recipient`, `body` are case-insensitive
  substring matches. Date (`received_after`/`received_before`) and `unread_only` go through
  Outlook's fast index; `body` is slow (it reads each message body) — only set it when you must.
- **Newest first.** Results are sorted by received time, newest first, capped at `max_results`
  (default 50, max 500).
- **Sender** matches against display name *or* email address.

## Reading a message

`outlook_get_email` needs the `entry_id` from a search result and (when the message lives in a
non-default mailbox) the same `store` name — EntryIDs are only unique within their store. It
returns subject, sender, to/cc/bcc, sent/received times, the plain-text body (truncated to
`max_body_chars`, default 20000), and the attachment list (name + size, not content).

## Limits & gotchas

- **Windows + classic Outlook only.** The "new" Outlook does not expose the same COM object;
  this plugin targets classic Outlook (the Office desktop app).
- **Outlook must be signed in** to the profile whose mail you want. The bridge attaches to (or
  starts) the running Outlook instance.
- **Object Model Guard.** Some antivirus/trust configurations pop a one-time
  "a program is trying to access e-mail address information" prompt the first time COM reads
  sender/recipient fields. Accept it once (or update AV definitions / mark Outlook trusted) and
  it stops. Reading subject/body/attachments of your own mailbox normally does not prompt.
- **Large scans are bounded.** A full all-mailbox scan is capped at 500 results and 120 s by
  default (raise `timeout_seconds` in the plugin config if needed). Prefer narrow filters over
  brute-force scans.
- **PST archives appear as their own "mailbox"** in `outlook_status` / `outlook_list_mailboxes`.
- Match the summary's language to the user's language.
