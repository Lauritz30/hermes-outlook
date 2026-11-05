# hermes-outlook

Read email from **classic Outlook (desktop)** for Hermes Agent — via a Windows PowerShell COM
bridge. Enumerate every mailbox the current Outlook profile can see (email accounts, shared
mailboxes, PST archives), list folders, scan/search messages across one or all mailboxes, and
read full emails. **No Graph API, no M365 admin consent.**

Windows-only. Talks to the local Outlook MAPI session through `Outlook.Application` COM.

## Tools

| Tool | What it does |
|---|---|
| `outlook_status` | Is Outlook reachable? Version, default mailbox, full store list. |
| `outlook_list_mailboxes` | Every mailbox/store with kind (mailbox / shared / PST / archive) and folder count. |
| `outlook_list_folders` | Folders in a mailbox (optionally recursive) with item counts. |
| `outlook_search_emails` | Search across one mailbox or **all** mailboxes; filters for subject/sender/recipient/body/unread/attachments/date range. |
| `outlook_get_email` | Full email by EntryID: recipients, body, attachments, categories, flags. |

## How it works

The plugin's Python handlers shell out to `powershell.exe`, which runs
[`scripts/outlook_com.ps1`](scripts/outlook_com.ps1). A JSON request is written to a temp file,
the script does the Outlook COM work, and writes a JSON result file back — file I/O sidesteps
PowerShell's OEM-codepage redirected-stdout problem entirely.

```
Hermes tool  →  Python handler  →  powershell.exe -File outlook_com.ps1  →  Outlook COM (MAPI)
```

## Install

```bash
# local (dev): drop the repo into the plugins dir and enable it
cp -r . "$HERMES_HOME/plugins/hermes-outlook"
hermes plugins enable hermes-outlook

# or from git
hermes plugins install Lauritz30/hermes-outlook --enable
```

Verify:

```bash
hermes plugins list          # hermes-outlook … enabled
hermes plugins doctor hermes-outlook
```

Then ask Hermes: *"scan all my mailboxes for unread mail from the last 7 days"*.

## Configuration

`config_schema` (see `plugin.yaml`):

- `powershell_path` — default `powershell.exe` (5.1). `pwsh` also works.
- `timeout_seconds` — default 120. Raise it for huge all-mailbox scans.

## Requirements & caveats

- **Classic Outlook** (the Office desktop app) must be installed and signed in. The "new"
  Outlook does not expose the same COM automation object.
- **Windows** only (COM + PowerShell).
- **Object Model Guard:** the first time COM reads sender/recipient fields, some trust/AV
  configurations show a one-time "a program is trying to access e-mail address information"
  prompt. Accept it once and it stops.
- Scans are bounded (500 results / 120 s by default) to avoid hanging on huge mailboxes.

## License

MIT.
