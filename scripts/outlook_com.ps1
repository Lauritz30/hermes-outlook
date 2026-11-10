# outlook_com.ps1 — Outlook (classic) COM bridge for the hermes-outlook plugin.
#
# Reads a JSON request from -RequestPath, performs the requested action against the
# local Outlook MAPI session via COM, and writes a JSON result to -ResultPath.
# Result files are UTF-8 (no BOM) so the Python side can read them losslessly.
#
# Never run this by hand with your own request file unless you know what it does:
# it enumerates mailboxes and reads email. Actions: status | list_mailboxes |
# list_folders | search_emails | get_email.

param(
    [Parameter(Mandatory = $true)][string]$RequestPath,
    [Parameter(Mandatory = $true)][string]$ResultPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 1

function Write-Result {
    param($Obj)
    try {
        $json = $Obj | ConvertTo-Json -Depth 20 -Compress
    } catch {
        $json = @{ ok = $false; error = "Could not serialize result: $($_.Exception.Message)" } | ConvertTo-Json -Depth 5 -Compress
    }
    [System.IO.File]::WriteAllText($ResultPath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Contains-Text {
    param([string]$Haystack, [string]$Needle)
    if ([string]::IsNullOrEmpty($Needle)) { return $true }
    if ([string]::IsNullOrEmpty($Haystack)) { return $false }
    return $Haystack -match [regex]::Escape($Needle)
}

# Names of folders skipped during a full-mailbox scan (noise, not mail content).
# Primary detection is EntryID comparison against the store's default system folders
# (localization-independent); this name list is only a fallback for special folders
# that GetDefaultFolder does not resolve (RSS/QuickSteps/Conversation settings).
$script:NoiseFolders = @(
    'Deleted Items', 'Gelöschte Elemente',
    'Junk E-Mail', 'Junk Email', 'Junk-E-Mail',
    'Outbox', 'Postausgang',
    'Drafts', 'Entwürfe',
    'Sync Issues', 'Synchronisierungsprobleme',
    'Conflicts', 'Konflikte',
    'Local Failures', 'Server Failures',
    'RSS Feeds', 'RSS-Feeds',
    'Quick Step Settings', 'Einstellungen für QuickSteps',
    'Conversation Action Settings'
)

# OlDefaultFolders values for the system folders we skip when scanning.
# 3=DeletedItems 4=Outbox 16=Drafts 23=Junk 19=Conflicts 20=SyncIssues
# 21=LocalFailures 22=ServerFailures 25=RssFeeds  (Sent Items=5 is intentionally kept.)
$script:SkipFolderTypes = @(3, 4, 16, 23, 19, 20, 21, 22, 25)

function Get-SystemFolderIds {
    param($Store)
    $ids = @{}
    foreach ($t in $script:SkipFolderTypes) {
        try {
            $f = $Store.GetDefaultFolder([int]$t)
            if ($null -ne $f) { $ids[[string]$f.EntryID] = $true }
        } catch {}
    }
    return ,$ids
}

# --- Store selection ---------------------------------------------------------

function Get-Store {
    param([string]$Name)
    $all = @($script:ns.Stores)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $script:ns.DefaultStore }
    foreach ($s in $all) {
        if ($s.DisplayName -eq $Name) { return $s }
    }
    foreach ($s in $all) {
        if ($s.DisplayName -like "*$Name*") { return $s }
    }
    $idx = 0
    if ([int]::TryParse($Name, [ref]$idx)) {
        if ($idx -ge 0 -and $idx -lt $all.Count) { return $all[$idx] }
    }
    return $null
}

function Get-StoreKind {
    param($Store)
    $fp = $null
    try { $fp = $Store.FilePath } catch {}
    if ($fp -and $fp.ToLower().EndsWith('.pst')) { return 'pst' }
    $et = -1
    try { $et = [int]$Store.ExchangeStoreType } catch {}
    switch ($et) {
        0 { return 'primary_mailbox' }
        1 { return 'exchange_mailbox' }
        2 { return 'public_folder' }
        3 { return 'shared_mailbox' }
        4 { return 'primary_archive' }
        5 { return 'additional_archive' }
        default { return 'unknown' }
    }
}

# --- Folder walking ----------------------------------------------------------

function Get-FolderByPath {
    param($Store, [string]$Path)
    $parts = @($Path -split '/' | Where-Object { $_ -ne '' })
    $folder = $Store.GetRootFolder()
    foreach ($part in $parts) {
        $found = $null
        foreach ($f in $folder.Folders) {
            if ($f.Name -eq $part) { $found = $f; break }
        }
        if ($null -eq $found) { return $null }
        $folder = $found
    }
    return $folder
}

function Get-MailFolders {
    # Emits pscustomobject { Folder, Path } for every mail folder under $Folder.
    param($Folder, [string]$Path, [bool]$IncludeSub, [bool]$SkipNoise, [hashtable]$SkipIds, [int]$Depth)
    foreach ($f in $Folder.Folders) {
        $p = if ($Path) { "$Path/$($f.Name)" } else { $f.Name }
        $isMail = $true
        try { $isMail = ([int]$f.DefaultItemType -eq 0) } catch {}
        $isNoise = $false
        if ($SkipNoise) {
            $isNoise = ($null -ne $SkipIds -and $SkipIds.ContainsKey([string]$f.EntryID)) -or
                       ($script:NoiseFolders -contains $f.Name)
        }
        if ($isMail -and -not $isNoise) {
            Write-Output ([pscustomobject]@{ Folder = $f; Path = $p })
        }
        if ($IncludeSub -and $Depth -lt 12 -and -not $isNoise) {
            Get-MailFolders $f $p $IncludeSub $SkipNoise $SkipIds ($Depth + 1)
        }
    }
}

# --- Item extraction ---------------------------------------------------------

function Get-MailItemSummary {
    param($Item, [string]$StoreName, [string]$FolderPath)
    $received = ''
    try { $received = $Item.ReceivedTime.ToString('yyyy-MM-ddTHH:mm:ss') } catch {}
    $senderName = ''
    try { $senderName = [string]$Item.SenderName } catch {}
    $senderEmail = ''
    try { $senderEmail = [string]$Item.SenderEmailAddress } catch {}
    $to = ''
    try { $to = [string]$Item.To } catch {}
    $attCount = 0
    try { $attCount = [int]$Item.Attachments.Count } catch {}
    $size = 0
    try { $size = [int64]$Item.Size } catch {}
    $cats = @()
    try { $cats = @(($Item.Categories -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } catch {}
    $unread = $false
    try { $unread = [bool]$Item.Unread } catch {}

    [pscustomobject]@{
        entry_id          = [string]$Item.EntryID
        subject           = [string]$Item.Subject
        sender            = $senderName
        sender_email      = $senderEmail
        to                = $to
        received_time     = $received
        unread            = $unread
        has_attachments   = ($attCount -gt 0)
        attachment_count  = $attCount
        size              = $size
        categories        = $cats
        store             = $StoreName
        folder            = $FolderPath
    }
}

# --- Actions ----------------------------------------------------------------

function Invoke-Status {
    $wasRunning = $false
    try { $wasRunning = ($null -ne (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) } catch {}

    $storeList = @()
    foreach ($s in $script:ns.Stores) {
        $fp = $null
        try { $fp = $s.FilePath } catch {}
        $storeList += [pscustomobject]@{
            name     = [string]$s.DisplayName
            store_id = [string]$s.StoreID
            kind     = (Get-StoreKind $s)
            file_path = $fp
        }
    }

    $accountList = @()
    try {
        foreach ($a in $script:ns.Accounts) {
            $accountList += [pscustomobject]@{
                display_name = [string]$a.DisplayName
                smtp_address = [string]$a.SmtpAddress
                user_name    = [string]$a.UserName
            }
        }
    } catch {}

    [pscustomobject]@{
        ok               = $true
        action           = 'status'
        outlook_version  = [string]$script:outlook.Version
        outlook_running  = $wasRunning
        default_store    = [string]$script:ns.DefaultStore.DisplayName
        store_count      = $storeList.Count
        stores           = $storeList
        account_count    = $accountList.Count
        accounts         = $accountList
    }
}

function Invoke-ListMailboxes {
    $storeList = @()
    foreach ($s in $script:ns.Stores) {
        $fp = $null
        try { $fp = $s.FilePath } catch {}
        $folderCount = 0
        try { $folderCount = @($s.GetRootFolder().Folders).Count } catch {}
        $storeList += [pscustomobject]@{
            name         = [string]$s.DisplayName
            store_id     = [string]$s.StoreID
            kind         = (Get-StoreKind $s)
            file_path    = $fp
            top_folder_count = $folderCount
        }
    }
    [pscustomobject]@{
        ok = $true
        action = 'list_mailboxes'
        mailbox_count = $storeList.Count
        mailboxes = $storeList
    }
}

function Invoke-ListFolders {
    param($P)
    $storeName = $P.store
    $recursive = $false
    if ($null -ne $P.recursive) { $recursive = [bool]$P.recursive }

    $store = Get-Store $storeName
    if ($null -eq $store) { throw "Mailbox '$storeName' not found." }

    $root = $store.GetRootFolder()
    $folders = @()
    foreach ($f in $root.Folders) {
        $itemCount = 0
        try { $itemCount = [int]$f.Items.Count } catch {}
        $folders += [pscustomobject]@{
            name        = [string]$f.Name
            path        = [string]$f.Name
            folder_id   = [string]$f.EntryID
            item_count  = $itemCount
        }
    }
    if ($recursive) {
        $sub = @(Get-MailFolders $root '' $true $false @{} 0)
        # only keep deeper paths (depth 2+), avoid duplicates with the top list
        foreach ($fo in $sub) {
            if ($fo.Path -match '/') {
                $itemCount = 0
                try { $itemCount = [int]$fo.Folder.Items.Count } catch {}
                $folders += [pscustomobject]@{
                    name       = [string]$fo.Folder.Name
                    path       = [string]$fo.Path
                    folder_id  = [string]$fo.Folder.EntryID
                    item_count = $itemCount
                }
            }
        }
    }

    [pscustomobject]@{
        ok = $true
        action = 'list_folders'
        store = [string]$store.DisplayName
        folder_count = $folders.Count
        folders = $folders
    }
}

function Invoke-Search {
    param($P)

    $storeName = $P.store
    $folderPath = $P.folder
    $subject = $P.subject
    $sender = $P.sender
    $recipient = $P.recipient
    $body = $P.body
    $unreadOnly = $false
    if ($null -ne $P.unread_only) { $unreadOnly = [bool]$P.unread_only }
    $hasAttach = $false
    if ($null -ne $P.has_attachments -and $P.has_attachments) { $hasAttach = $true }
    $includeSub = $true
    if ($null -ne $P.include_subfolders) { $includeSub = [bool]$P.include_subfolders }
    $skipNoise = $true
    if ($null -ne $P.skip_system_folders) { $skipNoise = [bool]$P.skip_system_folders }
    $max = 50
    if ($null -ne $P.max_results) { $max = [int]$P.max_results }
    if ($max -lt 1) { $max = 1 }
    if ($max -gt 500) { $max = 500 }

    $after = $null
    if ($null -ne $P.received_after -and $P.received_after) {
        $after = [datetime]::Parse([string]$P.received_after, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    $before = $null
    if ($null -ne $P.received_before -and $P.received_before) {
        $before = [datetime]::Parse([string]$P.received_before, [System.Globalization.CultureInfo]::InvariantCulture)
    }

    # Build an Outlook Restrict filter for the parts the MAPI layer can do fast.
    $restrictParts = @()
    if ($null -ne $after) { $restrictParts += "[ReceivedTime] >= '$($after.ToString('MM/dd/yyyy HH:mm'))'" }
    if ($null -ne $before) { $restrictParts += "[ReceivedTime] <= '$($before.ToString('MM/dd/yyyy HH:mm'))'" }
    if ($unreadOnly) { $restrictParts += "[Unread] = True" }
    $restrict = $null
    if ($restrictParts.Count -gt 0) { $restrict = '(' + ($restrictParts -join ' AND ') + ')' }

    # Resolve target stores.
    $stores = @()
    if ($storeName) {
        $s = Get-Store $storeName
        if ($null -eq $s) { throw "Mailbox '$storeName' not found." }
        $stores = @($s)
    } else {
        $stores = @($script:ns.Stores)
    }

    $results = @()
    $truncated = $false

    foreach ($store in $stores) {
        $sName = [string]$store.DisplayName
        $targets = @()
        if ($folderPath) {
            $f = Get-FolderByPath $store $folderPath
            if ($null -ne $f) { $targets = @([pscustomobject]@{ Folder = $f; Path = $folderPath }) }
        } else {
            $root = $store.GetRootFolder()
            $skipIds = if ($skipNoise) { Get-SystemFolderIds $store } else { @{} }
            $targets = @(Get-MailFolders $root '' $includeSub $skipNoise $skipIds 0)
        }

        foreach ($fo in $targets) {
            if ($results.Count -ge $max) { $truncated = $true; break }
            $items = $fo.Folder.Items
            if ($null -ne $restrict) {
                try { $items = $items.Restrict($restrict) } catch { }
            }
            try { $items.Sort('[ReceivedTime]', $true) } catch {}
            foreach ($item in $items) {
                if ($results.Count -ge $max) { $truncated = $true; break }
                try {
                    if ([int]$item.Class -ne 43) { continue }  # 43 = olMail
                } catch { continue }

                $summary = Get-MailItemSummary $item $sName $fo.Path

                if (-not (Contains-Text $summary.subject $subject)) { continue }
                if (-not (Contains-Text ("$($summary.sender) $($summary.sender_email)") $sender)) { continue }
                if (-not (Contains-Text $summary.to $recipient)) { continue }
                if ($hasAttach -and -not $summary.has_attachments) { continue }
                if ($body) {
                    $bodyText = ''
                    try { $bodyText = [string]$item.Body } catch {}
                    if (-not (Contains-Text $bodyText $body)) { continue }
                }
                $results += $summary
            }
        }
        if ($results.Count -ge $max) { $truncated = $true; break }
    }

    [pscustomobject]@{
        ok = $true
        action = 'search_emails'
        result_count = $results.Count
        truncated = $truncated
        items = $results
    }
}

function Invoke-GetEmail {
    param($P)
    $entryId = [string]$P.entry_id
    $storeName = $P.store
    $maxBody = 20000
    if ($null -ne $P.max_body_chars) { $maxBody = [int]$P.max_body_chars }
    if ($maxBody -lt 0) { $maxBody = 0 }
    $includeHtml = $false
    if ($null -ne $P.include_html) { $includeHtml = [bool]$P.include_html }

    if ([string]::IsNullOrWhiteSpace($entryId)) { throw "entry_id is required." }

    $storeId = $null
    if ($storeName) {
        $s = Get-Store $storeName
        if ($null -eq $s) { throw "Mailbox '$storeName' not found." }
        $storeId = $s.StoreID
    }

    $item = $null
    if ($null -ne $storeId) {
        $item = $script:ns.GetItemFromID($entryId, $storeId)
    } else {
        $item = $script:ns.GetItemFromID($entryId)
    }
    if ($null -eq $item) { throw "No email found for that entry_id." }

    $to = @(); $cc = @(); $bcc = @()
    try {
        foreach ($r in $item.Recipients) {
            $addr = ''; try { $addr = [string]$r.Address } catch {}
            $nm = ''; try { $nm = [string]$r.AddressEntry.Name } catch {}
            $type = 0; try { $type = [int]$r.Type } catch {}
            $rec = [pscustomobject]@{ name = $nm; address = $addr }
            if ($type -eq 1) { $to += $rec }
            elseif ($type -eq 2) { $cc += $rec }
            elseif ($type -eq 3) { $bcc += $rec }
        }
    } catch {}

    $atts = @()
    try {
        foreach ($a in $item.Attachments) {
            $atts += [pscustomobject]@{
                file_name    = [string]$a.FileName
                display_name = [string]$a.DisplayName
                size         = [int64]$a.Size
            }
        }
    } catch {}

    $bodyText = ''
    try { $bodyText = [string]$item.Body } catch {}
    $bodyTruncated = $false
    if ($bodyText.Length -gt $maxBody) {
        $bodyText = $bodyText.Substring(0, $maxBody)
        $bodyTruncated = $true
    }

    $html = $null
    if ($includeHtml) {
        try { $html = [string]$item.HTMLBody } catch {}
    }

    $sent = ''
    try { $sent = $item.SentOn.ToString('yyyy-MM-ddTHH:mm:ss') } catch {}
    $received = ''
    try { $received = $item.ReceivedTime.ToString('yyyy-MM-ddTHH:mm:ss') } catch {}

    $categories = @()
    try { $categories = @(($item.Categories -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } catch {}

    $result = [pscustomobject]@{
        ok = $true
        action = 'get_email'
        email = [pscustomobject]@{
            entry_id       = [string]$item.EntryID
            subject        = [string]$item.Subject
            sender_name    = ([string]$item.SenderName)
            sender_email   = ([string]$item.SenderEmailAddress)
            to             = $to
            cc             = $cc
            bcc            = $bcc
            sent_time      = $sent
            received_time  = $received
            body           = $bodyText
            body_truncated = $bodyTruncated
            body_length    = $bodyText.Length
            html           = $html
            attachments    = $atts
            attachment_count = $atts.Count
            categories     = $categories
            importance     = [int]$item.Importance
            unread         = [bool]$item.Unread
            size           = [int64]$item.Size
        }
    }
    return $result
}

function Invoke-SendEmail {
    param($P)
    $to = [string]$P.to
    if ([string]::IsNullOrWhiteSpace($to)) { throw "to is required." }

    $send = $false
    if ($null -ne $P.send) { $send = [bool]$P.send }

    # Optional: pick which of the user's email accounts sends the message.
    $acct = $null
    if ($null -ne $P.account -and $P.account) {
        $want = [string]$P.account
        foreach ($a in $script:ns.Accounts) {
            $smtp = ''; try { $smtp = [string]$a.SmtpAddress } catch {}
            $disp = ''; try { $disp = [string]$a.DisplayName } catch {}
            if ($smtp -eq $want -or $disp -eq $want -or $smtp -like "*$want*" -or $disp -like "*$want*") {
                $acct = $a
                break
            }
        }
        if ($null -eq $acct) {
            $avail = (@($script:ns.Accounts | ForEach-Object { [string]$_.SmtpAddress }) -join ', ')
            throw "Sender account '$want' not found. Available: $avail"
        }
    }

    $mail = $script:outlook.CreateItem(0)   # 0 = olMailItem
    $mail.Subject = [string]$P.subject
    $mail.To = $to
    if ($null -ne $P.cc -and $P.cc) { $mail.CC = [string]$P.cc }
    if ($null -ne $P.bcc -and $P.bcc) { $mail.BCC = [string]$P.bcc }
    if ($null -ne $acct) { $mail.SendUsingAccount = $acct }

    $body = [string]$P.body
    $html = $false
    if ($null -ne $P.body_is_html) { $html = [bool]$P.body_is_html }
    if ($html) { $mail.HTMLBody = $body } else { $mail.Body = $body }

    if ($null -ne $P.importance) {
        $imp = [int]$P.importance
        if ($imp -ge 0 -and $imp -le 2) { $mail.Importance = $imp }
    }

    $attPaths = @()
    if ($null -ne $P.attachments) { $attPaths = @($P.attachments) }
    foreach ($p in $attPaths) {
        $ap = [string]$p
        if (-not (Test-Path -LiteralPath $ap -PathType Leaf)) { throw "Attachment not found: $ap" }
        $null = $mail.Attachments.Add($ap)
    }

    $subj = [string]$mail.Subject
    $acctAddr = ''
    if ($null -ne $acct) { try { $acctAddr = [string]$acct.SmtpAddress } catch {} }
    if ($send) {
        $mail.Send()
        [pscustomobject]@{
            ok = $true; action = 'send_email'; sent = $true; mode = 'send'
            subject = $subj; to = $to; account = $acctAddr
        }
    } else {
        $mail.Display()
        [pscustomobject]@{
            ok = $true; action = 'send_email'; sent = $false; mode = 'draft'
            subject = $subj; to = $to; account = $acctAddr
        }
    }
}

function Invoke-ReplyEmail {
    param($P)
    $entryId = [string]$P.entry_id
    if ([string]::IsNullOrWhiteSpace($entryId)) { throw "entry_id is required." }

    $storeName = $P.store
    $storeId = $null
    if ($storeName) {
        $s = Get-Store $storeName
        if ($null -eq $s) { throw "Mailbox '$storeName' not found." }
        $storeId = $s.StoreID
    }

    $item = $null
    if ($null -ne $storeId) { $item = $script:ns.GetItemFromID($entryId, $storeId) }
    else { $item = $script:ns.GetItemFromID($entryId) }
    if ($null -eq $item) { throw "No email found for that entry_id." }

    $replyAll = $false
    if ($null -ne $P.reply_all) { $replyAll = [bool]$P.reply_all }
    $send = $false
    if ($null -ne $P.send) { $send = [bool]$P.send }

    # Reply()/ReplyAll() pre-fill To/Subject/quote AND carry the correct From identity,
    # including shared mailboxes the user can answer but not send from directly.
    $reply = if ($replyAll) { $item.ReplyAll() } else { $item.Reply() }

    $body = [string]$P.body
    if ($body) {
        $reply.Body = $body + "`r`n`r`n" + $reply.Body
    }

    $subj = [string]$reply.Subject
    $sendAcct = ''
    try { if ($reply.SendUsingAccount) { $sendAcct = [string]$reply.SendUsingAccount.SmtpAddress } } catch {}
    $onBehalf = ''
    try { $onBehalf = [string]$reply.SentOnBehalfOfName } catch {}

    if ($send) {
        $reply.Send()
        [pscustomobject]@{
            ok = $true; action = 'reply_email'; sent = $true; mode = 'send'
            subject = $subj; reply_all = $replyAll; account = $sendAcct; on_behalf_of = $onBehalf
        }
    } else {
        $reply.Display()
        [pscustomobject]@{
            ok = $true; action = 'reply_email'; sent = $false; mode = 'draft'
            subject = $subj; reply_all = $replyAll; account = $sendAcct; on_behalf_of = $onBehalf
        }
    }
}

# --- Main -------------------------------------------------------------------

try {
    $req = Get-Content -Raw -Path $RequestPath | ConvertFrom-Json
    $action = [string]$req.action
    $p = $req.params
    if ($null -eq $p) { $p = @{} }

    $script:outlook = New-Object -ComObject Outlook.Application
    $script:ns = $script:outlook.GetNamespace('MAPI')

    switch ($action) {
        'status'          { $result = Invoke-Status }
        'list_mailboxes'  { $result = Invoke-ListMailboxes }
        'list_folders'    { $result = Invoke-ListFolders $p }
        'search_emails'   { $result = Invoke-Search $p }
        'get_email'       { $result = Invoke-GetEmail $p }
        'send_email'      { $result = Invoke-SendEmail $p }
        'reply_email'     { $result = Invoke-ReplyEmail $p }
        default           { throw "Unknown action: $action" }
    }
    Write-Result $result
} catch {
    Write-Result @{ ok = $false; error = $_.Exception.Message }
}
