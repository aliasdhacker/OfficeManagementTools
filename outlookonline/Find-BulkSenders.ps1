<#
  Discover bulk/unsolicited email senders in your mailbox.
  Searches for messages containing "unsubscribe" and groups by sender.
  Run with: .\Find-BulkSenders.ps1
  Limit results: .\Find-BulkSenders.ps1 -Top 50
#>
param(
    [int]$Top = 100
)

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$ctx = Get-MgContext
if (-not $ctx) {
    Connect-MgGraph -Scopes "Mail.Read" | Out-Null
}
Write-Host "Signed in as: $((Get-MgContext).Account)`n"

$SearchHeaders = @{ "ConsistencyLevel" = "eventual" }

# -------------------------------------------------------
# Phase 1: Find all messages with "unsubscribe" in body
# -------------------------------------------------------
Write-Host "=== Searching for messages containing 'unsubscribe' ===" -ForegroundColor Cyan
Write-Host "This may take a minute...`n"

# Strategy: Use inferenceClassification eq 'other' via $filter to get
# all messages Exchange classified as bulk/low-priority (Focused Inbox "Other" tab).
# This catches far more than keyword search since Exchange uses ML + headers.
# Then also run a keyword search for "unsubscribe" to catch any that slipped through.

$allMessages = @{}  # hashtable keyed by message id for dedup

# --- Method 1: All "Other" (non-Focused) messages via $filter ---
Write-Host "  Method 1: Fetching 'Other' (non-Focused) messages..." -ForegroundColor Yellow
$uri = "/v1.0/me/messages?`$filter=inferenceClassification eq 'other'&`$top=200&`$select=id,from,inferenceClassification&`$count=true"
$page = 0

do {
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $SearchHeaders -OutputType PSObject
        if ($resp.value) {
            foreach ($m in $resp.value) {
                $allMessages[$m.id] = $m
            }
        }
        $page++
        Write-Host "    Page $page - total unique: $($allMessages.Count)"
        $uri = $resp.'@odata.nextLink'
    } catch {
        Write-Host "    Error on page ${page}: $($_.Exception.Message)" -ForegroundColor Red
        Start-Sleep -Seconds 5
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $SearchHeaders -OutputType PSObject
            if ($resp.value) {
                foreach ($m in $resp.value) { $allMessages[$m.id] = $m }
            }
            $uri = $resp.'@odata.nextLink'
        } catch {
            Write-Host "    Retry failed, stopping pagination." -ForegroundColor Red
            $uri = $null
        }
    }
    # Small delay every 10 pages to avoid throttling
    if ($page % 10 -eq 0) { Start-Sleep -Milliseconds 500 }
} while ($uri)

$otherCount = $allMessages.Count
Write-Host "  Method 1 found: $otherCount messages" -ForegroundColor Cyan

# --- Method 2: Keyword search for "unsubscribe" ---
Write-Host "`n  Method 2: Searching for 'unsubscribe' keyword..." -ForegroundColor Yellow
$uri = "/v1.0/me/messages?`$search=`"unsubscribe`"&`$top=200&`$select=id,from,inferenceClassification&`$count=true"
$page = 0

do {
    try {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $SearchHeaders -OutputType PSObject
        if ($resp.value) {
            foreach ($m in $resp.value) {
                $allMessages[$m.id] = $m
            }
        }
        $page++
        Write-Host "    Page $page - total unique: $($allMessages.Count)"
        $uri = $resp.'@odata.nextLink'
    } catch {
        Write-Host "    Error on page ${page}: $($_.Exception.Message)" -ForegroundColor Red
        Start-Sleep -Seconds 5
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -Headers $SearchHeaders -OutputType PSObject
            if ($resp.value) {
                foreach ($m in $resp.value) { $allMessages[$m.id] = $m }
            }
            $uri = $resp.'@odata.nextLink'
        } catch {
            Write-Host "    Retry failed, stopping pagination." -ForegroundColor Red
            $uri = $null
        }
    }
    if ($page % 10 -eq 0) { Start-Sleep -Milliseconds 500 }
} while ($uri)

$unsubCount = $allMessages.Count - $otherCount
Write-Host "  Method 2 added: $unsubCount new messages" -ForegroundColor Cyan

# Convert hashtable to array
$messageList = $allMessages.Values

Write-Host "`nTotal unique bulk messages found: $($messageList.Count)" -ForegroundColor Yellow

if ($messageList.Count -eq 0) {
    Write-Host "No bulk messages found."
    return
}

# -------------------------------------------------------
# Phase 2: Group by sender address
# -------------------------------------------------------
Write-Host "`n=== Grouping by sender ===" -ForegroundColor Cyan

$senderGroups = $messageList |
    Where-Object { $_.from -and $_.from.emailAddress } |
    Group-Object { $_.from.emailAddress.address.ToLower() } |
    Sort-Object Count -Descending |
    Select-Object -First $Top

# -------------------------------------------------------
# Phase 3: Check which senders are already in Sort-BulkMail
# -------------------------------------------------------
$knownSenders = @()
$sortScript = Join-Path $PSScriptRoot "Sort-BulkMail.ps1"
if (Test-Path $sortScript) {
    $content = Get-Content $sortScript -Raw
    # Extract email addresses from the script
    $matches = [regex]::Matches($content, "'([^']+@[^']+)'")
    foreach ($m in $matches) {
        $knownSenders += $m.Groups[1].Value.ToLower()
    }
}

# -------------------------------------------------------
# Output
# -------------------------------------------------------
Write-Host "`n=== Top $Top Bulk Senders ===" -ForegroundColor Cyan
Write-Host ("{0}  {1}  {2}" -f "Count".PadLeft(6), "Status".PadRight(10), "Sender")
Write-Host ("{0}  {1}  {2}" -f "-----".PadLeft(6), "------".PadRight(10), "------")

$newSenders = @()

foreach ($g in $senderGroups) {
    $addr = $g.Name
    $cnt = $g.Count
    if ($knownSenders -contains $addr) {
        $status = "[KNOWN]"
        $color = "DarkGray"
    } else {
        $status = "[NEW]"
        $color = "White"
        $newSenders += [pscustomobject]@{ Sender = $addr; Count = $cnt }
    }
    Write-Host ("{0}  {1}  " -f $cnt.ToString().PadLeft(6), $status.PadRight(10)) -NoNewline
    Write-Host $addr -ForegroundColor $color
}

# -------------------------------------------------------
# Summary
# -------------------------------------------------------
$totalKnown = ($senderGroups | Where-Object { $knownSenders -contains $_.Name }).Count
$totalNew = $senderGroups.Count - $totalKnown

Write-Host "`n==============================" -ForegroundColor Green
Write-Host "Total senders shown: $($senderGroups.Count)"
Write-Host "  Already in Sort-BulkMail: $totalKnown" -ForegroundColor DarkGray
Write-Host "  NEW (not yet categorized): $totalNew" -ForegroundColor Yellow
Write-Host "=============================="

# -------------------------------------------------------
# Export new senders to file
# -------------------------------------------------------
$outFile = Join-Path $env:USERPROFILE "Desktop\bulk_senders_new.txt"
$lines = @()
foreach ($s in $newSenders) {
    $lines += $s.Sender
}
$lines | Out-File -FilePath $outFile -Encoding UTF8

Write-Host "`nNew senders exported to: $outFile (clean list, one address per line)"
Write-Host "Copy any addresses you want into Sort-BulkMail.ps1 categories."
