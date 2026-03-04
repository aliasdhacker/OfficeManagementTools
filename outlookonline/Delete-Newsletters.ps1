<#
  Bulk delete newsletters from Outlook mailbox via Microsoft Graph.
  Messages are moved to Deleted Items (recoverable), NOT permanently deleted.
  Run with: .\Delete-Newsletters.ps1
  Dry-run only (list without deleting): .\Delete-Newsletters.ps1 -DryRun
#>
param(
    [switch]$DryRun
)

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "Mail.ReadWrite" | Out-Null
}
Write-Host "Signed in as: $((Get-MgContext).Account)`n"

$SearchHeaders = @{ "ConsistencyLevel" = "eventual" }

# -------------------------------------------------------
# Newsletter sender domains / patterns to target
# Add or remove entries as needed
# -------------------------------------------------------
$newsletterSenders = @(
    # --- Newsletters / Content ---
    'newsletters@emails.discoveryplus.com'
    'packtdatapro1@substack.com'
    'membership@outside.plus'
    'email@washingtonpost.com'
    'newsletter@mail.coinbase.com'
    'fandango@movies.fandango.com'
    'support@email.masterclass.com'
    'info@perforce.com'
    'justin.reock@perforce.com'

    # --- Ads / Promos / Marketing ---
    'hft@em.harborfreight.com'
    'OmahaSteaks@mail.omahasteaks.com'
    'GameStop@emails.gamestop.com'
    'samsung@innovations.samsungusa.com'
    'sales@impactguns.com'
    'zenhelp@omcaviar.com'
    'norton@secure.norton.com'
    'Autodesk@autodeskcommunications.com'
    'Fidelity.Investments@mail.fidelity.com'
    'reply@email-firsthorizon.com'
    'help@selflender.com'

    # --- Capital One ads/offers ---
    'hello@capitaloneshopping.com'
    'capitalone@offer.capitalone.com'
    'capitalone@message.capitalone.com'
    'donotreply@cardmessage.capitalone.com'

    # --- Veterans United drip campaigns ---
    'clientcare@vu.com'
    'education@vu.com'
    'team@vu.com'

    # --- Recruiting spam / job boards ---
    'dailyjobalert@postjobfree.com'
    'donotreply@upwork.com'
    'enrollment@wgu.edu'
    'team@email.rocketlawyer.com'

    # --- Misc marketing ---
    'newsletter@email.lifecare-news.com'
    'contact@swingbyswing.com'
    'course@golffacility.com'
    'membership@gaiagps.com'
    'john@predictabledesigns.com'
    'sales@bitraser.com'
    'quicken@mail.quicken.com'
    'VSPVisionCareVCM@e.vsp.com'
    'eDelivery@etradefrommorganstanley.com'
)

# -------------------------------------------------------
# Phase 1: Discover messages per sender
# -------------------------------------------------------
Write-Host "=== Phase 1: Scanning for newsletters ===" -ForegroundColor Cyan
$allIds = @()
$senderSummary = @()

foreach ($sender in $newsletterSenders) {
    $kql = "from:$sender"
    $uri = "/v1.0/me/messages?`$count=true&`$search=""$kql""&`$top=200&`$select=id"
    $ids = @()
    $next = $uri
    do {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $SearchHeaders -OutputType PSObject
            if ($resp.value) {
                $ids += $resp.value | ForEach-Object { $_.id }
            }
            $next = $resp.'@odata.nextLink'
        } catch {
            Write-Host "  Error searching $sender : $($_.Exception.Message)" -ForegroundColor Red
            $next = $null
        }
    } while ($next)

    if ($ids.Count -gt 0) {
        Write-Host "  $($ids.Count.ToString().PadLeft(5))  $sender"
        $allIds += $ids
        $senderSummary += [pscustomobject]@{ Sender = $sender; Count = $ids.Count }
    }
}

# Deduplicate (shouldn't be needed, but just in case)
$uniqueIds = $allIds | Select-Object -Unique

Write-Host "`n---------------------------------------"
Write-Host "Total newsletters found: $($uniqueIds.Count)" -ForegroundColor Yellow
Write-Host "---------------------------------------"

if ($uniqueIds.Count -eq 0) {
    Write-Host "Nothing to delete. Exiting."
    return
}

# Show summary table
Write-Host "`nBreakdown by sender:"
$senderSummary | Sort-Object Count -Descending | Format-Table -AutoSize

# -------------------------------------------------------
# Phase 2: Delete (or dry-run)
# -------------------------------------------------------
if ($DryRun) {
    Write-Host "DRY RUN mode - no messages were deleted." -ForegroundColor Green
    Write-Host "Run without -DryRun to actually delete."
    return
}

Write-Host "These messages will be moved to Deleted Items (recoverable)." -ForegroundColor Yellow
$confirm = Read-Host "Type YES to proceed with deletion"
if ($confirm -ne 'YES') {
    Write-Host "Aborted. No messages were deleted."
    return
}

Write-Host "`n=== Phase 2: Deleting $($uniqueIds.Count) messages ===" -ForegroundColor Cyan

# Use batch API for speed (20 per batch)
$batchSize = 20
$deleted = 0
$failed = 0

for ($i = 0; $i -lt $uniqueIds.Count; $i += $batchSize) {
    $end = [Math]::Min($i + $batchSize, $uniqueIds.Count)
    $requests = @()
    for ($j = $i; $j -lt $end; $j++) {
        $requests += @{
            id     = "$j"
            method = "DELETE"
            url    = "/me/messages/$($uniqueIds[$j])"
        }
    }

    $batch = @{ requests = $requests } | ConvertTo-Json -Depth 5
    try {
        $resp = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/$batch' -Body $batch -ContentType 'application/json' -OutputType PSObject
        foreach ($r in $resp.responses) {
            if ($r.status -eq 204) {
                $deleted++
            } else {
                $failed++
            }
        }
    } catch {
        Write-Host "  Batch error: $($_.Exception.Message)" -ForegroundColor Red
        $failed += ($end - $i)
    }

    # Progress
    $pct = [Math]::Round(($end / $uniqueIds.Count) * 100)
    Write-Host "  Progress: $end / $($uniqueIds.Count) ($pct%) - Deleted: $deleted, Failed: $failed"

    # Small delay to avoid throttling
    if ($end -lt $uniqueIds.Count) { Start-Sleep -Milliseconds 500 }
}

Write-Host "`n==============================" -ForegroundColor Green
Write-Host "Done! Deleted: $deleted | Failed: $failed"
Write-Host "Messages are in your Deleted Items folder (recoverable)."
Write-Host "=============================="
