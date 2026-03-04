<#
  Sort bulk mail into folders: "Bulk Newsletters", "Bulk Ads", "Bulk Other"
  Creates the folders if they don't exist, then moves messages into them.
  Run with: .\Sort-BulkMail.ps1
  Dry-run only (count without moving): .\Sort-BulkMail.ps1 -DryRun
#>
param(
    [switch]$DryRun
)

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Ensure we have Mail.ReadWrite (not just Mail.Read from a previous session)
$ctx = Get-MgContext
if (-not $ctx -or ($ctx.Scopes -notcontains 'Mail.ReadWrite')) {
    if ($ctx) {
        Write-Host "Current session lacks Mail.ReadWrite - reconnecting..." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
    }
    Connect-MgGraph -Scopes "Mail.ReadWrite" | Out-Null
}
Write-Host "Signed in as: $((Get-MgContext).Account)`n"

$SearchHeaders = @{ "ConsistencyLevel" = "eventual" }

# -------------------------------------------------------
# Categorized sender lists
# -------------------------------------------------------
$categories = @{
    'Bulk Newsletters' = @(
        'newsletters@emails.discoveryplus.com'
        'packtdatapro1@substack.com'
        'membership@outside.plus'
        'email@washingtonpost.com'
        'newsletter@mail.coinbase.com'
        'support@email.masterclass.com'
        'info@perforce.com'
        'justin.reock@perforce.com'
        'newsletter@email.lifecare-news.com'
        'membership@gaiagps.com'
        'john@predictabledesigns.com'
    )
    'Bulk Ads' = @(
        # Retail / shopping
        'hft@em.harborfreight.com'
        'OmahaSteaks@mail.omahasteaks.com'
        'GameStop@emails.gamestop.com'
        'samsung@innovations.samsungusa.com'
        'sales@impactguns.com'
        'zenhelp@omcaviar.com'
        'fandango@movies.fandango.com'
        'contact@swingbyswing.com'
        'course@golffacility.com'
        # Financial promos
        'hello@capitaloneshopping.com'
        'capitalone@offer.capitalone.com'
        'capitalone@message.capitalone.com'
        'donotreply@cardmessage.capitalone.com'
        'Fidelity.Investments@mail.fidelity.com'
        'reply@email-firsthorizon.com'
        'help@selflender.com'
        'eDelivery@etradefrommorganstanley.com'
        # Drip / nurture campaigns
        'clientcare@vu.com'
        'education@vu.com'
        'team@vu.com'
        # Software promos
        'norton@secure.norton.com'
        'Autodesk@autodeskcommunications.com'
        'quicken@mail.quicken.com'
        'VSPVisionCareVCM@e.vsp.com'
    )
    'Bulk Other' = @(
        # Recruiting spam / job boards
        'dailyjobalert@postjobfree.com'
        'donotreply@upwork.com'
        'enrollment@wgu.edu'
        'team@email.rocketlawyer.com'
        # Misc bulk
        'sales@bitraser.com'
        'noreply@github.com'
        'noreply@medium.com'
        'noreply@quora.com'
        'support@nordvpn.com'
		'jobalerts-noreply@linkedin.com'
		'shop@email.sharperimage.com'
	   'updates@email.gunzonedeals.com'
	   'petco@e.petco.com'
	   'gamestop@em.gamestop.com'
	   'noreply@palmettostatearmory.com'
	   'concierge@imperiacaviar.com'
	   'postmaster@email.booksamillion.com'
	   'rewards@e.officedepot.com'
	   'mail@promotional.omahasteaks.com'
	   'notifications-noreply@linkedin.com'
	   'contact@packtpub.com'
	   'hft@em.harborfreight.com'
	   'chewy@woof.chewy.com'
	   'rescuesquad@pawboost.com'
	   'sales@bladeops.com'
	   'noreply@ethermine.org'
	   'news.1@e.pier1.com'
	   'basspro@basspronews.com'
	   'cabelas@emails.cabelas.com'
	   'noreply@starz.com'
	   'hello@shop.thriftbooks.com'
	   'updates-noreply@linkedin.com'
	   'jobs-listings@linkedin.com'
	   'goplay@email.golfnow.com'
	   'info@i.drop.com'
	   'customerservice@parachutehome.com'
	   'info@arindustriesllc.com'
	   'newsletters@em.walmart.com'
	   'mail@mail.adobe.com'
	   'no-reply@e.allegiant.com'
	   'helzberg@emails.helzberg.com'
	   'noreply@cbholsters.com'
	   'info@notifications.acorns.com'
	   'notification@service.tiktok.com'
	   'pier1@news.pier1.com'
	   'promo@promo.newegg.com'
	   'sharper_image@email.sharperimage.com'
	   'news@marketing.us.swann.com'
	   'loyalty@loyalty.ms.aa.com'
	   'emarketing@e.netgear.com'
	   'reply-to@e.digikey.com'
	   'info@email.purple.com'
	   'noreply@emails.creditonebank.com'
	   'producer@gunstuff.tv'
	   'reply@e.thenorthface.com'
	   'donotreply@e.jostens.com'
	   'targetnews@em.target.com'
	   'contact@berettausa.com'
	   'tiffany@tco.tiffany.com'
	   'harborfreight@e.harborfreight.com'
	   'no-reply@mail.ring.com'
	   'info@nanuk.com'
	   'yourbookstore@em.efollett.com'
	   'newsletter@gatdeals.com'
	   'noreply@sharperimageonline.com'
	   'news@communication.tiffany.com'
	   'magic@magic.wizards.com'
	   'store+x-plane.org@ccsend.com'
	   'hello@stix.golf'
	   'contact@mail.mejuri.com'
	   'josabank@shop.josbank.com'
	   'psb@deltateamtactical.com'
	   'petsmart@emails.petsmart.com'
	   'newsdigest@insideapple.apple.com'
	   'lifetouch@e.lifetouch.com'
	   'noreply@updates.freelancer.com'
	   'news@tcgplayer.com'
	   'shop@beauty.sephora.com'
	   'no-reply@neighborhoods.ring.com'
	   'mail@email.adobe.com'
	   'e.mariadb@mariadb.com'
	   'sales@tacticalsolutionsgroupllc.com'
	   'avery@em.avery.com'
	   'marketing@sw.solarwinds.com'
	   'globalgetaways@email-marriott.com'
	   'newsletter@patriotdailypress.com'
	   'info@wgu.edu'
	   'venmo@email.venmo.com'
	   'news@e.sunglasshut.com'
	   'store@x-plane.org'
	   'help@parachutehome.com'
	   'etrade@e.etradefinancial.com'
	   'email@a.grubhub.com'
	   'hello@colemanfurniture.com'
	   'paypal@emails.paypal.com'
	   'e.turner@f5.com'
	   'reply@engage.amd.com'
    )
}

# -------------------------------------------------------
# Helper: Get or create a mail folder
# -------------------------------------------------------
function Get-OrCreateFolder {
    param([string]$FolderName)

    # Check if folder exists
    $uri = "/v1.0/me/mailFolders?`$filter=displayName eq '$FolderName'"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    if ($resp.value -and $resp.value.Count -gt 0) {
        Write-Host "  Folder '$FolderName' exists (id: $($resp.value[0].id))"
        return $resp.value[0].id
    }

    # Create it
    $body = @{ displayName = $FolderName } | ConvertTo-Json
    try {
        $created = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/me/mailFolders' -Body $body -ContentType 'application/json' -OutputType PSObject
    } catch {
        Write-Host "  FAILED to create folder '$FolderName': $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Make sure you have Mail.ReadWrite permission. Try: Disconnect-MgGraph; then re-run." -ForegroundColor Red
        throw
    }
    if (-not $created.id) {
        throw "Folder creation returned no ID for '$FolderName'"
    }
    Write-Host "  Created folder '$FolderName' (id: $($created.id))" -ForegroundColor Green
    return $created.id
}

# -------------------------------------------------------
# Helper: Move messages to folder via batch API
# -------------------------------------------------------
function Move-MessagesBatch {
    param(
        [string[]]$MessageIds,
        [string]$FolderId,
        [string]$FolderName
    )

    $batchSize = 20
    $moved = 0
    $failed = 0

    for ($i = 0; $i -lt $MessageIds.Count; $i += $batchSize) {
        $end = [Math]::Min($i + $batchSize, $MessageIds.Count)
        $requests = @()
        for ($j = $i; $j -lt $end; $j++) {
            $requests += @{
                id     = "$j"
                method = "POST"
                url    = "/me/messages/$($MessageIds[$j])/move"
                headers = @{ "Content-Type" = "application/json" }
                body   = @{ destinationId = $FolderId }
            }
        }

        $batch = @{ requests = $requests } | ConvertTo-Json -Depth 5
        try {
            $resp = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/$batch' -Body $batch -ContentType 'application/json' -OutputType PSObject
            foreach ($r in $resp.responses) {
                if ($r.status -eq 201 -or $r.status -eq 200) {
                    $moved++
                } else {
                    $failed++
                }
            }
        } catch {
            Write-Host "    Batch error: $($_.Exception.Message)" -ForegroundColor Red
            $failed += ($end - $i)
        }

        $pct = [Math]::Round(($end / $MessageIds.Count) * 100)
        Write-Host "    [$FolderName] $end / $($MessageIds.Count) ($($pct)%) - Moved: $moved, Failed: $failed"

        if ($end -lt $MessageIds.Count) { Start-Sleep -Milliseconds 500 }
    }
    return @{ Moved = $moved; Failed = $failed }
}

# -------------------------------------------------------
# Phase 1: Scan all categories
# -------------------------------------------------------
Write-Host "=== Phase 1: Scanning for bulk mail ===" -ForegroundColor Cyan

$categoryResults = @{}

foreach ($cat in $categories.Keys) {
    Write-Host "`n--- $cat ---" -ForegroundColor Yellow
    $catIds = @()
    $catSummary = @()

    foreach ($sender in $categories[$cat]) {
        $kql = "from:$sender"
        $uri = "/v1.0/me/messages?`$count=true&`$search=`"$kql`"&`$top=200&`$select=id"
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
                Write-Host "  Error: $sender - $($_.Exception.Message)" -ForegroundColor Red
                $next = $null
            }
        } while ($next)

        if ($ids.Count -gt 0) {
            Write-Host "  $($ids.Count.ToString().PadLeft(5))  $sender"
            $catIds += $ids
            $catSummary += [pscustomobject]@{ Sender = $sender; Count = $ids.Count }
        }
    }

    $unique = $catIds | Select-Object -Unique
    $categoryResults[$cat] = @{
        Ids     = $unique
        Summary = $catSummary
    }
    Write-Host "  Total for ${cat}: $($unique.Count)" -ForegroundColor Cyan
}

# Grand total
$grandTotal = ($categoryResults.Values | ForEach-Object { $_.Ids.Count } | Measure-Object -Sum).Sum
Write-Host "`n======================================="
Write-Host "Grand total messages to sort: $grandTotal" -ForegroundColor Yellow
Write-Host "======================================="

if ($grandTotal -eq 0) {
    Write-Host "Nothing to sort. Exiting."
    return
}

# Summary table
foreach ($cat in $categoryResults.Keys) {
    $count = $categoryResults[$cat].Ids.Count
    Write-Host "`n${cat} ($count messages):" -ForegroundColor Yellow
    $categoryResults[$cat].Summary | Sort-Object Count -Descending | Format-Table -AutoSize
}

# -------------------------------------------------------
# Phase 2: Create folders and move (or dry-run)
# -------------------------------------------------------
if ($DryRun) {
    Write-Host "DRY RUN mode - no messages were moved." -ForegroundColor Green
    Write-Host "Run without -DryRun to sort messages into folders."
    return
}

$confirm = Read-Host "`nType YES to create folders and move messages"
if ($confirm -ne 'YES') {
    Write-Host "Aborted. No messages were moved."
    return
}

Write-Host "`n=== Phase 2: Creating folders and moving messages ===" -ForegroundColor Cyan

$totalMoved = 0
$totalFailed = 0

foreach ($cat in $categoryResults.Keys) {
    $ids = $categoryResults[$cat].Ids
    if ($ids.Count -eq 0) { continue }

    $count = $ids.Count
    Write-Host "`nProcessing: ${cat} ($count messages)" -ForegroundColor Yellow
    $folderId = Get-OrCreateFolder -FolderName $cat
    $result = Move-MessagesBatch -MessageIds $ids -FolderId $folderId -FolderName $cat
    $totalMoved += $result.Moved
    $totalFailed += $result.Failed
}

Write-Host "`n==============================" -ForegroundColor Green
Write-Host "Done! Moved: $totalMoved | Failed: $totalFailed"
Write-Host "Check your Outlook folders: Bulk Newsletters, Bulk Ads, Bulk Other"
Write-Host "=============================="
