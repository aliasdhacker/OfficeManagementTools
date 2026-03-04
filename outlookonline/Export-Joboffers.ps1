<#
Export job-offer related messages since 2019-01-01 for mailbox me@acarr.org
Outputs: Desktop\job_offers_since_2019.csv
#>

# --------------------------
# Config
# --------------------------
$MailboxUpn = "me@acarr.org"
$SinceLocal = Get-Date "2019-01-01"
$SinceUtc   = ([DateTimeOffset]$SinceLocal.ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ssZ")
$OutPath    = Join-Path $env:USERPROFILE "Desktop\job_offers_since_2019.csv"

# AQS queries  (NO leading parentheses; use simple OR chains; keep phrases quoted)
$AqsOffers = 'offer OR offer letter OR employment agreement OR employment contract OR compensation summary OR comp package OR total rewards OR compensation plan OR base salary OR sign-on OR signing bonus OR equity OR RSU OR stock options'
$AqsAccept = 'from:me I accept OR from:me accepted OR from:me signed offer OR from:me offer acceptance'
$AqsEsign  = 'docusign OR DocuSign System OR adobesign OR Adobe Sign OR echosign OR HelloSign OR Dropbox Sign OR envelope completed OR completed your document OR has been completed'

# Graph header required for $search
$SearchHeaders = @{ "ConsistencyLevel" = "eventual" }

# --------------------------
# Minimal Graph connect
# --------------------------
if (-not (Get-Module Microsoft.Graph.Authentication -ListAvailable)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "Mail.Read" | Out-Null
}
Write-Host "Signed in as: $((Get-MgContext).Account)"

# --------------------------
# Helpers
# --------------------------
function New-SearchUri {
    param(
        [Parameter(Mandatory=$true)][string]$Aqs,
        [Parameter(Mandatory=$true)][ValidateSet('received','sent')]$DateField
    )
    $select = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,hasAttachments,webLink,bodyPreview,internetMessageId,conversationId"
    # Date constraint goes inside KQL because $filter can't be combined with $search
    $kql = "($Aqs) AND $DateField>=2019-01-01"
    "/v1.0/me/messages?`$count=true&`$search=""$kql""&`$top=50&`$select=$select"
}

function Invoke-GraphedQuery {
    param([Parameter(Mandatory=$true)][string]$Uri)
    $all = @()
    $next = $Uri
    do {
        try {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $SearchHeaders -OutputType PSObject
        } catch {
            $retryAfter = 5
            if ($_.Exception.Response -and $_.Exception.Response.Headers.'Retry-After') {
                [int]::TryParse($_.Exception.Response.Headers.'Retry-After', [ref]$retryAfter) | Out-Null
            }
            Write-Warning "Throttled or transient failure. Sleeping $retryAfter sec then retrying..."
            Start-Sleep -Seconds $retryAfter
            $resp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $SearchHeaders -OutputType PSObject
        }
        if ($resp.value) { $all += $resp.value }
        $next = $resp.'@odata.nextLink'
    } while ($next)
    ,$all
}

# --------------------------
# Run the three searches
# --------------------------
$Results = @()

# A) Offer / agreements (received since date)
$uriA = New-SearchUri -Aqs $AqsOffers -DateField received
Write-Host "`nQuery A (Offers/Agreements) ..."
$itemsA = Invoke-GraphedQuery -Uri $uriA
$Results += $itemsA
Write-Host ("  + {0} messages" -f $itemsA.Count)

# B) Your acceptances (SENT by you since date)
$uriB = New-SearchUri -Aqs $AqsAccept -DateField sent
Write-Host "`nQuery B (Your Acceptances) ..."
$itemsB = Invoke-GraphedQuery -Uri $uriB
$Results += $itemsB
Write-Host ("  + {0} messages" -f $itemsB.Count)

# C) E-sign completions (received since date)
$uriC = New-SearchUri -Aqs $AqsEsign -DateField received
Write-Host "`nQuery C (E-sign completions) ..."
$itemsC = Invoke-GraphedQuery -Uri $uriC
$Results += $itemsC
Write-Host ("  + {0} messages" -f $itemsC.Count)

# Deduplicate by InternetMessageId
$Dedup = $Results | Group-Object internetMessageId | ForEach-Object { $_.Group[0] }
Write-Host ("`nTotal unique messages: {0}" -f $Dedup.Count)

# --------------------------
# Expand attachment names
# --------------------------
$Rows = @()

foreach ($m in $Dedup) {
    $attachmentNames = @()
    if ($m.hasAttachments -eq $true) {
        $attsUri = "/v1.0/me/messages/$($m.id)/attachments?`$select=name,size,contentType,isInline"
        try {
            $attsAll = @()
            $next = $attsUri
            do {
                $attResp = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $SearchHeaders -OutputType PSObject
                if ($attResp.value) { $attsAll += $attResp.value }
                $next = $attResp.'@odata.nextLink'
            } while ($next)
            $attachmentNames = ($attsAll | ForEach-Object { $_.name }) -join "; "
        } catch {
            $attachmentNames = "[error retrieving attachments: $($_.Exception.Message)]"
        }
    }

    $Rows += [pscustomobject]@{
        Subject            = $m.subject
        From               = if ($m.from) { $m.from.emailAddress.address } else { "" }
        ReceivedDateTime   = $m.receivedDateTime
        SentDateTime       = $m.sentDateTime
        To                 = ($m.toRecipients | ForEach-Object { $_.emailAddress.address }) -join "; "
        Cc                 = ($m.ccRecipients | ForEach-Object { $_.emailAddress.address }) -join "; "
        HasAttachments     = $m.hasAttachments
        AttachmentNames    = $attachmentNames
        BodyPreview        = $m.bodyPreview
        WebLink            = $m.webLink
        InternetMessageId  = $m.internetMessageId
        ConversationId     = $m.conversationId
        QueryMatched       = $true
    }
}

# Sort & export
$Rows | Sort-Object { $_.ReceivedDateTime -as [datetime] } |
    Export-Csv -Path $OutPath -Encoding UTF8 -NoTypeInformation

Write-Host "`nExport complete: $OutPath"