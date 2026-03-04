<#
  Fetch emails by OWA ItemID using Graph $batch API, dump to timeline file
#>

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "Mail.Read" | Out-Null
}

# These are the URL-encoded ItemIDs from the OWA links
$encodedIds = @(
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEJAAABZbpUBipyRaDUF%2FjXG7uLAAFUovaKAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAHAXAIHAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAHvgdjbAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAH6brFPAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAIPAvH6AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAIq6DJ1AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAALnIc%2B%2BAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAL0wZgxAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAL0wZg1AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAFgAAABZbpUBipyRaDUF%2FjXG7uLAANJnDUEAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAMQxFhOAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEJAAABZbpUBipyRaDUF%2FjXG7uLAAMQxaIPAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAN0blPCAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAOXCR%2BgAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAPrMDnJAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAQDSgYsAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAQJ1VUbAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEJAAABZbpUBipyRaDUF%2FjXG7uLAAPvJvzoAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEJAAABZbpUBipyRaDUF%2FjXG7uLAARCvwiQAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAARJYdFNAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAATWmltEAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAUwlhCPAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAUwlhCRAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEJAAABZbpUBipyRaDUF%2FjXG7uLAAU5O6DCAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAXTO8o1AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAXTO8pBAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAXegSyJAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAX827I3AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAaLija5AAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAaLijbUAAA%3D'
    'AAMkAGRkYzMxNjFjLTk2MzktNGNjMi05MmZkLTVmODE2MTdhZDVjOABGAAAAAACzOa9zhJHRSIVec1zJYXguBwABZbpUBipyRaDUF%2FjXG7uLAAAAAAEMAAABZbpUBipyRaDUF%2FjXG7uLAAaL1BmNAAA%3D'
)

$total = $encodedIds.Count
Write-Host "Fetching $total emails via Graph batch API..."

# Graph $batch supports up to 20 requests per batch
$results = @()
$batchSize = 20

for ($start = 0; $start -lt $total; $start += $batchSize) {
    $end = [Math]::Min($start + $batchSize, $total)
    $requests = @()
    for ($j = $start; $j -lt $end; $j++) {
        # Keep the ID URL-encoded in the path so / and = don't break routing
        $requests += @{
            id     = "$j"
            method = "GET"
            url    = "/me/messages/$($encodedIds[$j])?`$select=id,subject,from,toRecipients,receivedDateTime,sentDateTime,bodyPreview"
        }
    }

    $batch = @{ requests = $requests } | ConvertTo-Json -Depth 5
    try {
        $resp = Invoke-MgGraphRequest -Method POST -Uri '/v1.0/$batch' -Body $batch -ContentType 'application/json' -OutputType PSObject
        foreach ($r in $resp.responses) {
            $idx = [int]$r.id
            if ($r.status -eq 200) {
                $msg = $r.body
                $results += [pscustomobject]@{
                    Index   = $idx + 1
                    Date    = if ($msg.receivedDateTime) { $msg.receivedDateTime } else { $msg.sentDateTime }
                    From    = if ($msg.from) { $msg.from.emailAddress.address } else { "unknown" }
                    Subject = $msg.subject
                    Preview = $msg.bodyPreview
                }
                Write-Host "  [$($idx+1)/$total] OK - $($msg.subject)"
            } else {
                Write-Host "  [$($idx+1)/$total] HTTP $($r.status): $($r.body.error.message)" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "Batch request failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Sort by date and output
$sorted = $results | Sort-Object { $_.Date -as [datetime] }

$outFile = Join-Path $env:USERPROFILE "Desktop\email_timeline.txt"
$lines = @()
foreach ($m in $sorted) {
    $lines += "===== $($m.Date) ====="
    $lines += "From: $($m.From)"
    $lines += "Subject: $($m.Subject)"
    $lines += "Preview: $($m.Preview)"
    $lines += ""
}
$lines | Out-File -FilePath $outFile -Encoding UTF8

Write-Host "`nFetched $($results.Count) of $total emails."
Write-Host "Timeline saved to: $outFile"
