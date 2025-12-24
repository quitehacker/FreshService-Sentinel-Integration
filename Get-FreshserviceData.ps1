<#
.SYNOPSIS
    Fetches tickets from Freshservice, enriches them with Agent and Group names, and exports to JSON/CSV.

.DESCRIPTION
    This script performs the following:
    1. Fetches all Agents and Groups to build lookup tables (ID -> Name).
    2. Fetches Tickets (with pagination and optional time filter).
    3. Enriches each ticket by replacing IDs with readable Names.
    4. Exports the enriched data to a local file.

.PARAMETER FreshserviceDomain
    The Freshservice domain, e.g., "yourcompany.freshservice.com".

.PARAMETER FreshserviceApiKey
    The API Key for Freshservice authentication.

.PARAMETER LookbackMinutes
    Optional. Fetch tickets updated in the last N minutes. Use 0 for ALL tickets. Default is 0.

.PARAMETER OutputPath
    Full path to save the output file (should end in .json or .csv).

.EXAMPLE
    .\Get-FreshserviceData.ps1 -FreshserviceDomain "acme.freshservice.com" -FreshserviceApiKey "szx..." -OutputPath ".\tickets.json"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FreshserviceDomain,

    [Parameter(Mandatory = $true)]
    [string]$FreshserviceApiKey,

    [int]$LookbackMinutes = 0,

    [string]$OutputPath = ".\FreshserviceTickets.json"
)

# -----------------------------------------------------------------------------
# Helper: Get-FreshserviceResource
# -----------------------------------------------------------------------------
function Get-FreshserviceResource {
    param(
        [string]$Url,
        [hashtable]$Headers
    )
    
    $page = 1
    $morePages = $true
    $items = @()

    # Determine query separator
    $connector = if ($Url -match "\?") { "&" } else { "?" }

    do {
        $fullUrl = "${Url}${connector}per_page=100&page=${page}"
        Write-Verbose "Fetching $fullUrl"
        
        try {
            $response = Invoke-RestMethod -Uri $fullUrl -Headers $Headers -Method Get
            
            # API returns { "tickets": [...] } or { "agents": [...] }
            # We need to find the array property dynamically or assume based on endpoint
            $props = $response | Get-Member -MemberType NoteProperty
            $arrayProp = $props | Where-Object { $response.($_.Name) -is [System.Array] } | Select-Object -First 1
            
            if ($arrayProp) {
                $batch = $response.($arrayProp.Name)
                if ($batch.Count -gt 0) {
                    $items += $batch
                    $page++
                }
                else {
                    $morePages = $false
                }
            }
            else {
                # Fallback if structure is different or empty
                $morePages = $false
            }

            Start-Sleep -Milliseconds 100
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                # Blind retry after 10s
                Start-Sleep -Seconds 10
            }
            else {
                Write-Error "Failed to fetch page $page from $Url : $($_.Exception.Message)"
                $morePages = $false
            }
        }

    } while ($morePages)

    return $items
}

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
try {
    Write-Host "Starting Freshservice Data Export..." -ForegroundColor Cyan

    $encodedApiKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($FreshserviceApiKey):X"))
    $headers = @{
        "Authorization" = "Basic $encodedApiKey"
        "Content-Type"  = "application/json"
    }

    # 1. Fetch Agents for Lookup
    Write-Host "Fetching Agents..." -ForegroundColor Yellow
    $agents = Get-FreshserviceResource -Url "https://$FreshserviceDomain/api/v2/agents" -Headers $headers
    $agentLookup = @{}
    foreach ($a in $agents) { $agentLookup[$a.id] = "$($a.first_name) $($a.last_name)".Trim() }
    Write-Host "Loaded $($agentLookup.Count) agents." -ForegroundColor Green

    # 2. Fetch Groups for Lookup
    Write-Host "Fetching Groups..." -ForegroundColor Yellow
    $groups = Get-FreshserviceResource -Url "https://$FreshserviceDomain/api/v2/groups" -Headers $headers
    $groupLookup = @{}
    foreach ($g in $groups) { $groupLookup[$g.id] = $g.name }
    Write-Host "Loaded $($groupLookup.Count) groups." -ForegroundColor Green

    # 3. Fetch Tickets
    Write-Host "Fetching Tickets..." -ForegroundColor Yellow
    $ticketUrl = "https://$FreshserviceDomain/api/v2/tickets?include=requester"
    
    if ($LookbackMinutes -gt 0) {
        $startTime = (Get-Date).AddMinutes(-$LookbackMinutes).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $ticketUrl += "&updated_since=$([uri]::EscapeDataString($startTime))"
    }

    $tickets = Get-FreshserviceResource -Url $ticketUrl -Headers $headers

    if ($tickets.Count -eq 0) {
        Write-Warning "No tickets found."
        exit
    }

    Write-Host "Found $($tickets.Count) tickets. Enriching data..." -ForegroundColor Cyan

    # 4. Enrich Data
    $enrichedTickets = @()
    foreach ($t in $tickets) {
        # Create a custom object to control order and added fields
        $enriched = [PSCustomObject]@{
            TicketId       = $t.id
            Subject        = $t.subject
            Status         = $t.status
            Priority       = $t.priority
            CreatedAt      = $t.created_at
            UpdatedAt      = $t.updated_at
            
            # Resolved Names
            AgentName      = if ($t.responder_id) { $agentLookup[$t.responder_id] } else { "Unassigned" }
            GroupName      = if ($t.group_id) { $groupLookup[$t.group_id] } else { "Unassigned" }
            
            # Requester (from include=requester)
            RequesterName  = if ($t.requester) { "$($t.requester.first_name) $($t.requester.last_name)".Trim() } else { "Unknown" }
            RequesterEmail = if ($t.requester) { $t.requester.email } else { $null }

            # Original Data (Optional: keep original IDs)
            AgentId        = $t.responder_id
            GroupId        = $t.group_id
            Description    = $t.description_text # or description for HTML
        }
        $enrichedTickets += $enriched
    }

    # 5. Export
    if ($OutputPath.EndsWith(".csv")) {
        $enrichedTickets | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    }
    else {
        $enrichedTickets | ConvertTo-Json -Depth 5 -Compress | Out-File -FilePath $OutputPath -Encoding UTF8
    }

    Write-Host "Export completed to: $OutputPath" -ForegroundColor Green

}
catch {
    Write-Error "Script Failed: $($_.Exception.Message)"
}
