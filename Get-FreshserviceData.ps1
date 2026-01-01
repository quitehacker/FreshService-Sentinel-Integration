<#
.SYNOPSIS
    Fetches tickets from Freshservice, enriches them, and either exports locally OR ingests to Sentinel (DCR).

.DESCRIPTION
    1. Fetches Agents/Groups for enrichment.
    2. Fetches Tickets (incremental based on LookbackMinutes).
    3. Enriches data.
    4. If OutputPath is provided, saves JSON/CSV locally.
    5. If Azure params are provided, pushes to Azure Sentinel DCR.

.PARAMETER FreshserviceDomain
    Domain, e.g., "yourcompany.freshservice.com".

.PARAMETER FreshserviceApiKey
    API Key for Freshservice.

.PARAMETER TenantId
    Azure AD Tenant ID. Required for Sentinel ingestion.

.PARAMETER ClientId
    App Registration Client ID. Required for Sentinel ingestion.

.PARAMETER ClientSecret
    App Registration Client Secret. Required for Sentinel ingestion.

.PARAMETER DceEndpoint
    Data Collection Endpoint Logs Ingestion URL. Required for Sentinel ingestion.

.PARAMETER DcrImmutableId
    Immutable ID of the Data Collection Rule. Required for Sentinel ingestion.

.PARAMETER StreamName
    Input Stream Name in the DCR. Required for Sentinel ingestion.

.PARAMETER OutputPath
    Make this specified to save data to a local file (e.g., .\data.json). 
    If this is set, Azure Ingestion is SKIPPED unless you also provide Azure params.

.PARAMETER LookbackMinutes
    Fetch tickets updated in the last N minutes. 0 = ALL. Default 60.
#>

[CmdletBinding(DefaultParameterSetName = "LocalExport")]
param(
    [Parameter(Mandatory = $true)] [string]$FreshserviceDomain,
    [Parameter(Mandatory = $true)] [string]$FreshserviceApiKey,
    
    # Azure Auth
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$TenantId,
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$ClientSecret,
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$DceEndpoint,
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$DcrImmutableId,
    [Parameter(Mandatory = $true, ParameterSetName = "AzureIngestion")] [string]$StreamName,

    # Local Export
    [Parameter(ParameterSetName = "LocalExport")] 
    [string]$OutputPath = ".\FreshserviceData.json",

    [int]$LookbackMinutes = 60
)

# -----------------------------------------------------------------------------
# Function: Get-AzAccessToken
# -----------------------------------------------------------------------------
function Get-AzAccessToken {
    param($TenantId, $ClientId, $ClientSecret)
    $body = @{ grant_type = "client_credentials"; client_id = $ClientId; client_secret = $ClientSecret; scope = "https://monitor.azure.com//.default" }
    $uri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    try {
        $response = Invoke-RestMethod -Method Post -Uri $uri -Body $body -ErrorAction Stop
        return $response.access_token
    }
    catch { throw "Failed to get Azure Access Token: $($_.Exception.Message)" }
}

# -----------------------------------------------------------------------------
# Function: Send-AzMonitorIngestion
# -----------------------------------------------------------------------------
function Send-AzMonitorIngestion {
    param([string]$DceEndpoint, [string]$DcrId, [string]$Stream, [string]$AccessToken, [array]$Payload)
    $apiVersion = "2023-01-01"
    $uri = "$DceEndpoint/dataCollectionRules/$DcrId/streams/$($Stream)?api-version=$apiVersion"
    $headers = @{ "Authorization" = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    try {
        $json = $Payload | ConvertTo-Json -Depth 5 -Compress
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $json -ErrorAction Stop
        Write-Verbose "Ingestion Success. Code: $($response.StatusCode)"
    }
    catch {
        Write-Error "Ingestion Failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            Write-Error "Details: $($reader.ReadToEnd())"
        }
    }
}

# -----------------------------------------------------------------------------
# Helper: Get-FreshserviceResource
# -----------------------------------------------------------------------------
function Get-FreshserviceResource {
    param([string]$Url, [hashtable]$Headers)
    $page = 1; $morePages = $true; $items = @()
    $connector = if ($Url -match "\?") { "&" } else { "?" }
    do {
        $fullUrl = "${Url}${connector}per_page=100&page=${page}"
        Write-Verbose "Fetching $fullUrl"
        try {
            $response = Invoke-RestMethod -Uri $fullUrl -Headers $Headers -Method Get
            $props = $response | Get-Member -MemberType NoteProperty
            $arrayProp = $props | Where-Object { $response.($_.Name) -is [System.Array] } | Select-Object -First 1
            if ($arrayProp) {
                $batch = $response.($arrayProp.Name)
                if ($batch.Count -gt 0) { $items += $batch; $page++ } else { $morePages = $false }
            }
            else { $morePages = $false }
            Start-Sleep -Milliseconds 100
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) { Start-Sleep -Seconds 10 } 
            else { Write-Error "Fetch Error: $($_.Exception.Message)"; $morePages = $false }
        }
    } while ($morePages)
    return $items
}

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
try {
    Write-Host "Starting Ticket Sync (Window: Last $LookbackMinutes mins)..." -ForegroundColor Cyan

    $encodedApiKey = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($FreshserviceApiKey):X"))
    $fsHeaders = @{ "Authorization" = "Basic $encodedApiKey"; "Content-Type" = "application/json" }

    # 1. Fetch & Enrich
    Write-Host "Fetching Metadata (Agents/Groups)..." -ForegroundColor Yellow
    $agents = Get-FreshserviceResource -Url "https://$FreshserviceDomain/api/v2/agents" -Headers $fsHeaders
    $agentLookup = @{}; foreach ($a in $agents) { $agentLookup[$a.id] = "$($a.first_name) $($a.last_name)".Trim() }

    $groups = Get-FreshserviceResource -Url "https://$FreshserviceDomain/api/v2/groups" -Headers $fsHeaders
    $groupLookup = @{}; foreach ($g in $groups) { $groupLookup[$g.id] = $g.name }

    $ticketUrl = "https://$FreshserviceDomain/api/v2/tickets?include=requester"
    if ($LookbackMinutes -gt 0) {
        $startTime = (Get-Date).AddMinutes(-$LookbackMinutes).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $ticketUrl += "&updated_since=$([uri]::EscapeDataString($startTime))"
        Write-Host "Fetching tickets updated since $startTime" -ForegroundColor Cyan
    }
    $tickets = Get-FreshserviceResource -Url $ticketUrl -Headers $fsHeaders

    if ($tickets.Count -eq 0) { Write-Host "No logs/tickets found in the last $LookbackMinutes minutes." -ForegroundColor Green; exit }
    # 4. Enrich & Transform
    Write-Host "Enriching $($tickets.Count) tickets..." -ForegroundColor Cyan
    $payloadSet = @()
    foreach ($t in $tickets) {
        # 1. Start with a hash table of ALL properties
        $orderedProps = [ordered]@{}
        
        # Add Standard Fields from Ticket Object
        $t.PSObject.Properties | ForEach-Object {
            if ($_.Name -ne "custom_fields" -and $_.Name -ne "tags") {
                $orderedProps[$_.Name] = $_.Value
            }
        }

        # 2. Add Enrichment Fields (Resolved Names)
        $orderedProps["AgentName"] = if ($t.responder_id) { $agentLookup[$t.responder_id] } else { "Unassigned" }
        $orderedProps["GroupName"] = if ($t.group_id) { $groupLookup[$t.group_id] } else { "Unassigned" }
        $orderedProps["RequesterName"] = if ($t.requester) { "$($t.requester.first_name) $($t.requester.last_name)".Trim() } else { "Unknown" }
        $orderedProps["RequesterEmail"] = if ($t.requester) { $t.requester.email } else { "" }
        $orderedProps["TimeGenerated"] = $t.updated_at # For Sentinel

        # 3. Flatten Custom Fields (if any)
        if ($t.custom_fields) {
            $t.custom_fields.PSObject.Properties | ForEach-Object {
                $orderedProps["Custom_$($_.Name)"] = $_.Value
            }
        }

        # 4. Flatten Tags
        if ($t.tags) {
            $orderedProps["Tags"] = ($t.tags -join ", ")
        }

        # Convert to Object
        $payloadSet += [PSCustomObject]$orderedProps
    }
    Write-Host "Prepared $($payloadSet.Count) records." -ForegroundColor Green

    # 2. Local Export (Priority if ParameterSetName is LocalExport)
    if ($PSCmdlet.ParameterSetName -eq "LocalExport") {
        Write-Host "Saving to local file: $OutputPath" -ForegroundColor Yellow
        if ($OutputPath.EndsWith(".csv")) {
            $payloadSet | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        }
        else {
            $payloadSet | ConvertTo-Json -Depth 5 -Compress | Out-File -FilePath $OutputPath -Encoding UTF8
        }
        Write-Host "Export Complete. (Azure Ingestion Skipped)" -ForegroundColor Green
    }
    
    # 3. Azure Ingestion (Only if Azure Params Provided)
    elseif ($PSCmdlet.ParameterSetName -eq "AzureIngestion") {
        Write-Host "Authenticating to Azure..." -ForegroundColor Yellow
        $token = Get-AzAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

        $batchSize = 200
        for ($i = 0; $i -lt $payloadSet.Count; $i += $batchSize) {
            $endParams = [Math]::Min($i + $batchSize - 1, $payloadSet.Count - 1)
            $batch = $payloadSet[$i..$endParams]
            Write-Host "Ingesting batch ($($batch.Count) records) to DCR..." -ForegroundColor Cyan
            Send-AzMonitorIngestion -DceEndpoint $DceEndpoint -DcrId $DcrImmutableId -Stream $StreamName -AccessToken $token -Payload $batch
        }
        Write-Host "Ingestion Complete." -ForegroundColor Green
    }

}
catch {
    Write-Error "Script Failed: $($_.Exception.message)"
}
