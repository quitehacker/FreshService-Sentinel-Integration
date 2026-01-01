# Freshservice to Sentinel Integration (DCR)

This project syncs Freshservice tickets to either a **Local File** (for testing/backup) or **Microsoft Sentinel** (via DCR Logs Ingestion API).

## Features
- **Dual Mode**:
    - **Local Export**: Save enriched data to JSON/CSV.
    - **Azure Ingestion**: Push data directly to a Log Analytics table via DCR.
- **Enrichment**: Resolves Agent/Group IDs to names.
- **Secure**: Uses Azure AD Service Principal authentication.

## Setup
### 1. Azure Configuration (Ingestion Only)
If you intend to send data to Sentinel, you **must** configure Azure resources first. See the [Azure DCR Setup Guide](./Azure_DCR_Setup_Guide.md).
Required values:
- `TenantId`, `ClientId`, `ClientSecret` (App Registration)
- `DceEndpoint` (Data Collection Endpoint)
- `DcrImmutableId` (Data Collection Rule)
- `StreamName` (Table definition in DCR)

### 2. Freshservice Configuration
- `FreshserviceDomain`
- `FreshserviceApiKey`

## Usage

### Mode 1: Local Test / Export (Recommended First Step)
Use this to verify data before sending to Azure.
Fetches tickets updated in the last `N` minutes (default 60) and saves to a file.

**Command:**
```powershell
.\Sync-FreshserviceToSentinel.ps1 `
    -FreshserviceDomain "yourcompany.freshservice.com" `
    -FreshserviceApiKey "YOUR_API_KEY" `
    -OutputPath ".\FreshserviceData.json" `
    -LookbackMinutes 5
```
*Note: Azure parameters are NOT required for local export.*

### Mode 2: Azure Sentinel Ingestion
Use this for production automation.

**Command:**
```powershell
.\Sync-FreshserviceToSentinel.ps1 `
    -FreshserviceDomain "yourcompany.freshservice.com" `
    -FreshserviceApiKey "YOUR_API_KEY" `
    -TenantId "GUID" `
    -ClientId "GUID" `
    -ClientSecret "SECRET" `
    -DceEndpoint "https://xyz.eastus-1.ingest.monitor.azure.com" `
    -DcrImmutableId "dcr-..." `
    -StreamName "Custom-FreshserviceTickets_CL" `
    -LookbackMinutes 60
```

## Automation
Schedule Mode 2 in an **Azure Automation Runbook** (Hybrid Worker or Cloud Job). Store secrets in Automation Variables/Credentials and pass them to the script parameters.
