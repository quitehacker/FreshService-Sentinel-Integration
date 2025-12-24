# Freshservice Data Export

This project contains a PowerShell script [`Get-FreshserviceData.ps1`](./Get-FreshserviceData.ps1) that fetches tickets from Freshservice, enriches them with **Agent** and **Group** names, and exports the data to a local JSON or CSV file.

## Features
- **Data Enrichment**: Automatically replaces IDs (like `responder_id`, `group_id`) with readable names (`AgentName`, `GroupName`).
- **Includes Requester**: Fetches requester details (Name, Email) for each ticket.
- **Incremental Fetch**: Supports a `LookbackMinutes` parameter to fetch only recently updated tickets.
- **Local Export**: Saves the enriched data to your local machine for analysis.

## Prerequisites
-   **Freshservice Domain**: e.g., `yourcompany.freshservice.com`
-   **Freshservice API Key**: From Profile Settings in Freshservice.

## Usage

### Parameters
| Parameter | Description | Required | Default |
| :--- | :--- | :--- | :--- |
| `FreshserviceDomain` | Your Freshservice domain URL base. | Yes | - |
| `FreshserviceApiKey` | Your personal API Key. | Yes | - |
| `OutputPath` | Full path for the output file (`.json` or `.csv`). | No | `.\FreshserviceTickets.json` |
| `LookbackMinutes` | Fetch tickets updated in last N mins (0 for ALL). | No | `0` |

### Example 1: Export to JSON (Default)
```powershell
.\Get-FreshserviceData.ps1 `
    -FreshserviceDomain "mycompany.freshservice.com" `
    -FreshserviceApiKey "YOUR_API_KEY" `
    -OutputPath ".\MyTickets.json"
```

### Example 2: Export to CSV
```powershell
.\Get-FreshserviceData.ps1 `
    -FreshserviceDomain "mycompany.freshservice.com" `
    -FreshserviceApiKey "YOUR_API_KEY" `
    -OutputPath ".\MyTickets.csv"
```

### Example 3: Fetch Recent Changes (Last Hour)
```powershell
.\Get-FreshserviceData.ps1 `
    -FreshserviceDomain "mycompany.freshservice.com" `
    -FreshserviceApiKey "YOUR_API_KEY" `
    -LookbackMinutes 60
```
