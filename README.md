# AI Interactions Query

This project gives you a simple way to export Microsoft 365 Copilot interaction history from Microsoft Graph and combine it with Copilot Studio / Agent Builder exports if needed.

If you are new to PowerShell or Microsoft Graph, this guide walks you through the setup step by step.

## REST API option

The root-level script uses the Microsoft Graph PowerShell SDK. For a separate app-only implementation that calls the Graph REST API directly, see [rest-api/README.md](rest-api/README.md). The REST option defaults to Graph v1.0, supports client-credentials authentication, and requires `AiEnterpriseInteraction.Read.All` application permission with admin consent (plus `User.Read.All` only when enumerating all users).

## What this project does

This script will:
- check whether the required Microsoft Graph PowerShell modules are installed
- install missing Graph modules automatically from PowerShell Gallery
- connect to Microsoft Graph using your tenant account
- verify the required permissions before running the export
- pull Microsoft 365 Copilot interaction records for one or more users
- optionally merge in Copilot Studio / Agent Builder export files
- sort the results by interaction type and app class
- save the results as CSV and JSON files

## Parameter input guide

Use the following values for the main script variables:

- `TenantId` — optional Microsoft Entra tenant GUID. Example: `00000000-0000-0000-0000-000000000000`
- `UserIds` — optional list of exact users to export. Example: `"user1@contoso.com","user2@contoso.com"`
- `OutputDirectory` — folder to save export files. Example: `".\exports"`
- `Top` — number of records to fetch for each user. Example: `100`
- `UseBeta` — switch to use the Graph beta endpoint. Example: `-UseBeta`
- `SkipUserEnumeration` — switch to skip the default tenant-wide user lookup. Use only in specialized cases.
- `AppClassFilter` — optional Graph app class filter. Example: `"IPM.SkypeTeams.Message.Copilot.BizChat"`
- `CopilotStudioExportPath` — path to a CSV or JSON file or folder containing Copilot Studio / Agent Builder exports. Example: `".\sample-copilot-studio-export.json"`
- `IncludeCopilotStudio` — switch that tells the script to include Copilot Studio / Agent Builder exports.
- `IncludeRawJson` — switch that adds the original raw JSON row to each exported record.

These values must match the format expected by the script. For example:
- `TenantId` must be a GUID string, not a user email address.
- `UserIds` must be Microsoft 365 user identifiers such as UPNs or object IDs, depending on what `Get-MgUser` accepts in your tenant.
- `OutputDirectory` must be a valid folder path, not a file path.

## What you need before you start

Before using this script, make sure you have:
- PowerShell 7 or newer installed
- access to Microsoft Graph with the right permissions in your tenant
- a Microsoft 365 Copilot license that supports Graph-based interaction exports
- permission to read user data in your tenant as needed

If you are not sure whether your account has the right access, ask your Microsoft 365 or Entra administrator.

## Required Microsoft Graph permissions

This script requires the following Microsoft Graph permissions:
- `AiEnterpriseInteraction.Read.All`
- `User.Read.All`

These are the permissions the script checks before it starts exporting data.

## Important limitation to understand

The Microsoft Graph `getAllEnterpriseInteractions` API is only for Microsoft 365 Copilot interaction data. It does not include interactions created by Copilot Studio or Agent Builder experiences.

That means:
- Microsoft 365 Copilot interactions can be pulled from Microsoft Graph
- Copilot Studio / Agent Builder interactions must be imported separately from exported CSV or JSON files

This script handles both sources, but it does it honestly: Graph data and Copilot Studio data are treated as separate data sources that can be combined in one final export.

## How to run the script

Open PowerShell in the folder that contains the script, then run:

```powershell
pwsh -File .\Get-AIInteractions.ps1
```

That is the simplest way to run it. It is designed to work for all users in the tenant by default. If you do not pass `-UserIds`, the script enumerates every user and exports their Microsoft 365 Copilot interactions.

This means:
- default behavior = all users in the tenant
- `-UserIds` = only specific users
- `-SkipUserEnumeration` = disables the tenant-wide user list and is intended only for special cases

The script will:
- install missing modules if needed
- connect to Microsoft Graph
- collect all users in the tenant
- export their Microsoft 365 Copilot interactions

## Default behavior for all users

This script is built for full-tenant coverage. The default export path is tenant-wide. In other words:

```powershell
pwsh -File .\Get-AIInteractions.ps1 -OutputDirectory ".\exports"
```

will query all available users in the tenant unless you explicitly restrict it with `-UserIds`.

If you only want to validate the script against one or two test accounts, use:

```powershell
pwsh -File .\Get-AIInteractions.ps1 -UserIds "user1@contoso.com","user2@contoso.com" -OutputDirectory ".\exports"
```

## Tenant-wide export checklist

Use this checklist before running the full export for every user:

1. Confirm that your tenant account has the required delegated permission scopes:
   - `AiEnterpriseInteraction.Read.All`
   - `User.Read.All`
2. Connect to the correct tenant before running the script:
   ```powershell
   Connect-MgGraph -TenantId "<tenant-guid>" -Scopes "AiEnterpriseInteraction.Read.All","User.Read.All"
   ```
3. Do not include `-UserIds` unless you intentionally want to restrict the export to a subset of users.
4. Keep the default path as-is so the script uses `Get-MgUser -All` and processes all users in the tenant.
5. Choose a large enough output folder and make sure you have storage available for CSV and JSON export files.
6. If the tenant is large, consider testing with a few users first before running the full tenant export.
7. If you need to limit the data to a specific app class, use `-AppClassFilter`.
8. If you need a smaller test run, use `-Top 100` or `-UserIds` only for validation.

Recommended full-tenant command:

```powershell
pwsh -File .\Get-AIInteractions.ps1 -OutputDirectory ".\exports"
```

This is the standard all-users mode and is the correct command when you want the script to enumerate every user and export all available Microsoft 365 Copilot interactions.

## Common examples

### Example 1: run with a specific tenant

```powershell
pwsh -File .\Get-AIInteractions.ps1 -TenantId "<tenant-guid>" -OutputDirectory ".\exports" -Top 100
```

- `-TenantId` tells the script which tenant to use
- `-OutputDirectory` tells it where to save the exported files
- `-Top 100` limits the API to the first 100 records per user request

### Example 2: export for specific users only

```powershell
pwsh -File .\Get-AIInteractions.ps1 -UserIds "user1@contoso.com","user2@contoso.com" -OutputDirectory ".\exports"
```

This is useful when you only want a subset of users instead of all users in the tenant.

### Example 3: use the beta endpoint

```powershell
pwsh -File .\Get-AIInteractions.ps1 -UserIds "user1@contoso.com" -OutputDirectory ".\exports" -UseBeta
```

Use `-UseBeta` only if you specifically need the beta version of the Graph endpoint. Otherwise, the default v1.0 endpoint is usually the best choice.

### Example 4: filter by app class

```powershell
pwsh -File .\Get-AIInteractions.ps1 -AppClassFilter "IPM.SkypeTeams.Message.Copilot.BizChat" -OutputDirectory ".\exports"
```

This filters the results to a specific Copilot app class, such as Teams or Microsoft 365 Chat.

### Example 5: include Copilot Studio / Agent Builder exports

```powershell
pwsh -File .\Get-AIInteractions.ps1 -IncludeCopilotStudio -CopilotStudioExportPath ".\copilot-studio-exports" -OutputDirectory ".\exports"
```

This tells the script to look for CSV or JSON export files in the folder you provide and merge them into the results.

## What a Copilot Studio export should look like

The `CopilotStudioExportPath` can point to a folder containing exported CSV or JSON files. The script looks for files such as:
- `.csv`
- `.json`

The files should ideally contain fields like:
- `Prompt` or `UserPrompt`
- `Response` or `Answer`
- `CreatedDateTime` or `Timestamp`
- `AgentName`
- `UserPrincipalName`
- `SessionId`
- `InteractionId`

The script normalizes these records so they fit into the same export schema as the Microsoft Graph results.

## How the script connects to Microsoft Graph

When you run it, the script does the following:
1. checks whether the Graph modules are installed
2. installs them if they are missing
3. connects to Graph using `Connect-MgGraph`
4. checks that the required scopes are in the current session

If you are not already connected, the script prompts or connects using the required scopes automatically.

## If you get a permission error

If the script says permissions are missing, you may need to reconnect using:

```powershell
Connect-MgGraph -Scopes "AiEnterpriseInteraction.Read.All","User.Read.All"
```

If you are using a tenant-specific connection, you can also use:

```powershell
Connect-MgGraph -TenantId "<tenant-guid>" -Scopes "AiEnterpriseInteraction.Read.All","User.Read.All"
```

If your admin has not granted these permissions, the script cannot pull the data you need.

## Output files

The script writes the following files to the output directory:
- `AIInteractions_<timestamp>.csv`
- `AIInteractions_<timestamp>.json`

These files contain the combined interaction records, sorted by:
- `InteractionType`
- `AppClass`
- `CreatedDateTime`

## Why the CSV and JSON are useful

The exported data is helpful when you want to:
- analyze AI usage by type
- review prompts and responses
- group activity by app or conversation type
- compare Microsoft 365 Copilot traffic with Copilot Studio/Agent Builder usage
- build reports for compliance, usage analysis, or governance

## Tips for beginners

- Start with the default command first before changing filters.
- Use `-UserIds` if you want to test the export on a small set of users.
- Use `-OutputDirectory` to keep exported files organized in a dedicated folder.
- If you are unsure whether the script is working, run it with one user first.
- If you need to combine Copilot Studio exports, use `-IncludeCopilotStudio` and a folder path that contains the CSV/JSON files.

## Final reminder

This project is meant to help you gather enterprise AI interaction data in one place for analysis. It is especially useful for researchers, governance teams, or anyone tracking AI usage across Microsoft 365 Copilot and Copilot Studio experiences.

## Output summary

The script creates:
- a CSV export for spreadsheet review and sorting
- a JSON export for programmatic handling and analysis
- a summary of interaction counts by type in the terminal output

This makes it easy to inspect the data quickly without writing a second script.

## Sample files included in this repo

This repo includes sample Copilot Studio export files so you can see the kind of input the script expects:

- `sample-copilot-studio-export.csv`
- `sample-copilot-studio-export.json`

These sample files contain example data including:
- `UserPrincipalName`
- `AgentName`
- `SessionId`
- `InteractionId`
- `Prompt` or `UserPrompt`
- `Response` or `Answer`
- `CreatedDateTime`
- `InteractionType`
- `AppClass`
- `SourceApplication`

To test the merge logic with the CSV sample file, run:

```powershell
pwsh -File .\Get-AIInteractions.ps1 -IncludeCopilotStudio -CopilotStudioExportPath ".\sample-copilot-studio-export.csv" -OutputDirectory ".\exports"
```

To test the merge logic with the JSON sample file, run:

```powershell
pwsh -File .\Get-AIInteractions.ps1 -IncludeCopilotStudio -CopilotStudioExportPath ".\sample-copilot-studio-export.json" -OutputDirectory ".\exports"
```

You can also point `-CopilotStudioExportPath` at a folder containing multiple CSV or JSON files.

## Suggested folder layout

A simple layout for this project looks like this:

```text
aiinteractions_query/
├── Get-AIInteractions.ps1
├── README.md
├── sample-copilot-studio-export.csv
├── sample-export-folder/
│   └── README.txt
├── exports/
│   └── AIInteractions_20260915_143000.csv
│   └── AIInteractions_20260915_143000.json
└── ...
```

This helps keep your raw exports and processed results separated.
