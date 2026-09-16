# REST API implementation

`Get-AIInteractions-Rest.ps1` is a self-contained PowerShell implementation that calls Microsoft Graph directly with HTTPS. It does not use Microsoft.Graph PowerShell SDK cmdlets and does not change the root-level SDK script.

## Prerequisites and permissions

Create a Microsoft Entra app registration with **application** permission `AiEnterpriseInteraction.Read.All` and grant admin consent. The interaction-history API does not support delegated permissions. The tenant and target users also need the appropriate Microsoft 365 Copilot license/service plan.

When using `-AllUsers`, also grant **application** permission `User.Read.All` with admin consent. This permission is not needed if you provide explicit user IDs.

The API returns Microsoft 365 Copilot interactions, including prompts and responses. It does **not** return interactions from agents created with Copilot Studio.

See Microsoft's endpoint documentation: [getAllEnterpriseInteractions](https://learn.microsoft.com/graph/api/aiinteractionhistory-getallenterpriseinteractions?view=graph-rest-1.0).

## Authentication

Use app-only client credentials. Supply values with parameters or the `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, and `AZURE_CLIENT_SECRET` environment variables. Do not put a secret in a script or commit it. For example, set `AZURE_CLIENT_SECRET` in a secure CI secret store or prompt for it at runtime before invoking the script.

An `-AccessToken` can be supplied instead. It takes precedence over client credentials; use it only from a secure source and never log it. The token must be an app-only Graph token containing the approved application permissions.

## Examples

Explicit user IDs do not enumerate `/users`:

```powershell
$env:AZURE_TENANT_ID = '<tenant-id>'
$env:AZURE_CLIENT_ID = '<app-client-id>'
$env:AZURE_CLIENT_SECRET = '<retrieve this secret from a secure environment or secret store>'
pwsh -File .\Get-AIInteractions-Rest.ps1 -UserIds '<user-object-id>' -OutputDirectory ..\exports
```

For interactive use, pass the secret without saving it:

```powershell
$secret = Read-Host 'Client secret'
pwsh -File .\Get-AIInteractions-Rest.ps1 -TenantId '<tenant-id>' -ClientId '<app-client-id>' -ClientSecret $secret -UserIds '<user-object-id>'
Remove-Variable secret
```

Enumerate all users (requires `User.Read.All`):

```powershell
pwsh -File .\Get-AIInteractions-Rest.ps1 -AllUsers -OutputDirectory ..\exports
```

Filter by application class and a required closed created-date range:

```powershell
pwsh -File .\Get-AIInteractions-Rest.ps1 -UserIds '<user-object-id>' -AppClassFilter 'IPM.SkypeTeams.Message.Copilot.BizChat' -CreatedDateTimeStart '2026-01-01T00:00:00Z' -CreatedDateTimeEnd '2026-01-31T23:59:59Z' -Top 100
```

The script uses the stable `v1.0` endpoint by default. Add `-UseBeta` only when beta behavior is specifically required; beta APIs can change. It URL-encodes initial query parameters, requests up to 100 records per page by default, and follows Graph's `@odata.nextLink` unchanged for subsequent pages.

Results are normalized and written as `AIInteractions_<timestamp>.csv` and `.json`. Use `-IncludeRawJson` to add each original API object to the export.
