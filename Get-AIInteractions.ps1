<#
    AI Interactions Query
    --------------------
    This script exports Microsoft 365 Copilot interaction history for users in a tenant.
    It also supports adding Copilot Studio or Agent Builder export data so you can combine
    multiple AI interaction sources in one CSV or JSON output.

    Beginner notes:
    - PowerShell scripts accept parameters like -TenantId and -OutputDirectory.
    - If you do not have the required Graph modules, the script installs them automatically.
    - If you do not have the required Graph permissions, the script tells you what to connect with.
    - The Microsoft Graph endpoint gets Microsoft 365 Copilot data only.
    - Copilot Studio / Agent Builder data must come from exported files, because Graph does not currently
      return those interactions from the getAllEnterpriseInteractions endpoint.

    Typical usage examples:
    pwsh -File .\Get-AIInteractions.ps1
    pwsh -File .\Get-AIInteractions.ps1 -TenantId "<tenant-guid>" -OutputDirectory ".\exports"
    pwsh -File .\Get-AIInteractions.ps1 -UserIds "user1@contoso.com","user2@contoso.com" -OutputDirectory ".\exports"
    pwsh -File .\Get-AIInteractions.ps1 -IncludeCopilotStudio -CopilotStudioExportPath ".\copilot-studio-exports" -OutputDirectory ".\exports"
#>

<#!
.SYNOPSIS
Exports Microsoft 365 Copilot and optional Copilot Studio/Agent Builder interactions.

.DESCRIPTION
This script checks for required Microsoft Graph PowerShell modules, connects to Graph,
validates required permissions, enumerates users in the tenant by default, and exports
interaction records as CSV and JSON. It can optionally merge Copilot Studio or Agent Builder
exports from CSV or JSON files.

.PARAMETER TenantId
Optional GUID or tenant identifier to scope the Graph connection to a specific tenant.
Example: "00000000-0000-0000-0000-000000000000"

.PARAMETER UserIds
Optional array of specific users to export. If omitted, the script enumerates every user in the tenant.
Example: "user1@contoso.com","user2@contoso.com"

.PARAMETER OutputDirectory
Folder where the script writes the CSV and JSON export files.
Example: ".\exports"

.PARAMETER Top
Maximum number of records to request per user from the Graph interaction endpoint.
Typical value: 100

.PARAMETER UseBeta
Use the Microsoft Graph beta endpoint instead of the default v1.0 endpoint.
Only use this when you specifically need beta behavior.

.PARAMETER SkipUserEnumeration
Skips the default tenant-wide Get-MgUser -All process. Use only in special scenarios.

.PARAMETER AppClassFilter
Optional application class filter for Graph results.
Example: "IPM.SkypeTeams.Message.Copilot.BizChat"

.PARAMETER CopilotStudioExportPath
Path to a CSV or JSON file, or a folder containing those files, for Copilot Studio/Agent Builder data.
Example: ".\sample-copilot-studio-export.json"

.PARAMETER IncludeCopilotStudio
Enables merging Copilot Studio or Agent Builder export records into the final result set.

.PARAMETER IncludeRawJson
Includes the original raw JSON payload for each record in the export.

.EXAMPLE
pwsh -File .\Get-AIInteractions.ps1

.EXAMPLE
pwsh -File .\Get-AIInteractions.ps1 -OutputDirectory ".\exports"

.EXAMPLE
pwsh -File .\Get-AIInteractions.ps1 -UserIds "user1@contoso.com","user2@contoso.com" -OutputDirectory ".\exports"

.EXAMPLE
pwsh -File .\Get-AIInteractions.ps1 -IncludeCopilotStudio -CopilotStudioExportPath ".\sample-copilot-studio-export.json" -OutputDirectory ".\exports"
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [string[]]$UserIds,
    [string]$OutputDirectory = ".",
    [int]$Top = 100,
    [switch]$UseBeta,
    [switch]$SkipUserEnumeration,
    [string]$AppClassFilter,
    [string]$CopilotStudioExportPath,
    [switch]$IncludeCopilotStudio,
    [switch]$IncludeRawJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

<#
    Function: Ensure-GraphModules
    Purpose:
    Confirms the Microsoft Graph PowerShell modules are installed and available.
    If they are missing, this function installs them from PSGallery automatically.

    Beginner note:
    A module is a package of PowerShell commands. This script uses Graph commands such as
    Connect-MgGraph and Get-MgUser, which live in the Microsoft Graph modules.
#>
function Ensure-GraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users'
    )

    foreach ($moduleName in $requiredModules) {
        $installed = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
        if (-not $installed) {
            Write-Host "Installing missing Microsoft Graph module: $moduleName" -ForegroundColor Yellow
            Install-Module -Name $moduleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        }
    }

    Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users -Force -ErrorAction Stop
}

<#
    Function: Test-GraphPermissions
    Purpose:
    Verifies that the current Microsoft Graph session includes the required permissions.
    If scopes don't exist, the user receives a clear message explaining how to reconnect.
#>
function Test-GraphPermissions {
    $requiredScopes = @(
        'AiEnterpriseInteraction.Read.All',
        'User.Read.All'
    )

    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Warning "No active Microsoft Graph connection was found. Run Connect-MgGraph -Scopes 'AiEnterpriseInteraction.Read.All','User.Read.All'"
        return $false
    }

    $contextScopes = @($context.Scopes)
    $missing = @($requiredScopes | Where-Object { $contextScopes -notcontains $_ })

    if ($missing.Count -gt 0) {
        Write-Warning "The current Graph session is missing required scopes: $($missing -join ', ')"
        Write-Warning "Reconnect with: Connect-MgGraph -Scopes 'AiEnterpriseInteraction.Read.All','User.Read.All'"
        return $false
    }

    return $true
}

<#
    Function: Get-UsersToProcess
    Purpose:
    Returns the list of users to query.

    Default behavior is tenant-wide collection: if you do not specify -UserIds,
    the script enumerates every user in the tenant. You can restrict the export to
    a specific set by passing -UserIds, which is useful for testing or limited runs.
#>
function Get-UsersToProcess {
    param(
        [string[]]$ExplicitUsers,
        [switch]$SkipEnumeration
    )

    if ($ExplicitUsers -and $ExplicitUsers.Count -gt 0) {
        foreach ($userId in $ExplicitUsers) {
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            [pscustomobject]@{
                Id = $user.Id
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
            }
        }
        return
    }

    if ($SkipEnumeration) {
        return @()
    }

    Write-Host "Enumerating all users in the tenant..." -ForegroundColor Cyan
    Get-MgUser -All -Property Id, UserPrincipalName, DisplayName |
        Sort-Object UserPrincipalName |
        Select-Object Id, UserPrincipalName, DisplayName
}

<#
    Function: Get-AIInteractionHistoryForUser
    Purpose:
    Calls the Microsoft Graph interaction history endpoint for one user.
    It supports optional paging and app-class filtering.

    Beginner note:
    The endpoint is a REST API call against Microsoft Graph. This script uses Invoke-MgGraphRequest,
    which is the Microsoft Graph PowerShell method for calling Graph endpoints directly.
#>
function Get-AIInteractionHistoryForUser {
    param(
        [string]$UserId,
        [int]$PageSize,
        [string]$AppClassFilter,
        [switch]$UseBeta
    )

    $version = if ($UseBeta) { 'beta' } else { 'v1.0' }
    $baseUri = "https://graph.microsoft.com/$version/copilot/users/$UserId/interactionHistory/getAllEnterpriseInteractions"
    $uri = $baseUri

    $queryParts = @()
    if ($PageSize -gt 0) { $queryParts += "`$top=$PageSize" }
    if ($AppClassFilter) {
        $queryParts += "`$filter=appClass eq '$([System.Uri]::EscapeDataString($AppClassFilter))'"
    }

    if ($queryParts.Count -gt 0) { $uri = "$baseUri?$($queryParts -join '&')" }

    $allInteractions = @()
    $nextLink = $uri

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
        $items = @()
        if ($response.value) {
            $items = @($response.value)
        }
        elseif ($response -is [System.Collections.IEnumerable] -and -not ($response -is [string])) {
            $items = @($response)
        }
        elseif ($response) {
            $items = @($response)
        }

        foreach ($item in $items) {
            $content = if ($item.body) { $item.body.content } else { $null }
            $contentType = if ($item.body) { $item.body.contentType } else { $null }

            $fromApplication = $null
            if ($item.from -and $item.from.application) {
                $fromApplication = $item.from.application.displayName
            }

            $interaction = [pscustomobject]@{
                UserId = $UserId
                UserPrincipalName = $null
                DisplayName = $null
                InteractionId = $item.id
                SessionId = $item.sessionId
                RequestId = $item.requestId
                AppClass = $item.appClass
                InteractionType = $item.interactionType
                ConversationType = $item.conversationType
                CreatedDateTime = $item.createdDateTime
                Locale = $item.locale
                SourceApplication = $fromApplication
                ContentType = $contentType
                Content = $content
                Contexts = if ($item.contexts) { ($item.contexts | ConvertTo-Json -Depth 10 -Compress) } else { $null }
                Attachments = if ($item.attachments) { ($item.attachments | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                Mentions = if ($item.mentions) { ($item.mentions | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                Links = if ($item.links) { ($item.links | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                RawJson = if ($IncludeRawJson) { ($item | ConvertTo-Json -Depth 100 -Compress) } else { $null }
            }

            $allInteractions += $interaction
        }

        $nextLink = $response.'@odata.nextLink'
    } while ($nextLink)

    return $allInteractions
}

<#
    Function: Import-CopilotStudioExports
    Purpose:
    Reads CSV or JSON export files from Copilot Studio or Agent Builder and normalizes them
    into the same record shape used by the Graph export.

    Beginner note:
    Most Copilot Studio exports are not in the exact same format as Graph output. This function
    looks for common field names like Prompt, Response, AgentName, and UserPrincipalName, then
    maps them into a consistent structure for export.
#>
function Import-CopilotStudioExports {
    param(
        [string]$Path
    )

    if (-not $Path) {
        return @()
    }

    $resolvedPath = if ([System.IO.Path]::IsPathRooted($Path)) { $Path } else { (Join-Path (Get-Location) $Path) }

    if (-not (Test-Path -Path $resolvedPath)) {
        throw "Copilot Studio export path not found: $resolvedPath"
    }

    $files = @()
    if (Test-Path -Path $resolvedPath -PathType Leaf) {
        $files = @($resolvedPath)
    }
    else {
        $files = Get-ChildItem -Path $resolvedPath -Recurse -File | Where-Object { $_.Extension -in '.csv', '.json' }
    }

    if ($files.Count -eq 0) {
        Write-Warning "No CSV/JSON Copilot Studio export files were found in '$resolvedPath'."
        return @()
    }

    $normalized = @()
    foreach ($file in $files) {
        $extension = $file.Extension.ToLowerInvariant()
        $rows = @()

        if ($extension -eq '.json') {
            $jsonContent = Get-Content -Path $file.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 100
            $rows = @($jsonContent)
            if ($jsonContent -and $jsonContent.value) {
                $rows = @($jsonContent.value)
            }
        }
        else {
            $rows = @(Import-Csv -Path $file.FullName)
        }

        foreach ($row in $rows) {
            $prompt = $row.Prompt
            if (-not $prompt) { $prompt = $row.UserPrompt }
            if (-not $prompt) { $prompt = $row.UserMessage }
            if (-not $prompt) { $prompt = $row.Question }
            if (-not $prompt) { $prompt = $row.Input }

            $response = $row.Response
            if (-not $response) { $response = $row.Output }
            if (-not $response) { $response = $row.Answer }

            $interactionDate = $row.CreatedDateTime
            if (-not $interactionDate) { $interactionDate = $row.Timestamp }
            if (-not $interactionDate) { $interactionDate = $row.Created }
            if (-not $interactionDate) { $interactionDate = (Get-Date).ToString('o') }

            $normalized += [pscustomobject]@{
                UserId = $row.UserId
                UserPrincipalName = $row.UserPrincipalName
                DisplayName = $row.DisplayName
                InteractionId = if ($row.InteractionId) { $row.InteractionId } elseif ($row.Id) { $row.Id } else { "$($file.Name)-$($normalized.Count)" }
                SessionId = $row.SessionId
                RequestId = $row.RequestId
                AppClass = if ($row.AppClass) { $row.AppClass } elseif ($row.AgentType) { $row.AgentType } else { 'Copilot Studio / Agent Builder' }
                InteractionType = if ($row.InteractionType) { $row.InteractionType } elseif ($row.Type) { $row.Type } else { 'copilotstudio' }
                ConversationType = if ($row.ConversationType) { $row.ConversationType } elseif ($row.AgentName) { $row.AgentName } else { 'agent' }
                CreatedDateTime = $interactionDate
                Locale = $row.Locale
                SourceApplication = if ($row.SourceApplication) { $row.SourceApplication } elseif ($row.AgentName) { $row.AgentName } else { 'Copilot Studio / Agent Builder' }
                ContentType = if ($row.ContentType) { $row.ContentType } else { 'text' }
                Content = if ($response) { $response } else { $prompt }
                Contexts = if ($row.Contexts) { ($row.Contexts | ConvertTo-Json -Depth 10 -Compress) } else { $null }
                Attachments = if ($row.Attachments) { ($row.Attachments | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                Mentions = if ($row.Mentions) { ($row.Mentions | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                Links = if ($row.Links) { ($row.Links | ConvertTo-Json -Depth 20 -Compress) } else { $null }
                RawJson = if ($IncludeRawJson) { ($row | ConvertTo-Json -Depth 100 -Compress) } else { $null }
            }
        }
    }

    return $normalized
}

<#
    Main script flow:
    1. Make sure Graph modules are installed.
    2. Connect to Microsoft Graph.
    3. Confirm required permissions are present.
    4. Get the selected users.
    5. Pull Microsoft 365 Copilot interaction history for each user.
    6. Optionally import Copilot Studio or Agent Builder export data.
    7. Sort the records and save them to CSV and JSON files.
#>
Ensure-GraphModules

if (-not (Get-MgContext -ErrorAction SilentlyContinue)) {
    $graphScopes = @('AiEnterpriseInteraction.Read.All', 'User.Read.All')
    if ($TenantId) {
        Write-Host "Connecting to Microsoft Graph for tenant $TenantId using required scopes: $($graphScopes -join ', ')" -ForegroundColor Cyan
        Connect-MgGraph -TenantId $TenantId -Scopes $graphScopes
    }
    else {
        Write-Host "Connecting to Microsoft Graph using required scopes: $($graphScopes -join ', ')" -ForegroundColor Cyan
        Connect-MgGraph -Scopes $graphScopes
    }
}

if (-not (Test-GraphPermissions)) {
    throw "Graph permission check failed. Grant the needed application or delegated permissions and reconnect."
}

$users = Get-UsersToProcess -ExplicitUsers $UserIds -SkipEnumeration:$SkipUserEnumeration
if (-not $users) {
    throw "No users were found to process. Check the tenant or specify -UserIds."
}

$allInteractions = @()
foreach ($user in $users) {
    $userLabel = if ($user.UserPrincipalName) { $user.UserPrincipalName } elseif ($user.DisplayName) { $user.DisplayName } else { $user.Id }
    Write-Verbose "Processing user: $userLabel"
    $history = Get-AIInteractionHistoryForUser -UserId $user.Id -PageSize $Top -AppClassFilter $AppClassFilter -UseBeta:$UseBeta

    foreach ($record in $history) {
        $record.UserPrincipalName = $user.UserPrincipalName
        $record.DisplayName = $user.DisplayName
        $allInteractions += $record
    }
}

if ($IncludeCopilotStudio) {
    if (-not $CopilotStudioExportPath) {
        Write-Warning "IncludeCopilotStudio was requested, but no CopilotStudioExportPath was provided. The script will merge any Copilot Studio/Agent Builder exports from the folder you supply."
    }
    else {
        $copilotStudioInteractions = Import-CopilotStudioExports -Path $CopilotStudioExportPath
        if ($copilotStudioInteractions.Count -gt 0) {
            $allInteractions += $copilotStudioInteractions
        }
        else {
            Write-Warning "No Copilot Studio or Agent Builder interactions were imported from '$CopilotStudioExportPath'."
        }
    }
}

if ($allInteractions.Count -eq 0) {
    Write-Host "No AI interaction history records were returned for the selected users or exports." -ForegroundColor Yellow
    return
}

$sortedInteractions = $allInteractions |
    Sort-Object @{Expression = { $_.InteractionType } }, @{Expression = { $_.AppClass } }, @{Expression = { $_.CreatedDateTime } }

$baseOutputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path (Get-Location) $OutputDirectory }
New-Item -ItemType Directory -Force -Path $baseOutputPath | Out-Null

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvPath = Join-Path $baseOutputPath "AIInteractions_$timestamp.csv"
$jsonPath = Join-Path $baseOutputPath "AIInteractions_$timestamp.json"

<#
    Export phase:
    The script saves both a CSV and a JSON version so you can use Excel or write custom PowerShell/Python logic later.
#>
$sortedInteractions | Export-Csv -Path $csvPath -NoTypeInformation
$sortedInteractions | ConvertTo-Json -Depth 100 | Set-Content -Path $jsonPath -Encoding UTF8

$summary = $sortedInteractions |
    Group-Object InteractionType |
    Sort-Object Name |
    Select-Object @{Name='InteractionType'; Expression={ $_.Name } }, @{Name='Count'; Expression={ $_.Count } }

Write-Host "Export complete." -ForegroundColor Green
Write-Host "CSV: $csvPath" -ForegroundColor Cyan
Write-Host "JSON: $jsonPath" -ForegroundColor Cyan
Write-Host "Interaction counts by type:" -ForegroundColor Cyan
$summary | Format-Table -AutoSize

Write-Host "`nNote: `getAllEnterpriseInteractions` does not return Copilot Studio or Agent Builder interactions. If you need those included, use -IncludeCopilotStudio with a CSV/JSON export path from the Copilot Studio or Agent Builder environment." -ForegroundColor Yellow
