[CmdletBinding()]
param(
    [string]$TenantId = $env:AZURE_TENANT_ID,
    [string]$ClientId = $env:AZURE_CLIENT_ID,
    [string]$ClientSecret = $env:AZURE_CLIENT_SECRET,
    [string]$AccessToken,
    [string[]]$UserIds,
    [switch]$AllUsers,
    [string]$OutputDirectory = '.',
    [ValidateRange(1, 100)]
    [int]$Top = 100,
    [switch]$UseBeta,
    [string]$AppClassFilter,
    [datetime]$CreatedDateTimeStart,
    [datetime]$CreatedDateTimeEnd,
    [switch]$IncludeRawJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-GraphErrorMessage {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $message = $ErrorRecord.Exception.Message
    $response = $ErrorRecord.Exception.Response
    if ($response) {
        try {
            $reader = [System.IO.StreamReader]::new($response.GetResponseStream())
            $body = $reader.ReadToEnd()
            $reader.Dispose()
            $detail = $body | ConvertFrom-Json -ErrorAction Stop
            if ($detail.error.message) {
                $message = "$($response.StatusCode) $($response.StatusDescription): $($detail.error.message)"
            }
        }
        catch {
            $message = "$($response.StatusCode) $($response.StatusDescription): $message"
        }
    }

    return $message
}

function Invoke-GraphGet {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    try {
        return Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers -ErrorAction Stop
    }
    catch {
        throw "Microsoft Graph request failed: $(Get-GraphErrorMessage -ErrorRecord $_)"
    }
}

function Get-AccessToken {
    param(
        [string]$SuppliedAccessToken,
        [string]$DirectoryTenantId,
        [string]$ApplicationId,
        [string]$ApplicationSecret
    )

    if ($SuppliedAccessToken) {
        return $SuppliedAccessToken
    }

    if (-not $DirectoryTenantId -or -not $ApplicationId -or -not $ApplicationSecret) {
        throw 'Provide -AccessToken, or provide TenantId, ClientId, and ClientSecret parameters (or their AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables).'
    }

    try {
        $tokenResponse = Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$([System.Uri]::EscapeDataString($DirectoryTenantId))/oauth2/v2.0/token" `
            -ContentType 'application/x-www-form-urlencoded' `
            -Body @{
                client_id     = $ApplicationId
                client_secret = $ApplicationSecret
                scope         = 'https://graph.microsoft.com/.default'
                grant_type    = 'client_credentials'
            } `
            -ErrorAction Stop
    }
    catch {
        throw "Microsoft Entra token request failed: $(Get-GraphErrorMessage -ErrorRecord $_)"
    }

    if (-not $tokenResponse.access_token) {
        throw 'Microsoft Entra token request did not return an access token.'
    }

    return $tokenResponse.access_token
}

function Get-GraphUser {
    param([hashtable]$Headers)

    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName&$top=999'
    do {
        $response = Invoke-GraphGet -Uri $uri -Headers $Headers
        foreach ($user in @($response.value)) {
            [pscustomobject]@{
                Id                = $user.id
                UserPrincipalName = $user.userPrincipalName
                DisplayName       = $user.displayName
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)
}

function Get-AIInteractionHistoryForUser {
    param(
        [pscustomobject]$User,
        [hashtable]$Headers,
        [string]$Version,
        [int]$PageSize,
        [string]$Filter,
        [switch]$IncludeRawJson
    )

    $encodedUserId = [System.Uri]::EscapeDataString($User.Id)
    $uri = "https://graph.microsoft.com/$Version/copilot/users/$encodedUserId/interactionHistory/getAllEnterpriseInteractions"
    $query = @("`$top=$PageSize")
    if ($Filter) {
        $query += "`$filter=$([System.Uri]::EscapeDataString($Filter))"
    }
    $uri = "$uri?$($query -join '&')"

    do {
        $response = Invoke-GraphGet -Uri $uri -Headers $Headers
        foreach ($item in @($response.value)) {
            $content = if ($item.body) { $item.body.content } else { $null }
            $contentType = if ($item.body) { $item.body.contentType } else { $null }
            $sourceApplication = if ($item.from -and $item.from.application) { $item.from.application.displayName } else { $null }
            [pscustomobject]@{
                UserId                = $User.Id
                UserPrincipalName     = $User.UserPrincipalName
                DisplayName           = $User.DisplayName
                InteractionId         = $item.id
                SessionId             = $item.sessionId
                RequestId             = $item.requestId
                AppClass              = $item.appClass
                InteractionType       = $item.interactionType
                ConversationType      = $item.conversationType
                CreatedDateTime       = $item.createdDateTime
                Locale                = $item.locale
                SourceApplication     = $sourceApplication
                ContentType           = $contentType
                Content               = $content
                Contexts              = if ($item.contexts) { $item.contexts | ConvertTo-Json -Depth 10 -Compress } else { $null }
                Attachments           = if ($item.attachments) { $item.attachments | ConvertTo-Json -Depth 20 -Compress } else { $null }
                Mentions              = if ($item.mentions) { $item.mentions | ConvertTo-Json -Depth 20 -Compress } else { $null }
                Links                 = if ($item.links) { $item.links | ConvertTo-Json -Depth 20 -Compress } else { $null }
                RawJson               = if ($IncludeRawJson) { $item | ConvertTo-Json -Depth 100 -Compress } else { $null }
            }
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)
}

if ($CreatedDateTimeStart -and -not $CreatedDateTimeEnd) {
    throw 'CreatedDateTimeStart requires CreatedDateTimeEnd because Graph requires both createdDateTime filter boundaries.'
}
if ($CreatedDateTimeEnd -and -not $CreatedDateTimeStart) {
    throw 'CreatedDateTimeEnd requires CreatedDateTimeStart because Graph requires both createdDateTime filter boundaries.'
}
if ($CreatedDateTimeStart -gt $CreatedDateTimeEnd) {
    throw 'CreatedDateTimeStart must be earlier than or equal to CreatedDateTimeEnd.'
}
if ($UserIds -and $AllUsers) {
    throw 'Specify either UserIds or AllUsers, not both.'
}
if (-not $UserIds -and -not $AllUsers) {
    throw 'Specify one or more UserIds, or use -AllUsers to enumerate users.'
}

$accessToken = Get-AccessToken -SuppliedAccessToken $AccessToken -DirectoryTenantId $TenantId -ApplicationId $ClientId -ApplicationSecret $ClientSecret
$headers = @{ Authorization = ('Bearer ' + $accessToken) }
$version = if ($UseBeta) { 'beta' } else { 'v1.0' }

$filterParts = @()
if ($AppClassFilter) {
    $escapedAppClass = $AppClassFilter.Replace("'", "''")
    $filterParts += "appClass eq '$escapedAppClass'"
}
if ($CreatedDateTimeStart) {
    $start = $CreatedDateTimeStart.ToUniversalTime().ToString('o')
    $end = $CreatedDateTimeEnd.ToUniversalTime().ToString('o')
    $filterParts += "createdDateTime ge $start and createdDateTime le $end"
}
$filter = $filterParts -join ' and '

$users = if ($AllUsers) {
    Write-Information 'Enumerating users through Microsoft Graph...' -InformationAction Continue
    @(Get-GraphUser -Headers $headers)
}
else {
    @($UserIds | ForEach-Object {
        [pscustomobject]@{
            Id                = $_
            UserPrincipalName = $null
            DisplayName       = $null
        }
    })
}

$allInteractions = @()
foreach ($user in $users) {
    Write-Verbose "Processing user: $($user.Id)"
    $allInteractions += @(Get-AIInteractionHistoryForUser -User $user -Headers $headers -Version $version -PageSize $Top -Filter $filter -IncludeRawJson:$IncludeRawJson)
}

if ($allInteractions.Count -eq 0) {
    Write-Information 'No AI interaction history records were returned.' -InformationAction Continue
    return
}

$outputPath = if ([System.IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory } else { Join-Path (Get-Location) $OutputDirectory }
New-Item -ItemType Directory -Force -Path $outputPath | Out-Null
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$sortedInteractions = $allInteractions | Sort-Object InteractionType, AppClass, CreatedDateTime
$csvPath = Join-Path $outputPath "AIInteractions_$timestamp.csv"
$jsonPath = Join-Path $outputPath "AIInteractions_$timestamp.json"
$sortedInteractions | Export-Csv -Path $csvPath -NoTypeInformation
$sortedInteractions | ConvertTo-Json -Depth 100 | Set-Content -Path $jsonPath -Encoding utf8

Write-Information 'Export complete.' -InformationAction Continue
Write-Information "CSV: $csvPath" -InformationAction Continue
Write-Information "JSON: $jsonPath" -InformationAction Continue
