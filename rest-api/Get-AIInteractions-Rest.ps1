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
    [switch]$IncludeCopilotStudio,
    [string]$CopilotStudioEnvironmentUrl = $env:DATAVERSE_ENVIRONMENT_URL,
    [string]$CopilotStudioAccessToken = $env:DATAVERSE_ACCESS_TOKEN,
    [string]$CopilotStudioFilter,
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
        [string]$ApplicationSecret,
        [string]$Scope = 'https://graph.microsoft.com/.default'
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
                scope         = $Scope
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

function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [string[]]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($name in $PropertyName) {
        if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($name)) {
            return $InputObject[$name]
        }

        $property = $InputObject.PSObject.Properties[$name]
        if ($property) {
            return $property.Value
        }
    }

    return $null
}

function ConvertTo-CompressedJsonOrNull {
    param(
        [object]$InputObject,
        [int]$Depth = 20
    )

    if ($null -eq $InputObject) {
        return $null
    }

    return $InputObject | ConvertTo-Json -Depth $Depth -Compress
}

function ConvertFrom-JsonOrNull {
    param([object]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    try {
        return $Value | ConvertFrom-Json -Depth 100 -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-GraphUser {
    param([hashtable]$Headers)

    $uri = 'https://graph.microsoft.com/v1.0/users?$select=id,userPrincipalName,displayName&$top=999'
    do {
        $response = Invoke-GraphGet -Uri $uri -Headers $Headers
        if ($response.value) {
            foreach ($user in $response.value) {
                [pscustomobject]@{
                    Id                = $user.id
                    UserPrincipalName = $user.userPrincipalName
                    DisplayName       = $user.displayName
                }
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
        if ($response.value) {
            foreach ($item in $response.value) {
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
        }
        $uri = $response.'@odata.nextLink'
    } while ($uri)
}

function Convert-CopilotStudioTranscriptToInteraction {
    param(
        [object]$Transcript,
        [switch]$IncludeRawJson
    )

    $transcriptId = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('conversationtranscriptid', 'id')
    $createdOn = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('createdon', 'createdDateTime', 'created')
    $modifiedOn = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('modifiedon', 'modifiedDateTime', 'modified')
    $content = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('content', 'transcript')
    $metadata = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('metadata')
    $name = Get-ObjectPropertyValue -InputObject $Transcript -PropertyName @('name')
    $parsedContent = ConvertFrom-JsonOrNull -Value $content
    $parsedMetadata = ConvertFrom-JsonOrNull -Value $metadata
    $metadataForExport = if ($parsedMetadata) { $parsedMetadata } else { $metadata }

    $activities = @()
    if ($parsedContent) {
        $directActivities = Get-ObjectPropertyValue -InputObject $parsedContent -PropertyName @('activities')
        $transcriptActivities = Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $parsedContent -PropertyName @('transcript')) -PropertyName @('activities')
        $valueActivities = Get-ObjectPropertyValue -InputObject $parsedContent -PropertyName @('value')

        if ($directActivities) {
            $activities = @($directActivities)
        }
        elseif ($transcriptActivities) {
            $activities = @($transcriptActivities)
        }
        elseif ($valueActivities) {
            $activities = @($valueActivities)
        }
        elseif ($parsedContent -is [System.Collections.IEnumerable] -and $parsedContent -isnot [string]) {
            $activities = @($parsedContent)
        }
    }

    if ($activities.Count -gt 0) {
        foreach ($activity in $activities) {
            $activityType = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('type')
            $text = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('text', 'message', 'content')
            if ($activityType -and $activityType -ne 'message' -and -not $text) {
                continue
            }

            $from = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('from')
            $conversation = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('conversation')
            $activityId = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('id')
            $fromId = Get-ObjectPropertyValue -InputObject $from -PropertyName @('id')
            $fromName = Get-ObjectPropertyValue -InputObject $from -PropertyName @('name')
            $conversationId = Get-ObjectPropertyValue -InputObject $conversation -PropertyName @('id')
            $timestamp = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('timestamp', 'createdDateTime', 'created')
            $attachments = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('attachments')
            $locale = Get-ObjectPropertyValue -InputObject $activity -PropertyName @('locale')

            [pscustomobject]@{
                UserId                = $fromId
                UserPrincipalName     = $null
                DisplayName           = $fromName
                InteractionId         = if ($activityId) { $activityId } elseif ($transcriptId) { $transcriptId } else { [guid]::NewGuid().ToString() }
                SessionId             = if ($conversationId) { $conversationId } else { $transcriptId }
                RequestId             = $null
                AppClass              = 'Copilot Studio'
                InteractionType       = if ($activityType) { "copilotstudio.$activityType" } else { 'copilotstudio' }
                ConversationType      = 'agent'
                CreatedDateTime       = if ($timestamp) { $timestamp } elseif ($createdOn) { $createdOn } else { $modifiedOn }
                Locale                = $locale
                SourceApplication     = if ($name) { $name } else { 'Copilot Studio' }
                ContentType           = 'text'
                Content               = $text
                Contexts              = ConvertTo-CompressedJsonOrNull -InputObject $metadataForExport -Depth 20
                Attachments           = ConvertTo-CompressedJsonOrNull -InputObject $attachments -Depth 20
                Mentions              = $null
                Links                 = $null
                RawJson               = if ($IncludeRawJson) { $activity | ConvertTo-Json -Depth 100 -Compress } else { $null }
            }
        }

        return
    }

    [pscustomobject]@{
        UserId                = $null
        UserPrincipalName     = $null
        DisplayName           = $null
        InteractionId         = if ($transcriptId) { $transcriptId } else { [guid]::NewGuid().ToString() }
        SessionId             = $transcriptId
        RequestId             = $null
        AppClass              = 'Copilot Studio'
        InteractionType       = 'copilotstudio.transcript'
        ConversationType      = 'agent'
        CreatedDateTime       = if ($createdOn) { $createdOn } else { $modifiedOn }
        Locale                = $null
        SourceApplication     = if ($name) { $name } else { 'Copilot Studio' }
        ContentType           = if ($parsedContent) { 'application/json' } else { 'text' }
        Content               = $content
        Contexts              = ConvertTo-CompressedJsonOrNull -InputObject $metadataForExport -Depth 20
        Attachments           = $null
        Mentions              = $null
        Links                 = $null
        RawJson               = if ($IncludeRawJson) { $Transcript | ConvertTo-Json -Depth 100 -Compress } else { $null }
    }
}

function Get-CopilotStudioInteractionHistory {
    param(
        [string]$EnvironmentUrl,
        [hashtable]$Headers,
        [int]$PageSize,
        [string]$Filter,
        [datetime]$CreatedDateTimeStart,
        [datetime]$CreatedDateTimeEnd,
        [switch]$IncludeRawJson
    )

    $normalizedEnvironmentUrl = $EnvironmentUrl.TrimEnd('/')
    $uri = "$normalizedEnvironmentUrl/api/data/v9.2/conversationtranscripts"
    $query = @(
        "`$select=conversationtranscriptid,createdon,modifiedon,name,content,metadata",
        "`$orderby=createdon asc",
        "`$top=$PageSize"
    )
    $filterParts = @()
    if ($Filter) {
        $filterParts += $Filter
    }
    if ($CreatedDateTimeStart) {
        $start = $CreatedDateTimeStart.ToUniversalTime().ToString('o')
        $end = $CreatedDateTimeEnd.ToUniversalTime().ToString('o')
        $filterParts += "createdon ge $start and createdon le $end"
    }
    if ($filterParts.Count -gt 0) {
        $query += "`$filter=$([System.Uri]::EscapeDataString(($filterParts -join ' and ')))"
    }
    $uri = "$uri?$($query -join '&')"

    do {
        try {
            $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $Headers -ErrorAction Stop
        }
        catch {
            throw "Dataverse request for Copilot Studio transcripts failed: $(Get-GraphErrorMessage -ErrorRecord $_)"
        }

        if ($response.value) {
            foreach ($transcript in $response.value) {
                Convert-CopilotStudioTranscriptToInteraction -Transcript $transcript -IncludeRawJson:$IncludeRawJson
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
if (-not $UserIds -and -not $AllUsers -and -not $IncludeCopilotStudio) {
    throw 'Specify one or more UserIds, use -AllUsers to enumerate users, or use -IncludeCopilotStudio with -CopilotStudioEnvironmentUrl.'
}
if ($IncludeCopilotStudio -and -not $CopilotStudioEnvironmentUrl) {
    throw 'IncludeCopilotStudio requires CopilotStudioEnvironmentUrl or the DATAVERSE_ENVIRONMENT_URL environment variable.'
}

$queryMicrosoft365Copilot = $UserIds -or $AllUsers
$headers = $null
if ($queryMicrosoft365Copilot) {
    $accessToken = Get-AccessToken -SuppliedAccessToken $AccessToken -DirectoryTenantId $TenantId -ApplicationId $ClientId -ApplicationSecret $ClientSecret
    $headers = @{ Authorization = ('Bearer ' + $accessToken) }
}
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

$allInteractions = @()
if ($queryMicrosoft365Copilot) {
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

    foreach ($user in $users) {
        Write-Verbose "Processing user: $($user.Id)"
        $allInteractions += @(Get-AIInteractionHistoryForUser -User $user -Headers $headers -Version $version -PageSize $Top -Filter $filter -IncludeRawJson:$IncludeRawJson)
    }
}

if ($IncludeCopilotStudio) {
    Write-Information 'Retrieving Copilot Studio transcripts from Dataverse...' -InformationAction Continue
    $copilotStudioScope = "$($CopilotStudioEnvironmentUrl.TrimEnd('/'))/.default"
    $copilotStudioToken = Get-AccessToken -SuppliedAccessToken $CopilotStudioAccessToken -DirectoryTenantId $TenantId -ApplicationId $ClientId -ApplicationSecret $ClientSecret -Scope $copilotStudioScope
    $copilotStudioHeaders = @{
        Authorization      = ('Bearer ' + $copilotStudioToken)
        Accept             = 'application/json'
        'OData-MaxVersion' = '4.0'
        'OData-Version'    = '4.0'
        Prefer             = "odata.maxpagesize=$Top"
    }
    $allInteractions += @(Get-CopilotStudioInteractionHistory -EnvironmentUrl $CopilotStudioEnvironmentUrl -Headers $copilotStudioHeaders -PageSize $Top -Filter $CopilotStudioFilter -CreatedDateTimeStart $CreatedDateTimeStart -CreatedDateTimeEnd $CreatedDateTimeEnd -IncludeRawJson:$IncludeRawJson)
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
