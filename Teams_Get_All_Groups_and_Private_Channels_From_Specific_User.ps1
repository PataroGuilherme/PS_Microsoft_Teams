##############
## Command to run in CMD: powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\guilherme.pataro\Downloads\Get-UserTeamsPrivateChannels.ps1" -UserPrincipalName "guilherme.pataro@company.com"


<#
.SYNOPSIS
    Lists every Microsoft Teams team a user belongs to, plus each team's private channels,
    and exports everything to a structured Excel workbook.

.DESCRIPTION
    Performance design:
      * Only the lightweight Microsoft.Graph.Authentication module is used (raw REST calls).
      * Teams are discovered with ONE paged call (memberOf, $select, $top=999).
      * Private channels are fetched with Microsoft Graph JSON batching (20 requests per call)
        and a server-side $filter, instead of one call per team.
      * Throttling (HTTP 429/503/504) is handled with Retry-After based retries.
      * The Excel file is written with ImportExcel (no Excel installation required).

.PARAMETER UserPrincipalName
    UPN (or object ID) of the user to inspect.

.PARAMETER OutputPath
    Full path of the .xlsx file to create. Defaults to the current folder.

.PARAMETER TenantId / ClientId / CertificateThumbprint
    Optional. Provide all three to authenticate app-only (recommended, see NOTES).

.EXAMPLE
    .\Get-UserTeamsPrivateChannels.ps1 -UserPrincipalName john.doe@contoso.com

.EXAMPLE
    .\Get-UserTeamsPrivateChannels.ps1 -UserPrincipalName john.doe@contoso.com `
        -TenantId <tenant-guid> -ClientId <app-guid> -CertificateThumbprint <thumbprint>

.NOTES
    Required modules : Microsoft.Graph.Authentication, ImportExcel
        Install-Module Microsoft.Graph.Authentication, ImportExcel -Scope CurrentUser

    Permissions (delegated or application, admin consent required):
        User.Read.All, Group.Read.All, GroupMember.Read.All, Channel.ReadBasic.All

    IMPORTANT: with DELEGATED auth, Graph only returns private channels that the signed-in
    account is itself a member of. To list ALL private channels of every team, use
    APPLICATION permissions (app-only authentication with a certificate).
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$UserPrincipalName,

    [string]$OutputPath,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'
$graphRoot = 'https://graph.microsoft.com/v1.0'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

#region Helper functions

function Get-GraphPaged {
    # Follows @odata.nextLink and returns all items of a collection.
    param([Parameter(Mandatory = $true)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -Headers @{ ConsistencyLevel = 'eventual' }
        if ($response.value) { $items.AddRange([object[]]$response.value) }
        $next = $response.'@odata.nextLink'
    }
    return $items
}

function Invoke-GraphBatch {
    # Sends requests in JSON batches of 20 and retries throttled sub-requests.
    # Returns a hashtable: request id (string) -> sub-response.
    param([Parameter(Mandatory = $true)][System.Collections.IEnumerable]$Requests)

    $results = @{}
    $pending = New-Object System.Collections.Generic.List[object]
    foreach ($r in $Requests) { [void]$pending.Add($r) }
    $attempt = 0

    while ($pending.Count -gt 0 -and $attempt -lt 6) {
        $retry = New-Object System.Collections.Generic.List[object]
        $maxWait = 0

        for ($i = 0; $i -lt $pending.Count; $i += 20) {
            $count = [Math]::Min(20, $pending.Count - $i)
            $chunk = New-Object System.Collections.Generic.List[object]
            for ($k = 0; $k -lt $count; $k++) { [void]$chunk.Add($pending[$i + $k]) }

            $payload = [ordered]@{ requests = $chunk.ToArray() }
            $body = ConvertTo-Json -InputObject $payload -Depth 6 -Compress

            $batchResponse = Invoke-MgGraphRequest -Method POST -Uri "$graphRoot/`$batch" `
                -Body $body -ContentType 'application/json'

            foreach ($sub in @($batchResponse.responses)) {
                $statusCode = 0
                [void][int]::TryParse([string]$sub.status, [ref]$statusCode)
                $subId = [string]$sub.id

                if ($statusCode -eq 429 -or $statusCode -eq 503 -or $statusCode -eq 504) {
                    foreach ($candidate in $chunk) {
                        if ([string]$candidate.id -eq $subId) { [void]$retry.Add($candidate); break }
                    }
                    $retryAfter = $null
                    if ($sub.headers) { $retryAfter = $sub.headers['Retry-After'] }
                    if ($retryAfter) {
                        $parsed = 0
                        if ([int]::TryParse([string]$retryAfter, [ref]$parsed) -and $parsed -gt $maxWait) { $maxWait = $parsed }
                    }
                }
                else {
                    $results[$subId] = [pscustomobject]@{ id = $subId; status = $statusCode; body = $sub.body }
                }
            }
        }

        $pending = $retry
        if ($pending.Count -gt 0) {
            $attempt++
            $sleep = [Math]::Max($maxWait, [int][Math]::Pow(2, $attempt))
            Write-Warning "Throttled on $($pending.Count) request(s). Retrying in $sleep second(s) (attempt $attempt)."
            Start-Sleep -Seconds $sleep
        }
    }

    foreach ($left in $pending) {
        $results[[string]$left.id] = [pscustomobject]@{ id = [string]$left.id; status = 429; body = $null }
    }
    return $results
}

#endregion

#region Prerequisites and authentication

foreach ($moduleName in 'Microsoft.Graph.Authentication', 'ImportExcel') {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw "Required module '$moduleName' is not installed. Run: Install-Module $moduleName -Scope CurrentUser"
    }
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

if ($TenantId -and $ClientId -and $CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    $authMode = 'Application (app-only)'
}
else {
    Connect-MgGraph -Scopes 'User.Read.All', 'Group.Read.All', 'GroupMember.Read.All', 'Channel.ReadBasic.All' -NoWelcome
    $authMode = 'Delegated'
    Write-Warning 'Delegated authentication: only private channels where the signed-in account is a member will be returned. Use app-only authentication for a complete report.'
}

#endregion

try {
    #region Step 1 - Resolve the user
    Write-Host "Resolving user '$UserPrincipalName'..." -ForegroundColor Cyan
    $encodedUser = [uri]::EscapeDataString($UserPrincipalName)
    $user = Invoke-MgGraphRequest -Method GET -Uri "$graphRoot/users/$encodedUser`?`$select=id,displayName,userPrincipalName"
    #endregion

    #region Step 2 - Get all Teams-enabled groups of the user (single paged call)
    Write-Host 'Retrieving Teams memberships...' -ForegroundColor Cyan
    $groups = Get-GraphPaged -Uri "$graphRoot/users/$($user.id)/memberOf/microsoft.graph.group?`$select=id,displayName,resourceProvisioningOptions&`$top=999"

    # Graph returns hashtables, so build real objects first (Sort-Object cannot read hashtable keys).
    # De-duplicate by Object ID (never by name: different teams can share the same display name).
    $teamMap = @{}
    foreach ($group in $groups) {
        $provisioning = @($group['resourceProvisioningOptions'])
        if ($provisioning -contains 'Team') {
            $teamId = [string]$group['id']
            if (-not $teamMap.ContainsKey($teamId)) {
                $teamMap[$teamId] = [pscustomobject]@{ Id = $teamId; Name = [string]$group['displayName'] }
            }
        }
    }
    $teams = @($teamMap.Values | Sort-Object -Property Name)
    Write-Host "Total groups returned: $(@($groups).Count)" -ForegroundColor DarkGray
    Write-Host "Teams found: $($teams.Count)" -ForegroundColor Green
    #endregion

    #region Step 3 - Get private channels of every team using Graph batching
    $channelsByTeam = @{}
    $failedTeams = New-Object System.Collections.Generic.List[object]

    if ($teams.Count -gt 0) {
        Write-Host 'Retrieving private channels (batched)...' -ForegroundColor Cyan
        $requests = for ($i = 0; $i -lt $teams.Count; $i++) {
            @{
                id     = [string]$i
                method = 'GET'
                url    = "/teams/$($teams[$i].Id)/channels?`$filter=membershipType%20eq%20'private'&`$select=id,displayName,description,membershipType,createdDateTime,isArchived"
            }
        }

        try {
            $batchResults = Invoke-GraphBatch -Requests @($requests)
        }
        catch {
            Write-Warning "Batch request failed, falling back to one request per team (slower). Details: $($_.Exception.GetType().FullName): $($_.Exception.Message)"
            Write-Warning ($_.ScriptStackTrace)
            $batchResults = @{}
            for ($i = 0; $i -lt $teams.Count; $i++) {
                try {
                    $single = Get-GraphPaged -Uri "$graphRoot/teams/$($teams[$i].Id)/channels?`$filter=membershipType%20eq%20'private'&`$select=id,displayName,description,membershipType,createdDateTime,isArchived"
                    $batchResults[[string]$i] = [pscustomobject]@{ id = [string]$i; status = 200; body = [pscustomobject]@{ value = @($single) } }
                }
                catch {
                    $code = 0
                    if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
                    $batchResults[[string]$i] = [pscustomobject]@{ id = [string]$i; status = $code; body = $null }
                }
            }
        }

        for ($i = 0; $i -lt $teams.Count; $i++) {
            $sub = $batchResults[[string]$i]
            if ($sub -and $sub.status -eq 200) {
                $channelsByTeam[$teams[$i].Id] = @($sub.body.value)
            }
            else {
                $status = if ($sub) { $sub.status } else { 'n/a' }
                $failedTeams.Add([pscustomobject]@{ TeamName = $teams[$i].Name; TeamObjectId = $teams[$i].Id; HttpStatus = $status })
                $channelsByTeam[$teams[$i].Id] = @()
            }
        }
    }
    #endregion

    #region Step 4 - Build report data
    $teamRows = New-Object System.Collections.Generic.List[object]
    $channelRows = New-Object System.Collections.Generic.List[object]
    $fullRows = New-Object System.Collections.Generic.List[object]

    foreach ($team in $teams) {
        $channels = @($channelsByTeam[$team.Id] | Where-Object { $_ })

        $teamRows.Add([pscustomobject][ordered]@{
                TeamName            = $team.Name
                TeamObjectId        = $team.Id
                PrivateChannelCount = $channels.Count
            })

        if ($channels.Count -eq 0) {
            $fullRows.Add([pscustomobject][ordered]@{
                    TeamName            = $team.Name
                    TeamObjectId        = $team.Id
                    PrivateChannelName  = $null
                    PrivateChannelId    = $null
                    IsArchived          = $null
                    CreatedDateTimeUtc  = $null
                })
            continue
        }

        $sortedChannels = @(
            $channels | ForEach-Object {
                [pscustomobject]@{
                    Id              = [string]$_['id']
                    DisplayName     = [string]$_['displayName']
                    Description     = [string]$_['description']
                    IsArchived      = [bool]$_['isArchived']
                    CreatedDateTime = $_['createdDateTime']
                }
            } | Sort-Object -Property DisplayName
        )

        foreach ($channel in $sortedChannels) {
            $created = $null
            if ($channel.createdDateTime) { $created = ([datetime]$channel.createdDateTime).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }

            $channelRows.Add([pscustomobject][ordered]@{
                    TeamName            = $team.Name
                    TeamObjectId        = $team.Id
                    PrivateChannelName  = $channel.displayName
                    PrivateChannelId    = $channel.id
                    IsArchived          = [bool]$channel.isArchived
                    CreatedDateTimeUtc  = $created
                    Description         = $channel.description
                })

            $fullRows.Add([pscustomobject][ordered]@{
                    TeamName            = $team.Name
                    TeamObjectId        = $team.Id
                    PrivateChannelName  = $channel.displayName
                    PrivateChannelId    = $channel.id
                    IsArchived          = [bool]$channel.isArchived
                    CreatedDateTimeUtc  = $created
                })
        }
    }

    $summaryRows = @(
        [pscustomobject]@{ Property = 'User display name';           Value = $user.displayName }
        [pscustomobject]@{ Property = 'User principal name';         Value = $user.userPrincipalName }
        [pscustomobject]@{ Property = 'User object ID';              Value = $user.id }
        [pscustomobject]@{ Property = 'Total teams';                 Value = $teams.Count }
        [pscustomobject]@{ Property = 'Teams with private channels'; Value = @($teamRows | Where-Object { $_.PrivateChannelCount -gt 0 }).Count }
        [pscustomobject]@{ Property = 'Total private channels';      Value = $channelRows.Count }
        [pscustomobject]@{ Property = 'Teams that failed to query';  Value = $failedTeams.Count }
        [pscustomobject]@{ Property = 'Authentication mode';         Value = $authMode }
        [pscustomobject]@{ Property = 'Report generated (UTC)';      Value = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
    )
    #endregion

    #region Step 5 - Export to Excel
    if (-not $OutputPath) {
        $safeName = ($user.userPrincipalName -replace '[^a-zA-Z0-9\.\-_@]', '_')
        $OutputPath = Join-Path (Get-Location) ("TeamsPrivateChannelsReport_{0}_{1}.xlsx" -f $safeName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    }
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }

    $common = @{
        Path          = $OutputPath
        AutoSize      = $true
        FreezeTopRow  = $true
        BoldTopRow    = $true
        TableStyle    = 'Medium2'
    }

    $summaryRows | Export-Excel @common -WorksheetName 'Summary' -TableName 'tblSummary' -Append:$false
    if ($teamRows.Count -gt 0) { $teamRows    | Export-Excel @common -WorksheetName 'Teams'           -TableName 'tblTeams'           -Append }
    if ($channelRows.Count -gt 0) { $channelRows | Export-Excel @common -WorksheetName 'PrivateChannels' -TableName 'tblPrivateChannels' -Append }
    if ($fullRows.Count -gt 0) { $fullRows    | Export-Excel @common -WorksheetName 'FullReport'      -TableName 'tblFullReport'      -Append }
    if ($failedTeams.Count -gt 0) { $failedTeams | Export-Excel @common -WorksheetName 'Errors'         -TableName 'tblErrors'          -Append }
    #endregion

    $stopwatch.Stop()
    Write-Host ''
    Write-Host "Teams: $($teams.Count) | Private channels: $($channelRows.Count) | Errors: $($failedTeams.Count)" -ForegroundColor Green
    Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
    Write-Host ("Elapsed time: {0:N1} seconds" -f $stopwatch.Elapsed.TotalSeconds) -ForegroundColor Green
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
