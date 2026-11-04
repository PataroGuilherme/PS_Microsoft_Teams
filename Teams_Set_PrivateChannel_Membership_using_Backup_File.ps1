###########################################################
### CMD Command (Test Mode): powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\PrivateChannel_Membership_Manager.ps1" -InputPath ".\Book1(1).xlsx" -OutputPath ".\Test_Report.xlsx" -TestMode
### CMD Command (Prod): powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\PrivateChannel_Membership_Manager.ps1" -InputPath ".\Book1(1).xlsx" -OutputPath ".\Production_Report.xlsx"


#requires -Version 5.1
<#
.SYNOPSIS
    Adds users to Microsoft 365 parent groups and private Teams channels based on an Excel input file.

.INPUT
    Excel columns:
    Group Name, Group Object ID, Channel Name, Channel ID,
    User Display Name, User Object ID, User Email, Role

.ROLE
    Role must be Member or Owner.

.REQUIREMENTS
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Install-Module ImportExcel -Scope CurrentUser

.NOTES
    The script uses Microsoft Graph v1.0 and caches group/channel memberships
    to avoid repeated GET requests for rows sharing the same group or channel.
    Large-workload optimizations include throttling-aware retries, one console result
    line per input row, periodic progress reporting, and deferred Excel report export.
#>

[CmdletBinding()]
param(
    [string]$InputPath  = ".\Book1(1).xlsx",
    [string]$OutputPath = ".\PrivateChannel_Membership_Report.xlsx",
    [switch]$TestMode,
    [ValidateRange(10, 5000)]
    [int]$ProgressIntervalRows = 100,
    [switch]$DetailedOutput
)

$ErrorActionPreference = "Stop"
$GraphBase = "https://graph.microsoft.com/v1.0"

# --------------------------- Prerequisites ---------------------------

foreach ($moduleName in @("Microsoft.Graph.Authentication", "ImportExcel")) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        throw "Required module '$moduleName' is not installed. Run: Install-Module $moduleName -Scope CurrentUser"
    }
}

Import-Module Microsoft.Graph.Authentication
Import-Module ImportExcel

if (-not (Test-Path -LiteralPath $InputPath)) {
    throw "Input Excel file not found: $InputPath"
}

# Delegated permissions. The signed-in operator must also have sufficient
# Entra/Teams administrative permissions to perform these operations.
$requiredScopes = @(
    "Group.ReadWrite.All",
    "GroupMember.Read.All",
    "ChannelMember.ReadWrite.All"
)

if ($TestMode) {
    Write-Host "TEST MODE ENABLED: Graph membership will be read, but no changes will be submitted." -ForegroundColor Yellow
}
else {
    Write-Host "LIVE MODE: Changes will be applied to Microsoft 365." -ForegroundColor Red
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $requiredScopes -NoWelcome

# --------------------------- Helper functions ---------------------------

function Get-GraphHttpStatusCode {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $exception = $ErrorRecord.Exception
    foreach ($candidate in @(
        $(if ($null -ne $exception.Response) { $exception.Response.StatusCode }),
        $(if ($null -ne $exception.InnerException -and $null -ne $exception.InnerException.Response) { $exception.InnerException.Response.StatusCode })
    )) {
        if ($null -ne $candidate) {
            try { return [int]$candidate } catch { }
        }
    }

    $message = [string]$exception.Message
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $message += " " + $ErrorRecord.ErrorDetails.Message
    }
    if ($message -match '(?<!\d)(429|503|504|408)(?!\d)') { return [int]$Matches[1] }
    return 0
}

function Get-GraphRetryDelaySeconds {
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory)][int]$Attempt
    )

    # Honor Retry-After whenever the Graph/HTTP response exposes it.
    try {
        $headers = $ErrorRecord.Exception.Response.Headers
        if ($null -ne $headers) {
            $retryAfter = $headers['Retry-After']
            if ($retryAfter) {
                $seconds = 0
                if ([int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -ge 0) {
                    return [Math]::Min($seconds, 120)
                }
            }
        }
    } catch { }

    try {
        $delta = $ErrorRecord.Exception.Response.Headers.RetryAfter.Delta
        if ($null -ne $delta) { return [Math]::Min([Math]::Max(0, [int][Math]::Ceiling($delta.TotalSeconds)), 120) }
    } catch { }

    # Fallback exponential backoff with a cap to avoid rapid retry loops.
    return [Math]::Min([int][Math]::Pow(2, $Attempt), 60)
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [int]$MaxAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($Method -eq 'GET') {
                return Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
            }
            $json = $Body | ConvertTo-Json -Depth 8
            return Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $json -ContentType 'application/json' -OutputType PSObject
        }
        catch {
            $statusCode = Get-GraphHttpStatusCode -ErrorRecord $_
            $isTransientHttpError = $statusCode -in @(408, 429, 503, 504)

            # Retry only explicit transient HTTP responses. A transport failure with
            # an unknown outcome is not blindly retried for POST to avoid duplicates.
            if ($isTransientHttpError -and $attempt -lt $MaxAttempts) {
                $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
                if ($DetailedOutput) {
                    Write-Host "  INFO - Graph returned HTTP $statusCode. Retrying in $delay second(s) ($attempt/$MaxAttempts)..." -ForegroundColor Yellow
                }
                Start-Sleep -Seconds $delay
                continue
            }
            throw
        }
    }
}

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-GraphRequestWithRetry -Method GET -Uri $nextUri
        if ($null -ne $response.value) {
            foreach ($item in $response.value) { $items.Add($item) }
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return $items
}

function Invoke-GraphPost {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][object]$Body
    )

    Invoke-GraphRequestWithRetry -Method POST -Uri $Uri -Body $Body | Out-Null
}

function Invoke-ChannelMemberPostWithRetry {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][object]$Body,
        [int]$MaxAttempts = 4
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Invoke-GraphPost -Uri $Uri -Body $Body | Out-Null
            return
        }
        catch {
            $errorText = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $errorText += " " + $_.ErrorDetails.Message
            }

            $isRosterPropagationError = $errorText -match "UserNotFoundInTeamRoster|not part of the parent team roster"
            if ($isRosterPropagationError -and $attempt -lt $MaxAttempts) {
                if ($DetailedOutput) { Write-Host "  INFO - Team roster update is still propagating. Retrying channel addition ($attempt/$MaxAttempts)..." -ForegroundColor Yellow }
                Start-Sleep -Seconds 3
                continue
            }

            throw
        }
    }
}

function New-IdSet {
    # Prevent PowerShell from enumerating an empty HashSet into $null.
    return ,([System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase))
}

function Get-GroupCache {
    param([Parameter(Mandatory)][string]$GroupId)

    $key = $GroupId.ToLowerInvariant()
    if (-not $script:GroupCache.ContainsKey($key)) {
        if ($DetailedOutput) { Write-Host "  Reading parent group membership: $GroupId" -ForegroundColor DarkGray }

        $members = New-IdSet
        $owners  = New-IdSet
        if ($null -eq $members -or $null -eq $owners) {
            throw "Unable to initialize parent group membership cache."
        }

        # Use the Team roster as the authoritative membership source.
        # Private-channel membership requires the user to exist in this roster;
        # group membership alone may not yet be reflected in the Team roster.
        # Do not use $select=userId because the collection is typed as
        # conversationMember and userId belongs to a derived member type.
        $teamMembersUri = "$GraphBase/teams/$GroupId/members"

        foreach ($entry in (Invoke-GraphGetAll -Uri $teamMembersUri)) {
            if ($entry.userId) {
                $rosterUserId = [string]$entry.userId
                [void]$members.Add($rosterUserId)
                if (@($entry.roles) -contains "owner") {
                    [void]$owners.Add($rosterUserId)
                }
            }
        }

        $script:GroupCache[$key] = [pscustomobject]@{
            Members = $members
            Owners  = $owners
        }
    }

    return $script:GroupCache[$key]
}

function Get-ChannelCache {
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$ChannelId
    )

    $key = ($GroupId + "|" + $ChannelId).ToLowerInvariant()
    if (-not $script:ChannelCache.ContainsKey($key)) {
        if ($DetailedOutput) { Write-Host "  Reading private channel membership: $ChannelId" -ForegroundColor DarkGray }

        $userIds = New-IdSet
        if ($null -eq $userIds) {
            throw "Unable to initialize private channel membership cache."
        }
        $encodedChannelId = [System.Uri]::EscapeDataString($ChannelId)
        # Do not use $select=userId here: the collection is typed as
        # conversationMember, while userId belongs to aadUserConversationMember.
        # Graph rejects selecting that derived-type property on the base collection.
        $uri = "$GraphBase/teams/$GroupId/channels/$encodedChannelId/members"

        foreach ($entry in (Invoke-GraphGetAll -Uri $uri)) {
            if ($entry.userId) { [void]$userIds.Add([string]$entry.userId) }
        }

        $script:ChannelCache[$key] = $userIds
    }

    # Preserve the HashSet object itself; do not enumerate it into strings.
    return ,($script:ChannelCache[$key])
}

# --------------------------- Read input ---------------------------

$rows = @(Import-Excel -Path $InputPath -WorksheetName "Channel Members")
if ($rows.Count -eq 0) {
    throw "The worksheet 'Channel Members' contains no data rows."
}

$requiredColumns = @(
    "Group Name", "Group Object ID", "Channel Name", "Channel ID",
    "User Display Name", "User Object ID", "User Email", "Role"
)

$actualColumns = @($rows[0].PSObject.Properties.Name)
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })
if ($missingColumns.Count -gt 0) {
    throw "Missing required Excel columns: $($missingColumns -join ', ')"
}

$script:GroupCache = @{}
$script:ChannelCache = @{}
$report = [System.Collections.Generic.List[object]]::new()
$successCount = 0
$testCount = 0
$failedCount = 0
$skippedCount = 0
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host ""
Write-Host "Starting membership processing for $($rows.Count) row(s)..." -ForegroundColor Cyan
Write-Host "A result line will be displayed for every processed row. The Excel report is generated only after all rows finish."
Write-Host ""

# --------------------------- Main processing ---------------------------

$rowNumber = 1 # Header is row 1 in Excel

foreach ($row in $rows) {
    $rowNumber++

    $groupName = [string]$row.'Group Name'
    $groupId   = [string]$row.'Group Object ID'
    $channelName = [string]$row.'Channel Name'
    $channelId = [string]$row.'Channel ID'
    $displayName = [string]$row.'User Display Name'
    $userId    = [string]$row.'User Object ID'
    $userEmail = [string]$row.'User Email'
    $role      = ([string]$row.'Role').Trim()

    $groupStatus = "Not processed"
    $channelStatus = "Not processed"
    $overallStatus = "Skipped"
    $details = ""

    if ($DetailedOutput) { Write-Host "[$rowNumber/$($rows.Count + 1)] $displayName <$userEmail>" -ForegroundColor White }

    # Treat a partially completed row as invalid; do not guess missing values.
    $values = @($groupName, $groupId, $channelName, $channelId, $displayName, $userId, $userEmail, $role)
    if (@($values | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        $details = "Skipped: one or more required cells are blank."
        if ($DetailedOutput) { Write-Host "  SKIPPED - Missing required information." -ForegroundColor Yellow }
    }
    elseif ($role -notin @("Member", "Owner")) {
        $details = "Skipped: Role must be 'Member' or 'Owner'."
        if ($DetailedOutput) { Write-Host "  SKIPPED - Invalid Role '$role'." -ForegroundColor Yellow }
    }
    else {
        try {
            # Load/cache parent group members and owners only once per group.
            $groupState = Get-GroupCache -GroupId $groupId
            $isGroupMember = $groupState.Members.Contains($userId)
            $isGroupOwner  = $groupState.Owners.Contains($userId)

            if (-not ($isGroupMember -or $isGroupOwner)) {
                $teamRoles = [System.Collections.Generic.List[string]]::new()
                if ($role -eq "Owner") {
                    [void]$teamRoles.Add("owner")
                }

                $teamMemberBody = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = $teamRoles
                    "user@odata.bind" = "$GraphBase/users('$userId')"
                }
                $addTeamMemberUri = "$GraphBase/teams/$groupId/members"

                if ($TestMode) {
                    $groupStatus = "TEST - Would add to Team roster as $role"
                    if ($DetailedOutput) { Write-Host "  TEST - Would add to parent Team roster as $role (no change submitted)." -ForegroundColor Yellow }
                }
                else {
                    Invoke-GraphPost -Uri $addTeamMemberUri -Body $teamMemberBody | Out-Null
                    $groupStatus = "Success - Added to Team roster as $role"
                    if ($DetailedOutput) { Write-Host "  SUCCESS - Added to parent Team roster as $role." -ForegroundColor Green }
                }

                # Update the in-memory state in both modes to simulate the planned
                # roster addition and prevent duplicate planned additions on later rows.
                [void]$groupState.Members.Add($userId)
                if ($role -eq "Owner") {
                    [void]$groupState.Owners.Add($userId)
                }
            }
            else {
                $existingParentRole = if ($isGroupOwner) { "Owner" } else { "Member" }
                $groupStatus = "Already present as $existingParentRole"
                if ($DetailedOutput) { Write-Host "  INFO - Already present in parent group ($existingParentRole)." -ForegroundColor DarkCyan }
            }

            # Only attempt channel membership after parent-group membership is confirmed.
            $channelUsers = Get-ChannelCache -GroupId $groupId -ChannelId $channelId

            if ($channelUsers.Contains($userId)) {
                $channelStatus = "Already present - No change"
                $overallStatus = "Skipped"
                if ($DetailedOutput) { Write-Host "  SKIPPED - Already present in private channel; existing role preserved." -ForegroundColor Yellow }
            }
            else {
                $encodedChannelId = [System.Uri]::EscapeDataString($channelId)
                $channelUri = "$GraphBase/teams/$groupId/channels/$encodedChannelId/members"

                # The Graph API expects roles to be a JSON array of strings.
                # A PowerShell one-item array can be unwrapped by an if-expression
                # and serialized as a scalar string, which causes HTTP 400.
                $channelRoles = [System.Collections.Generic.List[string]]::new()
                if ($role -eq "Owner") {
                    [void]$channelRoles.Add("owner")
                }

                $channelBody = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = $channelRoles
                    "user@odata.bind" = "$GraphBase/users('$userId')"
                }

                if ($TestMode) {
                    $channelStatus = "TEST - Would add as $role"
                    $overallStatus = "Test - Would add"
                    if ($DetailedOutput) { Write-Host "  TEST - Would add to private channel as $role (no change submitted)." -ForegroundColor Yellow }
                }
                else {
                    Invoke-ChannelMemberPostWithRetry -Uri $channelUri -Body $channelBody
                    $channelStatus = "Success - Added as $role"
                    $overallStatus = "Success"
                    if ($DetailedOutput) { Write-Host "  SUCCESS - Added to private channel as $role." -ForegroundColor Green }
                }

                [void]$channelUsers.Add($userId)
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $errorMessage = "$errorMessage | Graph details: $($_.ErrorDetails.Message)"
            }
            $overallStatus = "Failed"
            if ($groupStatus -eq "Not processed") { $groupStatus = "Failed" }
            if ($channelStatus -eq "Not processed") { $channelStatus = "Failed" }
            $details = $errorMessage
            Write-Host "  FAILED - $errorMessage" -ForegroundColor Red
        }
    }

    $report.Add([pscustomobject]@{
        "Group Name" = $groupName
        "Group Object ID" = $groupId
        "Channel Name" = $channelName
        "Channel ID" = $channelId
        "User Display Name" = $displayName
        "User Object ID" = $userId
        "User Email" = $userEmail
        "Requested Role" = $role
        "Parent Group Status" = $groupStatus
        "Private Channel Status" = $channelStatus
        "Overall Status" = $overallStatus
        "Details" = $details
        "Processed At" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    })

    switch ($overallStatus) {
        'Success' { $successCount++ }
        'Test - Would add' { $testCount++ }
        'Failed' { $failedCount++ }
        'Skipped' { $skippedCount++ }
    }

    # Always emit exactly one concise completion line for each input row.
    # DetailedOutput may additionally show the individual Graph/cache operations.
    $rowResult = "[{0}/{1}] {2} | {3} <{4}> | Team: {5} | Channel: {6}" -f `
        $report.Count, $rows.Count, $overallStatus, $displayName, $userEmail, $groupStatus, $channelStatus
    $rowColor = switch ($overallStatus) {
        'Success' { 'Green' }
        'Test - Would add' { 'Yellow' }
        'Failed' { 'Red' }
        default { 'DarkGray' }
    }
    Write-Host $rowResult -ForegroundColor $rowColor

    if (($report.Count % $ProgressIntervalRows) -eq 0 -or $report.Count -eq $rows.Count) {
        $percentComplete = [int](($report.Count / $rows.Count) * 100)
        $elapsed = $stopwatch.Elapsed
        $rate = if ($elapsed.TotalSeconds -gt 0) { [Math]::Round($report.Count / $elapsed.TotalSeconds, 1) } else { 0 }
        Write-Progress -Activity 'Processing Teams private-channel memberships' `
            -Status "$($report.Count)/$($rows.Count) rows | $percentComplete% | $rate rows/sec | $($elapsed.ToString('hh\:mm\:ss')) elapsed" `
            -PercentComplete $percentComplete
        Write-Host ("Progress: {0}/{1} ({2}%) | {3} rows/sec | Failed: {4}" -f $report.Count, $rows.Count, $percentComplete, $rate, $failedCount)
    }
}

Write-Progress -Activity 'Processing Teams private-channel memberships' -Completed
$stopwatch.Stop()

# No report file is written during row processing. Export starts only after the
# complete input worksheet has been processed.
Write-Host "All $($rows.Count) input row(s) processed. Generating the final Excel report..." -ForegroundColor Cyan

# --------------------------- Export report ---------------------------

$summary = @(
    [pscustomobject]@{ Metric = "Execution Mode"; Value = $(if ($TestMode) { "TEST - No changes submitted" } else { "LIVE - Changes applied" }) }
    [pscustomobject]@{ Metric = "Total Rows"; Value = $report.Count }
    [pscustomobject]@{ Metric = "Successful Channel Additions"; Value = $successCount }
    [pscustomobject]@{ Metric = "Planned Channel Additions (Test Mode)"; Value = $testCount }
    [pscustomobject]@{ Metric = "Failed Rows"; Value = $failedCount }
    [pscustomobject]@{ Metric = "Skipped Rows"; Value = $skippedCount }
    [pscustomobject]@{ Metric = "Processing Duration"; Value = $stopwatch.Elapsed.ToString("hh\:mm\:ss") }
    [pscustomobject]@{ Metric = "Average Processing Rate (rows/sec)"; Value = $(if ($stopwatch.Elapsed.TotalSeconds -gt 0) { [Math]::Round($report.Count / $stopwatch.Elapsed.TotalSeconds, 2) } else { 0 }) }
    [pscustomobject]@{ Metric = "Unique Teams Cached"; Value = $script:GroupCache.Count }
    [pscustomobject]@{ Metric = "Unique Private Channels Cached"; Value = $script:ChannelCache.Count }
    [pscustomobject]@{ Metric = "Report Generated"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
)

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

# AutoSize is intentionally omitted: scanning 20k+ rows to calculate every
# column width can dominate report-generation time. Excel remains filterable.
$report | Export-Excel -Path $OutputPath -WorksheetName "Details" -TableName "MembershipResults" -AutoFilter -FreezeTopRow -BoldTopRow
$summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -TableName "ExecutionSummary" -AutoFilter -FreezeTopRow -BoldTopRow -Append

Write-Host ""
Write-Host "==================== EXECUTION SUMMARY ====================" -ForegroundColor Cyan
Write-Host "Total rows:                 $($report.Count)"
Write-Host "Processing duration:        $($stopwatch.Elapsed.ToString('hh\:mm\:ss'))"
Write-Host "Unique Teams cached:         $($script:GroupCache.Count)"
Write-Host "Unique channels cached:      $($script:ChannelCache.Count)"
Write-Host "Successful channel adds:    $successCount" -ForegroundColor Green
Write-Host "Planned channel adds:       $testCount" -ForegroundColor Yellow
Write-Host "Failed rows:                $failedCount" -ForegroundColor Red
Write-Host "Skipped rows:               $skippedCount" -ForegroundColor Yellow
Write-Host "Excel report:               $((Resolve-Path -LiteralPath $OutputPath).Path)" -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan
