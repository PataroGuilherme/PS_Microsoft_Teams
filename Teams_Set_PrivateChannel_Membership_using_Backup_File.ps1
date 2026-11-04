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
#>

[CmdletBinding()]
param(
    [string]$InputPath  = ".\Book1(1).xlsx",
    [string]$OutputPath = ".\PrivateChannel_Membership_Report.xlsx",
    [switch]$TestMode
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
    "TeamMember.ReadWrite.All",
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

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject
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

    $json = $Body | ConvertTo-Json -Depth 8
    Invoke-MgGraphRequest -Method POST -Uri $Uri -Body $json -ContentType "application/json" -OutputType PSObject
}

function New-IdSet {
    return [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
}

function Get-GroupCache {
    param([Parameter(Mandatory)][string]$GroupId)

    $key = $GroupId.ToLowerInvariant()
    if (-not $script:GroupCache.ContainsKey($key)) {
        Write-Host "  Reading parent group membership: $GroupId" -ForegroundColor DarkGray

        $members = New-IdSet
        $owners  = New-IdSet

        # The Teams membership endpoint returns both team members and owners
        # (owners are identified by the "owner" role).
        $teamMembersUri = "$GraphBase/teams/$GroupId/members?`$select=userId,roles&`$top=100"

        foreach ($entry in (Invoke-GraphGetAll -Uri $teamMembersUri)) {
            if ($entry.userId) {
                [void]$members.Add([string]$entry.userId)
                if (@($entry.roles) -contains "owner") {
                    [void]$owners.Add([string]$entry.userId)
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
        Write-Host "  Reading private channel membership: $ChannelId" -ForegroundColor DarkGray

        $userIds = New-IdSet
        $encodedChannelId = [System.Uri]::EscapeDataString($ChannelId)
        $uri = "$GraphBase/teams/$GroupId/channels/$encodedChannelId/members?`$select=userId,roles&`$top=100"

        foreach ($entry in (Invoke-GraphGetAll -Uri $uri)) {
            if ($entry.userId) { [void]$userIds.Add([string]$entry.userId) }
        }

        $script:ChannelCache[$key] = $userIds
    }

    return $script:ChannelCache[$key]
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

Write-Host ""
Write-Host "Starting membership processing for $($rows.Count) row(s)..." -ForegroundColor Cyan
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

    Write-Host "[$rowNumber/$($rows.Count + 1)] $displayName <$userEmail>" -ForegroundColor White

    # Treat a partially completed row as invalid; do not guess missing values.
    $values = @($groupName, $groupId, $channelName, $channelId, $displayName, $userId, $userEmail, $role)
    if (@($values | Where-Object { [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        $details = "Skipped: one or more required cells are blank."
        Write-Host "  SKIPPED - Missing required information." -ForegroundColor Yellow
    }
    elseif ($role -notin @("Member", "Owner")) {
        $details = "Skipped: Role must be 'Member' or 'Owner'."
        Write-Host "  SKIPPED - Invalid Role '$role'." -ForegroundColor Yellow
    }
    else {
        try {
            # Load/cache parent group members and owners only once per group.
            $groupState = Get-GroupCache -GroupId $groupId
            $isGroupMember = $groupState.Members.Contains($userId)
            $isGroupOwner  = $groupState.Owners.Contains($userId)

            if (-not ($isGroupMember -or $isGroupOwner)) {
                $teamRoles = if ($role -eq "Owner") { @("owner") } else { @() }
                $teamBody = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = $teamRoles
                    "user@odata.bind" = "$GraphBase/users('$userId')"
                }
                $addTeamUri = "$GraphBase/teams/$groupId/members"

                if ($TestMode) {
                    $groupStatus = "TEST - Would add as $role"
                    Write-Host "  TEST - Would add to parent team as $role (no change submitted)." -ForegroundColor Yellow
                }
                else {
                    Invoke-GraphPost -Uri $addTeamUri -Body $teamBody | Out-Null
                    $groupStatus = "Success - Added as $role"
                    Write-Host "  SUCCESS - Added to parent team as $role." -ForegroundColor Green
                }

                # Update the in-memory state in both modes to simulate the planned
                # membership and prevent duplicate planned additions on later rows.
                [void]$groupState.Members.Add($userId)
                if ($role -eq "Owner") { [void]$groupState.Owners.Add($userId) }
            }
            else {
                $existingParentRole = if ($isGroupOwner) { "Owner" } else { "Member" }
                $groupStatus = "Already present as $existingParentRole"
                Write-Host "  INFO - Already present in parent group ($existingParentRole)." -ForegroundColor DarkCyan
            }

            # Only attempt channel membership after parent-group membership is confirmed.
            $channelUsers = Get-ChannelCache -GroupId $groupId -ChannelId $channelId

            if ($channelUsers.Contains($userId)) {
                $channelStatus = "Already present - No change"
                $overallStatus = "Skipped"
                Write-Host "  SKIPPED - Already present in private channel; existing role preserved." -ForegroundColor Yellow
            }
            else {
                $encodedChannelId = [System.Uri]::EscapeDataString($channelId)
                $channelUri = "$GraphBase/teams/$groupId/channels/$encodedChannelId/members"

                $channelRoles = if ($role -eq "Owner") { @("owner") } else { @() }

                $channelBody = @{
                    "@odata.type" = "#microsoft.graph.aadUserConversationMember"
                    "roles" = $channelRoles
                    "user@odata.bind" = "$GraphBase/users('$userId')"
                }

                if ($TestMode) {
                    $channelStatus = "TEST - Would add as $role"
                    $overallStatus = "Test - Would add"
                    Write-Host "  TEST - Would add to private channel as $role (no change submitted)." -ForegroundColor Yellow
                }
                else {
                    Invoke-GraphPost -Uri $channelUri -Body $channelBody | Out-Null
                    $channelStatus = "Success - Added as $role"
                    $overallStatus = "Success"
                    Write-Host "  SUCCESS - Added to private channel as $role." -ForegroundColor Green
                }

                [void]$channelUsers.Add($userId)
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
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
}

# --------------------------- Export report ---------------------------

$successCount = @($report | Where-Object { $_.'Overall Status' -eq "Success" }).Count
$testCount    = @($report | Where-Object { $_.'Overall Status' -eq "Test - Would add" }).Count
$failedCount  = @($report | Where-Object { $_.'Overall Status' -eq "Failed" }).Count
$skippedCount = @($report | Where-Object { $_.'Overall Status' -eq "Skipped" }).Count

$summary = @(
    [pscustomobject]@{ Metric = "Execution Mode"; Value = $(if ($TestMode) { "TEST - No changes submitted" } else { "LIVE - Changes applied" }) }
    [pscustomobject]@{ Metric = "Total Rows"; Value = $report.Count }
    [pscustomobject]@{ Metric = "Successful Channel Additions"; Value = $successCount }
    [pscustomobject]@{ Metric = "Planned Channel Additions (Test Mode)"; Value = $testCount }
    [pscustomobject]@{ Metric = "Failed Rows"; Value = $failedCount }
    [pscustomobject]@{ Metric = "Skipped Rows"; Value = $skippedCount }
    [pscustomobject]@{ Metric = "Report Generated"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
)

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

$report | Export-Excel -Path $OutputPath -WorksheetName "Details" -TableName "MembershipResults" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$summary | Export-Excel -Path $OutputPath -WorksheetName "Summary" -TableName "ExecutionSummary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -Append

Write-Host ""
Write-Host "==================== EXECUTION SUMMARY ====================" -ForegroundColor Cyan
Write-Host "Total rows:                 $($report.Count)"
Write-Host "Successful channel adds:    $successCount" -ForegroundColor Green
Write-Host "Planned channel adds:       $testCount" -ForegroundColor Yellow
Write-Host "Failed rows:                $failedCount" -ForegroundColor Red
Write-Host "Skipped rows:               $skippedCount" -ForegroundColor Yellow
Write-Host "Excel report:               $((Resolve-Path -LiteralPath $OutputPath).Path)" -ForegroundColor Cyan
Write-Host "===========================================================" -ForegroundColor Cyan
