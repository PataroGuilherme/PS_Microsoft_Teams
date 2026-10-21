#################
# Command Line: 
# TESTMODE (limited to 20 groups): powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\guilherme.pataro\Downloads\Get-TeamsPrivateChannelReport.ps1" -TestMode
# OR
# Production: powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\guilherme.pataro\Downloads\Get-TeamsPrivateChannelReport.ps1"
# Report: The final report is saved on: C:\Users\guilherme.pataro



<#
.SYNOPSIS
    Scans all Microsoft Teams groups in the tenant, lists their Private Channels
    (with IDs, owners and members) and exports a structured Excel report.

.DESCRIPTION
    For every Teams-enabled group in the tenant the script:
        1. Lists all channels of the team.
        2. If the team has no Private Channel, moves on to the next group.
        3. If the team has at least one Private Channel, lists each one with its
           channel ID, owners and members.
        4. Prints every piece of information to the console as it is collected.

    At the end, an Excel report (.xlsx) is generated with these sheets:
        Summary          - run information and totals
        Groups           - ALL groups scanned (name + Object ID), including those without Private Channels
        Private Channels - one line per Private Channel (IDs, counts, owners, members, guests)
        Channel Members  - one line per user per Private Channel (flat, filter-friendly)
        Detailed Log     - every message produced during the execution

    Test mode (-TestMode) limits the scan to the first 20 groups. The report is still
    generated exactly the same way.

    Authentication:
        - Default: interactive sign-in (delegated permissions).
        - Optional app-only sign-in with a certificate (-TenantId, -ClientId, -CertificateThumbprint).
          App-only is recommended for a complete tenant-wide scan, because delegated access
          may only expose Private Channels the signed-in user can see.

    Required Microsoft Graph permissions (read-only):
        Group.Read.All, Channel.ReadBasic.All, ChannelMember.Read.All, Team.ReadBasic.All

.PARAMETER ReportPath
    Optional path of the final .xlsx report. Defaults to a timestamped file in the current folder.

.PARAMETER TestMode
    When specified, only the first 20 groups are scanned (test run).

.PARAMETER TenantId
    Optional. Tenant ID for app-only (certificate) authentication.

.PARAMETER ClientId
    Optional. App (client) ID for app-only (certificate) authentication.

.PARAMETER CertificateThumbprint
    Optional. Certificate thumbprint for app-only (certificate) authentication.

.EXAMPLE
    .\Get-TeamsPrivateChannelReport.ps1 -TestMode

.EXAMPLE
    .\Get-TeamsPrivateChannelReport.ps1 -ReportPath "C:\Temp\private_channels.xlsx"

.EXAMPLE
    .\Get-TeamsPrivateChannelReport.ps1 -TenantId "<tenant-id>" -ClientId "<app-id>" -CertificateThumbprint "<thumbprint>"
#>

[CmdletBinding()]
param(
    [string]$ReportPath = (Join-Path (Get-Location) ("TeamsPrivateChannels_Report_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date))),

    [switch]$TestMode,

    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint
)

# ----------------------------------------------------------------------------
# Settings and report data (collected during the run, exported at the end)
# ----------------------------------------------------------------------------
$TestGroupLimit = 20

$script:LogEntries    = New-Object System.Collections.Generic.List[object]
$script:GroupRows     = New-Object System.Collections.Generic.List[object]
$script:ChannelRows   = New-Object System.Collections.Generic.List[object]
$script:MemberRows    = New-Object System.Collections.Generic.List[object]
$script:ConnectedTenant  = ''
$script:ConnectedAccount = ''
$script:AuthMode         = ''

$stats = [ordered]@{
    GroupsScanned           = 0
    GroupsWithPrivate       = 0
    GroupsWithoutPrivate    = 0
    GroupsWithErrors        = 0
    PrivateChannelsFound    = 0
    OwnerEntries            = 0
    MemberEntries           = 0
    GuestEntries            = 0
    Errors                  = 0
}

# ----------------------------------------------------------------------------
# Logging (console + in-memory list that becomes the "Detailed Log" sheet)
# ----------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'SKIP', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [string]$Color
    )
    $timestamp = Get-Date
    $script:LogEntries.Add([pscustomobject]@{
        'Timestamp' = $timestamp.ToString('yyyy-MM-dd HH:mm:ss')
        'Level'     = $Level
        'Message'   = $Message
    })

    if (-not $Color) {
        $Color = switch ($Level) {
            'SUCCESS' { 'Green' }
            'SKIP'    { 'Cyan' }
            'WARN'    { 'Yellow' }
            'ERROR'   { 'Red' }
            default   { 'White' }
        }
    }
    Write-Host ("{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f $timestamp, $Level, $Message) -ForegroundColor $Color
}

function Get-AdditionalProperty {
    param($Object, [string]$Name)
    if ($Object.AdditionalProperties -and $Object.AdditionalProperties.ContainsKey($Name)) {
        return $Object.AdditionalProperties[$Name]
    }
    return $null
}

function Format-Date {
    param($Value)
    if ($null -eq $Value) { return '' }
    try { return ([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') } catch { return "$Value" }
}

# ----------------------------------------------------------------------------
# Excel report
# ----------------------------------------------------------------------------
function Export-FinalReport {
    $summary = @(
        [pscustomobject]@{ 'Metric' = 'Report generated';                      'Value' = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
        [pscustomobject]@{ 'Metric' = 'Run mode';                              'Value' = $(if ($TestMode) { "Test mode (limited to $TestGroupLimit groups)" } else { 'Full tenant scan' }) }
        [pscustomobject]@{ 'Metric' = 'Authentication';                        'Value' = $script:AuthMode }
        [pscustomobject]@{ 'Metric' = 'Tenant ID';                             'Value' = $script:ConnectedTenant }
        [pscustomobject]@{ 'Metric' = 'Executed by';                           'Value' = $script:ConnectedAccount }
        [pscustomobject]@{ 'Metric' = 'Groups scanned';                        'Value' = $stats.GroupsScanned }
        [pscustomobject]@{ 'Metric' = 'Groups with Private Channels';          'Value' = $stats.GroupsWithPrivate }
        [pscustomobject]@{ 'Metric' = 'Groups without Private Channels';      'Value' = $stats.GroupsWithoutPrivate }
        [pscustomobject]@{ 'Metric' = 'Groups with errors';                    'Value' = $stats.GroupsWithErrors }
        [pscustomobject]@{ 'Metric' = 'Private Channels found';                'Value' = $stats.PrivateChannelsFound }
        [pscustomobject]@{ 'Metric' = 'Owner entries (all Private Channels)';  'Value' = $stats.OwnerEntries }
        [pscustomobject]@{ 'Metric' = 'Member entries (all Private Channels)'; 'Value' = $stats.MemberEntries }
        [pscustomobject]@{ 'Metric' = 'Guest entries (all Private Channels)';  'Value' = $stats.GuestEntries }
        [pscustomobject]@{ 'Metric' = 'Errors';                                'Value' = $stats.Errors }
    )

    try {
        if (Test-Path $ReportPath) { Remove-Item $ReportPath -Force -ErrorAction Stop }

        # Sheet 1: Summary
        $summary | Export-Excel -Path $ReportPath -WorksheetName 'Summary' -TableName 'SummaryTable' -TableStyle Medium2 -AutoSize

        # Sheet 2: Groups (all groups scanned)
        if ($script:GroupRows.Count -gt 0) {
            $last = $script:GroupRows.Count + 1
            $rangeStatus = "H2:H$last"
            $cond = @(
                New-ConditionalText -Text 'Error'     -Range $rangeStatus -BackgroundColor LightPink  -ConditionalTextColor Black
                New-ConditionalText -Text 'Processed' -Range $rangeStatus -BackgroundColor LightGreen -ConditionalTextColor Black
            )
            $script:GroupRows | Export-Excel -Path $ReportPath -WorksheetName 'Groups' -TableName 'GroupsTable' -TableStyle Medium2 -AutoSize -FreezeTopRow -ConditionalText $cond
        }

        # Sheet 3: Private Channels
        if ($script:ChannelRows.Count -gt 0) {
            $script:ChannelRows | Export-Excel -Path $ReportPath -WorksheetName 'Private Channels' -TableName 'ChannelsTable' -TableStyle Medium2 -AutoSize -FreezeTopRow
        }

        # Sheet 4: Channel Members (flat)
        if ($script:MemberRows.Count -gt 0) {
            $script:MemberRows | Export-Excel -Path $ReportPath -WorksheetName 'Channel Members' -TableName 'MembersTable' -TableStyle Medium2 -AutoSize -FreezeTopRow
        }

        # Sheet 5: Detailed Log
        if ($script:LogEntries.Count -gt 0) {
            $lastLog = $script:LogEntries.Count + 1
            $rangeLevel = "B2:B$lastLog"
            $condLog = @(
                New-ConditionalText -Text 'ERROR'   -Range $rangeLevel -BackgroundColor LightPink  -ConditionalTextColor Black
                New-ConditionalText -Text 'WARN'    -Range $rangeLevel -BackgroundColor Khaki      -ConditionalTextColor Black
                New-ConditionalText -Text 'SUCCESS' -Range $rangeLevel -BackgroundColor LightGreen -ConditionalTextColor Black
            )
            $script:LogEntries | Export-Excel -Path $ReportPath -WorksheetName 'Detailed Log' -TableName 'LogTable' -TableStyle Medium2 -AutoSize -FreezeTopRow -ConditionalText $condLog
        }

        Write-Host "`nExcel report saved to: $ReportPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to create the Excel report: $($_.Exception.Message)" -ForegroundColor Red
        # Fallback so the collected data is never lost
        try {
            $csvBase = [System.IO.Path]::ChangeExtension($ReportPath, $null).TrimEnd('.')
            if ($script:GroupRows.Count   -gt 0) { $script:GroupRows   | Export-Csv -Path "${csvBase}_Groups.csv"          -NoTypeInformation -Encoding UTF8 }
            if ($script:ChannelRows.Count -gt 0) { $script:ChannelRows | Export-Csv -Path "${csvBase}_PrivateChannels.csv" -NoTypeInformation -Encoding UTF8 }
            if ($script:MemberRows.Count  -gt 0) { $script:MemberRows  | Export-Csv -Path "${csvBase}_ChannelMembers.csv"  -NoTypeInformation -Encoding UTF8 }
            $script:LogEntries | Export-Csv -Path "${csvBase}_Log.csv" -NoTypeInformation -Encoding UTF8
            Write-Host "Fallback CSV files saved with prefix: $csvBase" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Fallback CSV export also failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ----------------------------------------------------------------------------
# Prerequisites
# ----------------------------------------------------------------------------
$appOnlyParams = @($TenantId, $ClientId, $CertificateThumbprint) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
if ($appOnlyParams.Count -gt 0 -and $appOnlyParams.Count -lt 3) {
    Write-Host "For app-only authentication, provide all three: -TenantId, -ClientId and -CertificateThumbprint." -ForegroundColor Red
    return
}

$requiredModules = @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Groups', 'Microsoft.Graph.Teams')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Module '$m' not found. Installing for current user..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
}

Write-Log "===== Script started =====" 'INFO'
Write-Log "Report file: $ReportPath" 'INFO'
if ($TestMode) { Write-Log "TEST MODE enabled - only the first $TestGroupLimit groups will be scanned." 'WARN' }

# ----------------------------------------------------------------------------
# Connect to Microsoft Graph
# ----------------------------------------------------------------------------
try {
    if ($appOnlyParams.Count -eq 3) {
        $script:AuthMode = 'App-only (certificate)'
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
    }
    else {
        $script:AuthMode = 'Delegated (interactive sign-in)'
        Connect-MgGraph -Scopes 'Group.Read.All', 'Channel.ReadBasic.All', 'ChannelMember.Read.All', 'Team.ReadBasic.All' -NoWelcome -ErrorAction Stop
    }
    $ctx = Get-MgContext
    $script:ConnectedTenant  = $ctx.TenantId
    $script:ConnectedAccount = if ($ctx.Account) { $ctx.Account } else { "App: $($ctx.ClientId)" }
    Write-Log "Connected to Microsoft Graph. Auth: $($script:AuthMode) | Tenant: $($ctx.TenantId) | Account: $($script:ConnectedAccount)" 'INFO'
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" 'ERROR'
    $stats.Errors++
    Export-FinalReport
    return
}

# ----------------------------------------------------------------------------
# Get all Teams-enabled groups
# ----------------------------------------------------------------------------
Write-Log "Retrieving Teams-enabled groups from the tenant..." 'INFO'
try {
    $groupParams = @{
        Filter           = "resourceProvisioningOptions/Any(x:x eq 'Team')"
        Property         = @('Id', 'DisplayName', 'Visibility', 'CreatedDateTime')
        ConsistencyLevel = 'eventual'
        CountVariable    = 'groupCount'
        ErrorAction      = 'Stop'
    }
    if ($TestMode) { $groupParams['Top'] = $TestGroupLimit } else { $groupParams['All'] = $true }

    $groups = @(Get-MgGroup @groupParams)
    if ($TestMode) { $groups = @($groups | Select-Object -First $TestGroupLimit) }
}
catch {
    Write-Log "Failed to retrieve groups: $($_.Exception.Message)" 'ERROR'
    $stats.Errors++
    Disconnect-MgGraph | Out-Null
    Export-FinalReport
    return
}

$totalGroups = $groups.Count
Write-Log "Groups to scan: $totalGroups" 'INFO'

# ----------------------------------------------------------------------------
# Main loop: one iteration per group
# ----------------------------------------------------------------------------
$index = 0
foreach ($group in $groups) {
    $index++
    $groupName = $group.DisplayName
    $groupId   = $group.Id

    Write-Host ""
    Write-Log ("--- Group {0}/{1}: {2} | Object ID: {3}" -f $index, $totalGroups, $groupName, $groupId) 'INFO' -Color 'Magenta'
    $stats.GroupsScanned++

    $totalChannels = 0
    $privateCount  = 0
    $groupStatus   = 'Processed'
    $groupDetails  = ''

    try {
        # 1) List all channels of the team
        $channels      = @(Get-MgTeamChannel -TeamId $groupId -All -ErrorAction Stop)
        $totalChannels = $channels.Count
        $privateChannels = @($channels | Where-Object { $_.MembershipType -eq 'private' })
        $privateCount    = $privateChannels.Count

        Write-Log "Channels found: $totalChannels (private: $privateCount)" 'INFO'

        # 2) No Private Channels: move on to the next group
        if ($privateCount -eq 0) {
            Write-Log "No Private Channels in this group. Moving to the next group." 'SKIP'
            $stats.GroupsWithoutPrivate++
            $groupDetails = 'No Private Channels.'
        }
        else {
            # 3) Private Channels found: list each one with owners and members
            $stats.GroupsWithPrivate++
            $groupDetails = "$privateCount Private Channel(s) found."

            foreach ($channel in $privateChannels) {
                $channelName = $channel.DisplayName
                $channelId   = $channel.Id
                $stats.PrivateChannelsFound++

                Write-Log ("  Private Channel: {0} | Channel ID: {1}" -f $channelName, $channelId) 'INFO' -Color 'Cyan'

                $owners        = New-Object System.Collections.Generic.List[string]
                $members       = New-Object System.Collections.Generic.List[string]
                $guests        = New-Object System.Collections.Generic.List[string]
                $membersStatus = 'Retrieved'
                $totalMembers  = $null

                try {
                    $channelMembers = @(Get-MgTeamChannelMember -TeamId $groupId -ChannelId $channelId -All -ErrorAction Stop)
                    $totalMembers   = $channelMembers.Count
                    Write-Log "    Channel users found: $totalMembers" 'INFO'

                    foreach ($m in $channelMembers) {
                        $userId = Get-AdditionalProperty -Object $m -Name 'userId'
                        $email  = Get-AdditionalProperty -Object $m -Name 'email'
                        $roles  = @($m.Roles)

                        if     ($roles -contains 'owner') { $role = 'Owner';  $stats.OwnerEntries++  }
                        elseif ($roles -contains 'guest') { $role = 'Guest';  $stats.GuestEntries++  }
                        else                              { $role = 'Member'; $stats.MemberEntries++ }

                        $label = if ($email) { "$($m.DisplayName) <$email>" } else { "$($m.DisplayName)" }
                        switch ($role) {
                            'Owner'  { $owners.Add($label) }
                            'Member' { $members.Add($label) }
                            'Guest'  { $guests.Add($label) }
                        }

                        Write-Log ("    [{0}] {1} | Email: {2} | Object ID: {3}" -f $role, $m.DisplayName, $email, $userId) 'INFO'

                        $script:MemberRows.Add([pscustomobject]@{
                            'Group Name'         = $groupName
                            'Group Object ID'    = $groupId
                            'Channel Name'       = $channelName
                            'Channel ID'         = $channelId
                            'User Display Name'  = $m.DisplayName
                            'User Object ID'     = $userId
                            'User Email'         = $email
                            'Role'               = $role
                        })
                    }
                }
                catch {
                    $membersStatus = "Error: $($_.Exception.Message)"
                    Write-Log "    Failed to retrieve members of channel '$channelName': $($_.Exception.Message)" 'ERROR'
                    $stats.Errors++
                }

                $script:ChannelRows.Add([pscustomobject]@{
                    'Group Name'              = $groupName
                    'Group Object ID'         = $groupId
                    'Channel Name'            = $channelName
                    'Channel ID'              = $channelId
                    'Channel Created (UTC)'   = Format-Date $channel.CreatedDateTime
                    'Total Users'             = $totalMembers
                    'Owners Count'            = $(if ($membersStatus -eq 'Retrieved') { $owners.Count }  else { $null })
                    'Members Count'           = $(if ($membersStatus -eq 'Retrieved') { $members.Count } else { $null })
                    'Guests Count'            = $(if ($membersStatus -eq 'Retrieved') { $guests.Count }  else { $null })
                    'Owners'                  = ($owners  -join '; ')
                    'Members'                 = ($members -join '; ')
                    'Guests'                  = ($guests  -join '; ')
                    'Members Retrieval Status' = $membersStatus
                })
            }
        }
    }
    catch {
        $groupStatus  = 'Error'
        $groupDetails = $_.Exception.Message
        Write-Log "Failed to process group '$groupName': $groupDetails" 'ERROR'
        $stats.GroupsWithErrors++
        $stats.Errors++
    }

    # Every group is recorded, with or without Private Channels
    $script:GroupRows.Add([pscustomobject]@{
        'Group Name'            = $groupName
        'Group Object ID'       = $groupId
        'Visibility'            = $group.Visibility
        'Group Created (UTC)'   = Format-Date $group.CreatedDateTime
        'Total Channels'        = $totalChannels
        'Private Channels'      = $privateCount
        'Has Private Channels'  = $(if ($privateCount -gt 0) { 'Yes' } else { 'No' })
        'Status'                = $groupStatus
        'Details'               = $groupDetails
    })
}

# ----------------------------------------------------------------------------
# Summary and final report
# ----------------------------------------------------------------------------
Write-Host ""
Write-Log "===== Summary =====" 'INFO'
Write-Log "Groups scanned ......................... $($stats.GroupsScanned)" 'INFO'
Write-Log "Groups with Private Channels ........... $($stats.GroupsWithPrivate)" 'INFO'
Write-Log "Groups without Private Channels ........ $($stats.GroupsWithoutPrivate)" 'INFO'
Write-Log "Groups with errors ..................... $($stats.GroupsWithErrors)" 'INFO'
Write-Log "Private Channels found ................. $($stats.PrivateChannelsFound)" 'INFO'
Write-Log "Owner entries .......................... $($stats.OwnerEntries)" 'INFO'
Write-Log "Member entries ......................... $($stats.MemberEntries)" 'INFO'
Write-Log "Guest entries .......................... $($stats.GuestEntries)" 'INFO'
Write-Log "Errors ................................. $($stats.Errors)" 'INFO'
Write-Log "===== Script finished =====" 'INFO'

Disconnect-MgGraph | Out-Null
Export-FinalReport
