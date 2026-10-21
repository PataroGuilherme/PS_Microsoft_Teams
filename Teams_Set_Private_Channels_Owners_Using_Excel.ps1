### Para este código funcionar, é necessário um Excel que tenha a seguinte estrutura de colunas:
# Coluna A: Group ID
# Coluna B: Private Channel ID
# Coluna C: User mail address;

### Em seguida, este arquivo de PowerShell precisa ser baixado na pasta de Downloads, e deixado junto com o Excel. Ambos na mesma pasta.

### Por fim, o comando a ser usado no CMD para execução é este abaixo:
# powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\users\guilherme.pataro\downloads\Add-PrivateChannelOwners-Excel.ps1" -ExcelPath "C:\users\guilherme.pataro\downloads\groups.xlsx"

##########################################
############# FUNÇÕES
# 1. Este código identifica o Owner listado > Confere se ele já é Owner no grupo Pai > Se for, o adiciona diretamente no Private Channel como Owner também. Se não, primeiro o adiciona no grupo pai como owner e depois no Private Channel.
# 2. Todos os demais membros ficam intactos, da forma como já estavam.
# 3. Este script não remove o Onwer do grupo pai (para evitar que ele seja quitado também de todos os demais canais privados).

<#
.SYNOPSIS
    Adds users as Owner of Teams Private Channels based on an Excel file and
    produces a structured Excel report (English).

.DESCRIPTION
    Input Excel layout (first sheet by default; row 1 is the header and is skipped):
        Column A = Teams group (parent team) Object ID
        Column B = Private Channel ID
        Column C = User e-mail address

    For each row:
        1. Resolves the user in Entra ID.
        2. If the user is already Owner of the Private Channel, moves to the next row.
        3. If the user is NOT a member of the parent team, adds the user to the team as Owner.
        4. Adds the user as Owner of the Private Channel
           (or promotes the user to Owner if already a channel member).
        5. Writes the result to the console and to the final Excel report.

    Final report (.xlsx) sheets:
        Summary      - run information and totals
        Report       - one line per processed row with the action taken on the team and on the channel
        Detailed Log - every message produced during the execution, with timestamp and level

.PARAMETER ExcelPath
    Path to the input .xlsx file.

.PARAMETER SheetName
    Optional worksheet name. Defaults to the first worksheet.

.PARAMETER ReportPath
    Optional path of the final .xlsx report. Defaults to a timestamped file in the current folder.

.PARAMETER DryRun
    Simulates everything without making any changes.

.EXAMPLE
    .\Add-PrivateChannelOwners.ps1 -ExcelPath "C:\Temp\channels.xlsx"

.EXAMPLE
    .\Add-PrivateChannelOwners.ps1 -ExcelPath "C:\Temp\channels.xlsx" -ReportPath "C:\Temp\report.xlsx" -DryRun
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExcelPath,

    [string]$SheetName,

    [string]$ReportPath = (Join-Path (Get-Location) ("PrivateChannelOwners_Report_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date))),

    [switch]$DryRun
)

# ----------------------------------------------------------------------------
# Report data (collected during the run and exported at the end)
# ----------------------------------------------------------------------------
$script:LogEntries = New-Object System.Collections.Generic.List[object]
$script:ReportRows = New-Object System.Collections.Generic.List[object]
$script:TenantId   = ''
$script:Account    = ''
$total             = 0

$stats = [ordered]@{
    Processed          = 0
    TeamOwnerAdded     = 0
    ChannelOwnerAdded  = 0
    ChannelPromoted    = 0
    AlreadyOwner       = 0
    SkippedInvalid     = 0
    Errors             = 0
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

function Add-Result {
    param(
        [int]$ExcelRow,
        [string]$TeamId,
        [string]$ChannelId,
        [string]$Email,
        [string]$UserName,
        [string]$UserId,
        [string]$TeamAction,
        [string]$ChannelAction,
        [string]$Status,
        [string]$Details
    )
    $script:ReportRows.Add([pscustomobject]@{
        'Excel Row'          = $ExcelRow
        'Timestamp'          = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        'User Email'         = $Email
        'User Display Name'  = $UserName
        'User Object ID'     = $UserId
        'Team Object ID'     = $TeamId
        'Private Channel ID' = $ChannelId
        'Team Action'        = $TeamAction
        'Channel Action'     = $ChannelAction
        'Status'             = $Status
        'Details'            = $Details
    })
}

function Export-FinalReport {
    $dryRunRows = @($script:ReportRows | Where-Object { $_.Status -eq 'DryRun' }).Count

    $summary = @(
        [pscustomobject]@{ 'Metric' = 'Report generated';                        'Value' = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }
        [pscustomobject]@{ 'Metric' = 'Mode';                                    'Value' = $(if ($DryRun) { 'Dry run (no changes made)' } else { 'Live' }) }
        [pscustomobject]@{ 'Metric' = 'Source file';                             'Value' = $ExcelPath }
        [pscustomobject]@{ 'Metric' = 'Tenant ID';                               'Value' = $script:TenantId }
        [pscustomobject]@{ 'Metric' = 'Executed by';                             'Value' = $script:Account }
        [pscustomobject]@{ 'Metric' = 'Data rows in source file (header excluded)'; 'Value' = $total }
        [pscustomobject]@{ 'Metric' = 'Rows processed';                          'Value' = $stats.Processed }
        [pscustomobject]@{ 'Metric' = 'Rows skipped (invalid data)';             'Value' = $stats.SkippedInvalid }
        [pscustomobject]@{ 'Metric' = 'Users added as Team Owner';               'Value' = $stats.TeamOwnerAdded }
        [pscustomobject]@{ 'Metric' = 'Users added as Channel Owner';            'Value' = $stats.ChannelOwnerAdded }
        [pscustomobject]@{ 'Metric' = 'Users promoted to Channel Owner';         'Value' = $stats.ChannelPromoted }
        [pscustomobject]@{ 'Metric' = 'Users already Channel Owner (no action)'; 'Value' = $stats.AlreadyOwner }
        [pscustomobject]@{ 'Metric' = 'Rows simulated (dry run)';                'Value' = $dryRunRows }
        [pscustomobject]@{ 'Metric' = 'Errors';                                  'Value' = $stats.Errors }
    )

    try {
        if (Test-Path $ReportPath) { Remove-Item $ReportPath -Force -ErrorAction Stop }

        # Sheet 1: Summary
        $summary | Export-Excel -Path $ReportPath -WorksheetName 'Summary' -TableName 'SummaryTable' -TableStyle Medium2 -AutoSize

        # Sheet 2: Report (one line per processed row)
        if ($script:ReportRows.Count -gt 0) {
            $last = $script:ReportRows.Count + 1
            $rangeStatus = "J2:J$last"
            $cond = @(
                New-ConditionalText -Text 'Success' -Range $rangeStatus -BackgroundColor LightGreen -ConditionalTextColor Black
                New-ConditionalText -Text 'Error'   -Range $rangeStatus -BackgroundColor LightPink  -ConditionalTextColor Black
                New-ConditionalText -Text 'Skipped' -Range $rangeStatus -BackgroundColor LightBlue  -ConditionalTextColor Black
                New-ConditionalText -Text 'DryRun'  -Range $rangeStatus -BackgroundColor Khaki      -ConditionalTextColor Black
            )
            $script:ReportRows | Export-Excel -Path $ReportPath -WorksheetName 'Report' -TableName 'ReportTable' -TableStyle Medium2 -AutoSize -FreezeTopRow -ConditionalText $cond
        }

        # Sheet 3: Detailed Log
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
        # Fallback so the results are never lost
        try {
            $csvBase = [System.IO.Path]::ChangeExtension($ReportPath, $null).TrimEnd('.')
            if ($script:ReportRows.Count -gt 0) { $script:ReportRows | Export-Csv -Path "${csvBase}_Report.csv" -NoTypeInformation -Encoding UTF8 }
            $script:LogEntries | Export-Csv -Path "${csvBase}_Log.csv" -NoTypeInformation -Encoding UTF8
            Write-Host "Fallback CSV files saved next to: $csvBase" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Fallback CSV export also failed: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# ----------------------------------------------------------------------------
# Prerequisites
# ----------------------------------------------------------------------------
if (-not (Test-Path $ExcelPath)) {
    Write-Host "Excel file not found: $ExcelPath" -ForegroundColor Red
    return
}

$requiredModules = @('ImportExcel', 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Users', 'Microsoft.Graph.Teams')
foreach ($m in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Module '$m' not found. Installing for current user..." -ForegroundColor Yellow
        Install-Module -Name $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m -ErrorAction Stop
}

Write-Log "===== Script started =====" 'INFO'
Write-Log "Excel file: $ExcelPath" 'INFO'
Write-Log "Report file: $ReportPath" 'INFO'
if ($DryRun) { Write-Log "DRY RUN mode enabled - no changes will be made." 'WARN' }

# ----------------------------------------------------------------------------
# Connect to Microsoft Graph
# ----------------------------------------------------------------------------
try {
    Connect-MgGraph -Scopes 'User.Read.All', 'TeamMember.ReadWrite.All', 'ChannelMember.ReadWrite.All', 'Group.ReadWrite.All' -NoWelcome -ErrorAction Stop
    $ctx = Get-MgContext
    $script:TenantId = $ctx.TenantId
    $script:Account  = $ctx.Account
    Write-Log "Connected to Microsoft Graph. Tenant: $($ctx.TenantId) | Account: $($ctx.Account)" 'INFO'
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" 'ERROR'
    Export-FinalReport
    return
}

# ----------------------------------------------------------------------------
# Read Excel (no header assumption: columns are read as P1, P2, P3)
# ----------------------------------------------------------------------------
$importParams = @{ Path = $ExcelPath; NoHeader = $true }
if ($SheetName) { $importParams['WorksheetName'] = $SheetName }

try {
    $rows = @(Import-Excel @importParams)
}
catch {
    Write-Log "Failed to read Excel file: $($_.Exception.Message)" 'ERROR'
    Disconnect-MgGraph | Out-Null
    Export-FinalReport
    return
}

$totalInSheet = $rows.Count

# The first row of the worksheet is the header: skip it and start from row 2
$rows  = @($rows | Select-Object -Skip 1)
$total = $rows.Count
Write-Log "Rows found in worksheet (including header): $totalInSheet" 'INFO'
Write-Log "Header row (row 1) skipped. Data rows to process: $total" 'INFO'

# ----------------------------------------------------------------------------
# Helpers and caches
# ----------------------------------------------------------------------------
$guidRegex   = '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$userCache   = @{}   # email  -> user object
$teamMembers = @{}   # teamId -> array of team members

function Get-UserByEmail {
    param([string]$Email)

    if ($userCache.ContainsKey($Email)) { return $userCache[$Email] }

    $user = $null
    try { $user = Get-MgUser -UserId $Email -Property Id, DisplayName, UserPrincipalName, Mail -ErrorAction Stop } catch { }

    if (-not $user) {
        $safe = $Email.Replace("'", "''")
        $user = Get-MgUser -Filter "mail eq '$safe' or userPrincipalName eq '$safe'" -Property Id, DisplayName, UserPrincipalName, Mail -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -First 1
    }

    $userCache[$Email] = $user
    return $user
}

function Get-TeamMembersCached {
    param([string]$TeamId, [switch]$Refresh)

    if ($Refresh -or -not $teamMembers.ContainsKey($TeamId)) {
        $teamMembers[$TeamId] = @(Get-MgTeamMember -TeamId $TeamId -All -ErrorAction Stop)
    }
    return $teamMembers[$TeamId]
}

function New-OwnerBody {
    param([string]$UserId)
    return @{
        '@odata.type'     = '#microsoft.graph.aadUserConversationMember'
        'roles'           = @('owner')
        'user@odata.bind' = "https://graph.microsoft.com/v1.0/users('$UserId')"
    }
}

# ----------------------------------------------------------------------------
# Main loop
# ----------------------------------------------------------------------------
$rowNumber = 0
foreach ($row in $rows) {
    $rowNumber++

    $teamId    = ("$($row.P1)").Trim()
    $channelId = ("$($row.P2)").Trim()
    $email     = ("$($row.P3)").Trim()

    $excelRow = $rowNumber + 1   # +1 because row 1 of the sheet is the header
    $common   = @{ ExcelRow = $excelRow; TeamId = $teamId; ChannelId = $channelId; Email = $email }

    Write-Host ""
    Write-Log ("--- Item {0}/{1} (Excel row {2}) | Team: {3} | Channel: {4} | User: {5}" -f $rowNumber, $total, $excelRow, $teamId, $channelId, $email) 'INFO' -Color 'Magenta'

    # Skip empty / invalid rows
    if ($teamId -notmatch $guidRegex -or [string]::IsNullOrWhiteSpace($email)) {
        Write-Log "Excel row $excelRow skipped: invalid Team ID or empty e-mail." 'SKIP'
        Add-Result @common -TeamAction 'Not performed' -ChannelAction 'Not performed' -Status 'Skipped' -Details 'Invalid Team ID or empty e-mail.'
        $stats.SkippedInvalid++
        continue
    }
    if ([string]::IsNullOrWhiteSpace($channelId)) {
        Write-Log "Excel row $excelRow skipped: empty Channel ID." 'SKIP'
        Add-Result @common -TeamAction 'Not performed' -ChannelAction 'Not performed' -Status 'Skipped' -Details 'Empty Channel ID.'
        $stats.SkippedInvalid++
        continue
    }

    $stats.Processed++

    $user          = $null
    $userName      = ''
    $userId        = ''
    $teamAction    = 'Not performed'
    $channelAction = 'Not performed'

    try {
        # 1) Resolve user
        $user = Get-UserByEmail -Email $email
        if (-not $user) {
            Write-Log "User '$email' was not found in Entra ID. Row skipped." 'ERROR'
            Add-Result @common -TeamAction $teamAction -ChannelAction $channelAction -Status 'Error' -Details 'User not found in Entra ID.'
            $stats.Errors++
            continue
        }
        $userName = $user.DisplayName
        $userId   = $user.Id
        Write-Log "User resolved: $($user.DisplayName) <$($user.UserPrincipalName)> (ObjectId: $($user.Id))" 'INFO'

        # 2) Check private channel membership first: if already Owner, go to the next row
        $channelMembers  = @(Get-MgTeamChannelMember -TeamId $teamId -ChannelId $channelId -All -ErrorAction Stop)
        $existingChannel = $channelMembers | Where-Object { $_.AdditionalProperties['userId'] -eq $user.Id }

        if ($existingChannel -and ($existingChannel.Roles -contains 'owner')) {
            Write-Log "User is already Owner of the private channel. Moving to the next row." 'SKIP'
            Add-Result @common -UserName $userName -UserId $userId -TeamAction 'Not checked' -ChannelAction 'Already Owner (no action)' -Status 'Skipped' -Details 'User is already Owner of the private channel.'
            $stats.AlreadyOwner++
            continue
        }

        # 3) Check parent team membership
        $members  = Get-TeamMembersCached -TeamId $teamId
        $isMember = $members | Where-Object { $_.AdditionalProperties['userId'] -eq $user.Id }

        if ($isMember) {
            $roles = if ($isMember.Roles) { $isMember.Roles -join ',' } else { 'member' }
            Write-Log "User is already in the parent team (role: $roles). No change needed at team level." 'SKIP'
            $teamAction = "Already in team (role: $roles) - no action"
        }
        else {
            Write-Log "User is NOT in the parent team. Adding as Owner of the team..." 'INFO'
            if ($DryRun) {
                Write-Log "[DRY RUN] Would add $email as Owner of team $teamId." 'WARN'
                $teamAction = 'Would add as Owner (dry run)'
            }
            else {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/members" -Body ((New-OwnerBody -UserId $user.Id) | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                Write-Log "User $email added as Owner of team $teamId." 'SUCCESS'
                $teamAction = 'Added as Owner'
                $stats.TeamOwnerAdded++
                [void](Get-TeamMembersCached -TeamId $teamId -Refresh)
                Start-Sleep -Seconds 5   # allow propagation before touching the private channel
            }
        }

        # 4) Private channel: promote or add as Owner (membership was checked in step 2)
        if ($existingChannel) {
            Write-Log "User is a channel member (not Owner). Promoting to Owner..." 'INFO'
            if ($DryRun) {
                Write-Log "[DRY RUN] Would promote $email to Owner of channel $channelId." 'WARN'
                $channelAction = 'Would promote to Owner (dry run)'
            }
            else {
                $upd = @{
                    '@odata.type' = '#microsoft.graph.aadUserConversationMember'
                    'roles'       = @('owner')
                }
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$channelId/members/$($existingChannel.Id)" -Body ($upd | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                Write-Log "User $email promoted to Owner of private channel $channelId." 'SUCCESS'
                $channelAction = 'Promoted to Owner'
                $stats.ChannelPromoted++
            }
        }
        else {
            Write-Log "Adding user as Owner of the private channel..." 'INFO'
            if ($DryRun) {
                Write-Log "[DRY RUN] Would add $email as Owner of channel $channelId." 'WARN'
                $channelAction = 'Would add as Owner (dry run)'
            }
            else {
                $maxAttempts = 3
                for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
                    try {
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/teams/$teamId/channels/$channelId/members" -Body ((New-OwnerBody -UserId $user.Id) | ConvertTo-Json -Depth 5) -ContentType 'application/json' -ErrorAction Stop | Out-Null
                        Write-Log "User $email added as Owner of private channel $channelId." 'SUCCESS'
                        $channelAction = 'Added as Owner'
                        $stats.ChannelOwnerAdded++
                        break
                    }
                    catch {
                        if ($attempt -lt $maxAttempts) {
                            Write-Log "Attempt $attempt/$maxAttempts failed ($($_.Exception.Message)). Retrying in 10s..." 'WARN'
                            Start-Sleep -Seconds 10
                        }
                        else { throw }
                    }
                }
            }
        }

        # Row completed
        if ($DryRun) {
            Add-Result @common -UserName $userName -UserId $userId -TeamAction $teamAction -ChannelAction $channelAction -Status 'DryRun' -Details 'Simulated - no changes were made.'
        }
        else {
            Add-Result @common -UserName $userName -UserId $userId -TeamAction $teamAction -ChannelAction $channelAction -Status 'Success' -Details 'Completed successfully.'
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log "Excel row $excelRow failed for '$email': $errMsg" 'ERROR'
        Add-Result @common -UserName $userName -UserId $userId -TeamAction $teamAction -ChannelAction $channelAction -Status 'Error' -Details $errMsg
        $stats.Errors++
    }
}

# ----------------------------------------------------------------------------
# Summary and final report
# ----------------------------------------------------------------------------
Write-Host ""
Write-Log "===== Summary =====" 'INFO'
Write-Log "Data rows (header excluded) ........ $total" 'INFO'
Write-Log "Rows processed ...................... $($stats.Processed)" 'INFO'
Write-Log "Rows skipped (invalid data) ......... $($stats.SkippedInvalid)" 'INFO'
Write-Log "Users added as Team Owner ........... $($stats.TeamOwnerAdded)" 'INFO'
Write-Log "Users added as Channel Owner ........ $($stats.ChannelOwnerAdded)" 'INFO'
Write-Log "Users promoted to Channel Owner ..... $($stats.ChannelPromoted)" 'INFO'
Write-Log "Users already Channel Owner ......... $($stats.AlreadyOwner)" 'INFO'
Write-Log "Errors .............................. $($stats.Errors)" 'INFO'
Write-Log "===== Script finished =====" 'INFO'

Disconnect-MgGraph | Out-Null
Export-FinalReport
