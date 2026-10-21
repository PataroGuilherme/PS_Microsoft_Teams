### Para este código funcionar, é necessário um Excel que tenha a seguinte estrutura de colunas:
# Coluna A: Group ID
# Coluna B: Private Channel ID
# Coluna C: User mail address;

### Em seguida, este arquivo de PowerShell precisa ser baixado na pasta de Downloads, e deixado junto com o Excel. Ambos na mesma pasta.

### Por fim, o comando a ser usado no CMD para execução é este abaixo:
# powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Users\guilherme.pataro\Downloads\set-owners.ps1" -ExcelPath "C:\Users\guilherme.pataro\Downloads\users.xlsx" -ReportPath "C:\Users\guilherme.pataro\Downloads\OwnerReport-MS.csv"

##########################################
############# FUNÇÕES
# 1. Este código identifica o Owner listado > Remove do grupo Pai > Aguarda 5 segundos > Adiciona como Owner do Grupo Pai novamente > Aguarda 5 segundos > Adiciona como Owner do Private Channel
# 2. Todos os demais membros ficam intactos, da forma como já estavam.


# Requirements: Install-Module Microsoft.Graph, ImportExcel
param([Parameter(Mandatory)][string]$ExcelPath,
      [string]$ReportPath = ".\OwnerReport_$(Get-Date -Format yyyyMMdd_HHmm).csv")

Connect-MgGraph -Scopes "TeamMember.ReadWrite.All","ChannelMember.ReadWrite.All","User.Read.All" -NoWelcome

$WaitSeconds = 5
$GraphBase   = "https://graph.microsoft.com/v1.0"

# Returns the most useful error text (Graph API details when available)
function Get-ErrorText {
    param($ErrorRecord)
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        return $ErrorRecord.ErrorDetails.Message
    }
    return $ErrorRecord.Exception.Message
}

# Finds the membership entry of a user in a team or in a channel (returns $null if not found)
function Find-Member {
    param([string]$TeamId, [string]$ChannelId, [string]$UserId)

    $uri = if ($ChannelId) { "$GraphBase/teams/$TeamId/channels/$ChannelId/members" }
           else            { "$GraphBase/teams/$TeamId/members" }

    while ($uri) {
        $resp  = Invoke-MgGraphRequest -Method GET -Uri $uri
        $match = $resp.value | Where-Object { $_.userId -eq $UserId } | Select-Object -First 1
        if ($match) { return $match }
        $uri = $resp.'@odata.nextLink'
    }
    return $null
}

# Removes the user from the parent team if they are an Owner or Member.
# Returns: "Removed (Owner)", "Removed (Member)" or "Not a member"
function Remove-FromTeam {
    param([string]$TeamId, [string]$UserId)

    $existing = Find-Member -TeamId $TeamId -UserId $UserId
    if ($existing -and $existing.id) {
        $role = if ($existing.roles -contains "owner") { "Owner" } else { "Member" }
        Invoke-MgGraphRequest -Method DELETE -Uri "$GraphBase/teams/$TeamId/members/$($existing.id)" | Out-Null
        return "Removed ($role)"
    }
    return "Not a member"
}

# Adds the user as Owner of the team (or channel, when ChannelId is provided).
# If the user is already a member, promotes them to Owner.
function Set-Owner {
    param([string]$TeamId, [string]$ChannelId, [string]$UserId)

    $base = if ($ChannelId) { "$GraphBase/teams/$TeamId/channels/$ChannelId/members" }
            else            { "$GraphBase/teams/$TeamId/members" }

    $existing = Find-Member -TeamId $TeamId -ChannelId $ChannelId -UserId $UserId

    if ($existing -and $existing.id) {
        $body = @{
            "@odata.type" = "#microsoft.graph.aadUserConversationMember"
            roles         = @("owner")
        } | ConvertTo-Json
        Invoke-MgGraphRequest -Method PATCH -Uri "$base/$($existing.id)" -Body $body -ContentType "application/json" | Out-Null
    }
    else {
        $body = @{
            "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
            roles             = @("owner")
            "user@odata.bind" = "$GraphBase/users('$UserId')"
        } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST -Uri $base -Body $body -ContentType "application/json" | Out-Null
    }
}

$rows  = Import-Excel -Path $ExcelPath -NoHeader -StartRow 2   # P1 = Team ID, P2 = Channel ID, P3 = User
$total = @($rows).Count
$index = 0

$report = foreach ($r in $rows) {
    $index++
    $teamId    = "$($r.P1)".Trim()
    $channelId = "$($r.P2)".Trim()
    $upn       = "$($r.P3)".Trim()
    $removeStatus = "Not processed"; $teamStatus = "Not processed"; $channelStatus = "Not processed"

    Write-Host "`n[$index/$total] Processing Team: $teamId | Channel: $channelId | User: $upn" -ForegroundColor Cyan

    try {
        $user = Get-MgUser -UserId $upn -ErrorAction Stop
        Write-Host "  User found: $($user.DisplayName) ($($user.Id))"

        # Step 1: remove user from parent team (if Owner or Member), then wait
        try {
            $removeStatus = Remove-FromTeam -TeamId $teamId -UserId $user.Id
            Write-Host "  Remove from team: $removeStatus" -ForegroundColor Green
            Write-Host "  Waiting $WaitSeconds seconds..."
            Start-Sleep -Seconds $WaitSeconds
        }
        catch {
            $removeStatus = "Failed: $(Get-ErrorText $_)"
            Write-Host "  Remove from team: $removeStatus" -ForegroundColor Red
        }

        if ($removeStatus -like "Failed*") {
            $teamStatus    = "Skipped: removal step failed"
            $channelStatus = "Skipped: removal step failed"
            Write-Host "  Team owner / Channel owner: skipped" -ForegroundColor Yellow
        }
        else {
            # Step 2: add user as Owner of the parent team, then wait
            try {
                Set-Owner -TeamId $teamId -UserId $user.Id
                $teamStatus = "Success"
                Write-Host "  Team owner: Success" -ForegroundColor Green
                Write-Host "  Waiting $WaitSeconds seconds..."
                Start-Sleep -Seconds $WaitSeconds
            }
            catch {
                $teamStatus = "Failed: $(Get-ErrorText $_)"
                Write-Host "  Team owner: $teamStatus" -ForegroundColor Red
            }

            # Step 3: add user as Owner of the private channel
            if ($teamStatus -eq "Success" -and $channelId) {
                try {
                    Set-Owner -TeamId $teamId -ChannelId $channelId -UserId $user.Id
                    $channelStatus = "Success"
                    Write-Host "  Channel owner: Success" -ForegroundColor Green
                }
                catch {
                    $channelStatus = "Failed: $(Get-ErrorText $_)"
                    Write-Host "  Channel owner: $channelStatus" -ForegroundColor Red
                }
            }
            elseif (-not $channelId) {
                $channelStatus = "Skipped: no Channel ID"
                Write-Host "  Channel owner: $channelStatus" -ForegroundColor Yellow
            }
            else {
                $channelStatus = "Skipped: team owner step failed"
                Write-Host "  Channel owner: $channelStatus" -ForegroundColor Yellow
            }
        }
    }
    catch {
        $msg = "Failed: user lookup - $(Get-ErrorText $_)"
        $removeStatus = $teamStatus = $channelStatus = $msg
        Write-Host "  $msg" -ForegroundColor Red
    }

    [pscustomobject]@{
        TeamId               = $teamId
        ChannelId            = $channelId
        User                 = $upn
        RemoveFromTeamStatus = $removeStatus
        TeamOwnerStatus      = $teamStatus
        ChannelOwnerStatus   = $channelStatus
    }
}

# ---- Final report ----
$ok   = @($report | Where-Object { $_.RemoveFromTeamStatus -notlike "Failed*" -and $_.TeamOwnerStatus -eq "Success" -and $_.ChannelOwnerStatus -eq "Success" })
$fail = @($report | Where-Object { $_ -notin $ok })

Write-Host "`n===== PROCESSING REPORT =====" -ForegroundColor Cyan
Write-Host "Total rows: $(@($report).Count) | Fully successful: $($ok.Count) | With issues: $($fail.Count)"
if ($ok.Count -gt 0) {
    Write-Host "`nProcessed successfully:" -ForegroundColor Green
    $ok | Format-Table TeamId, ChannelId, User -AutoSize
}
if ($fail.Count -gt 0) {
    Write-Host "`nNot processed / with issues:" -ForegroundColor Yellow
    $fail | Format-Table -AutoSize -Wrap
}
$report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "Full report saved to: $ReportPath" -ForegroundColor Green
