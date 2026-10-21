# Para executar este código, é necessário:
# 1. Criar um arquivo .txt na pasta de Downloads, com nome "groups.txt";
# 2. Neste arquivo txt criado, temos que colocar todos os Objects IDs de todos os Grupos do Teams (e para coletar esses IDs, basta ir no Entra ID e baixar um report de todos os grupos do tipo M365).
# 3. Este arquivo PowerShell precisa ser colocado na pasta de Downloads junto com o arquivo txt.

# Comando de execução via CMD: 
# powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\users\guilherme.pataro\downloads\Teams_Set_Owner_to_Private_Channel_based_on_Group_ownership.ps1.ps1" -Path "C:\users\guilherme.pataro\downloads\groups.txt"

############################
##### ETAPAS:
# 1. Este arquivo: 
# - Lerá o arquivo TXT e tomará cada Object ID;
# - Se o grupo não tiver Private Channels, ele pula para o próximo.
# - Se tiver Private Channel, ele coleta o primeiro Owner do grupo Pai > Remove-o > Aguarda 3 segundos > Readiciona > Lista todos os Private Channels > Adiciona o Owner como novo Owner em cada Private Channel.



<#
.SYNOPSIS
    For each Team in a list: removes the first Owner, waits 3 seconds, reassigns that user
    as Owner, and sets the same user as Owner in all of the Team's Private Channels.

.DESCRIPTION
    For each group provided:
      1. Identifies the first Owner (user) of the group.
      2. Removes that user from the group's Owners.
      3. Waits 3 seconds.
      4. Reassigns the same user as Owner of the group.
         (Steps 2-4 are skipped when the group has only 1 owner.)
      5. Lists the Team's Private Channels (if any).
      6. Adds the same user as Owner in all Private Channels
         (if already a channel member, the user is promoted to Owner).

    Notes:
      - Only the Owner role is removed/re-added; the user remains a member of the group.
      - "First owner" = first user returned by Graph in /groups/{id}/owners.
      - If a group has only 1 owner, that owner is NOT removed from the parent group (Graph
        requires at least one owner); the user is added directly as Owner in all Private Channels.
      - Groups synced from on-premises AD are SKIPPED (owners cannot be edited via Graph).
      - An error in one group does not stop the others.
      - At the end, a CSV is exported with the selected user, the processed group, and the
        Private Channels the user was added to.

.PARAMETER GroupIds
    One or more group Object IDs. Any GUID found in the text is accepted.
    From cmd, prefer -Path when the list is large.

.PARAMETER Path
    Text/CSV file with the Object IDs (one per line or comma-separated).
    Lines starting with # are ignored.

.PARAMETER WaitSeconds
    Wait time between removing and re-adding the owner (default: 3).

.PARAMETER ReportPath
    Optional. Path of the .csv exported at the end. If omitted, the file
    Report_Owners_<date>_<time>.csv is written to the same folder as the script.
    The report is generated in ENGLISH, with one row per group/private channel and the columns:
    User, UserUPN, Group, GroupId, PrivateChannel, ChannelId, Result, GroupStatus.

.PARAMETER Delimiter
    CSV delimiter. Default: the Windows list separator (";" in pt-BR),
    so Excel opens the columns correctly.

.PARAMETER Force
    Skips the confirmation prompt before starting.

.PARAMETER WhatIf
    Simulates the run: only reads data and shows what would be done, without changing anything.

.REQUIREMENTS
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

    Graph permissions (admin consent required):
        Group.ReadWrite.All, GroupMember.ReadWrite.All,
        Channel.ReadBasic.All, ChannelMember.ReadWrite.All

.EXAMPLE
    .\Reset-FirstOwnerInTeams.ps1 -Path .\groups.txt -WhatIf

.EXAMPLE
    .\Reset-FirstOwnerInTeams.ps1 -GroupIds "00000000-0000-0000-0000-000000000000","11111111-1111-1111-1111-111111111111" -Force -ReportPath C:\Scripts\report.csv
#>

param(
    [string[]]$GroupIds,
    [string]$Path,
    [ValidateRange(0, 300)]
    [int]$WaitSeconds = 3,
    [string]$ReportPath,
    [string]$Delimiter = (Get-Culture).TextInfo.ListSeparator,
    [switch]$Force,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$base = "https://graph.microsoft.com/v1.0"

# ---------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------
function Get-GraphAll {
    # GET with paging (@odata.nextLink)
    param([string]$Uri)
    $items = @()
    while ($Uri) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri
        if ($resp.value) { $items += $resp.value }
        $Uri = $resp.'@odata.nextLink'
    }
    $items
}

function Get-ErrMsg {
    # Detailed error message (includes the Graph error body and the line number)
    param($Err)
    $m = $Err.Exception.Message
    if ($Err.ErrorDetails -and $Err.ErrorDetails.Message) { $m = "$m | $($Err.ErrorDetails.Message)" }
    "$m (line $($Err.InvocationInfo.ScriptLineNumber))"
}

function Invoke-WithRetry {
    # Retries the operation in case of replication delay
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        try {
            return (& $Action)
        }
        catch {
            if ($i -eq $MaxAttempts) { throw }
            Write-Host "   Tentativa $i falhou. Nova tentativa em $DelaySeconds s ..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

# ---------------------------------------------------------------
# 1) Group list
# ---------------------------------------------------------------
$guidPattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
$texto = ""

if ($Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Arquivo '$Path' nao encontrado." }
    $linhas = Get-Content -LiteralPath $Path | Where-Object { $_.Trim() -notmatch '^#' }
    $texto += ($linhas -join "`n") + "`n"
}
if ($GroupIds) { $texto += ($GroupIds -join "`n") }

$ids = @([regex]::Matches($texto, $guidPattern) | ForEach-Object { $_.Value.ToLower() } | Select-Object -Unique)

if ($ids.Count -eq 0) {
    throw "Nenhum Object ID valido informado. Use -GroupIds ou -Path."
}

Write-Host "`n$($ids.Count) grupo(s) na lista." -ForegroundColor Cyan

# ---------------------------------------------------------------
# Connection to Microsoft Graph
# ---------------------------------------------------------------
$scopes = @(
    "Group.ReadWrite.All",
    "GroupMember.ReadWrite.All",
    "Channel.ReadBasic.All",
    "ChannelMember.ReadWrite.All"
)

if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes $scopes -NoWelcome
}

if ($WhatIf) { Write-Host "`n*** MODO SIMULACAO (-WhatIf): nada sera alterado ***" -ForegroundColor Magenta }

if (-not $Force -and -not $WhatIf) {
    Write-Host "ATENCAO: o primeiro owner de cada um dos $($ids.Count) grupo(s) sera removido e readicionado como Owner," -ForegroundColor Yellow
    Write-Host "         e definido como Owner em todos os Private Channels." -ForegroundColor Yellow
    $resp = Read-Host "Digite SIM para continuar"
    if ($resp -ne "SIM") {
        Write-Host "Operacao cancelada." -ForegroundColor Yellow
        return
    }
}

# ---------------------------------------------------------------
# 2 to 6) Process each group
# ---------------------------------------------------------------
$resumo   = New-Object System.Collections.Generic.List[object]
$detalhes = New-Object System.Collections.Generic.List[object]
$n = 0

foreach ($gid in $ids) {
    $n++
    $grpNome     = "?"
    $ownerUpn    = "?"
    $ownerNome   = "?"
    $ownerId     = $null
    $removido    = $false
    $readicionado = $false
    $canaisTotal = 0
    $canaisOk    = 0
    $canaisErro  = 0
    $status      = "OK"
    $singleOwner = $false
    $notaUnico   = ""

    Write-Host "`n=== [$n/$($ids.Count)] Grupo $gid ===" -ForegroundColor Cyan

    try {
        # --- Group ---
        $group = Invoke-MgGraphRequest -Method GET -Uri ($base + '/groups/' + $gid + '?$select=id,displayName,onPremisesSyncEnabled')
        $grpNome = $group.displayName
        Write-Host "Grupo: $grpNome"

        if ($group.onPremisesSyncEnabled -eq $true) {
            throw "SKIPPED: group is synced from on-premises AD (owners cannot be edited via Graph)."
        }

        # --- Check Private Channels before processing the parent group ---
        $privateChannels = @(Get-GraphAll ($base + '/teams/' + $gid + '/channels') |
            Where-Object { $_.membershipType -eq "private" })
        $canaisTotal = $privateChannels.Count

        if ($canaisTotal -eq 0) {
            Write-Host "Grupo sem Private Channels. Pulando para o proximo grupo pai." -ForegroundColor DarkGray
            continue
        }

        # --- 2) First owner ---
        $owners = @(Get-GraphAll ($base + '/groups/' + $gid + '/owners?$select=id,displayName,userPrincipalName'))
        # With only one owner, Graph does not allow removing it (a group needs at least one owner):
        # keep it in the parent group and only add it to the private channels
        $singleOwner = ($owners.Count -le 1)

        $primeiro = $owners | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' } | Select-Object -First 1
        if (-not $primeiro) {
            throw "SKIPPED: no user-type owner found."
        }

        $ownerId  = $primeiro.id
        $ownerUpn = $primeiro.userPrincipalName
        $ownerNome = $primeiro.displayName
        if ($singleOwner) { $notaUnico = " (single owner: not removed from group)" }
        Write-Host "Primeiro owner: $($primeiro.displayName) <$ownerUpn> (ID: $ownerId)" -ForegroundColor Green

        if ($singleOwner) {
            # Only one owner: do not remove it from the parent group (skips steps 2-4)
            Write-Host "Grupo com apenas 1 owner: $ownerUpn NAO sera removido do grupo pai; sera adicionado diretamente nos Private Channels." -ForegroundColor Yellow
        }
        else {
            # --- 2) Remove the owner ---
            Write-Host "Removendo $ownerUpn da lista de owners ..." -ForegroundColor Cyan
            if ($WhatIf) {
                Write-Host "   (Simulado: remocao, espera de $WaitSeconds s e readicao como owner)" -ForegroundColor DarkGray
            }
            else {
                Invoke-MgGraphRequest -Method DELETE -Uri ($base + '/groups/' + $gid + '/owners/' + $ownerId + '/$ref') | Out-Null
                $removido = $true
                Write-Host "   Removido como Owner." -ForegroundColor Green

                # --- 3) Wait ---
                Write-Host "Aguardando $WaitSeconds segundos ..." -ForegroundColor Cyan
                Start-Sleep -Seconds $WaitSeconds

                # --- 4) Reassign as owner ---
                Write-Host "Reatribuindo $ownerUpn como owner ..." -ForegroundColor Cyan
                $refBody = '{"@odata.id":"' + $base + '/directoryObjects/' + $ownerId + '"}'
                Invoke-WithRetry -Action {
                    Invoke-MgGraphRequest -Method POST -Uri ($base + '/groups/' + $gid + '/owners/$ref') -Body $refBody -ContentType "application/json" | Out-Null
                }
                $readicionado = $true
                Write-Host "   Adicionado como Owner." -ForegroundColor Green
            }
        }

        # --- 5) List Private Channels ---
        Write-Host "Private Channels: $canaisTotal" -ForegroundColor Cyan
        if ($canaisTotal -gt 0) {
            $privateChannels |
                ForEach-Object { [PSCustomObject]@{ Canal = $_.displayName; ChannelId = $_.id } } |
                Format-Table -AutoSize
        }

        # --- 6) Add as Owner in all Private Channels ---
        foreach ($channel in $privateChannels) {
            $chStatus = $null
            $chUri    = $base + '/teams/' + $gid + '/channels/' + $channel.id + '/members'
            try {
                $membrosCanal = @(Get-GraphAll $chUri)
                $existente = $membrosCanal | Where-Object { $_.userId -eq $ownerId } | Select-Object -First 1

                if ($existente) {
                    if (@($existente.roles) -contains "owner") {
                        $chStatus = "Already owner (no action needed)"
                    }
                    elseif ($WhatIf) {
                        $chStatus = "Simulated: would promote to Owner"
                    }
                    else {
                        $patch = '{"@odata.type":"#microsoft.graph.aadUserConversationMember","roles":["owner"]}'
                        Invoke-MgGraphRequest -Method PATCH -Uri ($chUri + '/' + $existente.id) -Body $patch -ContentType "application/json" | Out-Null
                        $chStatus = "Promoted to Owner"
                    }
                }
                elseif ($WhatIf) {
                    $chStatus = "Simulated: would add as Owner"
                }
                else {
                    # JSON as plain text (ConvertTo-Json in PowerShell 5.1 can serialize arrays as an object)
                    $body = '{{"@odata.type":"#microsoft.graph.aadUserConversationMember","roles":["owner"],"user@odata.bind":"{0}/users(''{1}'')"}}' -f $base, $ownerId
                    Invoke-WithRetry -Action {
                        Invoke-MgGraphRequest -Method POST -Uri $chUri -Body $body -ContentType "application/json" | Out-Null
                    }
                    $chStatus = "Added as Owner"
                }
                $canaisOk++
            }
            catch {
                $chStatus = "ERROR: " + (Get-ErrMsg $_)
                $canaisErro++
            }

            $detalhes.Add([PSCustomObject]@{
                GrupoId   = $gid
                Grupo     = $grpNome
                Owner     = $ownerUpn
                Canal     = $channel.displayName
                ChannelId = $channel.id
                Resultado = $chStatus
            })

            $cor = if ($chStatus -like "ERROR*") { "Red" } else { "Green" }
            Write-Host "   [$($channel.displayName)] $chStatus" -ForegroundColor $cor
        }

        if ($canaisErro -gt 0) { $status = "COMPLETED WITH CHANNEL ERRORS" }
        $status = $status + $notaUnico
    }
    catch {
        $msg = Get-ErrMsg $_
        if ($removido -and -not $readicionado) {
            $msg = "WARNING - OWNER REMOVED AND NOT RE-ADDED! Re-add $ownerUpn manually. " + $msg
            Write-Host $msg -ForegroundColor Red
        }
        else {
            Write-Host $msg -ForegroundColor Yellow
        }
        $status = $msg
    }

    $resumo.Add([PSCustomObject]@{
        GrupoId       = $gid
        Grupo         = $grpNome
        Owner         = $ownerUpn
        OwnerNome     = $ownerNome
        Removido      = $removido
        Readicionado  = $readicionado
        CanaisPrivados = $canaisTotal
        CanaisOk      = $canaisOk
        CanaisErro    = $canaisErro
        Status        = $status
    })
}

# ---------------------------------------------------------------
# Final summary
# ---------------------------------------------------------------
Write-Host "`n================ RESUMO ================" -ForegroundColor Cyan
$resumo | Select-Object Grupo, Owner, Removido, Readicionado, CanaisPrivados, CanaisOk, CanaisErro | Format-Table -AutoSize

$problemas = @($resumo | Where-Object { $_.Status -notlike "OK*" })
if ($problemas.Count -gt 0) {
    Write-Host "Grupos com observacoes ($($problemas.Count)):" -ForegroundColor Yellow
    $problemas | Select-Object GrupoId, Grupo, Status | Format-List
}

$errosCanal = @($detalhes | Where-Object { $_.Resultado -like "ERROR*" })
if ($errosCanal.Count -gt 0) {
    Write-Host "Canais com erro ($($errosCanal.Count)):" -ForegroundColor Yellow
    $errosCanal | Select-Object Grupo, Canal, Resultado | Format-List
}

# ---------------------------------------------------------------
# Export the final CSV: selected user, processed group, and private channels
# ---------------------------------------------------------------
if (-not $ReportPath) {
    $pasta = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $ReportPath = Join-Path $pasta ("Report_Owners_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv")
}
# Ensures a full path (in case -ReportPath was given as a relative path)
$ReportPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportPath)

$relatorio = foreach ($r in $resumo) {
    $dets = @($detalhes | Where-Object { $_.GrupoId -eq $r.GrupoId })
    if ($dets.Count -gt 0) {
        foreach ($d in $dets) {
            [PSCustomObject]@{
                User           = $r.OwnerNome
                UserUPN        = $r.Owner
                Group          = $r.Grupo
                GroupId        = $r.GrupoId
                PrivateChannel = $d.Canal
                ChannelId      = $d.ChannelId
                Result         = $d.Resultado
                GroupStatus    = $r.Status
            }
        }
    }
    else {
        $res = if ($r.Status -like "OK*") { "No private channels" } else { $r.Status }
        [PSCustomObject]@{
            User           = $r.OwnerNome
            UserUPN        = $r.Owner
            Group          = $r.Grupo
            GroupId        = $r.GrupoId
            PrivateChannel = ""
            ChannelId      = ""
            Result         = $res
            GroupStatus    = $r.Status
        }
    }
}

try {
    $relatorio | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
}
catch {
    # E.g., file open in Excel -> write to a different name
    $alt = Join-Path ([System.IO.Path]::GetDirectoryName($ReportPath)) ("Report_Owners_" + (Get-Date -Format "yyyyMMdd_HHmmssfff") + ".csv")
    Write-Host "Nao foi possivel gravar '$ReportPath' ($($_.Exception.Message)). Tentando '$alt' ..." -ForegroundColor Yellow
    $relatorio | Export-Csv -LiteralPath $alt -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
    $ReportPath = $alt
}

Write-Host "`nCSV report exported to:`n  $ReportPath" -ForegroundColor Green
