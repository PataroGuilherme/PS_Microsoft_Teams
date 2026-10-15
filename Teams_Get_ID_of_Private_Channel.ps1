Connect-MgGraph -Scopes "Team.ReadBasic.All","Channel.ReadBasic.All"

# --- Parâmetros ---
$NomeDoTime  = "BRS-BE-Business Systems"
$NomeDoCanal = "Internals"

# 1) Localiza o time (Group ID)
$team = Get-MgTeam -Filter "displayName eq '$NomeDoTime'"
if (-not $team) { throw "Time não encontrado." }
Write-Host "Team ID (Group ID): $($team.Id)" -ForegroundColor Cyan

# 2) Localiza o canal privado dentro do time
$canal = Get-MgTeamChannel -TeamId $team.Id `
    -Filter "displayName eq '$NomeDoCanal' and membershipType eq 'private'"

if (-not $canal) { throw "Private Channel não encontrado." }

$canal | Select-Object Id, DisplayName, MembershipType, CreatedDateTime
Write-Host "Channel ID: $($canal.Id)" -ForegroundColor Green





# "BRS-BE-Business Systems" Object ID: 085cd020-9436-4ec8-a9bc-35be0cdc6855
# Channel ID: 19:173538c2a5084ca08c134872fa7a058a@thread.tacv2




Connect-ExchangeOnline

# --- Parâmetros ---
$Inicio     = (Get-Date).AddDays(-30)   # ajuste o período (retenção padrão: 180 dias)
$Fim        = Get-Date
$NomeCanal  = "Nome do Private Channel"
$UsuarioAfetado = "usuario@seudominio.com"   # opcional, para filtrar

# Busca eventos do Teams
$logs = Search-UnifiedAuditLog `
    -StartDate $Inicio `
    -EndDate   $Fim `
    -RecordType MicrosoftTeams `
    -ResultSize 5000 `
    -SessionCommand ReturnLargeSet

# Interpreta o JSON de AuditData
$parsed = foreach ($l in $logs) {
    $d = $l.AuditData | ConvertFrom-Json
    [PSCustomObject]@{
        DataHora     = $l.CreationDate
        Operacao     = $d.Operation
        QuemExecutou = $d.UserId
        Time         = $d.TeamName
        Canal        = $d.ChannelName
        ChannelId    = $d.ChannelId
        Membros      = ($d.Members | ForEach-Object { "$($_.UPN) [$($_.Role)]" }) -join "; "
        ClientIP     = $d.ClientIP
        AuditData    = $l.AuditData
    }
}

# Filtro pelo canal e/ou usuário afetado
$filtrado = $parsed | Where-Object {
    $_.Canal -eq $NomeCanal -or $_.AuditData -match [regex]::Escape($UsuarioAfetado)
} | Sort-Object DataHora

$filtrado | Format-Table DataHora, Operacao, QuemExecutou, Canal, Membros -AutoSize
$filtrado | Export-Csv .\AuditoriaPrivateChannel.csv -NoTypeInformation -Encoding UTF8