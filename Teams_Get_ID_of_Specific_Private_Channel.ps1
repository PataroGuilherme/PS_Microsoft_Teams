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
