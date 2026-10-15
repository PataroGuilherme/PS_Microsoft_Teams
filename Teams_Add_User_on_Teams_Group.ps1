## O código abaixo é para ADICIONAR um ÚNICO usuário em um GRUPO:

Connect-MicrosoftTeams

$nomeEquipe = "Nome da Equipe"
$usuario    = "usuario@empresa.com"

$team = Get-Team -DisplayName $nomeEquipe | Where-Object { $_.DisplayName -eq $nomeEquipe }

if ($team) {
    Add-TeamUser -GroupId $team.GroupId -User $usuario -Role Member #or Owner
    Write-Host "Usuário $usuario adicionado à equipe $nomeEquipe." -ForegroundColor Green
} else {
    Write-Host "Equipe '$nomeEquipe' não encontrada." -ForegroundColor Red
}


## O código abaixo é para ADICIONAR uma LISTA de usuários ao MESMO GRUPO

$groupId = "00000000-0000-0000-0000-000000000000"

Import-Csv "C:\temp\usuarios.csv" | ForEach-Object {
    try {
        Add-TeamUser -GroupId $groupId -User $_.Email
        Write-Host "OK: $($_.Email)" -ForegroundColor Green
    } catch {
        Write-Host "Erro em $($_.Email): $($_.Exception.Message)" -ForegroundColor Red
    }
}
