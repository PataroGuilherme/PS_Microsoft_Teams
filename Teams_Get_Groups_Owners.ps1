# Conectar ao Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All","User.Read.All"

# ===============================
# INFORME AQUI OS NOMES DOS GRUPOS
# ===============================
$GroupNames = @(
"Group_Name_1",
"Group_Name_2"
)

# Array para armazenar os resultados
$Results = @()

foreach ($GroupName in $GroupNames) {

    Write-Host "Processando grupo: $GroupName" -ForegroundColor Cyan

    # Buscar grupo pelo DisplayName
    $Group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ConsistencyLevel eventual

    if (-not $Group) {
        Write-Warning "Grupo não encontrado: $GroupName"
        continue
    }

    # Buscar Owners do grupo
    $Owners = Get-MgGroupOwner -GroupId $Group.Id

    if (-not $Owners) {
        $Results += [PSCustomObject]@{
            GroupName = $GroupName
            Owner     = "Sem Owners"
            Email     = ""
        }
        continue
    }

    foreach ($Owner in $Owners) {
        $Results += [PSCustomObject]@{
            GroupName = $GroupName
            Owner     = $Owner.AdditionalProperties.displayName
            Email     = $Owner.AdditionalProperties.mail
        }
    }
}

# Exibir resultado
$Results | Format-Table -AutoSize
