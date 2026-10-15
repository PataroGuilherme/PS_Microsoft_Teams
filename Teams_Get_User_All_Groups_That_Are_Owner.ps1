sad# Usuário que será consultado
$UserUPN = "user@domain.com"

# Conectar ao Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All"

# Obter o ID do usuário
$User = Get-MgUser -UserId $UserUPN

# Buscar grupos M365 onde o usuário é Owner
$OwnedGroups = Get-MgUserOwnedObject -UserId $User.Id -All |
    Where-Object {
        $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group"
    }

# Obter detalhes dos grupos
$Results = foreach ($Group in $OwnedGroups) {

    $GroupDetails = Get-MgGroup -GroupId $Group.Id

    # Considerar somente Microsoft 365 Groups
    if ($GroupDetails.GroupTypes -contains "Unified") {

        [PSCustomObject]@{
            GroupName     = $GroupDetails.DisplayName
            GroupId       = $GroupDetails.Id
            Mail          = $GroupDetails.Mail
            MailNickname  = $GroupDetails.MailNickname
        }
    }
}

# Exibir resultado
$Results | Format-Table -AutoSize

# Opcional: exportar para CSV
$Results | Export-Csv ".\Teams_Owned_By_User.csv" -NoTypeInformation -Encoding UTF8
