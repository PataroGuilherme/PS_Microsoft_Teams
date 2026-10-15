# Usuário que será consultado
$UserUPN = "usuario1@empresa.com"

# Conectar ao Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "Group.Read.All"

# Obter o ID do usuário
$User = Get-MgUser -UserId $UserUPN

# Buscar grupos dos quais o usuário é membro
$MemberOf = Get-MgUserMemberOf -UserId $User.Id -All |
    Where-Object {
        $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group"
    }

# Obter detalhes dos grupos
$Results = foreach ($Group in $MemberOf) {

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
$Results | Export-Csv ".\Teams_Member_Of_User.csv" -NoTypeInformation -Encoding UTF8
