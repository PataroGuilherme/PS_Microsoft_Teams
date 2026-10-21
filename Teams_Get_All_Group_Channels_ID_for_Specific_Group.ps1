Connect-MgGraph -Scopes 'Channel.ReadBasic.All' -NoWelcome



$TeamId = 'OBJECT-ID'   # Object ID do grupo

Get-MgTeamChannel -TeamId $TeamId -All |
    Select-Object DisplayName, MembershipType, Id



