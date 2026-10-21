Connect-MicrosoftTeams
 
$Results = foreach ($Team in Get-Team)
{
    $PrivateChannels = Get-TeamChannel -GroupId $Team.GroupId -MembershipType Private
 
    foreach ($Channel in $PrivateChannels)
    {
        [PSCustomObject]@{
            TeamName      = $Team.DisplayName
            TeamGroupId   = $Team.GroupId
            ChannelName   = $Channel.DisplayName
            ChannelId     = $Channel.Id
        }
    }
}
 
$Results | Export-Csv "C:\Temp\PrivateChannels.csv" -NoTypeInformation -Encoding UTF8
