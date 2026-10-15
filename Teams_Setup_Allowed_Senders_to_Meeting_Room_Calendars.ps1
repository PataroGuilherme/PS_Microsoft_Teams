### Verificar se a sala já possui alguma regra específica:

Connect-ExchangeOnline

Get-CalendarProcessing -Identity "ROOM_NAME" |
    Format-List
	
	
### Aplicar uma regra específica:

-- In this one we can apply individual people or distribution lists;
-- on Exchange Admin Panel, we need keep 'Accept or decline automatically' option enabled, because
the idea here is allow/decline automatically based on this policy.


Set-CalendarProcessing -Identity "ROOM_NAME"" `
    -AutomateProcessing AutoAccept `
    -BookInPolicy "SECURITY_GROUP_NAME" `
    -AllBookInPolicy $false `
    -AllRequestInPolicy $false `
    -AllRequestOutOfPolicy $false
	
	
	Set-CalendarProcessing -Identity "ROOM_NAME" -AutomateProcessing AutoAccept -BookInPolicy "SECURITY_GROUP_NAME" -AllBookInPolicy $false -AllRequestInPolicy $false -AllRequestOutOfPolicy $false
