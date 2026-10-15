# Este script é usado para Salas do tipo Room que por algum motivo mostram o nome do usuário que marcou a reunião no convite geral, não o nome da sala em si
# Este script somente é aplicado quando no tablet (TC10 ou outros) aparecem o erro demonstrado acima.

Connect-ExchangeOnline

Set-CalendarProcessing -Identity "NomeDaSala" `
  -AddOrganizerToSubject $false `
  -DeleteSubject $false `
  -RemovePrivateProperty $false
  
Get-CalendarProcessing -Identity "NomeDaSala" | Format-List Identity,AddOrganizerToSubject,DeleteSubject,RemovePrivateProperty
