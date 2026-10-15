#To create Room Resources the following information is needed:
#- Meeting Room Name
#- Country
#- City
#- Building Name
#- Building Address (only include street and number, not city or additional information)
#- Floor
#- Postal Code
#- Capacity (numbers only)
#- Wheel chair accessibility (yes or no)
#- Room display - Example "LG 55 4K" (optional)
#- Room audio device - Example "Bose 7.1" (optional)
#- Room video device - Example "Poly Studio P5" (optional)

Connect-ExchangeOnline

2) Create Mailbox / Room Properly:
New-Mailbox -Name "ROOM_NAME" -Room -DisplayName "ROOM_DISPLAY_NAME" -PrimarySmtpAddress "ROOM_NAME@domain.com"

3) Set Password: (ONLY IF AN MS TEAMS ROOM PRO LICENSE IS ASSIGNED)
Set-Mailbox -identity "ROOM_NAME" -RoomMailboxPassword (ConvertTo-SecureString -String 'Y2HnHGb8e' -AsPlainText -Force) -EnableRoomMailboxAccount $true

4) Add freshly created Room (Mailbox) to Resource List:
Add-DistributionGroupMember -Identity "LIST_NAME" -Member "ROOM_NAME"
# This List Name mentioned is the list that appears on Teams Console, if the company applied 'Rooms per location'

5) Set Room’s Metadata accordingly:
Set-Place -Identity "ROOM_NAME" -CountryOrRegion "FR" -City "City Name" -Street "Av. Example Name" -Building "FR Office" -Floor 5 -FloorLabel "5" -PostalCode "34854" -Capacity 4
Set-Place -Identity "ROOM_NAME" -IsWheelChairAccessible $True

6) Set needed MTR Attributes accordingly:
Set-CalendarProcessing -Identity 'ROOM_NAME' -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -AllowConflicts $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse 'MS Teams Room Enabled - Capacity: 4 People' -ProcessExternalMeetingMessages $true

7) Check Room Mailbox attributes
Get-CalendarProcessing -Identity 'ROOM_NAME@domain.com' | fl

8) Add a new equipment in a existing Room:
Set-Place -Identity "ROOM_NAME" -AudioDeviceName "" -DisplayDeviceName "" -VideoDeviceName ""

9) Check the Distribution Lists for Teams Rooms already created:
Get-DistributionGroup -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "RoomList"} | Format-Table DisplayName,Identity,PrimarySmtpAddress -AutoSize

10) Check Teams Room properties:
Get-Place -identity ROOM_NAME | Get-Place | Select-Object AudiodeviceName, DisplayDeviceName, VideoDeviceName, CountryOrRegion , City, Street, Capacity, Floor, PostalCode, Building, IsWheelChairAccessible

11) Get All Teams Rooms created:
Get-Mailbox  -ResultSize Unlimited | Where {$_.RecipientTypeDetails -eq "Room"} | Format-Table DisplayName,Identity,PrimarySmtpAddress –AutoSize

12) Create a Resource List for Meeting Rooms:
New-DistributionGroup -Name "DISTRIBUTION_LIST" -DisplayName "My Office" -ModerationEnabled $false -PrimarySmtpAddress "DISTRIBUTION_LIST@domain.com" -MemberDepartRestriction Closed -MemberJoinRestriction Closed -RoomList

13) Delete a Resource List:
Remove-DistributionGroup -Identity "DISTRIBUTION_LIST" -Confirm:$false

14) Modify Resource List for Teams Rooms:
Set-DistributionGroup -Name "DISTRIBUTION_LIST" -DisplayName "My New Office" -ModerationEnabled $false -PrimarySmtpAddress "DISTRIBUTION_LIST@domain.com" -MemberDepartRestriction Closed -MemberJoinRestriction Closed -RoomList

16) Assing Delegation for a specific Room:
Set-CalendarProcessing -Identity "ROOM_NAME_01" -AutomateProcessing AutoAccept -BookInPolicy "user1@domain.com","user2@domain.com","user3@domain.com","user4@domain.com" -AllBookInPolicy $false

17) Modify Meeting Acceptance for a Room:
Set-CalendarProcessing -Identity "ROOM_NAME" `
  -AutomateProcessing AutoAccept `
  -AllowConflicts $true `
  -BookingWindowInDays 365 `
  -DeleteSubject $false `
  -AddOrganizerToSubject $true

  
18) Verify Room's Calendar permission:
Get-MailboxFolderPermission -Identity "ROOM_NAME:\Calendar"
#The word 'Calendar' can be changed accordingly the local language. If the tenant is on Portuguese, could be 'Calendario'

19) Add full Room's Calendar permissions for a user :
Connect-ExchangeOnline -UserPrincipalName user1@domain.com
Add-MailboxFolderPermission -Identity "ROOM_NAME:\Calendar" -User user1@domain.com -AccessRights Editor

20) Remove full Room's Calendar permissions for a user
Remove-MailboxFolderPermission -Identity "ROOM_NAME:\Calendar" -User user1@domain.com
