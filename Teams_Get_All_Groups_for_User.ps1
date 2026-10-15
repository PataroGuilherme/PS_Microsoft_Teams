Install-Module Microsoft.Graph -Scope CurrentUser

Connect-MgGraph -Scopes "Group.Read.All"

# Ensure you are connected to Microsoft Graph
Connect-MgGraph -Scopes "Group.Read.All"

# Replace with the user's UPN (email address)
$userUPN = "user.name@email.com"


# Get the user's object ID
$user = Get-MgUser -Filter "userPrincipalName eq '$userUPN'" | Select-Object -First 1

if ($user -eq $null) {
    Write-Host "User not found. Please check the UPN."
    return
}

# Use Microsoft Graph API to query groups where the user is a member
$requestUrl = "https://graph.microsoft.com/v1.0/users/$($user.Id)/memberOf/microsoft.graph.group"
$response = Invoke-MgGraphRequest -Method GET -Uri $requestUrl

# Parse the response to get Teams
$groups = $response.value | Where-Object { $_.resourceProvisioningOptions -contains "Team" }

if ($groups.Count -eq 0) {
    Write-Host "No Teams found for the user."
} else {
    Write-Host "Teams the user has access to:"
    $groups | ForEach-Object { Write-Host "- $($_.displayName)" }
}
