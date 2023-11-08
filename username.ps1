$users = Import-Excel -Path .\new_users.xlsx

$accountinfo = @()
foreach($user in $users){
    $account = Get-ADUser -Filter "Displayname -eq '$($user.Name)'" -Properties * | Select-Object Name, SamAccountName, EmailAddress, UserPrincipalName 
    $accountinfo += $account
}

$accountinfo | Export-Excel -Path .\username.xlsx