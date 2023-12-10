
$users = Import-Excel -Path .\new_users.xlsx

$domain_username = Read-Host -Prompt "Enter YOUR ADMIN domain\username"
$credientials = Get-Credential -UserName $domain_username -Message 'Enter Admin Password'

foreach($user in $users){
    try {
        $name = $user.Name.split(" ")
        $firstname =  $name[0]
        $lastname = $name[1]
        
        $Name = $user.Name
        $FirstName = $firstname
        $LastName  = $lastname
        $Password = "Welcome1"
        $Username = $firstname[0] + $lastname
        $Email = $user.Email
        $OU = "OU=ExternalTempUsers,OU=DOI Users,DC=DistinguishedName"
        $AccountEnabled = $true

        $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

        New-ADUser -Name $Name -UserPrincipalName "$Username@doi.nyc.gov" -SamAccountName $Username -EmailAddress $Email -AccountPassword $SecurePassword -Enabled $AccountEnabled -Path $OU -GivenName $FirstName -Surname $LastName -DisplayName $Name -Credential $credientials -ErrorAction Stop
    }
    catch {
        $Username = $firstname[0][1] + $lastname
        New-ADUser -Name $Name -UserPrincipalName "$Username@doi.nyc.gov" -SamAccountName $Username -EmailAddress $Email -AccountPassword $SecurePassword -Enabled $AccountEnabled -Path $OU -GivenName $FirstName -Surname $LastName -DisplayName $Name -Credential $credientials -ErrorAction Stop
    }
    
}


