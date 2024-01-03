$users = Import-Excel -Path .\new_users.xlsx

$authenticate = $true
$attempts = 3
while ($authenticate) {
    $domain_username = Read-Host -Prompt "Enter YOUR ADMIN domain\username"
    $credientials = Get-Credential -UserName $domain_username -Message 'Enter Admin Password'
    try {
        $session = New-PSSession -ComputerName 'servername' -Credential $credientials -ErrorAction Stop
        Remove-PSSession $session
        Write-Host "Authentication successful" -ForegroundColor Green
        $authenticate = $false
    } catch {
        $attempts = $attempts - 1
        if ($attempts -eq 0){
            Write-Host "Too many failed attempts. Exiting console." -ForegroundColor Red
            exit
        }
        Write-Host "Failed to authenticate please try again. $attempts attempts remaining." -ForegroundColor Red
    }
}

foreach($user in $users){
    try {
        $name = $user.Name.split(" ")
        $FirstName = $name[0]
        $LastName  = $name[1]
        $Name = $user.Name
        $Password = "Welcome1"
        $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
        $Username = $Firstname[0].tostring() + $LastName
        $Email = $user.Email
        $AccountEnabled = $true

        if ($users.Squad -eq "squad1"){$OU = "OU=Squad1,OU=DOI Users,DC=DOI,DC=NYCNET"}
        elseif ($users.Squad -eq "squad2"){$OU = "OU=Squad2,OU=DOI Users,DC=DOI,DC=NYCNET"}
        elseif ($users.Squad -eq "squad3"){$OU = "OU=Squad3,OU=DOI Users,DC=DOI,DC=NYCNET"}
        elseif ($users.Squad -eq "squad4"){$OU = "OU=Squad4,OU=DOI Users,DC=DOI,DC=NYCNET"}
        elseif ($users.Squad -eq "squad5"){$OU = "OU=Squad5,OU=DOI Users,DC=DOI,DC=NYCNET"} 
        elseif ($users.Squad -eq "squad6"){$OU = "OU=Squad6,OU=DOI Users,DC=DOI,DC=NYCNET"}
        else {$OU = "OU=TestUsers,OU=DOI Users,DC=DOI,DC=NYCNET"}

        #Sets up New User.
        New-ADUser `
            -Name $Name `
            -UserPrincipalName "$Username@doi.nyc.gov" `
            -SamAccountName $Username `
            -EmailAddress $Email `
            -AccountPassword $SecurePassword `
            -Enabled $AccountEnabled `
            -Path $OU `
            -GivenName $FirstName `
            -Surname $LastName `
            -DisplayName $Name `
            -Credential $credentials `
            -ErrorAction Stop

            #Adds the user to their Main group
            if ($OU -eq "OU=Squad1,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad1 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad2,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad2 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad3,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad3 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad4,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad4 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad5,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad5 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad6,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad6 -Members $Username -Credential $credientials}
            else {$OU = "OU=TestUsers,OU=DOI Users,DC=DOI,DC=NYCNET"}
    
        #Enables Remote Mailbox on Account.
        Invoke-Command -ComputerName "servername" -Credential $credentials -ScriptBlock {
            $aduser = Get-Aduser -Identity $using:Username -Properties *
            Enable-RemoteMailbox -Identity $aduser.DisplayName -RemoteRoutingAddress $using:Username@nycdoi365.mail.onmicrosoft.com -Credential $credientials
        }
    }
    catch {
        #If the username already exist then this block is ran.
        $Username = $Firstname[0][1].tostring() + $LastName
        New-ADUser `
            -Name $Name `
            -UserPrincipalName "$Username@doi.nyc.gov" `
            -SamAccountName $Username `
            -EmailAddress $Email `
            -AccountPassword $SecurePassword `
            -Enabled $AccountEnabled `
            -Path $OU `
            -GivenName $FirstName `
            -Surname $LastName `
            -DisplayName $Name `
            -Credential $credentials

            #Adds the user to their Main group
            if ($OU -eq "OU=Squad1,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad1 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad2,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad2 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad3,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad3 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad4,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad4 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad5,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad5 -Members $Username -Credential $credientials}
            elseif ($OU -eq "OU=Squad6,OU=DOI Users,DC=DOI,DC=NYCNET"){Add-ADGroupMember -Identity Squad6 -Members $Username -Credential $credientials}
            else {$OU = "OU=TestUsers,OU=DOI Users,DC=DOI,DC=NYCNET"}

        #Enables Remote Mailbox on Account.
        Invoke-Command -ComputerName "servername" -Credential $credentials -ScriptBlock {
            $aduser = Get-Aduser -Identity $using:Username -Properties *
            Enable-RemoteMailbox -Identity $aduser.DisplayName -RemoteRoutingAddress $using:Username@nycdoi365.mail.onmicrosoft.com -Credential $credientials
        } 
    }
}



