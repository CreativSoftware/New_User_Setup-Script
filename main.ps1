#Pulls information from Excel spreadsheet
$users = Import-Excel -Path .\new_users.xlsx

#Verifies Authentication
$authenticate = $true
$attempts = 3
while ($authenticate) {
    $domain_username = Read-Host -Prompt "Enter YOUR ADMIN domain\username"
    $credentials = Get-Credential -UserName $domain_username -Message 'Enter Admin Password'
    try {
        $session = New-PSSession -ComputerName '' -Credential $credentials -ErrorAction Stop
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

#Global Varibles
$Password = ""
$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
$AccountEnabled = $true
$Street = ''
$City = ''
$Zip = ''
$Country = ""
$State = ''
$Company = ''

#Reads each Row in Spreadsheet and creates the variables for the account creation
foreach($user in $users){
    $Name = $user.Name
    $Email = $user.Email
    $JobTitle = $user.Title
    $Department = $user.Squad

    $fullname = $user.Name.split(" ")
    $FirstName = $fullname[0]
    $LastName  = $fullname[1]
    $Username = $FirstName[0].tostring() + $LastName
    
    $Mname = $user.Manager.split(" ")
    $Mfirst = $Mname[0]
    $Mlast = $Mname[1]
    $Manager = Get-ADUser -Filter {GivenName -eq $Mfirst -and Surname -eq $Mlast} | Select-Object -First 1 |Select-Object -ExpandProperty SamAccountName

    if ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""} 
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    elseif ($Department -eq ""){$OU = ""}
    else {$OU = ""}

    #Creates New ADUser Account
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
        -StreetAddress $Street `
        -City $City `
        -PostalCode $Zip `
        -Country $Country `
        -State $State `
        -Title $JobTitle `
        -Department $Department `
        -Manager $Manager `
        -Company $Company `
        -HomeDirectory "\\home_folder\$Username" `
        -HomeDrive 'I:' `
        -Credential $credentials
        
    #Adds Specific Membership Groups
    if ($OU -eq ""){
        $groups = @('', '', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }    
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }  
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('','', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('', '', '', '')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    elseif ($OU -eq ""){
        $groups = @('')
        foreach($group in $groups){
            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
        }
    }
    
    #Adds General Membership Groups
    $groups = @('', '', '', '', '', '')
    foreach($group in $groups){
        Add-ADGroupMember `
            -Identity $group `
            -Members @($Username) `
            -Credential $credentials
    }

    #Connects to Azure and performs an ADsync
    Invoke-Command -ComputerName server -Credential $credentials -ScriptBlock {
        Start-ADSyncSyncCycle -PolicyType Delta
    }
    
    #Pauses script for 30 seconds
    Start-Sleep -Seconds 30

    #Connects to the Exchange Server and Enables Remote mailbox 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://server/PowerShell/ -Authentication Kerberos -Credential $credentials
    Import-PSSession $Session -DisableNameChecking

    Enable-RemoteMailbox -Identity $Name -RemoteRoutingAddress $Username@domain.mail.onmicrosoft.com
  }

#Removes Connection to Exchange Server
Remove-PSSession $Session



  
  
