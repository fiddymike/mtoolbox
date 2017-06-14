####################################################################
#NOTES:
#1 - Need to clean up the credential prompts for Office 365 and Active Directoy... and I think there is a duplicate entry for
#    Office 365 in a different function where I was trying to use a function to connect to and import the module for office 365.
#2 - Error Handling: Check all functions and commands for better error handling.
#
#
####################################################################

#Prompt to load profile
#----------------------------------------------------------################################################################## 
function Load-Profile ()
{
Write-Host ""
Write-Host "Welcome "$myname.toupper()" -----------------------------------" -ForegroundColor Blue -BackgroundColor Cyan
Write-Host ""
Write-Host "Load HelpDesk Profile?" -ForegroundColor Cyan
Write-Host ""
Switch -Regex ( Read-Host "[y or n]" )
    {
        y { Set-Profile }
        n { Write-Host ""
            Write-Host "Skipped loading HelpDesk profile." -ForegroundColor Yellow
            Write-Host "" }
        default { Set-Profile }
    }
 }

#Set Profile
#----------------------------------------------------------################################################################## 
function Set-Profile ()
{
Set-MyVariables
set-profilePrompt

}

#Set Profile Prompt
#----------------------------------------------------------################################################################## 
Function Set-ProfilePrompt () 
{
#Set User
#-----------------------------

Write-Host "
Administrator?" -ForegroundColor Cyan

switch ( Read-Host "
    1. Fiddymike
    2. Paul
    3. Jason
    4. Other

    [1, 2, 3 or 4]" ) 
    {
        1 { $Global:NTENTAdminAD = 'convera.com\madmin' 
            $Global:NTENTAdminO365 = 'mbennett@ntent.com' }
        2 { $Global:NTENTAdminAD = 'convera.com\pchavez-admin'
            $Global:NTENTAdminO365 = 'pchavez@ntent.com' }
        3 { $Global:NTENTAdminAD = 'convera.com\jjohnson-admin'
            $Global:NTENTAdminO365 = 'jjohnson@ntent.com' }
        4 { $Global:NTENTAdminAD = Read-Host "Enter Domain\Username"
            $Global:NTENTAdminO365 = Read-Host "Enter Office 365 Username" }
  Default { $Global:NTENTAdminAD = 'convera.com\madmin' 
            $Global:NTENTAdminO365 = 'mbennett@ntent.com' }

    }

#Set Server
#-----------------------------
Write-Host "
Server?" -ForegroundColor Cyan

switch (Read-Host "
    1. DC1
    2. DC2
    3. DC3
    4. DC4
    5. DC5
    6. Other

    1, 2, 3, 4, 5 or 6?") 
    {
        1 { $Global:NADServer = 'dc1.convera.com' }
        2 { $Global:NADServer = 'vsw-dc2.convera.com' }
        3 { $Global:NADServer = 'vsw-dc3.convera.com' }
        4 { $Global:NADServer = 'vsw-dc4.convera.com' }
        5 { $Global:NADServer = 'vsw-dc5.convera.com' }
        6 { $Global:NADServer = Read-Host "Enter Server Name" }
  default { $Global:NADServer = 'vsw-dc3.convera.com' }    
    }

#Load Active Directory
#-----------------------------

Write-Host "
Load Active Directory Module?" -ForegroundColor Cyan

switch (Read-Host "
    [y or n]") 
    {
        y { $Script:NLoadAD = $True }
        n { $Script:NLoadAD = $False
            Write-host ""
            Write-host "Skipping Active Directory" -ForegroundColor Red
          }
  default { $Script:NLoadAD = $True }
    }

#Load Office 365
#-----------------------------
Write-Host "
Load Office 365 Module?" -ForegroundColor Cyan

switch (Read-Host "
    [y or n]") 
    {
        y { $Script:NLoadO365 = $True }
        n { $Script:NLoadO365 = $False
            Write-host ""
            Write-host "Skipping Office 365" -ForegroundColor Red
          }
  default { $Script:NLoadO365 = $True }
    }

If ( $NLoadAD -eq $True ) { Connect-AD }
If ( $NLoadO365 -eq $true ) { Connect-O365 }
}

# Module: Set-ADCreds
#----------------------------------------------------------##################################################################
function Set-ADCreds ()
{ 
$global:ADCred = Get-Credential $NTENTAdminAD -Message "$myname, Please enter your AD Admin credentials."
}

# Module: Set-O365Creds
#----------------------------------------------------------##################################################################
function Set-O365Creds ()
{ 
$global:O365Cred = Get-Credential $NTENTAdminO365 -Message "$myname, Please enter your O365 Admin credentials."
}

#Connect Active Directory Module
#----------------------------------------------------------################################################################## 
function Connect-AD  ()
{   
    #Set Credential Variable
    Set-ADCreds
    
    #Import AD Module
    import-module ActiveDirectory -Global
 }

#Connect Active Directory Module [REMOTE]
#----------------------------------------------------------################################################################## 
function Connect-AD-Remote  ()
{
    #Set Credential Variable
    Set-ADCreds
    
    #Create Session for AD
    $Script:ADSession = New-PSSession -ComputerName $NADServer -Credential $ADCred
        
    #Invoke AD from Remote Host
    import-module (Import-PSSession -Session $ADSession -Module ActiveDirectory ) -Global

}


#Connect Office 365 and MSOLService
#----------------------------------------------------------################################################################## 
function Connect-O365 ()
{
    #Set Credential Variable
    Set-O365Creds
    
    #Create Session for Office 365
    $Script:O365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Authentication Basic -AllowRedirection -Credential $O365Cred

    #Import Office 365 Session
    Import-module ( Import-PSSession $O365Session -AllowClobber ) -global

    #Connect-MsolService
    Connect-MsolService -Credential $o365cred
}  

#Connect Module via Screen Prompt
#----------------------------------------------------------################################################################## 
function Connect-Mod ()
{
    #Set ModServer Variable
    $ModServer = Read-Host "Enter Server Name (i.e. vsw-dc3.convera.com)"
 
    #Set Credential Variable
    $Script:ModCred = Get-Credential $NTENTAdminAD -Message "$myname, Please enter your Admin credentials."
    
    #Create Session for AD
    $Script:ModSession = New-PSSession -ComputerName $ModServer -Credential $ModCred

    #Get Available Modules on Selected Server
    invoke-command -ComputerName $Modserver -Credential $ModCred { Get-Module -ListAvailable }

    #Set Module Name
    $ModName = Read-Host "Enter the Module Name (i.e. ActiveDirectory)"        

    #Invoke AD from Remote Host
    import-module (Import-PSSession -Session $ModSession -Module $ModName ) -Global

} 

#SET FULL USER
#----------------------------------------------------------################################################################## 
function Set-UserFull ()
{
    Write-Host "
    Test script with pre-populated test-user data? (mtest3)" -Foregroundcolor Cyan
    switch -regex (Read-Host "[y/n]") {
        "y" { Write-Host "
            Using Test-User data...
            " -ForegroundColor Green
            $Script:NFirstName = "Mike"
            $Script:NLastName = "Test3"
            $Script:NFullName = "Mike Test3"
            $script:NUsername = "mtest3"
            $script:NEmailAddress = "$NUsername@ntent.com"
            $Script:NTitle = "IT Bot"
            $Script:NDept = "IT"
            $Script:NDirectNumber = "760-555-1234"
            $Script:NMobileNumber ="760-555-9876"
            }
        "n" { Write-Host "
            Enter Information for the following:
            " -ForegroundColor Yellow
            $Script:NFirstName = Read-Host "First Name"
            $Script:NLastName = Read-Host "Last Name"
            $Script:NFullName = "$NFirstName $NLastName"
            $Script:NUsername = Read-Host "Username"
            $Script:NEmailAddress = "$Nusername@ntent.com"
            $Script:NTitle = Read-Host "Title"
            $Script:NDept = Read-Host "Department"
            $Script:NDirectNumber = Read-Host "Direct Phone Number"
            $Script:NMobileNumber = Read-Host "Mobile Phone Number"
            }
    }
    
}

# Module: Set-UserName
#----------------------------------------------------------##################################################################
function Set-UserName ()
{
$script:NUsername = Read-Host "Username"
}

# SET USER INFORMATION
#----------------------------------------------------------##################################################################
function Set-UserInfo ()
{
$script:NUsername = Read-Host "Username"

$script:NEmailAddress = "$NUsername@ntent.com"
if (-not(Get-MailBox -Identity $NEmailAddress)) { "Could not find a mailbox with email address $EmailAddress"; exit 1}

$script:dn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUsername).DistinguishedName
}

#SET USER EMAIL ADDRESS
#----------------------------------------------------------##################################################################
function Set-UserEmailAddress ()
{
    $Script:NEmailAddress = "$Nusername@ntent.com"
    if (-not(Get-MailBox -Identity $NEmailAddress)) { "Could not find a mailbox with email address $NEmailAddress"}
    else { Write-Host "Found" -ForegroundColor Green } 
}

#SET USER'S MANAGER'S EMAIL ADDRESS
#----------------------------------------------------------##################################################################
function Set-MgrEmail ()
{
$Script:NManagerName = read-host "Please enter Manager's alias"
$SCript:NManagerEmailAddress = "$NManagerName@ntent.com"
if (-not(Get-MailBox -Identity $NManagerEmailAddress)) { "Could not find a mailbox with email address $NManagerEmailAddress"}
}

# SET USER LOCATION
#----------------------------------------------------------##################################################################
function Set-UserLoc ()
{
Write-Host "
Office Location" -ForegroundColor Cyan

    Write-Host "
    1. Carlsbad
    2. New York
    3. Vienna
    4. Newport Beach
    5. UK" 
    $Script:NUserDeets = read-host "
    Enter 1, 2, 3, 4 or 5"

switch ($NUserDeets) {
    "1" { 
        #Carlsbad
        $Script:prefix = "CB"
        }
    "2" { 
        #New York
        $Script:prefix = "NY"
        }
    "3" { 
        #Vienna
        $Script:prefix = "Vi"
        }
    "4" { 
        #New Port
        $Script:prefix = "NB"
        }
    "5" { 
        #UK 
        $Script:prefix = "UK"
        }
}

$Script:NStreetAddress = (get-variable ($prefix + 'StreetAddress')).Value
$Script:Ncity = (get-variable ($prefix + 'City')).Value
$Script:NState = (get-variable ($prefix + 'State')).Value 
$Script:NPostalCode = (get-variable ($prefix + 'PostalCode')).Value
$Script:NCountry = (get-variable ($prefix + 'Country')).value
$Script:NDistName = (get-variable ($prefix + 'DistName')).value
$Script:NOUPath = (get-variable ($prefix + 'OUPath')).value
$Script:NOUXPath = (get-variable ($prefix + 'OUXPath')).value
$Script:NADServer = (get-variable ($prefix + 'ADServer')).value
}

#GET MY VARIABLES
#----------------------------------------------------------################################################################## 
function Set-MyVariables () 
{
# LOCATION VARIABLES
#---------------------------------------------------------- 
$Global:NCompany = "NTENT"
$Global:NADDomain = "@verticalsearchworks.com"
$Global:logfile = $MyP + "\logs\logfile.txt"
$Global:i        = 0
$Global:date     = Get-Date
$Global:MyPC = hostname

$Global:NVoiceMailSetup = 'https://confluence.ntent.com/display/DOC/Setup+Voicemail'
$Global:NPidginDomain = 'chat.verticalsearchworks.com'
$Global:NPidginPW = 'NTENTftw'
$Global:NTempPWNewUser = '2doSomething'


# Carlsbad
#---------------------------------------------------------- 
$Global:CBStreetAddress = "1808 Aston Avenue, Suite 170"
$Global:CBCity = "Carlsbad"
$Global:CBState = "CA"
$Global:CBPostalCode = "92008"
$Global:CBcountry = "US"
$Global:CBDistName = "carlsbademployees"
$Global:CBOUPath = "OU=carlsbad,OU=Domain Users,DC=convera,DC=com"
$Global:CBOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:CBADServer = "vsw-dc3.convera.com"

# New York
#---------------------------------------------------------- 
$Global:NYStreetAddress = "342 W 37th Street, Suite 100 Ground Floor"
$Global:NYCity = "New York"
$Global:NYState = "NY"
$Global:NYPostalCode = "10018"
$Global:NYcountry = "US"
$Global:NYDistName = "nyemployees"
$Global:NYOUPath = "OU=NY,OU=Domain Users,DC=convera,DC=com"
$Global:NYOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:NYADServer = "vsw-dc4.convera.com"

# Vienna
#---------------------------------------------------------- 
$Global:ViStreetAddress = "1919 Gallows Road, Suite 1050"
$Global:ViCity = "Vienna"
$Global:ViState = "VA"
$Global:ViPostalCode = "22182"
$Global:Vicountry = "US"
$Global:ViDistName = "viennaemployees"
$Global:ViOUPath = "OU=Vienna,OU=Domain Users,DC=convera,DC=com"
$Global:ViOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:ViADServer = "vsw-dc2.convera.com"

# Newport Beach
#----------------------------------------------------------         
$Global:NBStreetAddress = "4590 MacArthur Boulevard"
$Global:NBCity = "Newport Beach"
$Global:NBState = "CA"
$Global:NBPostalCode = "92660"
$Global:NBcountry = "US"
$Global:NBDistName = "newportbeachemployees"
$Global:NBOUPath = "OU=Newport,OU=Domain Users,DC=convera,DC=com"
$Global:NBOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:NBADServer = "vsw-dc3.convera.com"

# UK 
#---------------------------------------------------------- 
$Global:UKStreetAddress = "14 Great College Street"
$Global:UKCity = "Westminster"
$Global:UKState = "London"
$Global:UKPostalCode = "SW1P 3RX"
$Global:UKcountry = "UK"
$Global:UKDistName = "ukemployees@ntent.com"
$Global:UKOUPath = "OU=UK,OU=Domain Users,DC=convera,DC=com"
$Global:UKOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:UKADServer = "vsw-dc4.convera.com"

# MY CSS 
#----------------------------------------------------------
$Global:mycss = "<style>
table
    { 
    Margin: 0px 0px 0px 4px;
    Border: 1px solid #bebebe;
    Font-Family: Helvetica;
    Font-Size: 10pt;
    Background-Color: #fff;
    }
tr:hover td
    {
    Background-Color: #cd3427;
    Color: #fff;
    }
tr:nth-child(even)
    {
    Background-Color: #f2f2f2;
    }
th
    {
    Text-Align: Left;
    Color: #fff;
    Padding: 1px 4px 1px 4px;
	Background-Color: #cd3427;
    }
td
    {
    Vertical-Align: Top;
    Padding: 1px 4px 1px 4px;
    }
</style>"
}

#CREATE NEW USER
#----------------------------------------------------------################################################################## 
function New-User ()
{
# Call Functions
#----------------------------------------------------------
set-myvariables
set-userloc
set-userFull

#FUNCTIONS AND VARIABLES 
#----------------------------------------------------------

$SetPassword = ConvertTo-SecureString -AsPlainText $NTempPWNewUser -force

$NewUserSettings = [Ordered]@{
    UserPrincipalName = $NEmailAddress
    FirstName = $NFirstName
    LastName = $NLastName
    Displayname = $NFullName
    Title = $NTitle
    Department = $NDept
    PhoneNumber =  $NDirectNumber
    MobilePhone = $NMobileNumber
    StreetAddress = $NStreetAddress
    City = $NCity
    State = $NState
    PostalCode = $NPostalCode
    UsageLocation = $NCountry
    Office = $NCity
    Country = $NCountry
    Password = $NTempPWNewUser
    PasswordNeverExpires = $True
    ForceChangePassword = $False
    LicenseAssignment = "convera:ENTERPRISEPACK"
    }

Write-Host "
    This will effectively create the following:" -ForegroundColor DarkYellow
Write-Host "
OFFICE 365 USER
" -ForegroundColor Magenta
$NewUserSettings

$NADNewUser = [Ordered]@{
        SamAccountName = $NUsername 
        Name = $NFirstName
        server = $NADServer
        Credential = $ADCred
        GivenName = $NFirstName
        ChangePasswordAtLogon = $FALSE
        Surname = $NLastName
        DisplayName = $NFullName
        Office = $NCity
        Description = $NTitle
        EmailAddress = $NEmailAddress
        StreetAddress = $NStreetAddress
        City = $NCity
        state = $NState
        PostalCode = $NPostalCode
        Country = $NCountry
        UserPrincipalName = $NEmailAddress
        Company = $NCompany
        Department = $NDept
        enabled = $TRUE
        Title = $NTitle
        OfficePhone = $NDirectNumber
        MobilePhone = $NMobileNumber
        AccountPassword = $SetPassword
        }

Write-Host "
ACTIVE DIRECTORY USER" -ForegroundColor Magenta
$NAdNewUser

# VERIFY & RUN
#----------------------------------------------------------
Write-Host "
    Create Office 365 User?" -ForegroundColor Yellow
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Creating Office 365 user
        " -ForegroundColor Green
        New-MsolUser @NewUserSettings | out-null
        }
    "n" { Write-Host "
        Opted not to Create Office 365 User" -ForegroundColor Red; }
}

Write-Host "
    Create Active Directory User?" -ForegroundColor Yellow
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Creating Active Directory User
        " -ForegroundColor Green
        New-ADUser @NADNewUser | out-null
        }
    "n" { Write-Host "
        Opted not to Create Office 365 User" -ForegroundColor Red; exit 0; }
}


#Define DN to use in the  Move-ADObject command

$dn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUserName).DistinguishedName
 
# Move the users to the OU set above. 

Move-ADObject -server $NADServer -Credential $ADCred -Identity $dn -TargetPath $NOUPath 
 
#Rename the object to a good looking name to avoid displaying sAMAccountNames (eg tests1.user1)

$newdn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUsername).DistinguishedName
Rename-ADObject -server $NADServer -Credential $ADCred -Identity $newdn -NewName $NFullname

Write-Host = "
Complete
" -ForegroundColor Green
}

#REMOVE USER
#----------------------------------------------------------################################################################## 
function Remove-User ()
{
# Call Scripts
#----------------------------------------------------------
C:\users\MBennett\Documents\WindowsPowerShell\modules\NTENTVariables.ps1

# Call Functions
#----------------------------------------------------------
get-userloc
get-userinfo

$NManagerName = read-host "Please enter Manager's alias"
if (!$NManagerName) {$NManagerName = "mbennett"}
$NManagerEmailAddress = "$NManagerName@ntent.com"
if (-not(Get-MailBox -Identity $NManagerEmailAddress)) { "Could not find a mailbox with email address $NManagerEmailAddress"; exit 1}

$NPassword = Read-host "Password"
if (!$Npassword) {$NPassword = "G000dby3"}

$DGs = Get-DistributionGroup

# Variables for Active Directory
#----------------------------------------------------------
$Date = Read-Host 'Enter Term Date (i.e. mm/dd)'
if (!$Date) {$Date = Get-Date -Format "MM/dd"}

$setpassword = ConvertTo-SecureString -AsPlainText $NPassword -force

$NDescription = "-term $Date - email fwd to $NManagerEmailAddress"

# Pull User Account List
#----------------------------------------------------------
get-Mailbox -identity $NUserName | Select @{ l="Name"; e={ $_.DisplayName } }, @{ l="Active Directory"; e={ $_.CustomAttribute1 } }, @{ l="Office365"; e={ $_.CustomAttribute2 } }, @{ l="Sales Force"; e={ $_.CustomAttribute3 } }, @{ l="Google Apps"; e={ $_.CustomAttribute4 } }, @{ l="Podio"; e={ $_.CustomAttribute5 } }, @{ l="Atlassian"; e={ $_.CustomAttribute6 } }, @{ l="Secure NTENT"; e={ $_.CustomAttribute7 } }, @{ l="Ad Support"; e={ $_.CustomAttribute8 } }, @{ l="Tableau"; e={ $_.CustomAttribute9 } }, @{ l="Great Plains"; e={ $_.CustomAttribute10 } } |out-gridview

#----------------------------------------------------------
Write-Host "
    This will effectively disable $NUserName, forward email to $NManagerEmailAddress and change password to $NPassword ?
    " -ForegroundColor Yellow
Write-Host "Disable user in Office 365?"
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Office 365 user will be disabled
        " -ForegroundColor Cyan
        $NDisableO365 = $True
        }
    "n" { Write-Host "
        Opted not to disable Office 365 User" -ForegroundColor Red;
        $NDisableO365 = $False
        }
        }


Write-Host "Disable user in Active Directory?"
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Active Directory user will be disabled
        " -ForegroundColor Cyan
        $NDisableAD = $True
        }
    "n" { Write-Host "
        Opted not to disable Active Directory User" -ForegroundColor Red;
        $NDisableAD = $False
        }
        }


If ($NDisableO365 -eq $True) 
        { 
        #Hide User From Global Address List
        Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $true
        Write-host "-Completed - O365: Hide from Global Address List" -ForegroundColor Green

        #Disable Active Sync, OWA, POP, IMAP and MAPI
        Set-CASMailbox -Identity $NEmailAddress -ActiveSyncEnabled $False -OWAEnabled $False -OWAforDevicesEnabled $false -PopEnabled $False -ImapEnabled $False -MAPIEnabled $False
        Write-host "-Completed - O365: Disable Active Sync" -ForegroundColor Green

        #Set sign-in status to Blocked
        Set-MsolUser -UserPrincipalName $NEmailAddress -BlockCredential $True
        Write-host "-Completed - O365: Sign-in Status changed to 'Blocked'" -ForegroundColor Green

        #Forward email to Manager
        Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $true -ForwardingAddress $NManagerEmailAddress
        Write-host "-Completed - O365: Forward Email" -ForegroundColor Green

        #Set Password
        Set-MsolUser  -UserPrincipalName $NEmailAddress -StrongPasswordRequired $False
        Set-MsolUserPassword -UserPrincipalName $NEmailAddress -NewPassword $NPassword -ForceChangePassword $false | Out-Null
        Write-host "-Completed - O365: Changed Password" -ForegroundColor Green
        }

        <#
        #Change License to E01
        #Set-MsolUserLicense -AddLicenses convera:EXCHANGESTANDARD | Get-MsolUser $NEmailAddress
        Write-host "Completed: License Change"

        #Remove User from Distribution Groups
        foreach( $dg in $DGs)
        {
            if((Get-DistributionGroupMember -Identity $dg.identity).name -contains $NEmailAddress)
            {
                Remove-DistributionGroupMember -Identity $dg.identity -Member $NEmailAddress -confirm
            }
            else
            {
                "Could not find the user $Fullname in group $($dg.identity)"
            }
        }
        Write-host "Completed: Remove User from Distribution Groups"
        #>



If ($NDisableAD -eq "$True") 
        { 
        #Get Admin account credential
        $dn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUserName).DistinguishedName

        #Change Description
            Write-Host "-Changing AD description for $NUserName to $NDescription" -ForegroundColor Green
        Set-ADUser -server $NADServer -Credential $ADCred -Identity $dn -Description $NDescription -ea SilentlyContinue -ev err
        
        #Change ADUser Password
            Write-Host "-Changing AD password to $NPassword" -ForegroundColor Green
        Set-ADAccountPassword -server $NADServer -Credential $ADCred -Identity $dn -NewPassword $SetPassword -ea SilentlyContinue -ev +err

        #Disabe ADUser Account
            Write-Host "-Disabling $NUserName's User Account" -ForegroundColor Green
        Disable-ADAccount -Identity $dn -Credential $ADCred
                
        #Moving User to To Be Deleted OU
            Write-Host "-Moving $NUserName to the 'To Be Delete' OU ( $NOUXPath )" -ForegroundColor Green
        Move-ADObject -server $NADServer -Credential $ADCred -Identity $dn -TargetPath $NOUXPath -ea SilentlyContinue -ev +err
        }

# UserStatus
#----------------------------------------------------------
#. ( join-path $MyP\modules\o365\ userstatus.ps1


# Pull User Account List
#----------------------------------------------------------
get-Mailbox -identity $NUserName | Select @{ l="Name"; e={ $_.DisplayName } }, @{ l="Active Directory"; e={ $_.CustomAttribute1 } }, @{ l="Office365"; e={ $_.CustomAttribute2 } }, @{ l="Sales Force"; e={ $_.CustomAttribute3 } }, @{ l="Google Apps"; e={ $_.CustomAttribute4 } }, @{ l="Podio"; e={ $_.CustomAttribute5 } }, @{ l="Atlassian"; e={ $_.CustomAttribute6 } }, @{ l="Secure NTENT"; e={ $_.CustomAttribute7 } }, @{ l="Ad Support"; e={ $_.CustomAttribute8 } }, @{ l="Tableau"; e={ $_.CustomAttribute9 } }, @{ l="Great Plains"; e={ $_.CustomAttribute10 } } |out-gridview


Write-Host "$NUsername is a member of the following O365 groups:" -ForegroundColor Cyan
foreach ($group in get-distributiongroup -resultsize unlimited){

if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $dn){Write-Host $group.name}

}


$err

Write-host "
S C R I P T    C O M P L E T E * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"  -ForegroundColor DarkBlue -BackgroundColor Cyan 
}

#Add USER TO DISTRIBUTION GROUP(S)
#----------------------------------------------------------##################################################################
function Add-UserDistributionGroup ()
{
set-username
Set-UserEmailAddress

#Remove User from Distribution Groups
do 
    {
    $DistributionGroup = Read-Host "Please enter the distribution Name"
    if ($DistributionGroup -ne "")
        {Add-DistributionGroupMember -Identity $DistributionGroup -member $NUsername }
    } 
    while ($DistributionGroup -ne "")
}

#REMOVE USER FROM DISTRIBUTION GROUP(S)
#----------------------------------------------------------##################################################################
function Remove-UserDistributionGroup ()
{
set-username
Set-UserEmailAddress

#Remove User from Distribution Groups
do
    {
    $DistributionGroup = read-host "Please enter Distribution Group"
    if ($DistributionGroup -ne "")
        {Remove-DistributionGroupMember -Identity $DistributionGroup -Member $NEmailAddress}
    } 
    while ($DistributionGroup -ne "")
}

#USER STATUS
#----------------------------------------------------------################################################################## 
function Get-UserStatus ()
{
<#
  .SYNOPSIS
  This function retrieves user access information related to Office 365 and Active Directory
  .DESCRIPTION
  The primary use-case for this function is to check the status of departed users to ensure that all access has been disabled or terminated.
  .EXAMPLE
  get-userstatus -NUserName mbennett

  Name              : Mike Bennett
HiddenFromGAL     : False
ForwardMail       : 
ForwardSMTPMail   : 
DeliverTo         : False
ActiveSync        : True
OWA               : True
OWAdevices        : True
POP               : True
IMAP              : True
MAPI              : True
CredentialBlocked : False
IsLicensed        : True

Description
-----------

Get access properties of the user with NUserName 'mbennett'.
  .PARAMETER computername
  The computer name to query. Just one.
  .PARAMETER logname
  The name of a file to write failed computer names to. Defaults to errors.txt.
 #>

    param (
    [parameter(Mandatory=$True,
                Valuefrompipeline=$True,
                ValueFromPipelineByPropertyName=$true)]

    [String]$NUsername)

    PROCESS {
        
        $NEmailAddress = "$NUserName@ntent.com"
        
        $mb = get-mailbox $NUsername
            $mb_DisplayName = $mb.name
            $mb_Hidden = $mb.hiddenFromAddressListsEnabled
            $mb_Forward = $mb.ForwardingAddress
            $mb_ForwardSMTP = $mb.ForwardingSmtpAddress
            $mb_Delivery = $mb.DeliverToMailboxAndForward

        $casmb = get-casmailbox $NEmailAddress
            $casmb_activesync = $casmb.ActiveSyncEnabled
            $casmb_owa = $casmb.OWAEnabled
            $casmb_owa_devices = $casmb.OWAforDevicesEnabled
            $casmb_pop = $casmb.PopEnabled
            $casmb_imap = $casmb.ImapEnabled
            $casmb_mapi = $casmb.MAPIEnabled
        $msolu = get-msoluser -UserPrincipalName $NEmailAddress
            $msolu_credential = $msolu.BlockCredential
            $msolu_licensed = $msolu.IsLicensed

        $ad = get-aduser $NUsername -pr * -server $NADServer -Credential $ADCred
  <#      
        $Stuff =[Ordered]@{
        'Name' = $mb_DisplayName
        
        'Active Sync Enabled' = $casmb_activesync
        'OWA Enabled' = $casmb_owa
        'OWAdevices Enabled' = $casmb_owa_devices
        'POP Enabled' = $casmb_pop
        'IMAP Enabled' = $casmb_imap
        'MAPI Enabled' = $casmb_mapi
        'Credential Blocked' = $msolu_credential
        'Is Licensed' = $msolu_licensed
        'Hidden From GAL' = $mb_Hidden
        'Deliver To Enabled' = $mb_Delivery  
        'Forward Mail' = $mb_Forward
        'Forward SMTP Mail' = $mb_ForwardSMTP
        }

        $Ntest = New-Object psobject -Property $stuff

        Write-Output = $Ntest
#>
        Write-host ""
        Write-Host "Office 365" -ForegroundColor Blue -BackgroundColor Cyan
        Write-host ""
        If ($mb_Hidden -eq $true) {"[  TRUE   ] Hidden from GAL "}
        Else {Write-Host "[  FALSE  ] Hidden from GAL" -ForegroundColor Red}
        If ($casmb_activesync -eq $true) {Write-Host "[  FALSE  ] Active Sync Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] Active Sync Disabled"}
        If ($casmb_owa -eq $true) {Write-Host "[  FALSE  ] OWA Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] OWA Disabled"}
        If ($casmb_owa_devices -eq $true) {Write-Host "[  FALSE  ] OWA for Devices Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] OWA for Devices Disabled"}
        If ($casmb_pop -eq $true) {Write-Host "[  FALSE  ] POP Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] POP Disabled"}
        If ($casmb_imap -eq $true) {Write-Host "[  FALSE  ] IMAP Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] IMAP Disabled"}
        If ($casmb_mapi -eq $true) {Write-Host "[  FALSE  ] MAPI Disabled" -ForegroundColor Red}
        Else {"[  TRUE   ] MAPI Disabled"}
        If ($mb_Forward -eq $null) {Write-Host "[  FALSE  ] Mail Forwarded [Global Admin]" -ForegroundColor Red}
        Else {"[  TRUE   ] Mail Forwarded [Global Admin]"}
        If ($mb_ForwardSMTP -eq $null) {Write-Host "[  FALSE  ] Mail Forwarded [User's OWA]" -ForegroundColor Red}
        Else {"[  TRUE   ] Mail Forwarded [User's OWA]"}
        If ($mb_Delivery -eq $true) {"[  TRUE   ] Mail to Mailbox & FWD ADD"}
        Else {Write-Host "[  FALSE  ] Mail to Mailbox & FWD ADD" -ForegroundColor Red}
        If ($msolu_credential -eq $true) {"[  TRUE   ] Credential Blocked"}
        Else {Write-Host "[  FALSE  ] Credential Blocked" -ForegroundColor Red}
        If ($msolu_licensed -eq $false) {"[  TRUE   ] License Revoked"}
        Else {Write-Host "[  FALSE  ] License Revoked" -ForegroundColor Red}

        Write-host ""
        Write-host "Mailbox forwarded [Global Admin] to:"$mb.ForwardingAddress
        Write-host "Mailbox forwarded [User's OWA] to:"$mb.ForwardingSmtpAddress
        Write-host ""

        Write-Host "Active Directory" -ForegroundColor Blue -BackgroundColor Cyan
        Write-host ""

        If ($ad.enabled -eq $false) {"[  TRUE   ] Account Disabled"}
        Else {Write-Host "[  FALSE  ] Account Disabled" -ForegroundColor Red}

        Write-host ""
        Write-host "AD Description changed to:"$ad.Description

<#
        $NUserStatus = New-Object psobject
        

        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Active Sync Enabled' -Value $casmb_activesync
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'OWA Enabled' -Value $casmb_owa
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'OWAdevices Enabled' -Value $casmb_owa_devices
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'POP Enabled' -Value $casmb_pop
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'IMAP Enabled' -Value $casmb_imap
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'MAPI Enabled' -Value $casmb_mapi
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Credential Blocked' -Value $msolu_credential
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Is Licensed' -Value $msolu_licensed
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Hidden From GAL' -Value $mb_Hidden
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Deliver To Enabled' -Value $mb_Delivery
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Forward Mail' -Value $mb_Forward
        $NUserStatus | Add-Member -MemberType NoteProperty -Name 'Forward SMTP Mail' -Value $mb_ForwardSMTP
        
        write-output $NUserStatus
#>
      
    }
}

#GET USER DIRECTORY
#----------------------------------------------------------################################################################## 
function Get-CoDir ()
{
Write-Host "
1. CSV
2. HTML 
3. GridView" -ForegroundColor Cyan
switch -Regex (Read-Host "
View?") {
    "1" { $NOut = 'CSV' }
    "2" { $NOut = 'HTML' }
    "3" { $NOut = 'Gridview' }
        }

Write-Host "
Creating $NOut" -foregroundColor Green

$NGetDirectoryUSers = get-mailbox | ? { !$_.HiddenFromAddressListsEnabled } | % { 
    Get-MsolUser -UserPrincipalName $_.UserPrincipalName | ? { ($_.UserType -eq 'Member') -and $_.IsLicensed -and ($_.Country) } 
        } | Select DisplayName, PhoneNumber, MobilePhone, Title, UserPrincipalName, City, State | Sort-Object -Property DisplayName 
<#
$GetDirectoryUsers = Get-MsolUser | ? { ($_.UserType -eq 'Member') -and $_.IsLicensed -and ($_.Country -eq 'US') } | % {
    get-mailbox | ? { $_.HiddenFromAddressListsEnabled }
    } | Select Name, objectid | Sort Name
#>

If ( $NOut -eq 'CSV' ) { $NGetDirectoryUSers | Export-Csv "$mreports\NTENT Dir $(Get-Date -f "MM-dd hhmm-ss").csv" -NoTypeInformation }
If ( $NOut -eq 'HTML' ) { $NGetDirectoryUSers | ConvertTo-Html -body $mycss | Set-content "$mreports\NTENT Dir $(Get-Date -f "MM-dd hhmm-ss").html" }
If ( $NOut -eq 'GridView' ) { $NGetDirectoryUSers | Out-GridView }

}

#CHANGE OFFICE 365 PASSWORD
#----------------------------------------------------------################################################################## 
function Change-PWO365 ()
{ 
Set-UserName
Set-UserEmailAddress

$c = Get-Credential -Message "Enter password for this user" -username $NEmailAddress
$Password = $C.getnetworkcredential().password

Set-MsolUser  -UserPrincipalName $NEmailAddress -StrongPasswordRequired $False
Set-MsolUserPassword -UserPrincipalName $NEmailAddress -NewPassword $Password -ForceChangePassword $false | Out-Null
}

#CHANGE ACTIVE DIRECTORY PASSWORD
#----------------------------------------------------------################################################################## 
function Change-PWAD ()
{ 
Set-UserName

$Password = (Read-Host -Prompt "Provide New Password" -AsSecureString)

Set-ADAccountPassword -identity $NUsername -NewPassword $Password
}

#HIDE USER FROM GAL
#----------------------------------------------------------################################################################## 
function Hide-UserGAL ()
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $true
}

#SHOW USER IN GAL
#----------------------------------------------------------################################################################## 
function Show-UserGAL ()
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $False
}

#FORWARD USER'S EMAIL
#----------------------------------------------------------################################################################## 
function Set-EmailFwd () 
{
Set-UserName
Set-UserEmailAddress
Set-MgrEmail

Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $true -ForwardingAddress $NManagerEmailAddress
}

#REMOVE FORWARD FROM USER
#----------------------------------------------------------################################################################## 
function Remove-EmailFwd () 
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $False -ForwardingAddress $Null
}

# Resize the standard console window
#----------------------------------------------------------################################################################## 
Function Resize-ConsoleWindow
{
<#  .Synopsis Resize PowerShell console window .Description Resize PowerShell console window. Make it bigger, smaller or increase / reduce the width and height by a specified number .Parameter -Bigger Increase the window's both width and height by 10. .Parameter -Smaller Reduce the window's both width and height by 10. .Parameter Width Resize the window's width by passing in an integer. .Parameter Height Resize the window's height by passing in an integer. .Example # Make the window bigger. Resize-Console -bigger  .Example # Make the window smaller. Resize-Console -smaller  .Example # Increase the width by 15. Resize-Console -Width 15  .Example # Reduce the Height by 10. Resize-Console -Height -10  .Example # Reduce the Width by 5 and Increase Height by 10. Resize-Console -Width -5 -Height 10 #>
 
[CmdletBinding()]
PARAM (
[Parameter(Mandatory=$false,HelpMessage="Increase Width and Height by 10")][Switch] $B,
[Parameter(Mandatory=$false,HelpMessage="Reduce Width and Height by 10")][Switch] $S,
[Parameter(Mandatory=$false,HelpMessage="Increase / Reduce Width" )][Int32] $Width,
[Parameter(Mandatory=$false,HelpMessage="Increase / Reduce Height" )][Int32] $Height
)
 
#Get Current Buffer Size and Window Size
$bufferSize = $Host.UI.RawUI.BufferSize
$WindowSize = $host.UI.RawUI.WindowSize
If ($B -and $S)
{
Write-Error "Please make up your mind, you can't go bigger and smaller at the same time!"
} else {
if ($B)
{
$NewWindowWidth = $WindowSize.Width + 40
$NewWindowHeight = $WindowSize.Height + 20
 
#Buffer size cannot be smaller than Window size
If ($bufferSize.Width -lt $NewWindowWidth)
{
$bufferSize.Width = $NewWindowWidth
}
if ($bufferSize.Height -lt $NewWindowHeight)
{
$bufferSize.Height = $NewWindowHeight
}
$WindowSize.Width = $NewWindowWidth
$WindowSize.Height = $NewWindowHeight
 
} elseif ($S)
{
$NewWindowWidth = $WindowSize.Width - 10
$NewWindowHeight = $WindowSize.Height - 10
$WindowSize.Width = $NewWindowWidth
$WindowSize.Height = $NewWindowHeight
}
 
if ($Width)
{
#Resize Width
$NewWindowWidth = $WindowSize.Width + $Width
If ($bufferSize.Width -lt $NewWindowWidth)
{
$bufferSize.Width = $NewWindowWidth
}
$WindowSize.Width = $NewWindowWidth
}
if ($Height)
{
#Resize Height
$NewWindowHeight = $WindowSize.Height + $Height
If ($bufferSize.Height -lt $NewWindowHeight)
{
$bufferSize.Height = $NewWindowHeight
}
$WindowSize.Height = $NewWindowHeight
 
}
#commit resize
$host.UI.RawUI.BufferSize = $buffersize
$host.UI.RawUI.WindowSize = $WindowSize
}
 
}
 
New-Alias -Name rcw -Value Resize-ConsoleWindow





#LIST O365 GROUPS A USER IS A MEMBER OF
#----------------------------------------------------------################################################################## 
function Get-UserMembership ()
{
    $NUserDN = (get-mailbox $Nusername).distinguishedname

    "User " + $NUsername + " is a member of the following groups:"

    foreach ($group in get-distributiongroup -resultsize unlimited)
    {
    if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $NUserDN){$group.name}
    }    
}

#REMOVE USER FROM ALL O365 GROUPS
#----------------------------------------------------------################################################################## 
function Remove-UserMembership ()
{
    Set-UserName
    Set-UserEmailAddress

    Write-Host "$NEmailAddress"
    
    $DGs= Get-DistributionGroup | where { (Get-DistributionGroupMember $_ | foreach {$_.PrimarySmtpAddress}) -contains $NEmailAddress}

    foreach( $dg in $DGs)
    {
    Remove-DistributionGroupMember $dg -Member user@domain.com
    }
}

#REMOVE USER FROM ALL O365 GROUPS
#----------------------------------------------------------################################################################## 
function Get-MyCommand () 
{
get-command -module mtoolbox
}

Export-ModuleMember -function Load-Profile
Export-ModuleMember -function Set-Profile
Export-ModuleMember -function Connect-AD
Export-ModuleMember -function Connect-AD-Remote
Export-ModuleMember -function Connect-O365
Export-ModuleMember -function Connect-Mod
Export-ModuleMember -function Set-MyVariables
Export-ModuleMember -function New-User
Export-ModuleMember -function Remove-User
Export-ModuleMember -function Add-UserDistributionGroup
Export-ModuleMember -function Remove-UserDistributionGroup
Export-ModuleMember -function Get-UserStatus
Export-ModuleMember -function Get-CoDir
Export-ModuleMember -function Change-PWO365
Export-ModuleMember -function Change-PWAD
Export-ModuleMember -function Hide-UserGAL
Export-ModuleMember -function Show-UserGAL
Export-ModuleMember -function Add-EmailFwd
Export-ModuleMember -function Remove-EmailFwd
Export-ModuleMember -function Resize-ConsoleWindow
Export-ModuleMember -function Get-UserMembership
Export-ModuleMember -function Remove-UserMembership
Export-ModuleMember -function Get-MyCommand

Export-ModuleMember -alias rcw