####################################################################
#NOTES:
#1 - Need to clean up the credential prompts for Office 365 and Active Directoy... and I think there is a duplicate entry for
#    Office 365 in a different function where I was trying to use a function to connect to and import the module for office 365.
#2 - Error Handling: Check all functions and commands for better error handling.
#
#
####################################################################

#Prompt to load profile
#============================================================================================================================================================================================ 
function Load-Profile ()
{
Write-Host ""
Write-Host "Welcome "$myname.toupper()" -----------------------------------" -ForegroundColor DarkMagenta  -BackgroundColor White
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
#============================================================================================================================================================================================ 
function Set-Profile ()
{
Set-MyVariables
set-profilePrompt

}

#Set Profile Prompt
#============================================================================================================================================================================================ 
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
    1. DC1 ( dc1.convera.com ):VEGAS
    2. DC2 ( vn-dc2.convera.com ):VIENNA
    5. DC5 ( ny-dc5.convera.com ):NEW YORK
    7. DC7 ( dc7.convera.com ):CARLSBAD
    8. Other

    1, 2, 3... ?") 
    {
        1 { $Global:NADServer = 'dc1.convera.com' }
        2 { $Global:NADServer = 'vn-dc2.convera.com' }
        5 { $Global:NADServer = 'ny-dc5.convera.com' }
        7 { $Global:NADServer = 'dc7.convera.com' }
        8 { $Global:NADServer = Read-Host "Enter Server Name" }
  default { $Global:NADServer = 'dc7.convera.com' }    
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
#============================================================================================================================================================================================
function Set-ADCreds ()
{ 
$global:ADCred = Get-Credential $NTENTAdminAD -Message "$myname, Please enter your AD Admin credentials."
}

# Module: Set-O365Creds
#============================================================================================================================================================================================
function Set-O365Creds ()
{ 
$global:O365Cred = Get-Credential $NTENTAdminO365 -Message "$myname, Please enter your O365 Admin credentials."
}

#Connect Active Directory Module
#============================================================================================================================================================================================ 
function Connect-AD ()
{
    #Set Credential Variable
    Set-ADCreds
    
    #Import Active Directory Module
    import-module ActiveDirectory -Global  

} 

#Connect Active Directory Module [REMOTE]
#============================================================================================================================================================================================ 
function Connect-AD-Remote ()
{
    #Set Credential Variable
    Set-ADCreds
    
    #Create Session for AD
    $Script:ADSession = New-PSSession -ComputerName $NADServer -Credential $ADCred
        
    #Invoke AD from Remote Host
    #import-module (Import-PSSession -Session $ADSession -Module ActiveDirectory ) -Global

} 

#Connect Office 365 and MSOLService
#============================================================================================================================================================================================ 
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
#============================================================================================================================================================================================ 
function Connect-Mod ()
{
    #Set ModServer Variable
    $ModServer = Read-Host "Enter Server Name (i.e. dc7.convera.com)"
 
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
#============================================================================================================================================================================================ 
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
#============================================================================================================================================================================================
function Set-UserName ()
{
$script:NUsername = Read-Host "Username"
}

# SET USER INFORMATION
#============================================================================================================================================================================================
function Set-UserInfo ()
{
$script:NUsername = Read-Host "Username"

$script:NEmailAddress = "$NUsername@ntent.com"
    if (-not(Get-MailBox -Identity $NEmailAddress)) { "Could not find a mailbox with email address $EmailAddress"}
    else { Write-Host "Found" -ForegroundColor Green }

$script:dn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUsername).DistinguishedName
}

#SET USER EMAIL ADDRESS
#============================================================================================================================================================================================
function Set-UserEmailAddress ()
{
    $Script:NEmailAddress = "$Nusername@ntent.com"
    if (-not(Get-MailBox -Identity $NEmailAddress)) { "Could not find a mailbox with email address $NEmailAddress"}
    else { Write-Host "Found" -ForegroundColor Green } 
}

#SET USER'S MANAGER'S EMAIL ADDRESS
#============================================================================================================================================================================================
function Set-MgrEmail ()
{
$Script:NManagerName = read-host "Please enter Manager's alias"
$SCript:NManagerEmailAddress = "$NManagerName@ntent.com"
    if (-not(Get-MailBox -Identity $NManagerEmailAddress)) { "Could not find a mailbox with email address $NManagerEmailAddress"}
    else { Write-Host "Found" -ForegroundColor Green }
}

# SET USER LOCATION
#============================================================================================================================================================================================
function Set-UserLoc ()
{
Write-Host "
Office Location" -ForegroundColor Cyan

    Write-Host "
    1. Carlsbad
    2. New York
    3. Vienna
    4. UK - London
    5. ES - Spain" 
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
        #UK 
        $Script:prefix = "UK"
        }
    "5" {
        #Spain
        $Script:prefix = "ES"
        
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
#============================================================================================================================================================================================ 
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
$Global:CBADServer = "dc7.convera.com"

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
$Global:NYADServer = "ny-dc5.convera.com"

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
$Global:ViADServer = "vn-dc2.convera.com"

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
$Global:UKADServer = "dc7.convera.com"

# ES - Spain
#---------------------------------------------------------- 
$Global:ESStreetAddress = ""
$Global:ESCity = "Barcelona"
$Global:ESState = ""
$Global:ESPostalCode = ""
$Global:EScountry = "ES"
$Global:ESDistName = "spainemployees@ntent.com"
$Global:ESOUPath = "OU=SPAIN,OU=Domain Users,DC=convera,DC=com"
$Global:ESOUXPath = "OU=To Be Deleted,OU=Domain Users,DC=convera,DC=com"
$Global:ESADServer = "dc7.convera.com"



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
#============================================================================================================================================================================================ 
function new-user ()
{
# Check Office 4365 Licenses
#----------------------------------------------------------
$NO365Lics = Get-MsolAccountSku
Write-Host "Checking on Office 365 Licenses...
$NO365Lics | Select AccountSKUID, ActiveUnits, ConsumedUnits"

# Call Functions
#----------------------------------------------------------
Set-MyVariables
Set-UserLoc
Set-UserFull

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
        UserPrincipalName = "$NUsername@convera.com"
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
    Create Office 365 User?" -ForegroundColor Cyan
switch -regex (Read-Host "    [y/n]") {
    "y" {Write-Host ""
        Write-Host "Creating Office 365 user...
        " -ForegroundColor Yellow
        New-MsolUser @NewUserSettings | out-null
        Write-Host "
        O365 COMPLETE!
        " -ForegroundColor Green        
        }
    "n" {Write-Host ""
        Write-Host "Opted not to Create Office 365 User" -ForegroundColor Red; 
        }
}

Write-Host "
    Create Active Directory User?" -ForegroundColor Cyan
switch -regex (Read-Host "    [y/n]") {
    "y" {Write-Host ""
        Write-Host "Creating Active Directory User on: $nadserver...
        " -ForegroundColor Yellow
        New-ADUser @NADNewUser | out-null
        Set-DN
        Move-UserOU
        Set-ObjectName
        Write-Host "
        AD COMPLETE!
        " -ForegroundColor Green
        }
    "n" {Write-Host ""
        Write-Host "Opted not to Create Office 365 User" -ForegroundColor Red;
        }
}

Write-Host "
    Create New User Email?" -ForegroundColor Cyan
switch -regex (Read-Host "    [y/n]") {
    "y" {Write-Host ""
        Write-Host "Creating New User Email...
        " -ForegroundColor Yellow
        send-nuemail
        Write-Host "
        Email COMPLETE!
        " -ForegroundColor Green
        }
    "n" {Write-Host ""
        Write-Host "Opted not to Send New User Email" -ForegroundColor Red; 
        }
}

Write-Host "
    Run set-managerattr Script?" -ForegroundColor Cyan
switch -regex (Read-Host "    [y/n]") {
    "y" {Write-Host ""
        Write-Host "Setting Manager attribute...
        " -ForegroundColor Yellow
        set-managerattr
        Write-Host "
        Manager Attibute COMPLETE!
        " -ForegroundColor Green
        }
    "n" {Write-Host ""
        Write-Host "Opted not to run set-managerattr script" -ForegroundColor Red; 
        }
}
}


#Set Alternate O365 Contact Info
#============================================================================================================================================================================================
function Set-AltContactInfoO365 ()
{
    Set-UserName
    Set-UserEmailAddress
    
    $script:AltEmailO365 = Read-Host "Alternate Emaill Address?"

    Set-MsolUser -UserPrincipalName $NEmailAddress -AlternateEmailaddresses $script:AltEmailO365
}


#Set Manager Attribute in Office 365 and AD
#============================================================================================================================================================================================
function Set-ManagerAttr ()
{
    Set-MyVariables
    Set-UserInfo
    Set-MgrEmail

    set-user -Identity $NUsername -Manager $NManagerEmailAddress
}

#REENABLE USER FOR EMAIL BACKUP
#============================================================================================================================================================================================
function Reenable-O365forPST ()
{
    Set-MyVariables
    Set-UserInfo

    Enable-EmailProtocols
    Set-SignIn-Allowed
    Reset-PWO365

    Get-UserStatus $NUserName
}

#DISABLE USER FOR EMAIL BACKUP
#============================================================================================================================================================================================
function Disable-O365PostPST ()
{
    Set-MyVariables
    Set-UserInfo

    Disable-EmailProtocols
    Set-SignIn-Blocked
    Reset-PWO365

    Get-UserStatus $NUserName
}

#REENABLE USER
#============================================================================================================================================================================================ 
function Reenable-User ()
{
# Call Scripts
#----------------------------------------------------------
Set-MyVariables
Set-UserLoc
Set-UserFull

$NPassword = Read-host "Password"
if (!$Npassword) {$NPassword = "abc123!!"}

$DGs = Get-DistributionGroup

# Variables for Active Directory
#----------------------------------------------------------

$setpassword = ConvertTo-SecureString -AsPlainText $NPassword -force

$NDescription = Read-Host "Title"

Write-Host "
    This will effectively reenable $NUserName, and change the password to $NPassword ?
    " -ForegroundColor Yellow
Write-Host "Enable user in Office 365?"
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Office 365 user will be reenabled
        " -ForegroundColor Cyan
        $NReenaableO365 = $True
        }
    "n" { Write-Host "
        Opted not to disable Office 365 User" -ForegroundColor Red;
        $NReenableO365 = $False
        }
        }


Write-Host "Enable user in Active Directory?"
switch -regex (Read-Host "    [y/n]") {
    "y" { Write-Host "
        Active Directory user will be reenabled
        " -ForegroundColor Cyan
        $NReenableAD = $True
        }
    "n" { Write-Host "
        Opted not to disable Active Directory User" -ForegroundColor Red;
        $NReenableAD = $False
        }
        }


If ($NReenableO365 -eq $True) 
        { 
        Set-HideFromGAL-Off
        Enable-EmailProtocols
        Set-SignIn-Blocked
        Set-FwdAddress
        Set-PWO365
        }

If ($NReenableAD -eq "$True") 
        { 
        Reset-Description-AD
        Reset-PW-AD
        Reenable-UserAccount-AD
        Move-User-AD+
        }

Get-UserStatus $NUsername


# Pull User Account List
#----------------------------------------------------------
get-Mailbox -identity $NUserName | Select @{ l="Name"; e={ $_.DisplayName } }, @{ l="Active Directory"; e={ $_.CustomAttribute1 } }, @{ l="Office365"; e={ $_.CustomAttribute2 } }, @{ l="Sales Force"; e={ $_.CustomAttribute3 } }, @{ l="Google Apps"; e={ $_.CustomAttribute4 } }, @{ l="Podio"; e={ $_.CustomAttribute5 } }, @{ l="Atlassian"; e={ $_.CustomAttribute6 } }, @{ l="Secure NTENT"; e={ $_.CustomAttribute7 } }, @{ l="Ad Support"; e={ $_.CustomAttribute8 } }, @{ l="Tableau"; e={ $_.CustomAttribute9 } }, @{ l="Great Plains"; e={ $_.CustomAttribute10 } } |out-gridview


# Pull User's Distribution Group Memberships
#----------------------------------------------------------
Get-UserDL-


$err

Write-host "
S C R I P T    C O M P L E T E * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"  -ForegroundColor DarkBlue -BackgroundColor Cyan 
}

#DISABLE USER
#============================================================================================================================================================================================ 
function Disable-User ()
{
# Call Scripts
#----------------------------------------------------------
Set-MyVariables
Set-UserLoc
Set-UserInfo
Set-MgrEmail

$NPassword = Read-host "Password"
if (!$Npassword) {$NPassword = "G000dby3"}

$DGs = Get-DistributionGroup

# Variables for Active Directory
#----------------------------------------------------------

$Date = Read-Host 'Enter Term Date (i.e. mm/dd)'
if (!$Date) {$Date = Get-Date -Format "MM/dd"}

$setpassword = ConvertTo-SecureString -AsPlainText $NPassword -force

$NDescription = "-term $Date - email fwd to $NManagerEmailAddress"

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
        Set-HideFromGAL-On
        Disable-EmailProtocols
        Set-SignIn-Blocked
        Set-FwdAddress
        Set-PWO365
        }

If ($NDisableAD -eq "$True") 
        { 
        Set-Description-AD
        Set-PW-AD
        Disable-UserAccount-AD
        Move-User-AD
        }

Get-UserStatus $NUsername


# Pull User Account List
#----------------------------------------------------------
get-Mailbox -identity $NUserName | Select @{ l="Name"; e={ $_.DisplayName } }, @{ l="Active Directory"; e={ $_.CustomAttribute1 } }, @{ l="Office365"; e={ $_.CustomAttribute2 } }, @{ l="Sales Force"; e={ $_.CustomAttribute3 } }, @{ l="Google Apps"; e={ $_.CustomAttribute4 } }, @{ l="Podio"; e={ $_.CustomAttribute5 } }, @{ l="Atlassian"; e={ $_.CustomAttribute6 } }, @{ l="Secure NTENT"; e={ $_.CustomAttribute7 } }, @{ l="Ad Support"; e={ $_.CustomAttribute8 } }, @{ l="Tableau"; e={ $_.CustomAttribute9 } }, @{ l="Great Plains"; e={ $_.CustomAttribute10 } } |out-gridview


# Pull User's Distribution Group Memberships
#----------------------------------------------------------
Get-UserDL-


$err

Write-host "
S C R I P T    C O M P L E T E * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *"  -ForegroundColor DarkBlue -BackgroundColor Cyan 
}

#Add USER TO DISTRIBUTION GROUP(S)
#============================================================================================================================================================================================
function Add-UserDistributionGroup ()
{
set-username
Set-UserEmailAddress

#Add User from Distribution Groups
do 
    {
    $DistributionGroup = Read-Host "Please enter the distribution Name"
    if ($DistributionGroup -ne "")
        {Add-DistributionGroupMember -Identity $DistributionGroup -member $NUsername }
    } 
    while ($DistributionGroup -ne "")
}

#REMOVE USER FROM DISTRIBUTION GROUP(S)
#============================================================================================================================================================================================
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
#============================================================================================================================================================================================ 
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
        If ($mb_ForwardSMTP -eq $null) {"[  FALSE  ] Mail Forwarded [User's OWA]"}
        Else {Write-Host "[  TRUE   ] Mail Forwarded [User's OWA]" -ForegroundColor Red}
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
#============================================================================================================================================================================================ 
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
If ( $NOut -eq 'HTML' ) { $NGetDirectoryUSers | ConvertTo-Html -body $mycss | Set-content "C:\Users\madmin\Documents\WindowsPowerShell\Reports\NTENTDir" }
If ( $NOut -eq 'GridView' ) { $NGetDirectoryUSers | Out-GridView }

}

#GET USER DIRECTORY FOR IGLOO
#============================================================================================================================================================================================ 
function Get-CoDirIgloo ()
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
        } | Select FirstName, LastName, userPrincipalName, StreetAddress, City, State, PostalCode, Country, MobilePhone, PhoneNumber, Title | Sort-Object -Property DisplayName 
<#
$GetDirectoryUsers = Get-MsolUser | ? { ($_.UserType -eq 'Member') -and $_.IsLicensed -and ($_.Country -eq 'US') } | % {
    get-mailbox | ? { $_.HiddenFromAddressListsEnabled }
    } | Select Name, objectid | Sort Name
#>

If ( $NOut -eq 'CSV' ) { $NGetDirectoryUSers | Export-Csv "c:\users\madmin\desktop\NTENT Dir $(Get-Date -f "MM-dd hhmm-ss").csv" -NoTypeInformation }
If ( $NOut -eq 'HTML' ) { $NGetDirectoryUSers | ConvertTo-Html -body $mycss | Set-content "$mreports\NTENT Dir $(Get-Date -f "MM-dd hhmm-ss").html" }
If ( $NOut -eq 'GridView' ) { $NGetDirectoryUSers | Out-GridView }
}

#CHANGE USER PASSWORD
#============================================================================================================================================================================================ 
function Change-UserPW ()
{ 
Set-UserName
Set-UserEmailAddress

$c = Get-Credential -Message "Enter password for this user" -username $EmailAddress
$Password = $C.getnetworkcredential().password

Set-MsolUser  -UserPrincipalName $EmailAddress -StrongPasswordRequired $False
Set-MsolUserPassword -UserPrincipalName $EmailAddress -NewPassword $Password -ForceChangePassword $false | Out-Null
}

#HIDE USER FROM GAL
#============================================================================================================================================================================================ 
function Hide-UserGAL ()
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $true
}

#SHOW USER IN GAL
#============================================================================================================================================================================================ 
function Show-UserGAL ()
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $False
}

#FORWARD USER'S EMAIL
#============================================================================================================================================================================================ 
function Set-EmailFwd () 
{
Set-UserName
Set-UserEmailAddress
Set-MgrEmail

Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $true -ForwardingAddress $NManagerEmailAddress
}

#REMOVE FORWARD FROM USER
#============================================================================================================================================================================================ 
function Remove-EmailFwd () 
{
Set-UserName
Set-UserEmailAddress

Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $False -ForwardingAddress $Null
}

# Resize the standard console window
#============================================================================================================================================================================================ 
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
#============================================================================================================================================================================================ 
function Get-UserMembership ()
{
    $Script:NUserDN = (get-mailbox $Nusername).distinguishedname

    "User " + $NUsername + " is a member of the following groups:"

    foreach ($group in get-distributiongroup -resultsize unlimited)
    {
    if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $NUserDN){$group.name}
    }    
}

#REMOVE USER FROM ALL O365 GROUPS
#============================================================================================================================================================================================ 
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

#Set $DN Variable (Used for AD)
#============================================================================================================================================================================================ 
function set-DN ()
{
    $Script:dn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUserName).DistinguishedName
}

#Move User's AD Account to a Different OU
#============================================================================================================================================================================================ 
function Move-UserOU ()
{
    Move-ADObject -server $NADServer -Credential $ADCred -Identity $dn -TargetPath $NOUPath
}

#Rename the object to a good looking name to avoid displaying sAMAccountNames (eg tests1.user1)
#============================================================================================================================================================================================ 
function Set-ObjectName ()
{
    $Script:newdn = (Get-ADUser -server $NADServer -Credential $ADCred -Identity $NUsername).DistinguishedName
    Rename-ADObject -server $NADServer -Credential $ADCred -Identity $newdn -NewName $NFullname
}

#Hide User From Global Address List Enabled
#============================================================================================================================================================================================ 
function Set-HideFromGAL-On ()
{
   Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $true
   Write-host "-Completed - O365: Hide from Global Address List" -ForegroundColor Green
}

#Hide User From Global Address List Disabled
#============================================================================================================================================================================================ 
function Set-HideFromGAL-Off ()
{
   Set-Mailbox -Identity $NEmailAddress -HiddenFromAddressListsEnabled $false
   Write-host "-Completed - O365: Hide from Global Address List" -ForegroundColor Green
}


#Enable Active Sync, OWA, POP, IMAP and MAPI
#============================================================================================================================================================================================ 
function Enable-EmailProtocols ()
{
    Set-CASMailbox -Identity $NEmailAddress -ActiveSyncEnabled $True -OWAEnabled $True -OWAforDevicesEnabled $True -PopEnabled $True -ImapEnabled $True -MAPIEnabled $True
    Write-host "-Completed - O365: Enabled Active Sync and Email Protocols" -ForegroundColor Green
}

#Disable Active Sync, OWA, POP, IMAP and MAPI
#============================================================================================================================================================================================ 
function Disable-EmailProtocols ()
{
    Set-CASMailbox -Identity $NEmailAddress -ActiveSyncEnabled $False -OWAEnabled $False -OWAforDevicesEnabled $false -PopEnabled $False -ImapEnabled $False -MAPIEnabled $False
    Write-host "-Completed - O365: Disable Active Sync" -ForegroundColor Green
}

#Set sign-in status to Allowed
#============================================================================================================================================================================================ 
function Set-SignIn-Allowed ()
{
    Set-MsolUser -UserPrincipalName $NEmailAddress -BlockCredential $False
    Write-host "-Completed - O365: Sign-in Status changed to 'Allowed'" -ForegroundColor Green
}

#Set sign-in status to Blocked
#============================================================================================================================================================================================ 
function Set-SignIn-Blocked ()
{
    Set-MsolUser -UserPrincipalName $NEmailAddress -BlockCredential $True
    Write-host "-Completed - O365: Sign-in Status changed to 'Blocked'" -ForegroundColor Green
}

#Disable Email Forwarding to Manager
#============================================================================================================================================================================================ 
function Set-FwdAddress-Off ()
{
    Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $false -ForwardingAddress $Null
    Write-host "-Completed - O365: Disabled Email Forwarding" -ForegroundColor Green
}

#Forward email to Manager
#============================================================================================================================================================================================ 
function Set-FwdAddress ()
{
    Set-Mailbox -Identity $NEmailAddress -DeliverToMailboxAndForward $true -ForwardingAddress $NManagerEmailAddress
    Write-host "-Completed - O365: Forward Email" -ForegroundColor Green
}

#Reset Password
#============================================================================================================================================================================================ 
function Reset-PWO365 ()
{
    $Script:NResetPW = Read-Host 'New Password'

    Set-MsolUser  -UserPrincipalName $NEmailAddress -StrongPasswordRequired $False
    Set-MsolUserPassword -UserPrincipalName $NEmailAddress -NewPassword $NResetPW -ForceChangePassword $false | Out-Null
    Write-host "-Completed - O365: Changed Password" -ForegroundColor Green
}
#Set Password
#============================================================================================================================================================================================ 
function Set-PWO365 ()
{
    Set-MsolUser  -UserPrincipalName $NEmailAddress -StrongPasswordRequired $False
    Set-MsolUserPassword -UserPrincipalName $NEmailAddress -NewPassword $NPassword -ForceChangePassword $false | Out-Null
    Write-host "-Completed - O365: Changed Password" -ForegroundColor Green
}

#LIST USER'S DISTRIBUTION GROUPS
#============================================================================================================================================================================================ 
function Get-UserDL- () 
{
    $Script:DNO365 = (get-mailbox $NUserName).distinguishedname

    Write-Host "$NUsername is a member of the following O365 groups:" -ForegroundColor Cyan

    foreach ($group in get-distributiongroup -resultsize unlimited)
    {
        if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $DNO365){$group.name}
    }
}

#LIST USER'S DISTRIBUTION GROUPS w/ USERNAME PROMPT
#============================================================================================================================================================================================ 
function Get-UserDL () 
{
    $Script:NUserName = read-host -Prompt "Username" 

    $DNO365 = (get-mailbox $NUserName).distinguishedname

    Write-Host "$NUsername is a member of the following O365 groups:" -ForegroundColor Cyan

    foreach ($group in get-distributiongroup -resultsize unlimited)
    {
        if ((get-distributiongroupmember $group.identity | select -expand distinguishedname) -contains $DNO365){$group.name}
    }
}

#Reset AD Description
#============================================================================================================================================================================================ 
function Reset-Description-AD  ()
{
    Write-Host "-Changing AD description for $NUserName to $NDescription" -ForegroundColor Green
    Set-ADUser -server $NADServer -Credential $ADCred -Identity $dn -Description (Get-ADUser -identity mbennett -properties title).title $NDescription -ea SilentlyContinue -ev err
}

#Change AD Description
#============================================================================================================================================================================================ 
function Set-Description-AD  ()
{
    Write-Host "-Changing AD description for $NUserName to $NDescription" -ForegroundColor Green
    Set-ADUser -server $NADServer -Credential $ADCred -Identity $dn -Description $NDescription -ea SilentlyContinue -ev err
}


#Reset ADUser Password
#============================================================================================================================================================================================ 
function Reset-PW-AD ()
{
    $Script:ResetPassword = Read-Host "New Password"

    Write-Host "-Changing AD password to $NPassword" -ForegroundColor Green
    Set-ADAccountPassword -server $NADServer -Credential $ADCred -Identity $dn -NewPassword $ResetPassword -ea SilentlyContinue -ev +err
}
#Change ADUser Password
#============================================================================================================================================================================================ 
function Set-PW-AD ()
{
    Write-Host "-Changing AD password to $NPassword" -ForegroundColor Green
    Set-ADAccountPassword -server $NADServer -Credential $ADCred -Identity $dn -NewPassword $SetPassword -ea SilentlyContinue -ev +err
}

#Renable ADUser Account
#============================================================================================================================================================================================ 
function Renabe-UserAccount-AD ()
{
    Write-Host "-Reenabling $NUserName's User Account" -ForegroundColor Green
    Enable-ADAccount -Identity $dn -Credential $ADCred
}

#Disabe ADUser Account
#============================================================================================================================================================================================ 
function Disable-UserAccount-AD ()
{
    Write-Host "-Disabling $NUserName's User Account" -ForegroundColor Green
    Disable-ADAccount -Identity $dn -Credential $ADCred
}

#Moving User to UserInfoFull OU
#============================================================================================================================================================================================ 
function  Move-User-AD+ ()
{
    Write-Host "-Moving $NUserName to the 'To Be Delete' OU ( $NOUXPath )" -ForegroundColor Green
    Move-ADObject -server $NADServer -Credential $ADCred -Identity $dn -TargetPath $NOUPath -ea SilentlyContinue -ev +err
}

#Moving User to To Be Deleted OU
#============================================================================================================================================================================================ 
function  Move-User-AD ()
{
    Write-Host "-Moving $NUserName to the 'To Be Delete' OU ( $NOUXPath )" -ForegroundColor Green
    Move-ADObject -server $NADServer -Credential $ADCred -Identity $dn -TargetPath $NOUXPath -ea SilentlyContinue -ev +err
}

#AD USER LAST LOGIN
#============================================================================================================================================================================================
function Get-ADUserLogonStatus ()
{
    $UserName = Read-Host "UserName"

    $Expiry = Get-ADUser $Username –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
    $LastLogon = Get-ADUser $username –Properties * | Select-Object -Property lastlogondate

    $username
    Write-Host "Expiry Date: "$expiry.expirydate""
    Write-Host "Last Logon: "$LastLogon.lastlogondate""
  
}

#AD USER EXPIRY LIST
#============================================================================================================================================================================================
function Get-ADUserExpiryDates ()
{
    Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and passwordexpired -eq $false} –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" | Select-Object -Property "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}} | Out-GridView
}

#GET USER AUDIT [OFFICE 365]
#============================================================================================================================================================================================
function Get-UserAuditO3651 ()
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

$UserAuditO365 = Get-MsolUser | ? { ($_.UserType -eq 'Member') -and $_.IsLicensed } | % { get-mailbox -identity $_.displayname } | Select DisplayName, HiddenFromAddressListsEnabled, ForwardingAddress, ForwardingSmtpAddress, Office | Sort Office

If ( $NOut -eq 'CSV' ) { $UserAuditO365 | Export-Csv "C:\users\mbennett\Desktop\PSReports\UserAuditO365-- $(Get-Date -f "MM-dd hhmm-ss").csv" -NoTypeInformation }
If ( $NOut -eq 'HTML' ) { $UserAuditO365 | ConvertTo-Html -body $mycss | Set-content "C:\users\mbennett\Desktop\PSReports\UserAuditO365-- $(Get-Date -f "MM-dd hhmm-ss").html" }
If ( $NOut -eq 'GridView' ) { $UserAuditO365 | Out-GridView }

$UserAuditO365.count

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

#SEND NEW USER EMAIL
#----------------------------------------------------------################################################################## 
function Send-NUEmail ()
{
$ol = New-Object -comObject Outlook.Application

$muser = [Environment]::GetFolderPath("user")

$Nbody = @"
<!DOCTYPE html>
<html>
<head>
<style>
h2 { 
    display: block;
    font-size: 1.5em;
    margin-top: 0.67em;
    margin-bottom: 0.0em;
    margin-left: 0;
    margin-right: 0;
    font-weight: normal;
}
hr { 
width:100%;
height:1px;
background: #fff;
margin-top: 0.em;
margin-bottom: 0.5em; 
}
a,a:active,a:hover {
color: #cd3427;
}
body {
font-family: "Calibri";
}
</style>
</head>
<body>
Hi $NFirstName,<br>
<h2>Welcome to <span style="color: #cd3427;">$NCompany`</span>!</h2><br>
Below you will find helpful information to get you started...<br>
<br>
Your temporary password for all apps, unless otherwise specified, is: <span style="text-decoration: underline;"><span style="color: #cd3427; text-decoration: underline;">$NTempPWNewUser</span></span><br>
<hr>
<strong>Domain / Active Directory</strong><br>
Username: <span style="color: #cd3427;">$NUsername</span><br>
<br>
<strong>Wireless</strong><br>
NTENT Wireless: SSID = <span style="color: #cd3427;">Ntent Guest</span><br>
Password: <span style="color: #cd3427;">welcome2ntent</span><br>
<br>
<strong>Email</strong><br>
Web Access: <span style="color: #cd3427;">outlook.office365.com</span><br>
Username: <span style="color: #cd3427;">$NEmailAddress</span><br>
<br>
<strong>Jira / Confluence</strong><br>
Web Access: <a href="http://jira.ntent.com">jira.ntent.com</a> or <a href="http://confluence.ntent.com">confluence.ntent.com</a><br>
Username: <span style="color: #cd3427;">$Nusername</span><br>
<br>
<strong>Slack</strong><br>
URL: https://ntent.slack.com/signup<br>
Enter: <span style="color: #cd3427;">$NEmailAddress</span><br>
Check your email to verify your identity and set a password.<br>
<br>
<strong>Igloo (intranet)</strong><br>
Web Access: <a href="http://insidentent.com">insidentent.com</a><br>
USE SSO: Click on ‘Use: Network (Domain) Login’<br>
Username: <span style="color: #cd3427;">$NUsername</span><br>
<br>
<strong>Phone</strong><br>
Your Office Phone Number is: <span style="color: #cd3427;">$NDirectNumber</span><br>
<br>
<strong>Voicemail</strong><br>
To access voicemail, dial: <span style="color: #cd3427;">6000</span><br>
Your voicemail password is: <span style="color: #cd3427;">1111</span><br>
<br>
For instructions on how to setup and configure your voicemail, follow this link:<br>
$NVoiceMailSetup<br>
<hr>
If you have any trouble, navigate to helpdesk.ntent.com and submit a ticket through the Helpdesk Portal.<br>
<br>
Alternatively, you could email helpdesk@ntent.com<br>
<br>
<br>
<strong>Mike Bennett</strong> | Sr. Sys Admin | <span style="color: #cd3427;"><strong>NTENT</strong></span><br>
<a href="mailto:mbennett@ntent.com">mbennett@ntent.com</a> | 760.930.7687<br>
<a href="https://www.linkedin.com/company/ntent?utm_source=NTENT%20Email%201&amp;utm_medium=Email%20Signature&amp;utm_campaign=Email%20Sig%20-%20LinkedIn">LinkedIn</a> | <a href="https://twitter.com/withntent?utm_source=NTENT%20Email%201&amp;utm_medium=Email%20Signature&amp;utm_campaign=Email%20Sig%20-%20Twitter">Twitter</a></body><br>
</html>
"@

$mail = $ol.CreateItem(0)
$mail.Subject = "Welcome to $NCompany"
$mail.save()
$mail.to = "$NEmailAddress"
$mail.HTMLBody = $nbody

$inspector = $mail.GetInspector
$inspector.Display()

$NBOdy | out-file "$Muser\desktop\Welcome Email\Welcome Email $(Get-Date -f "MM-dd hhmm-ss").doc"
}

#SEND NEW USER EMAIL
#----------------------------------------------------------################################################################## 
function Get-NUEmail ()
{
Set-MyVariables
Set-UserLoc
Set-UserFull

Send-NUEmail

Invoke-Item "$Muser\desktop\Welcome Email"
}

#SEND NEW USER EMAIL
#----------------------------------------------------------################################################################## 
function Remove-MyPSS ()
{
Get-PSSession | Remove-PSSession
}

#New Stored Credential
#----------------------------------------------------------################################################################## 
Function New-StoredCredential {

    <#
    .SYNOPSIS
    New-StoredCredential - Create a new stored credential

    .DESCRIPTION 
    This function will save a new stored credential to a .cred file.

    .EXAMPLE
    New-StoredCredential

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path

        if (!(Test-Path $KeyPath)) {
        
            try {
                New-Item -ItemType Directory -Path $KeyPath -ErrorAction STOP | Out-Null
            }
            catch {
                throw $_.Exception.Message
            }           
        }
    }

    $Credential = Get-Credential -Message "Enter a user name and password"

    $Credential.Password | ConvertFrom-SecureString | Out-File "$($KeyPath)\$($Credential.Username).cred" -Force

}


#Get Stored Credential
#----------------------------------------------------------################################################################## 
Function Get-StoredCredential {

    <#
    .SYNOPSIS
    Get-StoredCredential - Retrieve or list stored credentials

    .DESCRIPTION 
    This function can be used to list available credentials on
    the computer, or to retrieve a credential for use in a script
    or command.

    .PARAMETER UserName
    Get the stored credential for the username

    .PARAMETER List
    List the stored credentials on the computer

    .EXAMPLE
    Get-StoredCredential -List

    .EXAMPLE
    $credential = Get-StoredCredential -UserName admin@tenant.onmicrosoft.com

    .EXAMPLE
    Get-StoredCredential -List

    .LINK
    https://practical365.com/saving-credentials-for-office-365-powershell-scripts-and-scheduled-tasks
    
    .NOTES
    Written by: Paul Cunningham

    Find me on:

    * My Blog:	http://paulcunningham.me
    * Twitter:	https://twitter.com/paulcunningham
    * LinkedIn:	http://au.linkedin.com/in/cunninghamp/
    * Github:	https://github.com/cunninghamp

    For more Office 365 tips, tricks and news
    check out Practical 365.

    * Website:	https://practical365.com
    * Twitter:	https://twitter.com/practical365
    #>

    param(
        [Parameter(Mandatory=$false, ParameterSetName="Get")]
        [string]$UserName,
        [Parameter(Mandatory=$false, ParameterSetName="List")]
        [switch]$List
        )

    if (!(Test-Path Variable:\KeyPath)) {
        Write-Warning "The `$KeyPath variable has not been set. Consider adding `$KeyPath to your PowerShell profile to avoid this prompt."
        $path = Read-Host -Prompt "Enter a path for stored credentials"
        Set-Variable -Name KeyPath -Scope Global -Value $path
    }


    if ($List) {

        try {
        $CredentialList = @(Get-ChildItem -Path $keypath -Filter *.cred -ErrorAction STOP)

        foreach ($Cred in $CredentialList) {
            Write-Host "Username: $($Cred.BaseName)"
            }
        }
        catch {
            Write-Warning $_.Exception.Message
        }

    }

    if ($UserName) {
        if (Test-Path "$($KeyPath)\$($Username).cred") {
        
            $PwdSecureString = Get-Content "$($KeyPath)\$($Username).cred" | ConvertTo-SecureString
            
            $Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $PwdSecureString
        }
        else {
            throw "Unable to locate a credential for $($Username)"
        }

        return $Credential
    }
}

Export-ModuleMember -function Load-Profile
Export-ModuleMember -function Set-Profile
Export-ModuleMember -function Connect-AD
Export-ModuleMember -function Connect-AD-Remote
Export-ModuleMember -function Connect-O365
Export-ModuleMember -function Connect-Mod
Export-ModuleMember -function Set-MyVariables
Export-ModuleMember -function New-User
Export-ModuleMember -function Set-AltContactInfoO365
Export-ModuleMember -function Set-ManagerAttr
Export-ModuleMember -function Reenable-O365forPST
Export-ModuleMember -function Disable-O365PostPST
Export-ModuleMember -function Reenable-User
Export-ModuleMember -function Disable-User
Export-ModuleMember -function Add-UserDistributionGroup
Export-ModuleMember -function Remove-UserDistributionGroup
Export-ModuleMember -function Get-UserStatus
Export-ModuleMember -function Get-CoDir
Export-ModuleMember -function Get-CoDirIgloo
Export-ModuleMember -function Change-UserPW
Export-ModuleMember -function Hide-UserGAL
Export-ModuleMember -function Show-UserGAL
Export-ModuleMember -function Add-EmailFwd
Export-ModuleMember -function Remove-EmailFwd
Export-ModuleMember -function Resize-ConsoleWindow
Export-ModuleMember -function Get-UserMembership
Export-ModuleMember -function Remove-UserMembership
Export-ModuleMember -function Set-DN
Export-ModuleMember -function Move-UserOU
Export-ModuleMember -function Set-ObjectName
Export-ModuleMember -function Set-HideFromGAL-On
Export-ModuleMember -function Set-HideFromGAL-Off
Export-ModuleMember -function Enable-EmailProtocols
Export-ModuleMember -function Disable-EmailProtocols
Export-ModuleMember -function Set-SignIn-Allowed
Export-ModuleMember -function Set-SignIn-Blocked
Export-ModuleMember -function Set-FwdAddress-Off
Export-ModuleMember -function Set-FwdAddress
Export-ModuleMember -function Reset-PWO365
Export-ModuleMember -function Set-PWO365
Export-ModuleMember -function Get-UserDL
Export-ModuleMember -function Get-UserDL-
Export-ModuleMember -function Reset-Description-AD
Export-ModuleMember -function Set-Description-AD
Export-ModuleMember -function Set-PW-AD
Export-ModuleMember -function Reenable-UserAccount-AD
Export-ModuleMember -function Disable-UserAccount-AD
Export-ModuleMember -function Move-User-AD+
Export-ModuleMember -function Move-User-AD
Export-ModuleMember -function Get-ADUserLogonStatus
Export-ModuleMember -function Get-ADUserExpiryDates
Export-ModuleMember -function Get-UserAuditO3651
Export-ModuleMember -function Change-PWO365
Export-ModuleMember -function Change-PWAD
Export-ModuleMember -function Send-NUEmail
Export-ModuleMember -function Get-NUEmail
Export-ModuleMember -function Remove-MyPSS
Export-ModuleMember -function New-StoredCredential
Export-ModuleMember -function Get-StoredCredential
Export-ModuleMember -alias rcw