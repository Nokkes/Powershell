. ..\..\Init.ps1
UniversalDate
TechDate

Winfo "=========================================================================================="
Winfo "Started $PSCommandPath by $env:USERNAME on $env:COMPUTERNAME at $UniversalDate"
Winfo "=========================================================================================="

StartWatch
$ErrorActionPreference = "SilentlyContinue"
$ScriptProps = @{
    EmailAddress = "arnovermeiren@hotmail.com"
    DDG = "ddg.gg" 
    CFL = "http://www.cheapflightslab.com/category/deals-from-europe"
    FlyNous = "http://www.flynous.com/Flights-from/benelux"
    CFLChanged = [bool]$CflChanged
    FlyNousChanged = [bool]$FlyNousChanged
    EmailBody = @()
    HeaderCFL = "<h1>Items have changed on CFL</h1>"
    HeaderFlyNous = "<h1>Items have changed on FlyNous</h1>"
}
$Creds = Import-PSCredential -Path C:\Scripts\PS\Hotmailcred

### Query CFL 1###
Winfo "============================"
Winfo "Querying CheapFlightsLab"
Winfo "============================"
Winfo "Starting $($ScriptProps.DDG)"
StartIE -URL $ScriptProps.DDG
WaitIE
Winfo "Navigating to $($ScriptProps.CFL)"
$IEwindow.GoTo($ScriptProps.CFL)
WaitIE
Winfo "Searching IE"
$DivCFL = SearchDivByClass -String "post-row"
$NewCFLDiv = $DivCFL.OuterText
$a = $NewCFLDiv -replace "`n|`r"
$b = $a.Replace('\s','')



# Dump Div Text to txt for compare as CFL changes HTML Props
Winfo "Writing Files"
$b > "C:\Scripts\PS\HTML\CheapFlightsLab_$TechDate.txt"
$DivCFL > "C:\Scripts\PS\HTML\CheapFlightsLab_$TechDate.html"

### Query FlyNous ###
Winfo "============================"
Winfo "Querying FlyNous"
Winfo "============================"
Winfo "Starting $($ScriptProps.DDG)"
StartIE -URL $ScriptProps.DDG
WaitIE
Winfo "Navigating to $($ScriptProps.Flynous)"
$IEwindow.GoTo($ScriptProps.FlyNous)
WaitIE
Winfo "Searching IE"
$DivFlyNous = SearchDivByClass -String "clearfix grid_11"
Winfo "Writing File"
$DivFlyNous.OuterHtml > "C:\Scripts\PS\HTML\FlyNous_$TechDate.html"



### Get 2 newest files per site ###
Winfo "Getting 2 most recent files per site"
$TwoNewestCFL = (gci -Path "C:\Scripts\PS\HTML\" -Filter "*cheapflightslab*.txt" | Sort-Object LastAccessTime -Descending | Select-Object -First 2).FullName
$TwoNewestFlyNous = (gci -Path "C:\Scripts\PS\HTML\" -Filter "*FlyNous*" | Sort-Object LastAccessTime -Descending | Select-Object -First 2).FullName


### Getting Hashes ###
Winfo "============================"
Winfo "MD5's"
Winfo "============================"
$MD5HashCFL1 = MD5Hash -Location "$($TwoNewestCFL[0])"
$MD5HashCFL2 = MD5Hash -Location "$($TwoNewestCFL[1])"
Winfo "Newest MD5 CFL = $MD5HashCFL1"
Winfo "Oldest MD5 CFL = $MD5HashCFL2"

$MD5HashFlyNous1 = MD5Hash -Location "$($TwoNewestFlyNous[0])"
$MD5HashFlyNous2 = MD5Hash -Location "$($TwoNewestFlyNous[1])"
Winfo "Newest MD5 Flynous = $MD5HashFlyNous1"
Winfo "Oldest MD5 Flynous = $MD5HashFlyNous2"


### Setting values based on MD5's ###
# CheapFlightLab #

if ($MD5HashCFL1 -ne $MD5HashCFL2) {
    $ScriptProps.CFLChanged = $true
}
else {
    $ScriptProps.CFLChanged = $false
}

# FlyNous #
if ($MD5HashFlyNous1 -ne $MD5HashFlyNous2) {
    $ScriptProps.FlyNousChanged = $true
}
else {
    $ScriptProps.FlyNousChanged = $false
}




if ($ScriptProps.CFLChanged -eq $true) {
    Winfo "CFL has changed, adding CFL webpage to EmailBody"
    $ScriptProps.EmailBody+=$ScriptProps.HeaderCFL+=$DivCFL.OuterHtml
}
else {
    Winfo "CFL hasn't changed"
}

if ($ScriptProps.FlyNousChanged -eq $true) {
    winfo "FlyNous has changed, adding FlyNous webpage to EmailBody"
    $ScriptProps.EmailBody+=$ScriptProps.HeaderFlyNous+=$DivFlyNous.OuterHtml
}
else {
    Winfo "FlyNous hasn't changed"
    
}

$messageParameters = @{
    Subject = "Digest: Items have changed"
    Body = "$($ScriptProps.EmailBody)"
    BodyAsHTML = $true
    From = "$($ScriptProps.emailaddress)"
    To = "$($ScriptProps.emailaddress)"
    SmtpServer = "smtp-mail.outlook.com" 
    Port = 587
    UseSsl = $true
    Verbose = $true
    Debug = $true
    Credential = $Creds
}

Winfo "Sending Email"
Send-MailMessage @messageParameters

Winfo "Stopping IE processes"
Get-Process iexplore | Stop-Process
StopWatch

Winfo "============================"
Winfo "Errors"
Winfo "============================"
winfo "$($Error.Count)"

pause
#Start-Job -InitializationScript $OpenVpn -ScriptBlock {OpenVPN -ConfigFile FR}