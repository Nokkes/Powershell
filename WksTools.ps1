##################################################################################
#                                                                                #
# Name      :    WksTool.ps1                                                     #
# Location  :    \\Jafile01\i_o_distributed_systems\Tools\Scripts\Arno\WksTool   #
# Created   :    08/12/2016                                                      #
# Modified  :    09/12/2016                                                      #
# Created by:    arno.vermeiren@aginsurance.be                                   #
#                                                                                #
##################################################################################

# Changing Location and parsing Functions
pushd \\Jafile01\i_o_distributed_systems\Tools\Scripts
. .\Functions.ps1

# Formatting Console window
$PsHost = Get-Host                  # Get the PowerShell Host.
$PsWindow = $PsHost.UI.RawUI        # Get the PowerShell Host's UI.
$BufferSize = $PsWindow.BufferSize  # Get the UI's current Buffer Size.
$BufferSize.Width = 80              # Set the new buffer's width to 80.
$BufferSize.Height = 50             # Set the new buffer's Height to 50.
$PsWindow.BufferSize = $BufferSize  # Set the new Buffer Size as active. 
$WindowSize = $PsWindow.WindowSize  # Get the UI's current Window Size.
$WindowSize.Width = 80              # Set the new Window Width to 80 columns.
$WindowSize.Height = 42             # Set the new Window Height to 42 columns.
$PsWindow.WindowSize = $WindowSize  # Set the new Window Size as active.

# Getting UserName and storing Credentias
$AdminUser =  whoami
$AdminUser = ($AdminUser).Replace(($AdminUser).Remove(4),"AG\L")
$Creds = Get-Credential -Message "Enter you Administrator credentials" -UserName $AdminUser

# Making Header
$Header = 
@"
================================================================================
                             Workstation Tool                                                                                           
================================================================================
"@

# Writing Header
Write-Output $Header

# Hostname input and formatting
do {Write-Output "Please enter a Workstation Hostname [PCxxxxxx]: "; $HostName = Read-Host} 
until ($Hostname -match 'PC\d\d\d\d\d\d')
$Hostname = $Hostname.ToUpper()

# Test if Hostname is pingable
if (Test-Connectivity -FunHostname $Hostname) {
    Write-Host -ForegroundColor Green "$Hostname is Online, continuing"
    Start-Sleep 2
}
else {
    Write-Host -ForegroundColor Red "$Hostname is Offline, quitting"
    Start-Sleep 2
    break
}

Clear-Host

#Gather Network Info
$NetworkCIM = Get-WmiObject -ClassName Win32_NetworkAdapterConfiguration -ComputerName $HostName -Credential $Creds | ? {$_.IPEnabled -eq $true}
$NetworkInfo = "$($NetworkCIM.Description)" + "`n              " + "IP  = $($NetworkCIM.IPAddress)" + "`n              " + "SN  = $($NetworkCIM.IPSubnet)" + "`n              " + "GW  = $($NetworkCIM.DefaultIPGateway)" + "`n              " + "MAC = $($NetworkCIM.MACAddress)"

#Gather System Info
$SystemCIM = Get-WmiObject -ClassName Win32_ComputerSystem -ComputerName $HostName -Credential $Creds
$SystemInfo = $SystemCIM.Manufacturer + " | " + $SystemCIM.Model

#Gather uptime
$UptimeCIM = (Get-WmiObject -ClassName Win32_OperatingSystem -ComputerName $Hostname -Credential $Creds).LastBootupTime

# Making Information Jumbotron
$Information = 
"Hostname    : $Hostname
Network     : $NetworkInfo
Systeminfo  : $Systeminfo
Uptime      : $Uptime
========================================
"

# Writing Header + Jumbotron
Write-Output $Header
Write-Output $Information

# Writing option list
[int]$Choice = 0
while ($Choice -lt 1 -or $Choice -gt 8 ) {
    Write-Output "1. Display Extended Information"
    Write-Output "2. Map Drive (K:)"
    Write-Output "3. Start Regedit"
    Write-Output "4. Start EventVwr"
    Write-Output "5. Start MSRA"
    Write-Output "6. List Processes"
    Write-Output "7. New Hostname"
    Write-Output "8. Quit and exit"
    Write-Output ""
    [Int]$Choice = Read-Host "Please select an option" 


    # Making switch to perform action on selection
    Switch ($Choice) {
            1 {Write-Host -ForegroundColor Green  "Showing Extended Info"}
            2 {Write-Host -ForegroundColor Green "Mapping Drive...";
               New-PSDrive -Name K -Root \\$Hostname\C$ -PSProvider FileSystem -Persist -Credential $Creds -ErrorAction SilentlyContinue | Out-Null;
               Start-Process explorer.exe k:;
               $Choice = 0;
               continue}
            3 {Write-Output "Starting Regedit... ($Hostname clipped)";
               $Hostname | clip.exe
               $Choice = 0
               continue
               }
            4 {Write-Host -ForegroundColor Green "Starting Eventvwr...";
               EventVwr $Hostname;
               $Choice = 0;
               continue}
            5 {Write-Host -ForegroundColor Green "Starting MSRA...";
               msra.exe /offerra $Hostname;
               $Choice = 0;
               continue}
            6 {Write-Host -ForegroundColor Green "List Processes"}
            7 {$Hostname = $null;
               $Hostname = Read-Host "Enter new Hostname"
               $Choice = 0
               }
            8 {exit}
            default {Write-Host -ForegroundColor Red  "Please enter a valid option"
               $Choice = 0;
               continue}
    }
    # Clear + print headers again
    Clear-Host
    $Header
    $Information   
}


<# Brol
explorer.exe /?
$Yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Yes"

$No = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "No"

$Options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$Result = $host.UI.PromptForChoice($Title, $Message, $Options, 0) 

switch ($Result)
    {
        0 {"Yes."}
        1 {"No."}
    }

#> 


Met vriendelijke groeten
Bien à vous
Kind regards

Vermeiren Arno
I&O Distributed Systems & Service Delivery - 1JB3E
Tel. +32 (0)2 664 75 66/ Internal number: 47566

AG Insurance
arno.vermeiren@aginsurance.be

Deze e-mail, met inbegrip van elk bijgevoegd document, is vertrouwelijk. Indien u niet de geadresseerde bent, is het openbaar maken, kopiëren of gebruik maken ervan verboden. Indien u dit bericht verkeerdelijk hebt ontvangen, gelieve het te vernietigen en de afzender onmiddellijk te verwittigen. De veiligheid en juistheid van email-berichten kunnen niet gewaarborgd worden, aangezien de informatie kan onderschept of gesaboteerd worden, verloren gaan of virussen kan bevatten. De afzender wijst bijgevolg elke aansprakelijkheid af in dergelijke gevallen. Indien een controle zich opdringt, gelieve een papieren kopie te vragen.
 
Ce message électronique, y compris tout document joint, est confidentiel. Si vous n'êtes pas le destinataire de ce message, toute divulgation, copie ou utilisation en est interdite. Si vous avez reçu ce message par erreur, veuillez le détruire et en informer immédiatement l'expéditeur. La sécurité et l'exactitude des transmissions de messages électroniques ne peuvent être garanties étant donné que les informations peuvent être interceptées, altérées, perdues ou infectées par des virus; l'expéditeur décline dès lors toute responsabilité en pareils cas. Si une vérification s'impose, veuillez demander une copie papier.
 
This email and any attached files are confidential and may be legally privileged. If you are not the addressee, any disclosure, reproduction, copying, distribution, or other dissemination or use of this communication is strictly prohibited. If you have received this transmission in error please notify the sender immediately and then delete this email. Email transmission cannot be guaranteed to be secure or error free as information could be intercepted, corrupted, lost, destroyed, arrive late or incomplete, or contain viruses. The sender therefore does not accept liability for any errors or omissions in the contents of this message which arise as a result of email transmission. If verification is required please request a hard copy version.
