### Setting standard vars & location###
cd "C:\Scripts\PS\GetIp"
$ErrorActionPreference = 'SilentlyContinue'
Start-Transcript -Path C:\Scripts\PS\GetIp\GetIP.log -Append
. ..\Init.ps1
TechDate
UniversalDate
$HomeDir = "C:\Scripts\PS\GetIp\"
$Creds = Import-PSCredential -Path "C:\Scripts\PS\Creds\credentials.enc.xml"

# Getting Last IP from FZ XML
$ServerConfig = "C:\Program Files (x86)\FileZilla Server\FileZilla Server.xml"
$ServerConfigContent = Get-Content $ServerConfig
$ServerXML = [XML]$ServerConfigContent
$ServerIP = $ServerXML.ChildNodes.Settings.ChildNodes | ? {$_.Name -eq "Custom PASV IP"}


# Getting IP from IP-api, need #try catch
$IpXml = Invoke-WebRequest -Uri http://ip-api.com/xml -UseBasicParsing
$IP = Select-Xml -Content $IpXml -XPath "//query" | 
    Select-Object -Last 1 -ExpandProperty Node


# If my external IP-api IP differs from the last IP in the Filezilla Server XML
if ($ServerIP.'#text' -ne $IP.'#cdata-section') {
    Wwarn "IPs have changed"
    Winfo "Building XML"
    $ServerXml.ChildNodes.Settings.ChildNodes | 
    ? { $_.name -eq 'Custom PASV IP' } | 
    % { $_.'#text' = $IP.'#cdata-section' }
    
    Winfo "Renaming FileZilla XML File"
    Rename-Item -LiteralPath "C:\Program Files (x86)\FileZilla Server\FileZilla Server.xml" -NewName "C:\Program Files (x86)\FileZilla Server\FileZilla Server$TechDate.xml"
    $ServerXML.Save("C:\Program Files (x86)\FileZilla Server\FileZilla Server.xml")
    
    Winfo "Starting FileZilla"
    Get-Service "FileZilla Server" | Start-Service

    winfo "Sending email"
    $EmailBody1 = "Hi <br></br>The IP of your FTP server has changed on $UniversalDate and is now ftp://nokkes.dynamic-dns.net"
    $EmailBody2 = ":49692"
    $EmailBody3 = "<br></br>--<br></br>Regards"
    $EmailBody = $EmailBody1+$EmailBody2+$EmailBody3
    $messageParameters = @{
    Subject = "Server IP has changed"
    Body = "$EmailBody"
    BodyAsHTML = $true
    From = "arnovermeiren@hotmail.com"
    Bcc = @("<arnovermeiren@hotmail.com>","<jona.simillion@gmail.com>","<sami_03_11_86@hotmail.com>")
    SmtpServer = "smtp-mail.outlook.com"
    Port = 587
    UseSsl = $true
    Verbose = $true
    Debug = $true
    Credential = $Creds}
    Send-MailMessage @messageParameters
}
else {
    Winfo "IPs haven't changed, starting FileZilla"
    Get-Service "FileZilla Server" | Start-Service
}

Start-Sleep 5
Stop-Transcript