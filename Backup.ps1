# Parse functions and assemblies
cd $env:SCRIPTZ\PS
. .\Init.ps1
$assEventlog = [System.Diagnostics.EventLog]
$ErrorActionPreference = "SilentlyContinue"
# Logpaths
$LogpathRoot = "C:\BackupLogs"
$LogpathScripts = "C:\BackupLogs\Backup_Scripts_$(TechDate).log"
$LogpathDocuments = "C:\BackupLogs\Backup_Documents_$(TechDate).log" 
$LogpathMusic = "C:\BackupLogs\Backup_Music_$(TechDate).log"
$LogpathPictures = "C:\BackupLogs\Backup_Pictures_$(TechDate).log"
$LogpathCrypto = "C:\BackupLogs\Backup_Crypto_$(TechDate).log"
$LogPathKeepass = "C:\BackupLogs\Backup_Keepass_$(TechDate).log"
$LogPathMoney = "C:\BackupLogs\Backup_Money_$(TechDate).log"
# From source directories
$ScriptsFrom = "E:\SCRIPTZ"

$DocsFrom = "F:\DOCZ"
$MusicFrom = "F:\MUZIC"
$PicsFrom = "F:\PICZ"
$Movies = "F:\VIDZ\MOVIEZ"
$TvShows = "F:\VIDZ\TVSHOWZ"
$MoneyFrom = "E:\MONEYZ"
$Crypto = "C:\keystore"
$Keepass = "C:\KeePass\"


# To target directories
$To1 = "F:\DOCZ\Scripts"
$To2 = "G:\DOCZ"
$To3 = "E:\MUZIC"
$To4 = "E:\PICZ"
$To5 = "E:\VIDZ"
$To6 = "E:\MONEYZ"
$To7 = "G:\SCRIPTZ"
$To8 = "D:\MONEYZ"

#Check folder on C
If ((Test-Path -LiteralPath $LogpathRoot) -eq $false) {
    New-Item -ItemType Directory -Force -Path $LogpathRoot
    Wwrite "BackupFolder created!" -newline $true
}
else {
    Wwrite "BackupFolder found!" -newline $true
}

#Check if Source "Backup" exists in eventvwr
If ($assEventlog::Exists("Backup") -eq $false) {
    $assEventlog::CreateEventSource("Backup","Application")
    Wwrite "Event Source created!" -newline $true
}
else {
    Wwrite "EventSource Backup Found!" -newline $true
}

# Backup files, copying recursively

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($ScriptsFrom) started.`nConsult <$($LogPathKeepass)> for more information." -Category 1
Copy-FileWithRobocopy -Source $ScriptsFrom -Destination $To7 -Logpath $LogpathScripts -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($MoneyFrom) started.`nConsult <$($LogPathKeepass)> for more information." -Category 1
Copy-FileWithRobocopy -Source $MoneyFrom -Destination $To8 -Logpath $LogPathMoney -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($DocsFrom) started.`nConsult <$($LogpathDocuments)> for more information." -Category 1
Copy-FileWithRobocopy -Source $DocsFrom -Destination $To2 -Logpath $LogpathDocuments -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($MusicFrom) started.`nConsult <$($LogpathMusic)> for more information." -Category 1
Copy-FileWithRobocopy -Source $MusicFrom -Destination $To3 -Logpath $LogpathMusic -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($PicsFrom) started.`nConsult <$($LogpathPictures)> for more information." -Category 1
Copy-FileWithRobocopy -Source $PicsFrom -Destination $To4 -Logpath $LogpathPictures -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($Crypto) started.`nConsult <$($LogpathCrypto)> for more information." -Category 1
Copy-FileWithRobocopy -Source $Crypto -Destination $To6 -Logpath $LogpathCrypto -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($Keepass) started.`nConsult <$($LogPathKeepass)> for more information." -Category 1
Copy-FileWithRobocopy -Source $Keepass -Destination $To6 -Logpath $LogPathKeepass -RetryCount 2

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Backup of $($ScriptsFrom) started.`nConsult <$($LogPathKeepass)> for more information." -Category 1
Copy-FileWithRobocopy -Source $ScriptsFrom -Destination $To7 -Logpath $LogpathScripts -RetryCount 2
# Dump Movies and TV Shows to CSV
Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "Writing Video Files (Movies & TV Shows) to file.`nConsult <$($To5)> for the CSV's." -Category 1
gci -Recurse -Path $Movies -File | Select-Object Name, FullName, CreationTime, Extension | Export-Csv -LiteralPath "$To5\Movies_$(TechDate).csv"
gci -Recurse -Path $TvShows -File | Select-Object Name, FullName, CreationTime, Extension | Export-Csv -LiteralPath "$To5\TvShows_$(TechDate).csv"

Write-EventLog -LogName Application -Source "Backup" -EventId 666 -Message "All Backups completed succesfully."

pause