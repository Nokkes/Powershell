$MyHome = "C:\Scripts\PS"
sl $MyHome
. .\Init.ps1


while ($true) {
    Start-Sleep 15
    UniversalDate
    $Connection = Get-SmbSession
    $Files = Get-SmbOpenFile
    $Info = $UniversalDate + ";" + $Files.Path + ";" + $Files.ClientComputerName + ";" + $Files.ClientUserName + ";" + $Connection.NumOpens
    if ($Info -notlike '*;;;;') {
        Add-Content $MyHome\log\SmbLog.log "$Info"
    }
}