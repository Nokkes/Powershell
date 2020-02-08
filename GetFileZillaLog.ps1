$DateToday = Get-Date
$DateMin7Days = $DateToday.AddDays(-6)
$DateMin6Days = $DateToday.AddDays(-5)
$DateMin5Days = $DateToday.AddDays(-4)
$DateMin4Days = $DateToday.AddDays(-3)
$DateMin3Days = $DateToday.AddDays(-2)
$DateMin2Days = $DateToday.AddDays(-1)
$DateMin1Days = $DateToday.AddDays(0)

$arrDate = $DateMin6Days, $DateMin5Days, $DateMin4Days, $DateMin3Days, $DateMin2Days, $DateMin1Days
$arrLogFiles = $arrDate | % { "fzs-" + $_.ToString('yyyy-MM-dd') + ".log"}



$arrLogFiles | % { Write-Host "$_";
                   Write-Host "-------------------";
                   cat "C:\Program Files (x86)\FileZilla Server\Logs\$_" } | 
               ? { $_ -like "*Connected on port 49692, sending welcome message..." -or 
                   $_ -like "*530 Login or password incorrect!" -or 
                   $_ -like "*230 Logged on" -or
                   $_ -like "*150 Opening data channel for file download from server of*" -or
                   $_ -like "*226 Successfully transferred*"
                   }
pause