cd "F:\VIDZ\TVSHOWZ\The Office (2001)\Season 02"
$Files = Get-ChildItem -Recurse 
                                     
                                     
$VideoFiles = $Files | ? {$_.Extension -eq ".mp4" -or
                          $_.Extension -eq ".mkv" -or
                          $_.Extension -eq ".avi"}
                                  

$VideoFiles | % {Rename-Item $_ -NewName ($_.Name -replace "S02.","S02E0") -Verbose }