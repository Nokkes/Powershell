$RootFolderPath = "C:\Users\Nokkes\Downloads"

$VideoFiles = (gci $RootFolderPath -Recurse) | ? {$_.Extension -eq ".mkv" -or 
                                                  $_.Extension -eq ".avi" -or 
                                                  $_.Extension -eq ".srt" -or 
                                                  $_.Extension -eq ".mp4"}

$VideoFiles | % {Move-Item -LiteralPath $_.FullName -Destination $RootFolderPath -Verbose}