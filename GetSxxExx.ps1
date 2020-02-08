$dir = Get-ChildItem "F:\VIDZ\TVSHOWZ\" -Recurse -Filter "*S??E??*"
$result = $dir | ? {$_.PSIsContainer -eq "False" -and
          $_.Extension -eq ".mkv" -or
          $_.Extension -eq ".mp4" -or
          $_.Extension -eq ".wmv" -or
          $_.Extension -eq ".avi"
            }

$result -notmatch 'S\d\dE\d\d' | select fullname | Out-GridView -Verbose