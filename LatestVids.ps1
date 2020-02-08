$Folder = "F:\VIDZ\TVSHOWZ"
$Date =  [datetime]"2019/03/15 00:00"
$Files = gci -Path $Folder -Recurse -Filter "*.mkv" -Force -ErrorAction SilentlyContinue | ? {$_.CreationTime -gt $Date}
$Files