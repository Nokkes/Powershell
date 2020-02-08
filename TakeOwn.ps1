$FromFolder = "C:\temp"
$TargetFolder = "F:\PICZ\"
# Folder to get the Acl from
$Acl = Get-Acl $FromFolder
$Acl.SetOwner([System.Security.Principal.NTAccount]"DESKTOPNOKKES\Nokkes")
(gci F:\PICZ\4chan -Recurse -Directory -ErrorAction SilentlyContinue).FullName | % {Set-Acl -LiteralPath $_ $Acl -Verbose}