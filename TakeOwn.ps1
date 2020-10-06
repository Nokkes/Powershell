$FromFolder = "C:\temp"
$TargetFolder = "D:\DOCZ\"
# Folder to get the Acl from
$Acl = Get-Acl $FromFolder
$Acl.SetOwner([System.Security.Principal.NTAccount]"DESKTOPNOKKES\Administrator")
(gci 'D:\DOCZ\Work' -Recurse -Directory -ErrorAction SilentlyContinue).FullName | % {Set-Acl -LiteralPath $_ $Acl -Verbose}
(gci "D:\DOCZ\Work" -Recurse -ErrorAction SilentlyContinue).FullName | % {Set-Acl -LiteralPath $_ $Acl -Verbose}