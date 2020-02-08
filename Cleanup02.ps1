#requires -Version 3 
#Requires -RunAsAdministrator 
# must run with admin privileges! 
 
# look at temp files older than 3 months 
$cutoff = (Get-Date).AddMonths(-3)
 
# use an ordered hash table to store logging info 
$sizes = [Ordered]@{}
 
# find all files in both temp folders recursively 
Get-ChildItem "$env:windir\temp", $env:temp -Recurse -Force -File |
# calculate total size before cleanup 
ForEach-Object { 
  $sizes['TotalSize'] += $_.Length 
  $_
} |
# take only outdated files 
Where-Object { $_.LastWriteTime -lt $cutoff } |
# try to delete. Add retrieved file size only 
# if the file could be deleted 
ForEach-Object {
  try
  { 
    $fileSize = $_.Length
    # ATTENTION: REMOVE -WHATIF AT OWN RISK
    # WILL DELETE FILES AND RETRIEVE STORAGE SPACE
    # ONLY AFTER YOU REMOVED -WHATIF
    Remove-Item -Path $_.FullName -ErrorAction SilentlyContinue -WhatIf
    $sizes['Retrieved'] += $fileSize
  }
  catch {}
}
 
 
# turn bytes into MB 
$Sizes['TotalSizeMB'] = [Math]::Round(($Sizes['TotalSize']/1MB), 1)
$Sizes['RetrievedMB'] = [Math]::Round(($Sizes['Retrieved']/1MB), 1)
 
New-Object -TypeName PSObject -Property $sizes 
