#requires -Version 3 
#Requires -RunAsAdministrator 
# must run with admin privileges! 
 
 
# use an ordered hash table to store logging info 
$sizes = [Ordered]@{}
 
 
 
Get-ChildItem -Path $env:windir\logs\cbs\ -Filter *.cab |
ForEach-Object {
  try
  { 
    $fileSize = $_.Length
    # ATTENTION: REMOVE -WHATIF AT OWN RISK
    # WILL DELETE FILES AND RETRIEVE STORAGE SPACE
    # ONLY AFTER YOU REMOVED -WHATIF
    Remove-Item -Path $_.FullName -ErrorAction SilentlyContinue
    $sizes['Retrieved'] += $fileSize
  }
  catch {}
}
 
# turn bytes into MB 
$Sizes['RetrievedMB'] = [Math]::Round(($Sizes['Retrieved']/1MB), 1)
 
New-Object -TypeName PSObject -Property $sizes 
