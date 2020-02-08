$path = "$env:windir\logs\cbs\"
 
$space = Get-ChildItem -Path $path -Filter cbspersist*.cab -File |
  Measure-Object -Property Length -Sum |
  Select-Object -Property Count, Sum
 
'{0} backed up log files eat up {0:n1} MB' -f $space.Count, ($space.Sum/1MB) 
$space | Remove-Item