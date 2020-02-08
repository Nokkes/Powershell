# look at temp files older than 3 months 
$cutoff = (Get-Date).AddMonths(-3)
 
$space = Get-ChildItem "$env:temp" -Recurse -Force |
  Where-Object { $_.LastWriteTime -lt $cutoff } |
  Measure-Object -Property Length -Sum |
  Select-Object -ExpandProperty Sum
 
'Taken space: {0:n1} MB' -f ($space/1MB) 
