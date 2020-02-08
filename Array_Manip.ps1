$ar = @()
 
# SLOW 
Measure-Command {
  for ($x = 1; $x -lt 10000; $x += 1) 
  {
    $ar += $x
  }
}
 
 
# FAST 
[System.Collections.ArrayList]$ar = @()
Measure-Command {
  for ($x = 1; $x -lt 10000; $x += 1) 
  {
    $null = $ar.Add($x)
  }
} 
