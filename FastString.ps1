# SLOW
$log = ''
Measure-Command {
  for ($x = 1; $x -lt 10000; $x += 1) 
  {
    $log += "I’ve just logged step $x"
  }
}
 
 
# FAST 
$log = [System.Text.StringBuilder]''
 
Measure-Command {
  for ($x = 1; $x -lt 10000; $x += 1) 
  {
    $log.AppendLine("I’ve just logged step $x")
  }
}
$log.ToString() 
