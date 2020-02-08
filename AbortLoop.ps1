try 
{
  do
  {
    Write-Host '.' -NoNewline
    Start-Sleep -Milliseconds 800
  } while ($true)
}
finally 
{
  $sapi = New-Object -ComObject Sapi.SpVoice
  $sapi.Speak('Hey, you aborted me dammit!')
} 
$PSVersionTable