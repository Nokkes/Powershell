
$preInit = @{
  Rate = 0
  Volume = 100
}

 
 
Add-Type -AssemblyName System.Speech
$speaker = [System.Speech.Synthesis.SpeechSynthesizer] $preInit
$speaker.SelectVoice('Microsoft Zira Desktop')

$null = $Speaker.Speak(“Go fuck yourself.”)
$speaker.Dispose()