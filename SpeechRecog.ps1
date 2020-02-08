#requires -Version 3 
cd E:\SCRIPTZ\PS\
Import-Module .\SpeechRecog.psm1

 Add-SpeechCommands @{ 
     "Check Inbox" = { Say "Checking for email" }
     "What time is it?" = { Say "It is $(Get-Date -f "h:mm tt")" } 
     "What day is it?"  = { Say $(Get-Date -f "dddd, MMMM dd") } 
     "Processes"  = { 
        $proc = ps | sort ws -desc 
        Say $("$($proc.Count) processes, including $($proc[0].name), which is using " + 
              "$([int]($proc[0].ws/1mb)) megabytes of memory") 
     } 
     
  } -Computer "computer" -Verbose  
  Add-SpeechCommands @{ "Run Notepad" = { & "C:\Programs\DevTools\Notepad++\notepad++.exe" }} -Computer "computer"
getspch

Start-Listening
Clear-SpeechCommands 