$source = "DataGrabberSvc"
[System.Diagnostics.EventLog]::SourceExists($source) 
[System.Diagnostics.EventLog]::DeleteEventSource($source)