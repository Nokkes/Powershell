. ..\..\Scripts\PS\Init.ps1
$ErrorActionPreference = "SilentlyContinue"

$Site = "https://www.google.be/?search=&start=0&hl=EN&gl=US&filter=0&complete=0"
Winfo "Site = $Site"

StartIE -URL $Site
WaitIE
$IEwindow.BringToFront()
#Start-Job -InitializationScript $OpenVpn -ScriptBlock {OpenVPN -ConfigFile FR}

$obj = $IEwindow.TextField($Wfind::ByName("q"))
$obj.TypeText("Arno Vermeiren")
WaitIE


$obj2 = $IEwindow.Button($Wfind::ByValue("Google Search"))
$obj2.Focus()
WaitIE
$obj2.Click()
WaitIE
$IEwindow.WaitForComplete()

$IEwindow.Link($Wfind::ByText("2"))