. ..\..\Scripts\PS\Init.ps1
$ErrorActionPreference = "SilentlyContinue"

$Site = "www.opodo.fr"
Winfo "Site = $Site"

StartIE -URL $Site -ToFront $false
WaitIE

#Start-Job -InitializationScript $OpenVpn -ScriptBlock {OpenVPN -ConfigFile FR}

$From = SearchIEByClass -String 'od-airportselector-input airportselector_input'
$From.Exists
$From.DoubleClick()
$From.Highlight()
$From.Focus()
$From.TypeText("Brussel")
$From.
$from.TypeText('Brussel')