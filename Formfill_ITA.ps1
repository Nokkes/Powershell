. ..\..\Scripts\PS\Init.ps1
$ErrorActionPreference = "SilentlyContinue"

$Site = "http://matrix.itasoftware.com/"
Winfo "Site = $Site"

StartIE -URL $Site -ToFront $true
WaitIE

#Start-Job -InitializationScript $OpenVpn -ScriptBlock {OpenVPN -ConfigFile FR}

$From = SearchIEById -String "cityPair-orig-0"
$To = SearchIEById -String "cityPair-dest-0"
$FromDate = SearchIEById -String "cityPair-outDate-0"
$ToDate = SearchIEById -String "cityPair-retDate-0"
$Curreny =  $IEwindow.TextField($Wfind::ByClass('gwt-SuggestBox IR6M2QD-q-a IR6M2QD-a-l'))
$Go = SearchButtonById -String 'searchButton-0'

$From
$From.TypeText("BRU")



$To.TypeText("BCN")
$IEwindow.BringToFront()

$FromDate.TypeText('10/21/2016')
$IEwindow.PressTab()
$ToDate.TypeText('10/23/2016')
$IEwindow.PressTab()
$Curreny.TypeText('EUR')
$IEwindow.PressTab()
$Go.Focus()
$Go.ClickNoWait()

<# For changing the +- Date
$Div = $IEwindow.Div($Wfind::ByClass('IR6M2QD-L-p'))
$t = $Div.Div($Wfind::ByClass('IR6M2QD-a-m IR6M2QD-a-u'))

$l = $t.SelectList($Wfind::ByText('On this day onlyOr day beforeOr day afterPlus/minus 1 dayPlus/minus 2 days'))
$l.SelectByValue("+-1")
#>


