# Starting miner
$MainDir = "C:\Users\Nokkes\Desktop\Mining"
cd $MainDir\EthMine\bin
#Start-Process "EthStartMineGPU_Dwarfpool.bat"

#Load Functions
cd "C:\Scripts\PS\"
. .\Init.ps1

$Csv = "C:\Users\Nokkes\Desktop\Mining\Poloniex.csv"
while ($true)
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $Json = FromUriToJson -uri https://poloniex.com/public?command=returnTicker

    <#
    $lineUSDT_BTC = $Json.USDT_BTC.Values -join ";"
    $line = $lineUSDT_BTC + ";" + $(UniversalDate)
    $line | Out-File -LiteralPath $Csv -Append -Encoding UTF8
    #>
}

[System.Windows.Documents.Serialization.SerializerWriter]