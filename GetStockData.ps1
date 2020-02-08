$resp = Invoke-WebRequest -Uri "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=5min&apikey=1PVZCB3K3UE8NKG2" -Method Get
$json = [Microsoft.PowerShell.Commands.JsonObject]::ConvertFromJson($resp.Content,[ref]$null)
$json = New-Object -ComObject 
# call.Get("https://www.alphavantage.co", string.Format("/query?function=TIME_SERIES_INTRADAY&symbol={0}&interval=5min&apikey=1PVZCB3K3UE8NKG2", symbol));
$es = Invoke-WebRequest -Uri http://ubuntu:9200/stocks/stock/1 -Method Put -Body $json
$tst = Invoke-WebRequest -Uri http://ubuntu:9200/stocks/_search -Method Get

$strGet = @'
{
    "query": {
        "query_string": {
            "query": "*"
        }
    }
}
'@
$headGet = @{"Authorization"="Basic $auth"}
$tst = Invoke-WebRequest -Uri http://ubuntu:9200/_cat/indices?v -Method get -ContentType "application/json"
$tst.Content