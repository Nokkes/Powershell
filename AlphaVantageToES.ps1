param(
    [Int32]$interval
)
$interval = 5
$indexSettings = @"
{
    "settings": {
        "index": {
            "number_of_replicas": "0"
        }
    },
    "mappings": {
        "properties": {
            "@timestamp": {
                "type": "date",
                "format": "yyyy-MM-dd HH:mm:ss"
            },
            "symbol": {
                "type": "text",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }
            },
			"high": {
                "type": "float",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }
            },
            "open": {
                "type": "float",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }
            },
            "low": {
                "type": "float",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }
            },
            "close": {
                "type": "float",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }

            },
            "volume": {
                "type": "integer",
                "fields": {
                    "keyword": {
                        "type": "keyword",
                        "ignore_above": 256
                    }
                }
            }
        }
    }
}
"@
$shardDate = Get-Date -f "yyyy.MM.dd"
$symbols = "MSFT","AMD","NVDA","CRBP","LMT","PANW"
$ErrorActionPreference = "SilentlyContinue"

try {
    Invoke-WebRequest -Uri "http://192.168.0.189:9200/stock-$($shardDate)/" -Body $($indexSettings) -Method PUT -ContentType "application/json" -Verbose
}

catch{
    $Error | % {
        Write-Host -f Yellow "$($_.Exception.Message)"
        $errMsgText = $_.ErrorDetails.Message
        $errMsgJson = $errMsgText | ConvertFrom-Json
        $errMsg = $errMsgJson.error.reason
        Write-Host -f Yellow $errMsg
    }
    $Error.Clear()
}


$symbols | % {
    $wr = Invoke-WebRequest -Uri "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=$_&interval=$($interval)min&apikey=1PVZCB3K3UE8NKG2"
    $json = ConvertFrom-Json -InputObject $wr.content
    $jsonIdx = "Time Series ($($interval)min)"
    $properties = $json.$jsonIdx |
        Get-Member -MemberType Properties |
        Select-Object -ExpandProperty Name

    foreach ($element in $properties) {
        $ht = @{}
        $ht.Add("@timestamp",$element)
        $ht.Add("symbol",$json.'Meta Data'.'2. Symbol')
        foreach ($stock in $json.$jsonIdx.$element) {
            $ht.Add("open", $stock.'1. open')
            $ht.Add("high", $stock.'2. high')
            $ht.Add("low", $stock.'3. low')
            $ht.Add("close", $stock.'4. close')
            $ht.Add("volume", $stock.'5. volume')
            $body = $ht | ConvertTo-Json
            Write-Host "Writing $($body)"
            $res = Invoke-WebRequest -Uri "http://192.168.0.189:9200/stock-$($shardDate)/_doc" -Body $body -Method POST -ContentType "application/json" -Verbose
            $res
        }
        $ht.Clear()
    }
    
}