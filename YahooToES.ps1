# param(
#     [Int32]$interval
# )
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
                "format": "yyyy-MM-dd HH:mm:ss||yyyy-MM-dd||epoch_millis"
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
Invoke-WebRequest -Uri "http://192.168.0.189:9200/stock-$($shardDate)/" -Body $($indexSettings) -Method PUT -ContentType "application/json" -Verbose


# catch{
#     $Error | % {
#         Write-Host -f Yellow "$($_.Exception.Message)"
#         $errMsgText = $_.ErrorDetails.Message
#         $errMsgJson = $errMsgText | ConvertFrom-Json
#         $errMsg = $errMsgJson.error.reason
#         Write-Host -f Yellow $errMsg
#     }
#     $Error.Clear()
# }

$symbols = "ABI.BR","MSFT"
$wr = Invoke-WebRequest -Uri "https://query1.finance.yahoo.com/v7/finance/quote?symbols=$($symbols -join ",")"
$json = ConvertFrom-Json -InputObject $wr.content
$jsonResult = $json.quoteResponse.result
$jsonCount = $jsonResult.Count

for ($i = 0; $i -lt $jsonCount; $i++) {
    $properties = $jsonResult[$i] | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
    $body = $jsonResult[$i] | ConvertTo-Json
    Invoke-WebRequest -Uri "http://192.168.0.189:9200/stock-$($shardDate)/_doc" -Body $body -Method POST -ContentType "application/json" -Verbose
}