. ..\..\Scripts\PS\Init.ps1

$UrlProd = "https://api.sabre.com"
$UrlTest = "https://api.test.sabre.com"

$URL = "/v2/auth/token"


$Token = @{"Authorization" = "Basic VmpFNk5UVjJjbTh6YWpZMk5tOWhiVFl3Y0RwRVJWWkRSVTVVUlZJNlJWaFVPalJDUmxWa1l6bDQ=";
"grant_type"="client_credentials";
"Accept" =  "*/*"
}

$URI = $UrlTest + $URL

$a = Invoke-RestMethod -Uri $URI -Headers $Token -Method Post -ContentType "application/x-www-form-urlencoded"

$a | fl *

# Yahoo
# Client ID (Consumer Key)
#    dj0yJmk9RjByZTc3NWs5eldEJmQ9WVdrOVUzcG1kWEE1TXpnbWNHbzlNQS0tJnM9Y29uc3VtZXJzZWNyZXQmeD0yOA--
# Client Secret (Consumer Secret)
#    93e12e0dbc8b25a0c80452f6fe4d1a03f50f02a8