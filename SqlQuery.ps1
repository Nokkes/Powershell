$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = "Data Source=RDUAFSSP;Initial Catalog=RDUAFSS_P000;User Id=%UserName%;Password=%Password%;"
$conn.ConnectionString = "Data Source=RDUAFSSP;Initial Catalog=RDUAFSS_P000;Integrated Security=SSPI;"

$conn.open()
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.connection = $conn 

foreach ($quota in $quotas)
{# Update DB
$path = $quota.path
$Size = [math]::Round($quota.Size/1MB,0)
$usage = [math]::Round($quota.Usage/1MB,0)
$cmd.commandtext = "Insert into QuotaR2(Server,Spacealloc,SpaceLimitMB,SizeUsedMB) Values('{0}','{1}','{2}','{3}')" -f $server, $path, $Size, $usage
           # write-host $cmd.commandtext 
            $rowsupdated = $cmd.executenonquery()       #| Out-Null
            if (-not($rowsupdated -eq 1)) {
            write-host "DB update failed " 
            write-host            $cmd.commandtext 
            }
} 
