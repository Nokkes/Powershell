
$Servers = @(
    "192.168.0.247",
    "192.168.0.149"
)
<#
# Construct table containing ServerName + ServiceName
$Services = $Servers | % {[PSCustomObject]@{ServerName = $_; ServiceName = Get-Service -Name "OMS_*" -ComputerName $_ -ErrorAction SilentlyContinue}}

# Stop Service for each entry in table

$Services | % {
    Invoke-Command -ComputerName $_.ServerName -ScriptBlock { net stop $_.ServiceName }   
}

workflow StopServices
{
    foreach -parallel($s in $server_list) { 
        inlinescript { (Get-Service -Name '*Service Name*' -PSComputerName $s) | Where-Object {$_.status -eq "Stopped"} | Set-Service -Status Running
        }
    }
}

StopServices
#>
$creds = Get-Credential

Invoke-Command -ComputerName "WINVMNOKKES" -Credential $creds -ScriptBlock { Get-Service }