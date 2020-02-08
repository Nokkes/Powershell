$a = Get-ChildItem -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall
foreach ($key in $a) 
{
    Get-ItemPropertyValue -Path $_ -Name DisplayName
}