##################################################################
#####                        Assemblies                      #####
##################################################################

Add-Type -Path "C:\Scripts\PS\Watin\bin\net40\WatiN.Core.dll" -Verbose
Add-Type -Path "C:\Scripts\PS\Watin\bin\net40\Interop.SHDocVw.dll" -Verbose
Add-Type -TypeDefinition '
using System;
using System.Runtime.InteropServices;
 
namespace Utilities {
   public static class Display
   {
      [DllImport("user32.dll", CharSet = CharSet.Auto)]
      private static extern IntPtr SendMessage(
         IntPtr hWnd,
         UInt32 Msg,
         IntPtr wParam,
         IntPtr lParam
      );
 
      public static void PowerOff ()
      {
         SendMessage(
            (IntPtr)0xffff, // HWND_BROADCAST
            0x0112,         // WM_SYSCOMMAND
            (IntPtr)0xf170, // SC_MONITORPOWER
            (IntPtr)0x0002  // POWER_OFF
         );
      }
   }
}
'

##################################################################
#####                        Functions                       #####
##################################################################

# Robocopy CmdLet
function Copy-FileWithRobocopy
{
  param
  (
    [Parameter(Mandatory)]
    [string]$Source,
    
    [Parameter(Mandatory)]
    [string]$Destination,
    
    [string]$Filter = '*',

    [string]$Logpath = '',
    
    [int]$RetryCount = 0,
    
    [string]$ExcludeDirectory = '',
    
    [switch]$Open,
        
    [switch]$NoRecurse
  )
  
    $Recurse = '/S'
    if ($NoRecurse) 
    {        
        $Recurse = ''
    }
    
    if ($Logpath -ne '')
    {
        robocopy.exe $Source $Destination $Filter /R:$RetryCount $Recurse /XD $ExcludeDirectory /log:$Logpath  /np
    }
    else 
    {
        robocopy.exe $Source $Destination $Filter /R:$RetryCount $Recurse /XD $ExcludeDirectory /np
    }

    if ($Open)
    {
        explorer $Destination
    }
} 


$Wfind = [WatiN.Core.Find]

$OpenVpn = {
    Function OpenVPN {
    [CmdLetBinding()]
        param (
            [Parameter(Mandatory = $true)]
            [string][ValidateSet("CA","FR","SE","US")]$ConfigFile
        )
        switch ($ConfigFile) {
            "CA" {$c = "C:\Program Files\OpenVPN\config\ca.openvpn.frootvpn.ovpn"}
            "FR" {$c = "C:\Program Files\OpenVPN\config\fr.openvpn.frootvpn.ovpn"}
            "SE" {$c = "C:\Program Files\OpenVPN\config\se.openvpn.frootvpn.ovpn"}
            "US" {$c = "C:\Program Files\OpenVPN\config\us.openvpn.frootvpn.ovpn"}
            default {$c = ""}
        }

        & 'C:\Program Files\OpenVPN\bin\openvpn.exe' "$c"
    }
}

function Disable-OneDrive {
    $regkey1 = 'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    $regkey2 = 'Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    Set-ItemProperty -Path $regkey1, $regkey2 -Name System.IsPinnedToNameSpaceTree -Value 0
}
  
function Enable-OneDrive {
    $regkey1 = 'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    $regkey2 = 'Registry::HKEY_CLASSES_ROOT\Wow6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'    
    Set-ItemProperty -Path $regkey1, $regkey2 -Name System.IsPinnedToNameSpaceTree -Value 1
} 

function Export-PSCredential {
        param ( $Credential = (Get-Credential), $Path = "credentials.enc.xml" )
 
        # Look at the object type of the $Credential parameter to determine how to handle it
        switch ( $Credential.GetType().Name ) {
                # It is a credential, so continue
                PSCredential            { continue }
                # It is a string, so use that as the username and prompt for the password
                String                          { $Credential = Get-Credential -credential $Credential }
                # In all other caess, throw an error and exit
                default                         { Throw "You must specify a credential object to export to disk." }
        }
       
        # Create temporary object to be serialized to disk
        $export = "" | Select-Object Username, EncryptedPassword
       
        # Give object a type name which can be identified later
        $export.PSObject.TypeNames.Insert(0,’ExportedPSCredential’)
       
        $export.Username = $Credential.Username
 
        # Encrypt SecureString password using Data Protection API
        # Only the current user account can decrypt this cipher
        $export.EncryptedPassword = $Credential.Password | ConvertFrom-SecureString
 
        # Export using the Export-Clixml cmdlet
        $export | Export-Clixml $Path
        Write-Host -foregroundcolor Green "Credentials saved to: " -noNewLine
 
        # Return FileInfo object referring to saved credentials
        Get-Item $Path
}
 
function Import-PSCredential {
        param ( $Path = "credentials.enc.xml" )
 
        # Import credential file
        $import = Import-Clixml $Path
       
        # Test for valid import
        if ( !$import.UserName -or !$import.EncryptedPassword ) {
                Throw "Input is not a valid ExportedPSCredential object, exiting."
        }
        $Username = $import.Username
       
        # Decrypt the password and store as a SecureString object for safekeeping
        $SecurePass = $import.EncryptedPassword | ConvertTo-SecureString
       
        # Build the new credential object
        $Credential = New-Object System.Management.Automation.PSCredential $Username, $SecurePass
        Write-Output $Credential
}

function MD5Hash {
    param (
    $Location
    )

    $FullPath = Resolve-Path $Location
    $MD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $File = [System.IO.File]::Open($FullPath,[System.IO.Filemode]::Open, [System.IO.FileAccess]::Read)
    [System.BitConverter]::ToString($MD5.ComputeHash($File))
    $File.Dispose()
}

Function StartIE {
    param (
        [string]$URL
    )

    $CIM = Get-CimInstance -ClassName Win32_VideoController
    $VertRes = $CIM.CurrentVerticalResolution
    $HorRes = $CIM.CurrentHorizontalResolution

    $Global:IEwindow = New-Object WatiN.Core.IE
    $IEwindow.ClearCache()
    $IEwindow.ClearCookies()
    $IEwindow.GoTo("$URL")
    $IEwindow.SizeWindow($HorRes,$VertRes)
}

Function WaitIE {
    $IEwindow.WaitForComplete()
}

Function StopIE {
    $IEwindow.Close()
}

Function SearchIEById {
    
    param (
        [CmdLetBinding()]
        [string]$String

    )
    $Global:Id = $IEwindow.TextField($Wfind::ById("$String"))
    if ($Id.Exists -eq $true) {
        return $Id
    }
    else {
        Wwarn "Id not found"
    }
}

Function SearchButtonById {
    
    param (
        [CmdLetBinding()]
        [string]$String

    )
    $Global:Button = $IEwindow.Button($Wfind::ById("$String"))
    if ($Button.Exists -eq $true) {
        return $Button
    }
    else {
        Wwarn "Button not found"
    }
}

Function SearchIEByName {
    
    param (
        [CmdLetBinding()]
        [string]$String

    )
    $Global:Name = $IEwindow.TextField($Wfind::ByName("$String"))
    if ($Name.Exists -eq $true) {
        return $Name
    }
    else {
        Wwarn "Name not found"
    }
}

Function SearchIEByValue {
    
    param (
        [CmdLetBinding()]
        [string]$String

    )
    $Global:Value = $IEwindow.TextField($Wfind::ByValue("$String"))
    if ($Value.Exists -eq $true) {
        return $Value
    }
    else {
        Wwarn "Value not found"
    }
}

Function SearchDivByClass {
    
    param (
        [CmdLetBinding()]
        [string]$String

    )
    $Global:ieClass = $IEwindow.Div($Wfind::ByClass("$String"))
    if ($ieClass.Exists -eq $true) {
        return $ieClass
    }
    else {
        Wwarn "Div not found"
    }
}

Function SearchAndClickButtonId {
        [CmdLetBinding()]
    param (
        [string]$ClickTargetButtonId
    )
    
    $ClickButton = $IEwindow.Button($Wfind::ById("$ClickTargetButtonId"))
    $ClickButton.Click()
}

Function SearchTableId {
    [CmdLetBinding()]
    param (
        [string]$ClickTable
    )
    $Table = $IEwindow.Table($Wfind::ById('sitelookup_dlResultsEach_ctl00_ddCategoriesEach'))
}

Function SearchClassId {

}



Function Count {
    
    param (
        [CmdletBinding()]
        [Parameter(Mandatory = $true)] 
        [int32]$Counter
    )
    ForEach ($Count in (1..$Counter)) {
        Winfo "Pausing: $Count/$Counter secs" -newline $true
        Start-Sleep 1
    } 
}

#progress function
#Write-Progress -Id 1 -Activity $Message -Status "$($Counter - $Count) seconds left" -PercentComplete (($Count / $Counter) * 100)
#Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed

Function Switch-DisplayOff {
    [Utilities.Display]::PowerOff()
}

#==============================================================================
# Date and time Functions
#==============================================================================

Function UniversalDate {
   return Get-Date -f "dd/MM/yyyy HH:mm:ss"
}

Function TechDate {
    return Get-Date -f "yyyy_MM_dd_HH_mm_ss"
}

Function StartWatch {
    $Global:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Winfo "StopWatch started"
}

Function StopWatch {
    $Stopwatch.Stop()
    Winfo "StopWatch took $($Stopwatch.Elapsed.TotalSeconds) seconds to complete"
}

Function Invoke-WithProgressBar {
  param
  (
    #[Parameter(Mandatory)]
    [Parameter(ValueFromPipeline)]
    [ScriptBlock]
    $Task
  )
  
  try
  {
    $ps = [PowerShell]::Create()
    $null = $ps.AddScript($Task)
    $handle = $ps.BeginInvoke()
  
    $i = 0
    while(!$handle.IsCompleted)
    {
      Write-Progress -Activity 'Hang in...' -Status $i -PercentComplete ($i % 100)
      $i++
      Start-Sleep -Milliseconds 300
    }
    Write-Progress -Activity 'Hang in...' -Status $i -Completed
    $ps.EndInvoke($handle)
  }
  finally
  {
    $ps.Stop()
    $ps.Runspace.Close()
    $ps.Dispose()
  } 
} 

Function Wwrite {
    param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
        [string]$txt,
        [bool]$newline
    )
    if ($newline -eq $true) {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host $txt
    }
    else {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host $txt -NoNewline
    }
}

Function Winfo {
    param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
        [string]$txt,
        [bool]$newline
    )
    if ($newline -eq $true) {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Green $txt
    }
    else {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Green $txt -NoNewline
    }
}

Function Wwarn {
    param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
        [string]$txt,
        [bool]$newline
    )
    if ($newline -eq $true) {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Yellow $txt
    }
    else {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Yellow $txt -NoNewline
    }
}

Function Werror {
    param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
        [string]$txt,
        [bool]$newline
    )
    if ($newline -eq $true) {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Red $txt
    }
    else {
        Write-Host "$(UniversalDate)" -NoNewline; Write-Host " :: " -NoNewline; Write-Host -ForegroundColor Red $txt -NoNewline
    }
}
       

#==============================================================================
# Write werrorLog message
#==============================================================================
function werrorLog
       {Param(
              [Parameter(Mandatory=$false,Position=1)] [string]$e)                 
              werror "Exception: $($e.Exception)"
              werror "ID: $($e.Exception.id)"
              werror "Message: $($e.Exception.Message)"
       }

#==============================================================================
# Logging (WIP)
#==============================================================================

Function Start-Logging {
    
    
    param (
    [CmdLetBinding()]
    [bool]$Debugging = $False, # Change to TRUE to output TEST logs.
    [string]$FileName = $TechDate,
    [string]$LogPath
    )
    If ($Debugging -eq $false) {

        Start-Transcript -Path "$LogPath'\'$($env:USERNAME)'\'$FileName.log'"

    } Else {
        Walert "Debugging is set to false"
    }
}

Function Winput { Param([Parameter(ValueFromPipeline=$true)]$Text)

    $Date = UniversalDate
    
    $Input = Read-Host "$Date $Text"

    Write-Output "$Date $($Text): $Input" | Out-File $strLogfile -Append

    Return $Input

}

#==============================================================================
# Check if a process is running and show info
#==============================================================================

Function CheckProcess { 
    Param (
        [Parameter(ValueFromPipeline=$true)]$Text
        )
    if (Get-Process -Name "$Text") {
    $Date = UniversalDate
    }
}


#==============================================================================
# Check if a user is Member Of an AD-Group
#==============================================================================

Function CheckAD {
    Param (
        [CmdletBinding()]
        [string]$UserName,
        [string]$GroupName,
        [switch]$Global:IsMember
        )
    if ((Get-ADPrincipalGroupMembership -Identity "$UserName").Name -contains "$GroupName") {
        $IsMember = $true
    }
    else {
        $IsMember = $false 
    }
} 

#==============================================================================
# Validate if a string  is a valid  IP address
#==============================================================================

Function Validate-IPAddress(){
       
       param(
       [Parameter(Position=0, Mandatory=$false)] [string]$IP)
       
       $IPAddress = $null
       return ([System.Net.IPAddress]::TryParse($IP,[ref]$IPAddress))

}

#==============================================================================
# Convert Decimal IP to Binary IP
#==============================================================================

Function ConvertTo-BinaryIP( [String]$IP ) {

  $IPAddress = [Net.IPAddress]::Parse($IP)

  Return [String]::Join('.', 
    $( $IPAddress.GetAddressBytes() | %{ 
      [Convert]::ToString($_, 2).PadLeft(8, '0') } ))
}

#==============================================================================
# Convert Binary IP to Decimal IP
#==============================================================================

Function ConvertTo-DecimalIP( [String]$IP ) {

  $IPAddress = [Net.IPAddress]::Parse($IP)
  $i = 3; $DecimalIP = 0;
  $IPAddress.GetAddressBytes() | %{
    $DecimalIP += $_ * [Math]::Pow(256, $i); $i-- }

  Return [UInt32]$DecimalIP
}

#==============================================================================
# Convert Binary IP to Decimal IP
#==============================================================================

Function ConvertTo-DottedDecimalIP( [String]$IP ) {

  Switch -RegEx ($IP) {
    "([01]{8}\.){3}[01]{8}" {

      Return [String]::Join('.', $( $IP.Split('.') | %{ 
        [Convert]::ToInt32($_, 2) } ))
    }
    "\d" {

      $IP = [UInt32]$IP
      $DottedIP = $( For ($i = 3; $i -gt -1; $i--) {
        $Remainder = $IP % [Math]::Pow(256, $i)
        ($IP - $Remainder) / [Math]::Pow(256, $i)
        $IP = $Remainder
       } )
       
      Return [String]::Join('.', $DottedIP)
    }
    default {
      Write-Error "Cannot convert this format"
    }
  }
}

function Get-IPLocation([Parameter(Mandatory)]$IPAddress)
{
    Invoke-RestMethod -Method Get -Uri "http://geoip.nekudo.com/api/$IPAddress" |
      Select-Object -ExpandProperty Country -Property City, IP, Location
} 

#==============================================================================
# Return Mask lenght from Mask
#==============================================================================

Function ConvertTo-MaskLength( [String]$Mask ) {

  $IPMask = [Net.IPAddress]::Parse($Mask)
  $Bits = "$( $IPMask.GetAddressBytes() | %{
    [Convert]::ToString($_, 2) } )" -Replace "[\s0]"

  Return $Bits.Length
}

#==============================================================================
# Return Mask from Mask lenght
#==============================================================================

Function ConvertTo-Mask( [Byte]$MaskLength ) {

  Return ConvertTo-DottedDecimalIP ([Convert]::ToUInt32(
    $(("1" * $MaskLength).PadRight(32, "0")), 2))
}

#==============================================================================
# Return Network Address from an IP with its Mask
#==============================================================================

Function Get-NetworkAddress( [String]$IP, [String]$Mask ) {

  Return ConvertTo-DottedDecimalIP $(
    (ConvertTo-DecimalIP $IP) -BAnd 
    (ConvertTo-DecimalIP $Mask))
}

#==============================================================================
# Return Broadcast Address from an IP with its Mask
#==============================================================================

Function Get-BroadcastAddress( [String]$IP, [String]$Mask ) {

  Return ConvertTo-DottedDecimalIP $(
    (ConvertTo-DecimalIP $IP) -BOr 
    ((-BNot (ConvertTo-DecimalIP $Mask)) -BAnd [UInt32]::MaxValue))
}

#==============================================================================
# Return Network info from IP with its Mask
#==============================================================================

Function Get-NetworkInfo( [String]$IP, [String]$Mask ) {
  If ($IP.Contains("/"))
  {
    $Temp = $IP.Split("/")
    $IP = $Temp[0]
    $Mask = $Temp[1]
  }

  If (!$Mask.Contains("."))
  {
    $Mask = ConvertTo-Mask $Mask
  }

  $DecimalIP = ConvertTo-DecimalIP $IP
  $DecimalMask = ConvertTo-DecimalIP $Mask
  
  $Network = $DecimalIP -BAnd $DecimalMask
  $Broadcast = $DecimalIP -BOr 
    ((-BNot $DecimalMask) -BAnd [UInt32]::MaxValue)
  $NetworkAddress = ConvertTo-DottedDecimalIP $Network
  $RangeStart = ConvertTo-DottedDecimalIP ($Network + 1)
  $RangeEnd = ConvertTo-DottedDecimalIP ($Broadcast - 1)
  $BroadcastAddress = ConvertTo-DottedDecimalIP $Broadcast
  $MaskLength = ConvertTo-MaskLength $Mask
  
  $BinaryIP = ConvertTo-BinaryIP $IP; $Private = $False
  Switch -RegEx ($BinaryIP)
  {
    "^1111"  { $Class = "E"; $SubnetBitMap = "1111" }
    "^1110"  { $Class = "D"; $SubnetBitMap = "1110" }
    "^110"   { 
      $Class = "C"
      If ($BinaryIP -Match "^11000000.10101000") { $Private = $True } }
    "^10"    { 
      $Class = "B"
      If ($BinaryIP -Match "^10101100.0001") { $Private = $True } }
    "^0"     { 
      $Class = "A" 
      If ($BinaryIP -Match "^00001010") { $Private = $True } }
   }   
   
  $NetInfo = New-Object Object
  Add-Member NoteProperty "Network" -Input $NetInfo -Value $NetworkAddress
  Add-Member NoteProperty "Broadcast" -Input $NetInfo -Value $BroadcastAddress
  Add-Member NoteProperty "Range" -Input $NetInfo `
    -Value "$RangeStart - $RangeEnd"
  Add-Member NoteProperty "Mask" -Input $NetInfo -Value $Mask
  Add-Member NoteProperty "MaskLength" -Input $NetInfo -Value $MaskLength
  Add-Member NoteProperty "Hosts" -Input $NetInfo `
    -Value $($Broadcast - $Network - 1)
  Add-Member NoteProperty "Class" -Input $NetInfo -Value $Class
  Add-Member NoteProperty "IsPrivate" -Input $NetInfo -Value $Private
  
  Return $NetInfo
}

#==============================================================================
# Validate if a string  is a valid  IP address
#==============================================================================

Function Get-NetworkRange( [String]$IP, [String]$Mask ) {
  If ($IP.Contains("/"))
  {
    $Temp = $IP.Split("/")
    $IP = $Temp[0]
    $Mask = $Temp[1]
  }

  If (!$Mask.Contains("."))
  {
    $Mask = ConvertTo-Mask $Mask
  }

  $DecimalIP = ConvertTo-DecimalIP $IP
  $DecimalMask = ConvertTo-DecimalIP $Mask
  
  $Network = $DecimalIP -BAnd $DecimalMask
  $Broadcast = $DecimalIP -BOr ((-BNot $DecimalMask) -BAnd [UInt32]::MaxValue)

  For ($i = $($Network + 1); $i -lt $Broadcast; $i++) {
    ConvertTo-DottedDecimalIP $i
  }
}


#==============================================================================
# Set Binding Order
#==============================================================================

function Set-BindingOrder
{
       param([Parameter(Mandatory = $true)] $FirstInterfaceName)
       
    $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}
       
       $process = start-process nvspbind.exe -ArgumentList "/++ ""$FirstInterfaceName"" ms_tcpip" -PassThru -Wait
       switch($process.ExitCode)
       {
              0 {$rt.status = $true}
              14 {$rt.status = $true;$rt.alreadydone=$true}
              default {$rt.status = $false}
       }
       return $rt
}      

#==============================================================================
# Rename network interface based on default gateway
#==============================================================================

function Rename-Interface
{
       param([Parameter(Mandatory = $true)] $IfDefGateway,
       [Parameter(Mandatory = $true)] $IfNewName
       )

    $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}

       $if_config = Get-NetIPConfiguration|where-object{$_.IPv4DefaultGateway.NextHop -eq $IfDefGateway}
       if($if_config)
       {
              $if_nic = Get-NetAdapter|where-object{$_.ifIndex -eq $($if_config.InterfaceIndex)}
              if($if_nic.Name -eq $IfNewName)
              {
                     $rt.status = $true
                     $rt.alreadydone=$true
              }
              else
              {
                     Rename-NetAdapter -InputObject $if_nic -NewName $IfNewName
                     $rt.status = $true
              }      
       }
       else
       {
              $rt.status = $false
              $rt.errormsg = "Cannot find NIC in subnet $IfDefGateway"
       }
       
       return $rt
}

#==============================================================================
# Check if specfied interface has specified IP address
#==============================================================================

function Check-IP
{
       param([Parameter(Mandatory = $true)] $IfName,
       [Parameter(Mandatory = $true)][ValidateScript({$_ -match [IPAddress]$_ })] [string] $IfIP
       )

    $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}

       $if_config = Get-NetIPConfiguration|where-object{$_.InterfaceAlias -eq $IfName}
       if($if_config)
       {
              $curIP = $if_config.IPv4Address.IPAddress
              if($curIP -eq $IfIP)
              {
                     if($if_config.NetIPv4Interface.Dhcp -eq "Enabled")
                     {
                           $rt.status = $false
                           $rt.errormsg = "Server has the foreseen IP $IfIP but received via DHCP. Please use static IP."
                     }
                     else
                     {
                           $rt.status = $true
                     }      
              }
              else
              {
                     $rt.status = $false
                     $rt.errormsg = "Server does not have foreseen IP $IfIP (current IP: $curIP)"
              }
       }
       else
       {
              $rt.status = $false
              $rt.errormsg = "Cannot find NIC with name $IfName"
       }
       
       return $rt
}


#
# Examples
#
#
#Function RunExample( $Expression )
#{
#  Write-Host "Running: $Expression" -ForegroundColor Green
#  Invoke-Expression $Expression
#}
#
#RunExample "ConvertTo-BinaryIP 192.168.1.1"
#RunExample "ConvertTo-DecimalIP 192.168.1.1"
#RunExample "ConvertTo-DottedDecimalIP 11000000.10101000.00000001.00000001"
#RunExample "ConvertTo-DottedDecimalIP 3232235777"
#RunExample "ConvertTo-MaskLength 255.255.128.0"
#RunExample "ConvertTo-Mask 17"
#
#RunExample "Get-NetworkAddress 192.168.1.1 255.255.255.0"
#Get-NetworkAddress 10.14.116.75 255.255.255.0
#RunExample "Get-BroadcastAddress 10.14.116.75 255.255.255.0"
#Get-BroadcastAddress 10.14.116.75 255.255.255.0
#RunExample "Get-NetworkInfo 229.168.1.1 255.255.248.0"
#RunExample "Get-NetworkInfo 172.16.1.243 18"
#RunExample "Get-NetworkInfo 10.0.0.3/14"
#
#RunExample "Get-NetworkRange 192.168.1.5 255.255.255.248"
#RunExample "Get-NetworkRange 172.18.0.23 30"
#RunExample "Get-NetworkRange 172.18.0.23/29"



#==============================================================================
# NAME: functions-General-v1
# DESCRIPTION: Contains basic functions 
# VERSION: 1.0
# AUTHOR: Berckmans Dieter
# CREATION DATE : 17/03/2014
# OWNER: Bercmans Dieter,Van Eyll Axel
# MODIFCATION DATE: 
#      V1 initial release
#==============================================================================


$global:WriteIndentation =""                    
$global:correlationID =""         
#$global:CurrentADdomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain() 

$global:ADS_GROUP_TYPE_DOMAIN_LOCAL_GROUP  = 4     
$global:ADS_GROUP_TYPE_GLOBAL_GROUP      = 2      
$global:ADS_GROUP_TYPE_LOCAL_GROUP      = 4     
$global:ADS_GROUP_TYPE_UNIVERSAL_GROUP    = 8       
$global:ADS_GROUP_TYPE_SECURITY_ENABLED    = -2147483648      
$global:ADS_GLOBALGROUP=($ADS_GROUP_TYPE_GLOBAL_GROUP  -bor $ADS_GROUP_TYPE_SECURITY_ENABLED)
$global:ADS_LOCALGROUP=($ADS_GROUP_TYPE_LOCAL_GROUP  -bor $ADS_GROUP_TYPE_SECURITY_ENABLED)




#==============================================================================
# Except for the base function winfo,wwarn,... preferable no console output
# in any of these functions
#==============================================================================

#==============================================================================
# Configures the output window to something more standard and usefull
#==============================================================================
function Set-DefaultHostConfig
       {Param([Parameter(Mandatory=$false,Position=1)][string]$title="POWERSHELL")
                     
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
        If($title -ne $null -and $title -ne "")
             {$Host.UI.RawUI.WindowTitle =$title}
              
       $host.ui.RawUI.ForegroundColor = "gray" 
        $host.ui.RawUI.BackgroundColor = "black" 
        $buffer = $host.ui.RawUI.BufferSize 
        $buffer.width = 200 
        $buffer.height = 3000 
        $host.UI.RawUI.Set_BufferSize($buffer) 
        $maxWS = $host.UI.RawUI.Get_MaxWindowSize() 
        $ws = $host.ui.RawUI.WindowSize 
        IF($maxws.width -ge 180) 
              { $ws.width = 180 } 
        Else { $ws.height = $maxws.height } 
        IF($maxws.height -ge 42) 
              { $ws.height = 42 } 
        Else { $ws.height = $maxws.height } 
        $host.ui.RawUI.Set_WindowSize($ws) 
        #Clear-Host
}


#==============================================================================
# Sets the onscreen cursor
#==============================================================================
function Set-ConsolePosition  
{Param([Parameter(Mandatory=$false,Position=1)] [int]$X,      
              [Parameter(Mandatory=$false)][int]$Y)

              
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
              
       $position=$host.ui.rawui.cursorposition
       # Store new X and Y Co-ordinates away
       $position.x=$x
       $position.y=$y
       # Place modified location back to $HOST
       $host.ui.rawui.cursorposition=$position}
       
#==============================================================================
# Waiting for a key being pressed
#==============================================================================
#Function PauseScript
#{     winfo "Press any key to continue ..."
#      $key = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")}

#==============================================================================
# Proceed code (ask a code to confirm something)
#==============================================================================
function ProceedCode
       {Param([Parameter(Mandatory=$false)] [int]$size=5)

       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"


       $set = "abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray() 
        for ($x = 0; $x -lt $size; $x++) {$proceedcode += $set | Get-Random}
       $try=0
       while ($try -le 4 -and $in -ne $proceedcode -and $in -ne "QUIT")
                    {
                     winfo "Confirmation code '$proceedcode', type code or QUIT to quit (try $try/3)" -NoNewline    
                     $in = Read-host "?"
                     $try+=1}
       If($in -eq "QUIT")
             {return $false}
       Else
             {return ($in -eq $proceedcode)}
      }
#==============================================================================
# Sleep indicator to let admin know something is happening
#==============================================================================
function SleepIndicator
       {Param([Parameter(Mandatory=$false)] [int]$timeinsec=5)
       $c=1
       $last="\"
       $position=$host.ui.rawui.cursorposition
       while ($c -le $timeinsec)
             {$host.ui.rawui.cursorposition=$position
              wwarn "($last)               " -nonewline -notimestamp
              $c+=1
              switch ($last)
                     {"\" {$last="|"}
                     "|" {$last="/"}
                     
                      "/" {$last="-"}
                     "-" {$last="\"}
                     }
                     Start-Sleep 1}}

#==============================================================================
#Functions WITHOUT console output
#      functions in these section SHOULD not have any console OUTPUT
#      this is the preffered method
#==============================================================================


#==============================================================================
# Write info message
#==============================================================================
function wtitle
       {Param(
              [Parameter(Mandatory=$false,Position=1)] [string]$message,    
              [Parameter(Mandatory=$false)][AllowEmptyString()] [string]$CorrelationID="",
              [Parameter(Mandatory=$false)] [string]$indentation,                  
              [Parameter(Mandatory=$false)] [switch]$NoTimeStamp,
              [Parameter(Mandatory=$false)] [switch]$NoNewLine,
              [Parameter(Mandatory=$false)] [int]$forcemaxchar=-1)
                     
              If (!($NoTimeStamp.IsPresent)) 
            {write-host -foregroundcolor "gray" -backgroundcolor "Black"  "$(Get-Date -format G) - " -nonewline
                   if (!($CorrelationID -eq ""))  
                {write-host -foregroundcolor "gray" -backgroundcolor "Black" "$CorrelationID - " -nonewline}
             else
                {if ($global:CorrelationID -ne "" -and $global:CorrelationID -ne $null)  
                    {write-host -foregroundcolor "gray" -backgroundcolor "Black" "$($global:CorrelationID) - " -nonewline}}}

              if ($indentation -ne "" -and $indentation -ne $null)  
                     {$message = "$($indentation)$($message)"}
              else
                     {if ($global:WriteIndentation -ne "" -and $global:WriteIndentation -ne $null)  
                           {$message = "$($global:WriteIndentation)$($message)"}}
              
              If($forcemaxchar -ne -1)
                     {If($message.length -gt $forcemaxchar)
                           {$message = $message.substring(0,$forcemaxchar)}
                     Else
                           {$message = $message + (" " * ($forcemaxchar-$message.length) )}}
              If($NoNewLine.IsPresent) 
                     {Write-Host -foregroundcolor "white" -backgroundcolor "Black" $message -NoNewline }
              Else
                     {Write-Host -foregroundcolor "white" -backgroundcolor "Black" $message}}


#==============================================================================
# Write gray message
#==============================================================================
function wgray
       {Param(
              [Parameter(Mandatory=$false,Position=1)] [string]$message,    
              [Parameter(Mandatory=$false)][AllowEmptyString()] [string]$CorrelationID="",      
              [Parameter(Mandatory=$false)] [string]$indentation,
              [Parameter(Mandatory=$false)] [switch]$NoTimeStamp,
              [Parameter(Mandatory=$false)] [switch]$NoNewLine,
        [Parameter(Mandatory=$false)] [switch]$NoMaxCharAlignment,
              [Parameter(Mandatory=$false)] [int]$forcemaxchar=-1)
              

              If (!($NoTimeStamp.IsPresent)) 
            {write-host -foregroundcolor "gray" -backgroundcolor "Black"  "$(Get-Date -format G) - " -nonewline
                   if (!($CorrelationID -eq ""))  
                {write-host -foregroundcolor "gray" -backgroundcolor "Black" "$CorrelationID - " -nonewline}
             else
                {if ($global:CorrelationID -ne "" -and $global:CorrelationID -ne $null)  
                    {write-host -foregroundcolor "gray" -backgroundcolor "Black" "$($global:CorrelationID) - " -nonewline}}}
              if ($indentation -ne "" -and $indentation -ne $null)  
                     {$message = "$($indentation)$($message)"}
              else
                     {if ($global:WriteIndentation -ne "" -and $global:WriteIndentation -ne $null)  
                           {$message = "$($global:WriteIndentation)$($message)"}}
              If($forcemaxchar -ne -1)
                     {If($message.length -gt $forcemaxchar)
                           {$message = $message.substring(0,$forcemaxchar)}
                     Else
                           {if (!$NoMaxCharAlignment.ispresent) {$message = $message + (" " * ($forcemaxchar-$message.length) )}}}
              If($NoNewLine.IsPresent) 
                     {Write-Host -foregroundcolor "gray" -backgroundcolor "Black" $message -NoNewline }
              Else
                     {Write-Host -foregroundcolor "gray" -backgroundcolor "Black" $message}
       }
                     

#==============================================================================
# Check if y have administrative rights
#==============================================================================
function Get-UserIsAdmin  
       {Param([Parameter(Mandatory=$false,Position=1)]
                     [boolean]$StopOnFailure=$true,
                     [switch]$nomessage)
                     
       $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    if (!($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)))
              {if ($StopOnFailure)
                     {werror "Cmdline needs to be in administrative mode, cannot continue"
                     exit 10}
              if (!$nomessage.ispresent)
                    {wwarn "Administrative mode required"}
              return $false}
       else
              {if (!$nomessage.ispresent)
                     {winfo "Script running in administrative mode credentials $($env:username)"}
              return $true}
       } 
       
       
#==============================================================================
#Get and format the date
#             26/10/2012 fixed to support 24h format
#==============================================================================
Function Get-FileFormatDate {  
     param( [DateTime]$Date = [DateTime]::now )  
     return $Date.ToUniversalTime().toString( "yyyy-MM-dd_HH-mm-ss" )} 

#==============================================================================
#Get TimeSpamp
#==============================================================================
Function Get-TimeStamp {         
     return  get-date -format "yyyyMMdd-HHmmss"} 

#==============================================================================
# Fileshare PERMISSIONS
#==============================================================================
Function New-FileShare
       {Param(
              [Parameter(Mandatory=$true,Position=1)] [string]$name,
              [Parameter(Mandatory=$true,Position=2)]  [string]$path,
              [Parameter(Mandatory=$true,Position=3)]  [string]$group
                     )                    
    $Class = "Win32_Share"
    $Method = "Create"
    $description = "This is shared for me to test"
    $sd = ([WMIClass] "\\.\root\cimv2:Win32_SecurityDescriptor").CreateInstance()
    $ACE = ([WMIClass] "\\.\root\cimv2:Win32_ACE").CreateInstance()
    $Trustee = ([WMIClass] "\\.\root\cimv2:Win32_Trustee").CreateInstance()
    $Trustee.Name = $group
    $Trustee.Domain = $Null
    $Trustee.SID = @(1, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0)
    $ace.AccessMask = 2032127
    $ace.AceFlags = 3
    $ace.AceType = 0
    $ACE.Trustee = $Trustee
    $sd.DACL += $ACE.psObject.baseobject 
    $mc = [WmiClass]"\\.\ROOT\CIMV2:$Class"
    $InParams = $mc.psbase.GetMethodParameters($Method)
    $InParams.Access = $SD
    $InParams.Description = $description
    $InParams.MaximumAllowed = $Null
    $InParams.Name = $name
    $InParams.Password = $Null
    $InParams.Path = $path
    $InParams.Type = [uint32]0
    $R = $mc.PSBase.InvokeMethod($Method, $InParams, $Null)
    switch ($($R.ReturnValue))
     {
         0 {Winfo  "Share:$name Path:$path Result:Success"; break}
         2 {Wwarn  "Share:$name Path:$path Result:Access Denied";break}
         8 {Wwarn  "Share:$name Path:$path Result:Unknown Failure";break}
         9 {Wwarn  "Share:$name Path:$path Result:Invalid Name";break}
         10 {Wwarn "Share:$name Path:$path Result:Invalid Level";break}
         21 {Wwarn "Share:$name Path:$path Result:Invalid Parameter";break}
         22 {Wwarn "Share:$name Path:$path Result:Duplicate Share";break}
         23 {Wwarn "Share:$name Path:$path Result:Reedirected Path";break}
         24 {Wwarn "Share:$name Path:$path Result:Unknown Device or Directory";break}
         25 {Wwarn "Share:$name Path:$path Result:Network Name Not Found";break}
         default {Werror "Share:$name Path:$path Result:*** Unknown Error ***";break}
    }
}


#==============================================================================
# Creates a scheduled task via schedtask and a xml file
#==============================================================================
Function Create-ScheduledTask
       {Param(
             [Parameter(Mandatory=$true,Position=1)]  [string]$account,
              [Parameter(Mandatory=$true,Position=2)]  [string]$userpass,
              [Parameter(Mandatory=$true,Position=3)]  [string]$taskname,
              [Parameter(Mandatory=$true,Position=4)]  [string]$xmlfile
                     )             
       $res=ntrights -u "$account" +r SeBatchLogonRight
       $curcomputer=get-content env:computername

       winfo "removing old scheduled task $taskname"
       $myCmd = "schtasks.exe /delete /TN ""$taskname"" /F"
       $res= Invoke-Expression $myCmd
       winfo "Create scheduled task $taskname runas $account "
       $myCmd = "schtasks.exe /Create /XML ""$xmlfile"" /TN ""$taskname"" /RU ""$account"" /RP ""$userpass"" /S $curcomputer"
       $res= Invoke-Expression $myCmd
       winfo "Update priviliges"
       $myCmd = "schtasks.exe /change /TN ""$taskname""  /RL highest /RP ""$userpass"" "
       $res= Invoke-Expression $myCmd}



#==============================================================================
# Retrieve DNS servers information from Network adapter
#==============================================================================
Function Get-DNSServers
       {Param(
              [Parameter(Mandatory=$true,Position=1)]  [string]$ComputerName)

                     try{
                           $Networks = Get-WmiObject -Class Win32_NetworkAdapterConfiguration `
                                                -Filter IPEnabled=TRUE `
                                                -ComputerName $ComputerName `
                                                -ErrorAction Stop}
                     catch{
                           Write-Verbose "Failed to Query $Computer. Error details: $_"
                           continue}
                     foreach($Network in $Networks) {
                           $DNSServers = $Network.DNSServerSearchOrder
                           $NetworkName = $Network.Description
                           If(!$DNSServers) {
                                  $PrimaryDNSServerIP = "Notset"
                    $PrimaryDNSServerName = "NotSet"
                                  $SecondaryDNSServerIP = "Notset"
                    $SecondaryDNSServerName = "Notset"}
                           Elseif($DNSServers.count -eq 1) {
                                  $PrimaryDNSServerIP = $DNSServers[0]
                                  $SecondaryDNSServerIP = "Notset"
                    $SecondaryDNSServerName = "Notset"}
                           Else{
                                  $PrimaryDNSServerIP = $DNSServers[0]
                                  $SecondaryDNSServerIP = $DNSServers[1]
                    $SecondaryDNSServerNAme = [System.Net.Dns]::gethostentry($SecondaryDNSServerIP)}
                           If($network.DHCPEnabled) {$IsDHCPEnabled = $true}
                           
                If($PrimaryDNSServerIP -ne "NotSet"){$PrimaryDNSServerName = [System.Net.Dns]::gethostentry($PrimaryDNSServerIP)}
                If($SecondaryDNSServerIP -ne "NotSet"){$SecondaryDNSServerName = ([System.Net.Dns]::gethostentry($SecondaryDNSServerIP))}
                         
                           $OutputObj  = New-Object -Type PSObject
                           $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ComputerName.ToUpper()
                           $OutputObj | Add-Member -MemberType NoteProperty -Name PrimaryDNSServerIP -Value $PrimaryDNSServerIP
                $OutputObj | Add-Member -MemberType NoteProperty -Name PrimaryDNSServerName -Value $PrimaryDNSServerName
                           $OutputObj | Add-Member -MemberType NoteProperty -Name SecondaryDNSServerIP -Value $SecondaryDNSServerIP
                $OutputObj | Add-Member -MemberType NoteProperty -Name SecondaryDNSServerName -Value $SecondaryDNSServerName
                           $OutputObj | Add-Member -MemberType NoteProperty -Name IsDHCPEnabled -Value $IsDHCPEnabled
                           $OutputObj | Add-Member -MemberType NoteProperty -Name NetworkName -Value $NetworkName
                           Return $OutputObj
                           
                     }

       }
       
#==============================================================================
# Delete Between (string operation)
#==============================================================================

Function Remove-StringBetween
       {Param(
              [Parameter(Mandatory=$true,Position=1)] [string]$startdelimiter,
              [Parameter(Mandatory=$true,Position=1)] [string]$stopdelimiter,
              [Parameter(Mandatory=$true,Position=1)]  [string]$what)
              
       If($what.Contains($startdelimiter) -And $what.Contains($stopdelimiter))
         {Return $what.Substring(0, $what.IndexOf($startdelimiter)) + $what.Substring($what.IndexOf($stopdelimiter) + $stopdelimiter.Length)}
     Else
         {Return $what}}
              
               
#==============================================================================
# Delete Between (string operation) iteratief
#==============================================================================
Function Remove-StringBetweenAll
       {Param(
              [Parameter(Mandatory=$true,Position=1)]  [string]$startdelimiter,
              [Parameter(Mandatory=$true,Position=1)]  [string]$stopdelimiter,
              [Parameter(Mandatory=$true,Position=1)]  [string]$what)
       $result = $what
     While ($result.Contains($startdelimiter) -and $result.Contains($stopdelimiter))
             {$result = $result.Substring(0, $result.IndexOf($startdelimiter)) + $result.Substring($result.IndexOf($stopdelimiter) + $stopdelimiter.Length)}
       Return $result}

#==============================================================================
# Get Correlation ID todo
#==============================================================================
Function New-CorrelationID
       {Param([Parameter(Mandatory=$false)] [switch]$nomessage)
       $set = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789".ToCharArray() 
        for ($x = 0; $x -lt 20; $x++) {$correlationID += $set | Get-Random}
       if (!$nomessage.ispresent)
              {winfo "New Correlation ID '$correlationid'"}
       return $correlationID}
              
#==============================================================================
# Start keeping track of counters
#             use SetCounter "Countername" "Description"
#             to increment use setcounter "countername"
#             initilize first $counters=@{}
#==============================================================================
function Set-Counter
       {Param([Parameter(Position=0,Mandatory=$true)] [string]$ccode,
                 [Parameter(Mandatory=$false)] [string]$desc="")

       If($counters -ne $null)
              {
              If($counters.containskey($ccode))
                     {$counters[$ccode].counter += 1}
              Else
              { $c=new-object psobject -property @{
                           Name=$ccode;                      
                           Counter=0;
                           description=$desc}
                 $counters.add($ccode,$c)}}}           
           
           
#==============================================================================
# Send mail via smtp
#==============================================================================

Function Send-Mail{
       param(
              [Parameter(Position=0, Mandatory=$true)] [String] $Mailsubject,
              [Parameter(Position=1, Mandatory=$true)] [String] $MailBody,
              [Parameter(Position=2, Mandatory=$true)] [string] $MailFrom,
              [Parameter(Position=3, Mandatory=$true)] [string[]] $MailTo,
              [Parameter(Position=4, Mandatory=$false)] [string[]] $MailCc=$null,
              [Parameter(Position=5, Mandatory=$false)] [switch]$htmlbody=$false)  
       
       $smtpServer          = "relayserver.resum.internal"
       $smtp                      = new-object Net.Mail.SmtpClient($smtpServer)
       
       $msg                       = new-object Net.Mail.MailMessage
       $msg.subject         = $Mailsubject
       $msg.body                  = $MailBody
       $msg.From                  = $MailFrom
       if($htmlbody){$msg.IsBodyHtml = $true}
       $MailTo|foreach {$msg.To.Add($_)}
       if($MailCc -ne $null)
       {
              $MailCc|foreach {$msg.Cc.Add($_)}
       }

       $smtp.Send($msg)

}



#==============================================================================
# Function to remove all lines containing the pattern from a logfiles
#==============================================================================

Function Clean-Logfile{

       param( [Parameter(Position=0, Mandatory=$true)] [String] $pattern,
                     [Parameter(Position=1, Mandatory=$true)] [string] $logfile,
                     [Parameter(Position=2, Mandatory=$true)] [string] $logfile_cleaned)
       
    $servername = [System.Net.Dns]::GetHostName()

    # Use StreamReader to process each line of logfile, one line at a time, comparing each line against
    # all the patterns, incrementing the counter of matches to each pattern.  Have to use StreamReader
    # because get-content and the Switch statement are extremely slow with large files.  
    Wgray "Cleaning $logfile"
    
    $reader = new-object System.IO.StreamReader -ArgumentList "$logfile"
    $writer = [System.IO.StreamWriter] $logfile_cleaned

    If(-not $?) { "`nERROR: Could not find file: $logfile`n" ; exit }

    while ( ([String] $line = $reader.readline()) -ne $null){
        #Ignore blank lines and comment lines.
        If($line.length -eq 0 ) { continue }
              if  (!($line -match $pattern)) {$writer.WriteLine($line)}
    }
    $writer.close()
    $reader.Close()
}

#==============================================================================
# Function to get a file attribute
# ex: Get-FileAttribute $item.FullName "compressed"
#==============================================================================
function Get-FileAttribute
{
       param( [Parameter(Position=0, Mandatory=$true)] [String] $file,
                     [Parameter(Position=1, Mandatory=$true)] [string] $attribute)

       $val = ([System.IO.FileAttributes]$attribute).value__;
       if(((Get-Item -force $file).Attributes -band $val) -eq $val){$true}
       Else{$false}
} 

#==============================================================================
# Function to set the "Compressed" file attibute to true
#==============================================================================

function Compress-File
{
       param( [Parameter(Position=0, Mandatory=$true)] [String] $file)
       
       If($file.StartsWith("\\")){# "UNC" path are not supported by CIM_DATAFILE
              $Array = $file.Split("\")
              $computername=$Array[2]
              $ShareName=$Array[3]
              $Share= Get-WmiObject win32_share -computer $computername | Where-Object { $_.Name -eq $ShareName }
        if($share.path.endsWith("\")){$Sharepath = $Share.path.Replace("\","")}
              Else{$Sharepath = $Share.path}
              $Parentpath= "\\" + [string]::join("\",$file.Split("\")[2..3])
              $path=$Sharepath + $file.substring($parentpath.length,$file.length-$parentpath.length)
              $wmifilepath = $path.Replace("\","\\")}
       Else{# local folder
              $computername = get-content env:computername
              $wmifilepath = $file.Replace("\","\\")}
       #check and see if file is already compressed
       $CIMFile = Get-WmiObject -Class CIM_DATAFILE -filter "Name='$wmifilepath'" -ComputerName $computername
       If(-Not $CIMFile.Compressed){
              Wgray "Compressing $file..."
                     $rc=$CIMFile.Compress()
                     If($rc.ReturnValue -ne 0){WError "Failed to compress $file."
                           $false}
                     Else{$true}}
       Else{$true}
}


#==============================================================================
# Display result
#==============================================================================

Function DisplayResult
{
       param( [Parameter(Position=0, Mandatory=$true)] $result)
       
       If($result.status -eq $false){
              werror "(Failed)" -notimestamp
              Werror "$($result.errormsg)" -notimestamp}
       ElseIf($result.alreadyexists -eq $true -or $result.alreadymember -eq $true -or $result.alreadydone -eq $true){winfo "(Already done)" -notimestamp
       }
       Else{wwarn "(Configured)" -notimestamp}
}

#==============================================================================
# Display Result
#==============================================================================

Function DisplayResult2
{
       param([Parameter(Position=0, Mandatory=$true)] $result,
       [Parameter(Position=1, Mandatory=$false)] [switch]$ExitOnError
       )
       
       #Write-Host "AlreadyDone: ""$($result.alreadydone)""" -NoNewline
       If($result.status -eq $false)
       {
              Werror "NOK" -notimestamp
              Werror "$($result.errormsg)" -notimestamp
              if($ExitOnError.IsPresent)
              {
                     stop-transcript
                     Exit
              }
       }
       ElseIf($result.alreadydone)
       {
              Wwarn "Already done" -notimestamp
       }
       ElseIf($result.alreadyExists)
       {
              Wwarn "Already done" -notimestamp
       }
       Else
       {
              Winfo "OK" -notimestamp
       }
}


#==============================================================================
# Validate if the string parameter is a valid IP address
#==============================================================================

Function Validate-IPAddress{

       param(
              [Parameter(Position=0, Mandatory=$true)] [String]      $IPAddress)

       try{
              $ip = [System.Net.IPAddress]::parse($ipAddress)
              $isValidIP = [System.Net.IPAddress]::tryparse([string]$ipAddress, [ref]$ip)}
       Catch{}
       
       If($isValidIP){return $true}
       Else{return $false}
}


#==================================================================================
# Create a New DNS record (only CNAME and A Record are supported in this version)
# ex: Create-dnsrecord JAROOT01 test.aginsurance.extranet 10.14.243.231 -arec
#==================================================================================


function Create-dnsrecord { 
       
       param(
              [Parameter(Position=0, Mandatory=$True)] [String]      $DNSserver,
              [Parameter(Position=1, Mandatory=$True)] [string]      $RecordName,
              [Parameter(Position=2, Mandatory=$False)] [string]     $IPaddress=$null,
              [Parameter(Position=3, Mandatory=$False)] [string]     $targethost,
              [Parameter(Position=4, Mandatory=$False)] [switch]     $arec,                            #For A record
              [Parameter(Position=5, Mandatory=$False)] [switch]     $cname)                    #For CNAME

       $rt = new-object psobject -property @{status=$false;alreadyexists=$false;errormsg=""}

       #parameters validation
       
       If($DNSServer -eq $null){
              $rt.errormsg="Missing parameter DNSServer"
              return $rt}
       Else{
              If(-not (Test-Connection -ComputerName $DNSserver -quiet)){
                     $rt.errormsg="DNS server not found"
                     return $rt}}
       
       If($RecordName -eq $null){
              $rt.errormsg="Missing parameter RecordName"
              return $rt}
       Else{
              $FLZoneName = $RecordName.substring($RecordName.IndexOf(".") +1, $RecordName.Length - $RecordName.IndexOf(".") -1)
              $FLZone=Get-WmiObject -ComputerName $DNSServer -Namespace "root\MicrosoftDNS" -Class MicrosoftDNS_Zone|Where{$_.ContainerName -match $FLZoneName}
              If($FLZone -eq $null){
                     $rt.errormsg="Zone ""$FLZoneName"" not found"
                     return $rt}}
       
       If($cname){
              If(($targethost -eq $null) -or ($targethost -eq "")){
                     $rt.errormsg="Missing parameter targethost"
                     return $rt}
              Else{
                     $ZoneName = $TargetHost.substring($TargetHost.IndexOf(".") +1, $TargetHost.Length - $TargetHost.IndexOf(".") -1)
                     $Zone=Get-WmiObject -ComputerName $DNSServer -Namespace "root\MicrosoftDNS" -Class MicrosoftDNS_Zone|Where{$_.ContainerName -match $ZoneName}
                     If($Zone -eq $null){
                           $rt.errormsg="Zone ""$ZoneName"" not found"
                           return $rt}}}
       
       If($arec){
              If(($IPAddress -eq $null) -or ($IPAddress -eq "")){
                     $rt.errormsg="Missing parameter IPAddress"
                     return $rt}
              Else{
                     try{
                           $ip = [System.Net.IPAddress]::parse($ipAddress)
                           $isValidIP = [System.Net.IPAddress]::tryparse([string]$ipAddress, [ref]$ip)}
                     Catch{$isValidIP=$false}
                     
                     If(-not($isValidIP)){
                           $rt.errormsg="IPAddress is not valid."
                           return $rt}}}
       $CreateRecord=$false
       $DNSRecord=Get-WmiObject -ComputerName $DNSServer -Namespace "root\MicrosoftDNS" -Class MicrosoftDNS_ResourceRecord  -Filter "ContainerName ='$($FLZone.name)'"|Where{$_.ownername -match $RecordName}
       If($DNSRecord -ne $null){
              $rt.alreadyexists=$true
              
              If($arec){ 
              If($DNSRecord.__CLASS -ne "MicrosoftDNS_AType"){
                           $rt.errormsg+="It is not a A record"
                           $rt.alreadyexists=$false} 
                     If(!($DNSRecord.RecordData -match $IPaddress)){
                           $rt.errormsg+="Another IP address is already set:$($DNSRecord.RecordData)"
                           $DNSRecord.psbase.Delete() 
                           $CreateRecord=$true
                           $rt.alreadyexists=$false}}
              
              
              If($cname){ 
              If($DNSRecord.__CLASS -ne "MicrosoftDNS_CNAMEType"){
                           $rt.errormsg+="It is not a CNAME record"
                           $rt.alreadyexists=$false} 
                     If($DNSRecord.RecordData -ne ($Targethost + ".")){
                           $rt.errormsg+="Another target host is already set:$($DNSRecord.RecordData)"
                           $DNSRecord.psbase.Delete() 
                           $CreateRecord=$true
                           $rt.alreadyexists=$false}}

              if($rt.alreadyexists){$rt.status=$true}
              
       }
       Else{$CreateRecord=$true}
       
       If($CreateRecord){
              $rec = [WmiClass]"\\$DNSserver\root\MicrosoftDNS:MicrosoftDNS_ResourceRecord"  

              ## A Record
              If($arec){ 
                     $text = "$RecordName IN A $IPaddress"  
                     $r=$rec.CreateInstanceFromTextRepresentation($DNSserver, $FLZone.name, $text)
                     $rt.status=$true}
              ## CNAME 
              If($cname){ 
                     $text = "$RecordName IN CNAME $targethost" + "."
                     $r=$rec.CreateInstanceFromTextRepresentation($DNSserver, $FLZone.name, $text)
                     $rt.status=$true}}
       
       return $rt
} 

#==================================================================================
# Return true if a port is opened
#==================================================================================

Function Test-OpenedPort{

       param (
       [Parameter(Position=0, Mandatory=$True)][string] $Hostname,
       [Parameter(Position=1, Mandatory=$True)][string] $Port)

       $TimeOut      = 1000
       $IPAddress    = [System.Net.Dns]::GetHostAddresses($Hostname)
       $Address      = [System.Net.IPAddress]::Parse($IPAddress)
       $Socket       = New-Object System.Net.Sockets.TCPClient
       $Connect      = $Socket.BeginConnect($Address,$Port,$null,$null)
       If($Connect.IsCompleted){
              $Wait = $Connect.AsyncWaitHandle.WaitOne($TimeOut,$false) 
              If(!$Wait){
                     $Socket.Close() 
                     return $false} 
              Else{
              $Socket.EndConnect($Connect)
              $Socket.Close()
              return $true}}
       Else{return $false}
}

#==============================================================================
# Get all databases from a server
#      Returns TRUE if the database does NOT exist
#==============================================================================
Function Get-SQLDataBase
{Param(
              [Parameter(Mandatory=$True,Position=1)] [string]$dbconnection,       
              [Parameter(Mandatory=$True,Position=2)] [string]$dbname)      

  $rt = [PSCustomObject]@{status=$false;result=$false;errormsg="";trace=""} 

   try
              {[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
            $srv = new-object ('Microsoft.SqlServer.Management.Smo.Server') $dbconnection

         $srv.ConnectionContext.LoginSecure=$true

         $db= $srv.Databases | ? { $_.name -like $dbname } 
         $rt.result = ($db -eq $null)
         $rt.status = $true
         $rt.trace="finished"     
 
         }

       catch
               {
         $rt.errormsg=$_.Exception.Message
         $rt.trace="unkown"
              $rt.status=$false}
  return $rt

  }

#==============================================================================
# Launch a query to a database server
#==============================================================================
Function Get-DbQuery
{Param(
              [Parameter(Mandatory=$True,Position=1)] [string]$dbserver,    
              [Parameter(Mandatory=$True,Position=2)] [string]$dbname,
              [Parameter(Mandatory=$True,Position=3)]  [string]$query)      

  $rt = [PSCustomObject]@{status=$false;dbserver=$dbserver;dbname=$dbname;query=$query;result=$null;errormsg=""}       


  try
              {[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | out-null
            $srv = new-object ('Microsoft.SqlServer.Management.Smo.Server') "$dbserver"  

         $srv.ConnectionContext.LoginSecure=$true
               write-verbose "Executing $query on  $dbserver/$dbname"

               $db = new-object ("Microsoft.sqlServer.Management.Smo.Database") ($dbserver, $dbname)
               $rt.result = $db.ExecuteWithResults($query)
         $rt.status=$true}
        catch
              {$rt.errormsg=$_.Exception.Message
              $rt.status=$false}
  return $rt } 

#==============================================================================
#Create ActiveDirectory group when not exists
#==============================================================================
Function new-AdGroup
       {Param(
              [Parameter(Mandatory=$True,Position=1)] [string]$ou,   
              [Parameter(Mandatory=$True,Position=2)]  [string]$groupname,
              [Parameter(Mandatory=$True,Position=3)]  [string]$desc,
              [Parameter(Mandatory=$false)] [switch]$global)                
                     
       $rxt=new-object psobject -property @{alreadyexists=$false;status=$false;groupname="";desc="";errormsg=""}
       $rxt.groupname=$groupname
       $rxt.desc=$desc
       $rxt.status=$false

       If ((ADGroupExists $groupname) -ne $null) 
        {$rxt.alreadyexists=$true
              $rxt.status=$true}
     Else 
        {If ($global.IsPresent)
                     {$groupType=$global:ADS_GLOBALGROUP}                          
               Else
                     {$groupType=$global:ADS_LOCALGROUP}

              try {$root = [ADSI] "LDAP://$ou"
                     $grp = $root.Create("group", "cn=$groupname")
                      $grp.Put("sAMAccountName", $groupname)
                     $grp.Put("groupType", $groupType )
                      $grp.Put("Description", $desc)
                     $grp.SetInfo()                                        
                     $rxt.status=$true}
              catch [Exception]
                     {$rxt.errormsg=$_.Exception.Message}}

    return $rxt}

#==============================================================================
#Create A User only when it does not yet exist
#==============================================================================
Function New-User
       {Param(
              [Parameter(Mandatory=$True,Position=1)] [string]$ou,   
              [Parameter(Mandatory=$True,Position=2)]  [string]$account,
              [Parameter(Mandatory=$True,Position=3)]  [string]$pass,
              [Parameter(Mandatory=$true,Position=4)]  [string]$comment)
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
        $rt= new-object psobject -property @{alreadyexists=$false;status=$false;username="";comment="";created=$false;error="";errormsg="";detailederrormsg="";pass="";action=""}
       $rt.comment = $comment
       $rt.username = $account
       $rt.pass = $pass
            
        If((ADUserExists $account) -ne $null) 
              {write-verbose "AD User $account already exist"
              $rt.alreadyexists=$true 
               $rt.status=$true 
               $rt.created=$false}
       Else 
              {write-verbose "AD user $account does not yet exist"
              $root = [ADSI] "LDAP://$ou"
              
               
               try {
                     write-verbose "CreateUser $account 1"
                     $User = $root.Create("user", "cn=$account")
                     write-verbose "CreateUser $account 2"
                     $User.Put("sAMAccountName", $account)
                     $User.Put("userPrincipalName", $account)
                     $User.Put("Description", $comment)
                     write-verbose "CreateUser $account 2"
                     $User.SetInfo()
                     write-verbose "CreateUser $account ok"
                     try {
                            $User.PsBase.Invoke("SetPassword", $pass)
                            $User.PsBase.InvokeSet("AccountDisabled", $false)
                           #$user.psbase.invokeset("accountdisabled",$false)
                            $User.SetInfo()
                           write-verbose "Password and account set ok"
                           try {
                                  New-Variable ADS_UF_DONT_EXPIRE_PASSWD 0x10000 -Option Constant
                                  #Marc's change request
                                  #New-Variable ADS_UF_PASSWD_NOTREQD 0x00020 -Option Constant
                                  New-Variable ADS_UF_PASSWD_NOTREQD 0x20 -Option Constant
                                  [int]$flag=$User.useraccountcontrol[0]
                                  $flag=$flag -bor $ADS_UF_DONT_EXPIRE_PASSWD
                                  $flag=$flag -bor $ADS_UF_PASSWD_NOTREQD
                                  $User.useraccountcontrol=$flag 
                                  $User.SetInfo()
                                  write-verbose "Expiration flag set ok"
                                  $rt.status=$true
                                  $rt.created=$true
                                  $rt.action="User created"}
                            catch [Exception]
                                  {write-verbose  "user expiration flag failed with $($_.Exception.Message)"
                                   $rt.status=$false
                                  $rt.detailederrormsg=$_.Exception.Message
                                  $rt.errormsg="Failed:Flags"}}
              catch [Exception]
                           {write-verbose "User password set failed with $($_.Exception.Message)"  
                            $rt.status=$false
                           $rt.detailederrormsg=$_.Exception.Message
                           $rt.errormsg="Failed:Password set"}}
              catch [Exception]
                            {write-verbose "User creation failed with $($_.Exception.Message)"  
                            $rt.status=$false
                           $rt.detailederrormsg=$_.Exception.Message
                           $rt.errormsg="Failed:create"}}
       return $rt}


#==============================================================================
#Find ActiveDirectory Group Object
#==============================================================================
Function ADGroupExists
       {Param([Parameter(Mandatory=$True,Position=1)] [string]$groupname)
         return ADObjectExists "(&(objectClass=group)(cn=$groupname))"} 
         
#==============================================================================
#Find ActiveDirectory User Object
#==============================================================================
Function ADUserExists 
       {Param([Parameter(Mandatory=$True,Position=1)] [string]$username)      
         return ADObjectExists "(&(objectClass=user)(sAMAccountName=$username))"} 


#==============================================================================
#Find ActiveDirectory Computer Object
#==============================================================================
Function ADComputerExists
       {Param([Parameter(Mandatory=$false,Position=1)] [string]$ComputerName=$env:computername)
       return ADObjectExists "(&(objectClass=computer)(cn=$ComputerName))"}

#==============================================================================
#Find ActiveDirectory Object
#==============================================================================
Function ADObjectExists
       {Param([Parameter(Mandatory=$true,Position=1)] [string]$adfilter)
         $root = [ADSI] "LDAP://$($global:CurrentADdomain)"
      $searcher = New-Object System.DirectoryServices.DirectorySearcher $root
         $searcher.filter = $adfilter
         $adobj = $searcher.FindOne()
         If($adobj  -eq $null) 
            {return $null} 
              Else
            {return $adobj.path}}   

#==============================================================================
#Add members to a Group
#==============================================================================
Function Add-ToAdGroup
       {Param(
              [Parameter(Mandatory=$True,Position=1)] [string]$groupname,   
              [Parameter(Mandatory=$True,Position=2)]  [string]$username,
        [Parameter(Mandatory=$False)]    [switch]$Group)

       $rt= new-object psobject -property @{status=$false;alreadydone=$false;errormsg="";rawerror=""}
       
        
     $destinationgroup=ADGroupExists $groupname
     switch ($true)
       {($group.ispresent)
            {$user=ADGroupExists $username}
        (!$group.ispresent)
            {$user=ADUserExists $username}
        ($destinationgroup -eq $null) 
            {$rt.errormsg = "Group '$groupname' doesnt exist"
             break}
           ($user -eq $null -and !$group.ispresent)  
                     {$rt.errormsg = "user '$username' doesnt exist"
             break}
           ($user -eq $null -and $group.ispresent)  
                     {$rt.errormsg = "Group '$username' doesnt exist"
             break}
              ($true)
                  {
             try 
                            {$root = [ADSI] "LDAP://$($global:CurrentADdomain)"
                          $adgroup = [ADSI]$destinationgroup
                          $aduser = [ADSI]$user                 
                 $r=$adgroup.add($aduser.psbase.path)
                           $rt.status = $true
                           return $rt}
                       catch [Exception]
                                   {$rt.errormsg = $_.Exception.Message
                     $rt.rawerror=$_.Exception.Message
                     If($_.Exception.Message -eq "The object already exists. (Exception from HRESULT: 0x80071392)")
                                         {write-verbose "already member of group"
                                         $rt.alreadydone=$true
                                         $rt.status=$true}
                                         Else
                                         {write-verbose "Failed with $($_.Exception.Message)"  }}}}
              return $rt }

#==============================================================================
#Add members to a LOCAL Group
#==============================================================================
Function Add-ToLocalGroup
       {Param(
              [Parameter(Mandatory=$True,Position=1)]  [string]$Account,    
              [Parameter(Mandatory=$false)] [string]$group="administrators",
              [Parameter(Mandatory=$false)] [string]$classtype="group")     
       $rt= new-object psobject -property @{status=$false;alreadymember=$false;errormsg="";rawerror="";stack=@()}
       try
             {$account = $account.replace("\","/")
              $adsi = [ADSI]"WinNT://$($env:computername)/$group,$classtype"
               $adsi.add("WinNT://$Account,$classtype")
              $adsi = $null
              $rt.status=$true}
       catch
             {$rt.rawerror = $_.exception.message
              If($rt.rawerror -like "*is already a member*")
                     {$rt.alreadymember=$true
                     $rt.status=$true}
              Else
                    {$rt.status=$false}
              $err = $_.Exception
              while ( $err.InnerException ) 
                       {$err = $err.InnerException
                        $rt.stack += $err.message}}
                           
       return $rt   
       }

#==============================================================================
#Remove members from a LOCAL Group
#==============================================================================
Function Remove-fromLocalGroup
       {Param(
              [Parameter(Mandatory=$True,Position=1)]  [string]$Account,    
              [Parameter(Mandatory=$false)] [string]$group="administrators",
              [Parameter(Mandatory=$false)] [string]$classtype="group")     
       $rt= new-object psobject -property @{status=$false;wasmember=$false;errormsg="";rawerror=""}
       try
             {$account = $account.replace("\","/")
              $adsi = [ADSI]"WinNT://$($env:computername)/$group,$classtype"
               $adsi.remove("WinNT://$Account,$classtype")
              $adsi = $null
              $rt.status=$true}
       catch
             {$rt.rawerror = [string]$_.exception.message
              $rt.status=$false
              $err = $_.Exception
              while ( $err.InnerException ) 
                       {$err = $err.InnerException
                        $rt.stack += $err.message}}           
        return $rt   }

                           
#==============================================================================
#Encrypt a password using a key
#==============================================================================
Function New-CryptedPassword
       {Param(
              [Parameter(Mandatory=$True,Position=1)]
                     [string]$password,   
              [Parameter(Mandatory=$True,Position=2)]
                     [string]$key)              
       $epassword= $password | ConvertTo-SecureString -AsPlainText -force
       $sepassword= $epassword | ConvertFrom-SecureString -Key ([Byte[]]$key.Split(" "))
       return $sepassword
       }

#==============================================================================
#Decrypt a password using a key
#==============================================================================
Function Get-CryptedPassword 
       {Param(
              [Parameter(Mandatory=$True,Position=1)]  [string]$password,   
              [Parameter(Mandatory=$True,Position=2)]  [string]$key) 
       $passwordSecure = ConvertTo-SecureString -String $password -Key ([Byte[]]$key.Split(" "))
       $npw=[System.Runtime.InteropServices.marshal]::PtrToStringAuto([System.Runtime.InteropServices.marshal]::SecureStringToBSTR($passwordSecure))
       return $npw
       }
       

                                  
#==============================================================================
# Checks if the variable is numeric
#==============================================================================
function isNumeric 
       {Param([Parameter(Mandatory=$true,Position=1)]  [string]$x)
       try 
              {0 + $x | Out-Null
              return $true    } 
        catch {return $false}}

#==============================================================================
# Calculate the MD5 of a file
#==============================================================================
Function Get-FileMD5
       {Param(
              [Parameter(Mandatory=$false,Position=1)] [string]$filename)

    $md5 = new-object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
    $hash = [System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($filename)))
    return $hash}

#==============================================================================
# Remove Certain XML lines
#      example:
#      $keyList=@("<IsUpgrade>","<URI>")
#      Remove-ContentLines ".\infile.xml" ".\outfile.xml" $keylist -force
#==============================================================================
Function Remove-ContentLines
       {Param(
              [Parameter(Mandatory=$false,Position=1)] [string]$filename,   
              [Parameter(Mandatory=$false,Position=2)] [string]$outfilename,       
              [Parameter(Mandatory=$false,Position=3)] [array]$keyList,
              [Parameter(Mandatory=$false)] [switch]$force)
              
       [string]$xdata = Get-Content $filename
       $result=""
       switch ($true)
             {((Test-Path $outfilename) -and $force.ispresent)
                    {Remove-Item $outfilename -force
                     Write-Host "File exists and was removed"}
              (Test-Path $outfilename)
                    {Write-Host "File exists cannot continue"
                     break}
              (! (Test-Path $outfilename))
                    {Write-Host "Creating file"
                     $reader = [System.IO.File]::OpenText($filename)
                     try {
                     for(;;) {
                           $line = $reader.ReadLine()
                                  
                           If($line -eq $null) { break }
                                  $Startswith=$false
                                  foreach ($key in $keyList)
                                         {If($line.toupper().trim().startswith($key.toupper()))
                                                {$Startswith = $true
                                                break}}
                                  
                                  If(! $Startswith)
                                         {"$line`n`r" | Out-File $outfilename -Encoding "UTF8" -Append }}}
                     finally {$reader.Close()}
                                  }}}


       
#==============================================================================
# Start logging
#==============================================================================
Function Start-AgLogging
       {Param(
              [Parameter(Mandatory=$true,Position=1)] [string]$logfile,     
              [Parameter(Mandatory=$false)] [switch]$notranscript)
              
    if (!$notranscript.ispresent) 
       {if ($Host.name -eq "Windows PowerShell ISE Host")
            {wwarn "Transcript not supported in ISE host"}
         else
            {$logpath = split-path $logfile
             if (!(Test-Path "$($logpath)")) {$null=New-Item "$($logpath)" -type directory -ea:silentlycontinue}
             Start-Transcript $logfile -Append -ea:silentlycontinue  
                winfo "logging to $logfile" }}

    winfo "Script build: $buildnummer Powershell: $($Host.Version) by $($env:username)"}

#==============================================================================
# End script
#==============================================================================          
Function EndScript
       {param([Parameter(Mandatory=$true,Position=1)][int]$exitcode, 
              [Parameter(Mandatory=$false,Position=2)] [string]$message="",
        [Parameter(Mandatory=$false)] [datetime]$starttime=$global:scriptstarttime,
        [Parameter(Mandatory=$false)] [switch]$nomessage,
        [Parameter(Mandatory=$false)] [switch]$notranscript,
        [Parameter(Mandatory=$false)] [switch]$keeppssession)
       
     switch ($true)
        {($message -ne "" -and !$nomessage.ispresent)
                {werror $message}      
            (!$notranscript.ispresent -and $Host.name -eq "Windows PowerShell ISE Host" -and !$nomessage.ispresent)
                {wwarn "Transcript not supported in ISE host"}
            (!$notranscript.ispresent -and $Host.name -ne "Windows PowerShell ISE Host")
                  {$n=Stop-Transcript -ea:silentlycontinue}
         ($starttime -ne $null -and !$nomessage.ispresent)          
                      {$timespend = New-TimeSpan $($scriptstarttime) $(Get-Date)
                       winfo "Finished execution (exitcode $exitcode) in $($timespend.totalseconds)s"}
         ($starttime -eq $null -and !$nomessage.ispresent)          
                     {winfo "Finished execution (exitcode $exitcode)"}}
     if (!$keeppssession.ispresent) {get-pssession|remove-pssession }
       exit $exitcode}
       
#==============================================================================
# Execution library options
#==============================================================================
if (!$NOUICONFIG.ispresent -and $host.Name -eq "ConsoleHost")              
       {(Get-Host).UI.RawUI.ForeGroundColor ="gray"
       Set-DefaultHostConfig $HostTitle}

if ($NeedAdmin.ispresent)                
       {$n=Get-UserIsAdmin}

#==============================================================================
# Wait for AD replication. Check will be timed out after 90 seconds by default.
# Return True when object is created on all DC
# Return False when, at least, one DC is not synchronized
# Currently it only check replication of an object creation. None check of attribute change  has been done.
#Ex:   Wait-ADReplication "FSL SQLSRC`$_RX" "Group"
#             Wait-ADReplication x30869 "User"
#==============================================================================

Function Wait-ADReplication()
{
       Param(
    [Parameter(Position=0,Mandatory=$True)] [String] $Objectname, 
    [Parameter(Position=1,Mandatory=$false)] [String] $ObjectClass="Group",
       [Parameter(Position=2,Mandatory=$False)] [String] $TimeoutLimit="300",
    [Parameter(Mandatory=$False)] [switch] $nomessage,
    [Parameter(Mandatory=$False)] [int] $waittimebetweentry=1000)

    $rt= [PSCustomObject]@{isReplicated=$false;NrOfDc="";dclist=$null;errormsg="";query=$null}

    If($ObjectClass -eq "Group") {$rt.query = "(&(objectClass=Group)(cn=$ObjectName))"}
       If($ObjectClass -eq "User")  {$rt.query = "(&(objectClass=user)(sAMAccountName=$ObjectName))"}

       $timeout = new-timespan -Seconds $TimeoutLimit
       $sw = [diagnostics.stopwatch]::StartNew()
    $rt.dclist= [DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().DomainControllers |select Name,isReplicated,errormsg
       
    $rt.NrOfDc = $dcs.count
    if ($rt.NrOfDc -gt 0)
           {
         Do{
            foreach ($dc in $rt.dclist|? {!$_.isReplicated})                         
                               {try
                         {$SearchRoot = [ADSI]"LDAP://$($dc.Name)"
                                     $Searcher = New-Object DirectoryServices.DirectorySearcher($SearchRoot, $rt.query)                           
                                     $res=$Searcher.FindOne()
                                     Write-Verbose "$DomainController : $($res.path)"
                                      If($res -ne $null){$dc.isReplicated=$true}}
                            catch
                          {$dc.errormsg =$_.Exception.Message}}

            Start-Sleep -Milliseconds $waittimebetweentry
            $rt.isReplicated = (($rt.dclist|? {!$_.isReplicated}).count -eq 0)}
              
            While(!$rt.isReplicated -and ($sw.elapsed -lt $timeout))
            $sw.stop()
            Write-Verbose "$($sw.elapsed)"}
    else
        {$rt.errormsg="No DC found !"}
    return $rt}



#==============================================================================
#Create A User with a random password
#==============================================================================
Function CreateorUpdateUserWithRandomPassword
       {Param(
              [Parameter(Mandatory=$True,Position=1)]
                     [string]$ou,  
              [Parameter(Mandatory=$True,Position=2)]
                     [string]$account,
              [Parameter(Mandatory=$true,Position=3)]
                     [string]$comment)
                     
       write-verbose "called CreateorUpdateUserWithRandomPassword $account" 
        $pass=New-Password 8
       write-verbose "Creating user $account with one time password"
       $rt=New-User $adou $account $pass $comment
       if ($rt.status -eq $true)
             {if ($rt.alreadyexists -eq $true)
                    {write-verbose "Account already exist and needs a password reset"
                     $rtReset=ResetUserPass $account $pass
                     if ($rtreset.status -eq $false)
                           {$rt.status = $false
                           $rt.errormsg = $rtReset.errormsg}
                     else
                           {$rt.action = "Password reset"}}}
       else
             {write-verbose "User could not be created"}
       return $rt}

#==============================================================================
#Generates a random password
#==============================================================================
Function New-Password ()
       {param ( [int]$Length = 8 )
       
       write-verbose "called New-Password with length $Length" 
       $Assembly = Add-Type -AssemblyName System.Web
       $RandomComplexPassword = [System.Web.Security.Membership]::GeneratePassword($Length,2)
       return $RandomComplexPassword}

#==============================================================================
#Reset User Return object
#==============================================================================
Function ResetUserReturn()
   {return new-object psobject -property @{status=$false;username="";desc="";errormsg="";pass=""}}


#==============================================================================
# Add ACL  (only tested on a folder currently)
#==============================================================================
function Add-ACLRight()
{
       param(
       [Parameter(Position=1, Mandatory=$true)] [String] $Path,
       [Parameter(Position=2, Mandatory=$true)] [String] $Account,
       [Parameter(Position=3, Mandatory=$true)] [string] $Right,
       [Parameter(Position=4, Mandatory=$false)]       [string] $Access="Allow",
       [Parameter(Position=5, Mandatory=$false)]       [string] $InheritanceFlags = "containerinherit, ObjectInherit",
       [Parameter(Position=6, Mandatory=$false)]       [string] $PropagationFlags = "None")
       
       $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg="";trace=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"  
       
       $objACL = Get-ACL $Path
    switch ($true)
         {($objACL -eq $null)
            {$rt.status = $false
                   $rt.errormsg = "Could not find the folder $path"
             break}
          ($objACL -ne $null)
            {$colRights = [System.Security.AccessControl.FileSystemRights]$Right
             $InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]$InheritanceFlags 
                $PropagationFlag = [System.Security.AccessControl.PropagationFlags]::$PropagationFlags 
                $objType =[System.Security.AccessControl.AccessControlType]::$Access 
                $objUser = New-Object System.Security.Principal.NTAccount($Account) 
                $objACE = New-Object System.Security.AccessControl.FileSystemAccessRule("$objUser", "$colRights", "$InheritanceFlag", "$PropagationFlag", "$objType") 
             $acelist = $objACL.access | ? {$_.FileSystemRights -eq $objACE.FileSystemRights -and 
                              $_.AccessControlType -eq $objACE.AccessControlType -and 
                              $_.IdentityReference -eq $objACE.IdentityReference -and 
                              $_.IsInherited -eq $objACE.IsInherited -and 
                              $_.InheritanceFlags -eq $objACE.InheritanceFlags -and 
                              $_.PropagationFlags -eq $objACE.PropagationFlags}}
        ($acelist -ne $null)
            {$rt.status = $true
             $rt.alreadydone = $true
             break}

        ($acelist -eq $null)
             {try
                {$objACL.AddAccessRule($objACE)
                           Set-ACL $Path $objACL -ea:stop
                 $rt.status = $true}
             catch
                {$rt.status = $false
                 $rt.errormsg = $_.Exception.Message
                 break}}}
    return $rt
}
#==============================================================================
# New-Folder Ag style
#==============================================================================
function New-AgFolder
    {param([Parameter(Position=1, Mandatory=$true)] [String] $Path)

    $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg="";trace=""}
    switch ($true)
        {(Test-Path $Path)
            {$rt.status=$true
             $rt.alreadydone=$true
             break}
         ($true)      
            {try 
                {New-Item $Path -type directory}
            catch
                {$rt.status = $false
                 $rt.errormsg = $_.Exception.Message
                 break}}
        (Test-Path $Path)
            {$rt.status=$true}}
    return $rt}



#==============================================================================
# Get hash of string 
# use SHA512 algorithm
#==============================================================================

function Get-textHash()
{
       param ([String]$textToHash)
    
       $hasher = new-object System.Security.Cryptography.SHA1Managed
    $toHash = [System.Text.Encoding]::UTF8.GetBytes($textToHash)
    $hashByteArray = $hasher.ComputeHash($toHash)
    foreach($byte in $hashByteArray){$res += $byte.ToString()}
    return $res;
}

#==============================================================================
# Get hash of Folder content based on Filter (array of regular expression)
# Filenames are included in the result hash
# use SHA512 algorithm
#==============================================================================

function Get-Folderhash()
{
       param 
       ([Parameter(Position=0, Mandatory=$true)]       [String] $FolderPath,
       [Parameter(Position=1, Mandatory=$false)]       [String[]] $arrFilter = "^(?i)(.*)")
       
       $Folderhash   = $null
       $Hasharr             = Get-ChildItem $FolderPath | Where-Object {!$_.psiscontainer } |Sort-Object name | get-hash -Algorithm SHA1
       $HasharrResult       = @()

       foreach($Hash in $Hasharr){
              $file = Get-ChildItem $Hash.path
              If((Select-String $arrFilter -inp $file.name) -ne $null)
              {
                     $HasharrResult += $Hash
                     $Folderhash += $Hash.hashstring
                     $Folderhash += (Get-textHash $file.name)}
       }
       
       
       $hashrt= new-object psobject -property @{
              Folderhash=$Folderhash 
              HashArray=$HasharrResult}
       
       return $hashrt
}


#==============================================================================
# Install a MSI
#==============================================================================

Function Install-MSI {


       param(
              [Parameter(Position=0, Mandatory=$true)] [String] $msiFile,
              [Parameter(Position=1, Mandatory=$true)] [String] $msiDisplayName,
              [Parameter(Position=2, Mandatory=$true)] [String] $msiDisplayVersion,
              [Parameter(Position=3, Mandatory=$false)] [String] $msitargetDir=$null)

       
       $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
       
       $rt = new-object psobject -property @{status=$false;alreadyexists=$false;errormsg=""}
       
       if (!(Test-Path $msiFile)){
              $rt.errormsg =  "Invalid MSI Filepath: $($msiFile)."
              return $rt}
       
       $arguments = @(
       "/i"
       "`"$msiFile`""
       "/qn"
       "/norestart")
       
       if ($msitargetDir -ne $null -and $msitargetDir -ne ""){
       
       if (!(Test-Path $msitargetDir)){
               $rt.errormsg =  "Invalid target directory: $($msitargetDir)."
                     return $rt}

           $arguments += "INSTALLDIR=`"$msitargetDir`""}

       $MSI = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall | Where {$_.GetValue("DisplayName") -eq $msiDisplayName }
       if ($MSI -ne $null) {
              $rt.alreadyexists = $true
              If ($MSI.GetValue("DisplayVersion")      -eq $msiDisplayVersion){
                     $rt.status = $true
                     $rt.errormsg = """$msiDisplayName - $msiDisplayVersion"" is already installed."}
              Else{$rt.errormsg = ("Another version of ""$msiDisplayName"" is already installed:" + $MSI.GetValue("DisplayVersion"))}
              return $rt}
       Else{
              $process = Start-Process -FilePath msiexec.exe -ArgumentList $arguments -Wait -PassThru
              if ($process.ExitCode -eq 0){
                     $rt.status = $true
                     $rt.errormsg = """$msiDisplayName - $msiDisplayVersion"" successfully installed"}
              else {$rt.errormsg =  "Installation of $($msifile) failed with exit code $($process.ExitCode)"}
              return $rt}
}


#==============================================================================
# Unstall a MSI
#=============================================================================

#Find key: 
# Don't use Win32_product to query http://blogs.technet.com/b/askds/archive/2012/04/19/how-to-not-use-win32-product-in-group-policy-filtering.aspx
# Use regedit: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{3965C9F9-9B9A-4391-AC4B-8388210D3AA0}

Function Uninstall-MSI {


       param(
              [Parameter(Position=0, Mandatory=$true)] [String] $Server,
              [Parameter(Position=1, Mandatory=$true)] [String] $msiDisplayName,
              [Parameter(Position=2, Mandatory=$true)] [String] $msiDisplayVersion,
              [Parameter(Position=3, Mandatory=$false)] [String] $msitargetDir=$null)

       
       $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"

       $rt = new-object psobject -property @{status=$false;alreadydone=$false;errormsg=""}
       
       $RegistryLocation = 'Software\Microsoft\Windows\CurrentVersion\Uninstall'    
       $RemoteRegistry = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)    
       $IdentifyingNumber = $RemoteRegistry.OpenSubKey($RegistryLocation).GetSubKeyNames()|Where{$RemoteRegistry.OpenSubKey("$RegistryLocation\$_").GetValue('DisplayName') -match "$msiDisplayName"} 
       
       Switch($true)
       {
              ($IdentifyingNumber -eq $null)
              {
                     $rt.status = $true
                     $rt.alreadydone = $true
                     break
              }
              ($IdentifyingNumber -ne $null){
                     Write-Verbose "$RegistryLocation\$IdentifyingNumber"
                     $RegVersion = $RemoteRegistry.OpenSubKey("$RegistryLocation\$IdentifyingNumber").GetValue('DisplayVersion')
              }      
              ($RegVersion -ne $msiDisplayVersion)
              {
                     Write-Verbose "Another version of $msiDisplayName is installed: $RegVersion"
                     $rt.status = $true
                     $rt.alreadydone = $true
                     break
              }
              ($RegVersion -eq $msiDisplayVersion)
              {
                     Write-verbose "uninstall version $RegVersion"
                     $classKey="IdentifyingNumber=`"$IdentifyingNumber`",Name=`"$DisplayName`",version=`"$DisplayVersion`""
                     $res=([wmi]"\\$server\root\cimv2:Win32_Product.$classKey").Uninstall()
                     $rt.status=$true
              }
       }
       
       return $rt
}


#==============================================================================
# Get-WMIkey
#==============================================================================
#https://gallery.technet.microsoft.com/scriptcenter/WMI-Helper-Module-for-90e4f22e
#http://blogs.technet.com/b/heyscriptingguy/archive/2011/12/14/use-powershell-to-find-and-uninstall-software.aspx

function Get-WmiKey 
{  
  <# 
   .Synopsis 
    This function returns the key property of a WMI class 
   .Description 
    This function returns the key property of a WMI class 
   .Example 
    Get-WMIKey win32_bios 
    Returns the key properties for the Win32_bios WMI class in root\ciimv2 
   .Example 
    Get-WmiKey -class Win32_product 
    Returns the key properties for the Win32_Product WMI class in root\cimv2 
   .Example 
    Get-WmiKey -class systemrestore -namespace root\default 
    Gets the key property from the systemrestore WMI class in the root\default 
    WMI namespace.  
   .Parameter Class 
    The name of the WMI class 
   .Parameter Namespace 
    The name of the WMI namespace. Defaults to root\cimv2 
   .Parameter Computer 
    The name of the computer. Defaults to local computer 
   .Role 
    Meta 
   .Component 
    HSGWMIModuleV6  
   .Notes 
    NAME:  Get-WMIKey 
    AUTHOR: ed wilson, msft 
    LASTEDIT: 10/18/2011 17:38:20 
    KEYWORDS: Scripting Techniques, WMI 
    HSG: HSG-10-24-2011 
   .Link 
     Http://www.ScriptingGuys.com 
 #Requires -Version 2.0 
 #> 
 Param( 
   [Parameter(Mandatory = $true,Position = 0)] 
   [string]$class, 
   [string]$namespace = "root\cimv2", 
   [string]$computer = $env:computername 
 ) 
  [wmiclass]$class = "\\{0}\{1}:{2}" -f $computer,$namespace,$class 
  $class.Properties |  
      Select-object @{Name="PropertyName";Expression={$_.name}} -ExpandProperty Qualifiers | Where-object {$_.Name -eq "key"} | ForEach-Object {$_.PropertyName}
} #end GetWmiKey 
 
 


#==============================================================================
# Analyse An URL and returns in diffrent parts
#==============================================================================

Function Get-UrlDetails
    {Param([Parameter(Mandatory=$true,Position=0)][AllowEmptyString()] [string] $url)

    $urlreg = new-object System.Text.RegularExpressions.Regex "^((?<protocol>http[s]?|ftp):\/)?\/?(?<domain>[^:\/\s]+)(?<port>:([^\/]*))?(?<path>(\/\w+)*\/)?$" 

    if ($url.split("/").count -le 3 -and $url -ne "")  
       {$url = $url+"/"}

    try
        {$tmpauth = $urlreg.match($url+"/")
        
        if (!$tmpauth.success)
            {$urlreg = new-object System.Text.RegularExpressions.Regex "^((?<protocol>http[s]?|ftp):\/)?\/?(?<domain>[^:\/\s]+)(?<port>:([^\/]*))?" 
            $tmpauth = $urlreg.match($url+"/")}
         $fullpath = $tmpauth.groups['path'].value.tostring()
         $tmppath = ($fullpath.substring(1).split("/"))[0] 
         $fullpath = $fullpath.substring(1).substring(0,$fullpath.length-2)}
    catch
        {$tmppath=""}

    $port = $tmpauth.groups['port'].value
    if ($url -ne "" -and ($port -eq "" -or $port -eq $null))
        {$port=80}

    Return [PSCustomObject]@{
            Url=AgToLower($url)
            domain=$tmpauth.groups['domain'].value.tostring()
            Path=$tmppath
            FullPath=$fullpath
            port= $port
            protocol=$tmpauth.groups['protocol'].value}}

#==============================================================================
# Lower a string, when null return an empty string
#==============================================================================

Function AgToLower 
    {param([Parameter(Position=0, Mandatory=$true)][AllowEmptyString()] [String] $str)
     if ($str -eq $null) 
        {return ""}
    else
        {return $str.tolower()}}

#==============================================================================
# Validate if a string  is a valid  IP address
#==============================================================================

Function Test-IPAddress(){
       
       param(
       [Parameter(Position=0, Mandatory=$false)] [string]$IP)
       
       $IPAddress = $null
       return ([System.Net.IPAddress]::tryparse($IP,[ref]$IPAddress) -and $IP -eq $IPaddress.tostring())
}

#==============================================================================
# Validate if a string  is a valid  DNS target
#==============================================================================
function Test-dnsrecord {
       
       param(
              [Parameter(Position=1, Mandatory=$True)] [string]      $RecordName)

       $rt = new-object psobject -property @{record=$null;status=$false;alreadyexists=$false;errormsg=""}
       $res=$null
       #parameters validation
$MethodDefinition = @'
       [DllImport("dnsapi.dll", EntryPoint = "DnsFlushResolverCache")]
       private static extern UInt32 DnsFlushResolverCache();
       public static void FlushDNSCache() //flush dns cache
       {
       UInt32 result = DnsFlushResolverCache();

       }
'@
       if (-not ([System.Management.Automation.PSTypeName]'Win32.dnsapi').Type)
       {
           $r = Add-Type -MemberDefinition $MethodDefinition -Name 'dnscache' -Namespace 'AG' -PassThru
       }

       [AG.dnscache]::FlushDNSCache()
       
       Try    {
              $res = [System.Net.Dns]::GetHostByName($RecordName)
              }
       catch
             {
              $rt.errormsg=$_.Exception.InnerException.message
             $rt.status=$false
              }

       If($res -ne $null){
              $rt.alreadyexists=$true
              $rt.status=$true
              $rt.record=$res}
       Else{$rt.status=$false}

       return $rt
} 

#==============================================================================
# Get all http SPNs
#==============================================================================

function Get-HttpSpn
{
    $serviceType="HTTP"
    $filter = "(servicePrincipalName=$serviceType/*)"
    $domain = New-Object System.DirectoryServices.DirectoryEntry
    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    $searcher.SearchRoot = $domain
    $searcher.PageSize = 1000
    $searcher.Filter = $filter
    $results = $searcher.FindAll()

    foreach ($result in $results) {
        $account = $result.GetDirectoryEntry()
        foreach ($spn in $account.servicePrincipalName.Value) {
            
                     if($spn -match "^http\/(?<hostheader>[^:]+)[^:]*(:{1}(?<port>\w+))?$") {
                           new-object psobject -property @{hostheader=$matches.hostheader;Port=$matches.port;AccountName=$($account.sAMAccountName);SPN=$spn}
                           } 
        }
    }
}

#==============================================================================
# Create SPN
#==============================================================================

function New-Spn
{
    param(
    [Parameter(Position=0, Mandatory=$true)] [string] $UserPath,
       [Parameter(Position=1, Mandatory=$true)] [string] $SPNValue)
       
       
       $ADS_PROPERTY_CLEAR = 1
       $ADS_PROPERTY_UPDATE = 2
       $ADS_PROPERTY_APPEND = 3
       $ADS_PROPERTY_DELETE = 4
       
       $ADUser = [ADSI]$UserPath
       $ADUser.PutEx($ADS_PROPERTY_APPEND, "servicePrincipalName", @($SPNValue))
       $ADUser.SetInfo()
} 

#==============================================================================
#Excecute SQL Query
#==============================================================================
function ExecuteSQLquery()
{
       param 
       ([String]$SQLInstanceName, 
       [String]$SQLDatabaseName,
       [String]$SQLQuery)
       
       try{#"function ExcecuteQuery"
              $SqlConnection = New-Object System.Data.SqlClient.SqlConnection 
              $SqlConnection.ConnectionString = "Server=$SQLInstanceName;Database=$SQLDatabaseName;Integrated Security=True" 
              $SqlCmd = New-Object System.Data.SqlClient.SqlCommand 
              $SqlCmd.CommandText = $SQLQuery 
              $SqlCmd.Connection = $SqlConnection 
              $SqlCmd.CommandTimeout = 0 
              $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter 
              $SqlAdapter.SelectCommand = $SqlCmd 
              $DataSet = New-Object System.Data.DataSet 
              $SqlAdapter.Fill($DataSet)|Out-Null
              ,$DataSet}
       Catch{
              throw "Failure - Cannot execute SQLQuery: " + $Error[0].exception.message}
       finally{
           if ($SqlConnection.State -eq 'Open'){
              $SqlConnection.Close()|Out-Null}}
}


#==============================================================================
#Foreach parallel
#==============================================================================

function ForEach-Parallel {
    param(
        [Parameter(Mandatory=$true,position=0)]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [PSObject]$InputObject,
        [Parameter(Mandatory=$false)]
        [int]$MaxThreads=5
    )
    BEGIN {
        $iss = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $pool = [Runspacefactory]::CreateRunspacePool(1, $maxthreads, $iss, $host)
        $pool.open()
        $threads = @()
        $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param(`$_)`r`n" + $Scriptblock.ToString())
    }
    PROCESS {
        $powershell = [powershell]::Create().addscript($scriptblock).addargument($InputObject)
        $powershell.runspacepool=$pool
        $threads+= @{
            instance = $powershell
            handle = $powershell.begininvoke()
        }
    }
    END {
        $notdone = $true
        while ($notdone) {
            $notdone = $false
            for ($i=0; $i -lt $threads.count; $i++) {
                $thread = $threads[$i]
                if ($thread) {
                    if ($thread.handle.iscompleted) {
                        $thread.instance.endinvoke($thread.handle)
                        $thread.instance.dispose()
                        $threads[$i] = $null
                    }
                    else {
                        $notdone = $true
                    }
                }
            }
        }
    }
}

#==============================================================================
# Test if local server is a cluster node
#==============================================================================

Function Test-isCluster{
       param([string]$serverName)
       
       $s = Get-WmiObject -Class Win32_SystemServices -ComputerName $serverName
       if ($s | select PartComponent | where {$_ -like "*ClusSvc*"}) {return $true}
       else { return $false}
}


#==============================================================================
# Create Mount points for SQL instance 
#==============================================================================

function Create-SQLMountPoints
{
    param(
              [Parameter(Position=0, Mandatory=$true)] [string] $SQLInstanceID,
              [Parameter(Position=1, Mandatory=$true)] [string[]] $SQLDisks,
              [Parameter(Position=2, Mandatory=$true)] [string] $InstanceRootFolderPath,
              [Parameter(Position=3, Mandatory=$true)] [string] $RootLetter,
              [Parameter(Position=4, Mandatory=$false)] $computername = $env:computername
       )
       
       $rt = new-object psobject -property @{status=$false;alreadydone=$false;errormsg=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
       $InstanceUserDataFolderPath       = "\MSSQL\UserData\"
       $InstanceUserLogFolderPath        = "\MSSQL\UserLog\"
       $InstanceUserBackupFolderPath     = "\MSSQL\UserBackup\"
    $InstanceTempDataFolderPath   = "\MSSQL\TempData\"
       $InstanceTempLogFolderPath        = "\MSSQL\TempLog\"
    
       
       $RootPrefix                                     = "Root"
       $UserDataFolderPrefix                    = "UserData"
       $UserLogFolderPrefix                     = "UserLog"
    $UserBackupFolderPrefix              = "UserBck"
       $TempDataFolderPrefix                    = "TempData"
       $TempLogFolderPrefix              = "TempLog"
       
       $RootExists                                     = $False
       $UserdataExists                                 = $False
       $UserLogExists                                  = $False
    $UserBackupExists                           = $False
       $TempdataExists                                 = $False
       $TempLogExists                                  = $False
       
       
       $RootDiskName = $SQLDisks|Where{$_ -match "root"}
       Write-Verbose "RootDiskname: '$RootDiskName'"
       
       if($RootDiskName -eq $null){
              $rt.errormsg = "Root drive not specified"
       }
       Else{
              $RootDisk = Get-Volume -cimsession $computername|where{$_.FileSystemLabel -match "$SQLInstanceID_$rootdiskname"}
              if($RootDisk -eq $null){
                     $rt.errormsg = "Root drive($rootdiskname)not found on $computername"
              }
              Else{
                     if($RootDisk.DriveLetter -ne $RootLetter){
                           $rt.errormsg = "Drive letter of $rootdiskname ($RootDisk.DriveLetter) doesn't match drive letter specified in naming ($RootLetter)" 
                     }
                     
                     $SQLDiskNames = $SQLDisks|Where{$_ -notmatch "root"}
                     
                     
                     Foreach($SQLDiskName in $SQLDiskNames){
                           $TempFolderPath = $RootLetter + ":" + $InstanceRootFolderPath + "." + $SQLInstanceID +  "\MSSQL\" + $SQLDiskName.substring(0, $SQLDiskName.length -2) + "\" + $SQLDiskName
                           $TempVolumeLabel = $RootLetter.substring(0,1) + "_" + $SQLInstanceID + "_" + $SQLDiskName
                           Write-verbose "FolderPath: $tempFolderPath"
                           Write-verbose "VolumeLable: $TempVolumeLabel"
                           $res=Create-MountPoint -diskname "$($TempVolumeLabel)" -folderpath "$($tempFolderPath)" -computername $computername
                           if($res.status){
                                  $FilesFolder = "\\" + $computername + "\" +  ($TempFolderPath -replace ":","$") + "\files"
                                  Write-Verbose "Create Filesfolder - $FilesFolder"
                                  If(!(Test-Path $FilesFolder)){$FilesFolder = New-Item -ItemType Directory -Path $FilesFolder}
                                  $rt.status=$true}
                           Else{$rt.status=$false
                                  $rt.errormsg = $res.errormsg}
                     }
              }
       
       }
       return $rt
}

#==============================================================================
# Convert a file to the target Encoding with no BOM
# Possible value: unknown | string | unicode | bigendianunicode | utf8 | utf7 | utf32 | ascii | default | oem
# See https://technet.microsoft.com/en-us/library/hh849882.aspx
#==============================================================================

Function Convert-FileEncoding ()
{

       Param(
       [Parameter(Position=0,Mandatory=$True)] [String] $Path,
              [Parameter(Position=1,Mandatory=$True)] [String] $TargetEncoding = 'UTF8'
       )
       
       $rt = [PSCustomObject]@{status=$false;alreadydone=$false;errormsg=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
       if(!(Test-Path $Path)){
              $rt.errormsg = "File $path doesn't exist"
              return $rt
       }
       
       #Get current encoding if BOM implemented. If not BOM, file encoding can't be determined
       
       [byte[]]$byte = get-content -Encoding byte -ReadCount 4 -TotalCount 4 -Path $Path
       
       Switch($true)
       {
              ( $byte[0] -eq 0xef -and $byte[1] -eq 0xbb -and $byte[2] -eq 0xbf ) 
       {
                     Write-verbose 'UTF8'
                     $CurrentEncoding = 'UTF8'
                     break
              }
              ($byte[0] -eq 0xfe -and $byte[1] -eq 0xff) 
              {
                     Write-verbose 'Unicode'
                     $CurrentEncoding = 'Unicode'
                     break
              } 
       ($byte[0] -eq 0 -and $byte[1] -eq 0 -and $byte[2] -eq 0xfe -and $byte[3] -eq 0xff) 
              {
                     Write-verbose 'UTF32'                    
                     $CurrentEncoding = 'UTF32'
                     break
              } 
       ($byte[0] -eq 0x2b -and $byte[1] -eq 0x2f -and $byte[2] -eq 0x76)
           { 
                     Write-verbose 'UTF7'
                     $CurrentEncoding = 'UTF7'
                     break
              } 
              ($byte[0] -eq 0x55 -and $byte[1] -eq 0x53 -and $byte[2] -eq 0x45 -and $byte[3 -eq 0x20])
       {
                     Write-verbose 'ANSI'
                     $CurrentEncoding = 'ANSI'
                     break
              }
              ($true)
              {
                     Write-verbose "unknown"
                     $CurrentEncoding = "unknown"
              }
       
       }

       if($CurrentEncoding -eq $TargetEncoding)
       {
              Write-Verbose "Encoding is already ""$TargetEncoding"""
              $rt.status = $true
              $rt.alreadydone = $true
       }
       Else{
              
              Write-Verbose "Change Encoding from ""$CurrentEncoding"" to ""$TargetEncoding"""
              
              try{
                     $File = Get-Content $Path
                     $File | Out-File -Encoding $TargetEncoding $Path #With BOM
                     #$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($False)
                     #[System.IO.File]::WriteAllLines($Path, $File, $Utf8NoBomEncoding)
                     $rt.status = $true
              }
              catch{
                     $rt.errormsg = "$($error[0].Exception.Message)"
                     $Error.Clear()
                     return $rt
              }

       }
       
       return $rt
}

#==============================================================================
# Get-ServicesInfo
#==============================================================================       

function Get-ServicesInfo2
{
       
       Param(
              [Parameter(Position=0,Mandatory=$false)] [String] $ComputerName=$null,
              [Parameter(Position=1,Mandatory=$false)] [String] $WQLFilter=$null
       )
       
       $rt= [PSCustomObject]@{status=$false;alreadydone=$false;errormsg="";computername=$ComputerName;services=$null;CIMError="";CIMErrormsg="";WMIError="";WMIErrormsg=""}
       
       If($verbosePreference -eq "Continue"){Write-host "(Verbose Mode)"}
       Write-Verbose "-------------------------------------"
       Write-Verbose "Called $($MyInvocation.MyCommand)"
       $columnWidth = $PsBoundParameters.Keys.length | Sort-Object| Select-Object -Last 1
       $PsBoundParameters.GetEnumerator() | ForEach-Object {Write-Verbose ("  {0,-$columnWidth} : {1}" -F $_.Key, $_.Value)}
       Write-Verbose "-------------------------------------"
       
       $Error.Clear()
       
       $CIMfailed = $false
       
       if($ComputerName -eq $null)
       {
              $ComputerName = $env:computername
       }
       
       try{
              
              
              $WQLquery  = "Select DisplayName,name,exitcode,state,startname,PathName from win32_service"
              if($WQLFilter -ne $null){$WQLquery +=  " Where $WQLFilter"}
              
              $Date = Get-date
              Write-Verbose "CIM Query - $Date"
              
              $rt.services = Get-CimInstance -Query $WQLQuery -ComputerName $ComputerName -OperationTimeoutSec 10 -ErrorAction stop| select PSComputerName, DisplayName,name, exitcode, state, startname, PathName ,@{name='FileVersion';expression={([System.Diagnostics.FileVersionInfo]::GetVersionInfo("\\$ComputerName\" + "$($_.PathName)".split('"')[1].replace(':','$'))).ProductVersion}}
              $rt.status = $true


       }
       catch{
              $Date = Get-date
              Write-Verbose "CIM error - $Date"
              $CIMfailed = $true
              $rt.errormsg =  "Cannot connect the computer $ComputerName - CIM"
              $rt.CIMError = $error.Exception.HResult
              $rt.CIMErrormsg = $error.Exception.Message
              Write-verbose "issue: $($error[0].Exception.InnerException)"
              Write-Verbose "$($error[0].Exception.InnerException)"
              Write-Verbose "Fullname:""$($error[0].Exception.fullanme)"""
              Write-Verbose "Fullobject:""$($error[0].Exception.failedobject)"""
              $Error.Clear()
       }
       
       #if http SPN set on another (service account) try via WMI accelarator
       if($CIMfailed){
              try
              {
                     $WMIsearcher = [wmisearcher]$WQLQuery
                     
                     $WMIsearcher.options.timeout='0:0:10' #set timeout to 5 seconds
                     $ConnectionOptions = New-Object Management.ConnectionOptions
              $ManagementScope = New-Object Management.ManagementScope("\\$ComputerName\root\cimv2", $ConnectionOptions)
              $WMISearcher.Scope = $ManagementScope

                     
                     #$WMIsearcher.scope.path = "\\$ComputerName\root\cimv2" 
                     $Date = Get-date
                     Write-Verbose "WMI Query - $Date"
                     #PSComputer is empty when using $WMIsearcher
                     $rt.services=$WMIsearcher.get() | select @{name='PSComputerName';expression={$ComputerName}}, DisplayName,name, exitcode, state, startname, PathName,@{name='FileVersion';expression={([System.Diagnostics.FileVersionInfo]::GetVersionInfo("\\$ComputerName\" + "$($_.PathName)".split('"')[1].replace(':','$'))).ProductVersion}}
                     $rt.status = $true
                     $rt.errormsg =  ""
              }
              Catch{
                     $Date = Get-date
                     Write-Verbose "WMI error - $Date"
                     $rt.errormsg =  "Cannot connect the computer $ComputerName - WMISearcher"
                     $rt.WMIError = $error.Exception.HResult
                     $rt.WMIErrormsg = $error.Exception.Message
                     Write-verbose "issue: $($error[0].Exception.InnerException)"
                     Write-Verbose "$($error[0].Exception.InnerException)"
                     Write-Verbose "Fullname:""$($error[0].Exception.fullanme)"""
                     Write-Verbose "Fullobject:""$($error[0].Exception.failedobject)"""
                     $Error.Clear()
              }
       }
       
       return $rt

}