### Preparation ###

#region Functions
Function UniversalDate {
   return Get-Date -f "dd/MM/yyyy HH:mm:ss.ffffff"
}

Function TechDate {
    return Get-Date -f "yyyy_MM_dd_HH_mm_ss.ffffff"
}

Function GetDate {
    return Get-Date
}

Function StartWatch {
    $Global:Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    Winfo "StopWatch started"
}

Function StopWatch {
    $Stopwatch.Stop()
    Winfo "StopWatch took $($Stopwatch.Elapsed.TotalSeconds) seconds to complete"
}
Function Winfo {
    param (
        [Parameter(Mandatory = $true,ValueFromPipeline = $true)]
        [string]$txt,
        [bool]$NewLine = $true
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
#endregion

# Location and stuff

Winfo "==================================================="
Winfo "Started $($PSCommandPath)"
Winfo "User: $env:USERNAME"
Winfo "Host: $env:COMPUTERNAME"
Winfo "==================================================="
cd E:\SCRIPTZ\PS\GetImmoWeb
$WatinAssemblies = ".\Watin\bin\net40\"

# Assemblies
Winfo "Importing Assemblies"
Add-Type -Path $WatiNAssemblies\WatiN.Core.dll -Verbose
Add-Type -Path $WatiNAssemblies\Interop.SHDocVw.dll -Verbose
Add-Type -Path $WatiNAssemblies\Microsoft.mshtml.dll

# Set script Props
$ErrorActionPreference = "Continue"
$VerbosePreference = "Continue"

$Props = @{
    DDG = "ddg.gg" 
    ImmoWeb = "https://www.immoweb.be/nl/zoek/appartement/te-koop/brussel/1000?minprice=75000&maxprice=260000&minroom=0&maxroom=3"
    SurfaceMin = 65
    SurfaceMax = 150
    EmailAddress = "arnovermeiren@hotmail.com"
    EmailBody = @()
}

# Kill IE and consorts if process running
Winfo "Stopping open IE instances"
Stop-Process -Name iexplore -ea SilentlyContinue

# Assembly Shortcuts
$Wfind = [WatiN.Core.Find]
$Wie = [WatiN.Core.IE]::new()

### Query Immoweb###

# Open Page
Winfo "Starting $($Props.ImmoWeb)"
$Wie.GoTo($Props.ImmoWeb)
$Wie.WaitForComplete()

# Maximize and Hide
Winfo "Sizing screens"
$Wie.ShowWindow("Maximize")
$Wie.WaitForComplete()
# $Wie.BringToFront()
$Wie.ShowWindow("Minimize")
$Wie.WaitForComplete()

# Find and fill in Values
Winfo "Finding and modifying criteria"
$tfSurfaceMin = $Wie.TextField($Wfind::ById("xsurfacehabitabletotale1"))
$Wie.WaitForComplete()
$tfSurfaceMax = $Wie.TextField($Wfind::ById("xsurfacehabitabletotale2"))
$Wie.WaitForComplete()
$tfSurfaceMin.Value = $Props.SurfaceMin
$Wie.WaitForComplete()
$tfSurfaceMax.Value = $Props.SurfaceMax
$Wie.WaitForComplete()
$tfSurfaceMin.Focus()
$Wie.WaitForComplete()
$Wie.PressTab()
$Wie.WaitForComplete()

# Refine result
Winfo "Filtering results"
$divVerfijn = $Wie.Div($Wfind::ByClass("btn-bleu-middle affiner"))
$Wie.WaitForComplete()
$lnkVerfijn = $divVerfijn.Link($Wfind::ByUrl("javascript:void(0);"))
$Wie.WaitForComplete()
$lnkVerfijn.Click()
$Wie.WaitForComplete()

# Sort result
Winfo "Apply sort"
$lstsort = $Wie.SelectList($Wfind::ByName("xorderby1"))
$Wie.WaitForComplete()
[WatiN.Core.OptionCollection]$opts = $lstsort.Options
$Wie.WaitForComplete()
$opts[1].Select()
$Wie.WaitForComplete()

# Get the List of Divs we want from result
$result = $Wie.Div($Wfind::ById("result"))
$data = $result.Div($Wfind::ByClass([regex]"\w-donnees-bien"))
$Wie.WaitForComplete()
$links = $result.Link($Wfind::By("target","IWEB-MAIN"))
$Wie.WaitForComplete()

$sb = [System.Text.StringBuilder]::new()
$ctr = 0

foreach ($elem in $data.Divs)
{
    [void]$sb.Append($elem.InnerHtml)
}


Set-Content -Path $PSScriptRoot\GetImmoWeb_$(TechDate).htm -Value $result.InnerHtml

Winfo "Closing"
$Wie.Close()