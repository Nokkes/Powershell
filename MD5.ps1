$files = Get-ChildItem -File -Recurse -LiteralPath "F:\VIDZ\MOVIEZ\Loro (2018)" | ? {$_.Extension -eq ".mkv"}
[System.Collections.ArrayList]$md5s = @()
$ctr = 0
$files | % {
    $md5temp = Get-FileHash -LiteralPath $_.FullName -Algorithm MD5
    $obj = [pscustomobject] @{
        "MD5Hash" = $md5temp.Hash;
        "Path" = $md5temp.Path}
    $md5s.Add($obj) | Out-Null
    $ctr++
    Write-Progress -Activity "Checking $($files.Count) Files" -Status "File $($ctr) of $($files.Count)" -PercentComplete ($ctr/$($files.Count)*100) -CurrentOperation "Checking $($_.FullName)"
}

Write-Host "Calculation done, sorting..."
$result = $md5s | Group-Object MD5hash | ? {$_.Count -gt 1} | % {$_.Group | ft}
$result | clip.exe
Write-Host "$($result.Count) double files found, results copied to clipboard!"