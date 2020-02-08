sl C:\Scripts\PS
. .\Memory-Tools.ps1


try {
    $proc = [System.Diagnostics.Process]::GetProcessById(12616)
    $module = $proc.MainModule
    $size = $module.ModuleMemorySize
    $base = $module.BaseAddress
    $secret_mult = $module.Size * 4096
    Check-MemoryProtection -Address $base
    Dump-Memory $base $size -DumpToFile .\out.exe

}
catch {
    $Error
    Dump-Memory $base $size -DumpToFile .\out.exe
}