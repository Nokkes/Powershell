Function QueryDB {
            Param([Parameter(Mandatory=$true)][string]$query,
                  [Parameter(Mandatory=$false)][array]$parameters=$null,
                  [Parameter(Mandatory=$true)][string]$connectionstring,
                  [Parameter(Mandatory=$false)][switch]$NeedResult,
                  [Parameter(Mandatory=$false)][switch]$DumpError,
                  [Parameter(Mandatory=$false)][string]$DumpFile)

            $rt = @{query=$query;parameters=$parameters;connectionstring=$connectionstring;result=$null;status=$false;errmsg=""}
            try {
                 $Connection = New-Object System.Data.SQLClient.SQLConnection
                 $Connection.ConnectionString = $connectionstring
                 $n=$Connection.Open()
                 $Command = New-Object System.Data.SQLClient.SQLCommand
                 $Command.Connection = $Connection
                 $Command.CommandText = $query

                 if ($parameters -ne $null) {
                    foreach ($para in $parameters) {
                        if ($para.value -ne $null) {
                            $parobj = New-Object System.Data.SQLClient.SqlParameter
                            $parobj.ParameterName = $para.name
                            if ($para.value.GetType().name -eq 'PSCustomObject') {
                                $parobj.value = $para.value.tostring()}
                            else {
                                $parobj.value = $para.value}
                            $command.Parameters.Add($parobj) }}}

                 if ($NeedResult.ispresent)  {
                    $reader=$Command.ExecuteReader()
                    $datatable= new-object System.Data.DataTable
                    $datatable.load($reader)
                    $rt.result = $datatable.select()}
                else {
                    $rt.result=$Command.ExecuteScalar() }
                $rt.status=$true }
            catch {
                Log "Problem updating database $($_.Exception.Message)"
                $rt.status=$false
                $rt.errmsg=$_.Exception.Message
                if ($DumpError.ispresent) {
                    $rt | export-Clixml "$DumpFile.failed" -Force
                    }                
                }
            finally
                {if ($Command -ne $null)
                    {$Command.Dispose()
                     $Command=$null}
                 if ($Connection -ne $null)
                    {$Connection.Close()
                     $Connection.Dispose()
                     $connection=$null}}
        return $rt}
                     
function TrimLower($s)
    {if ($s -eq $null) 
        {return ""}
     else
        {return $s.tolower().trim()}}

function Convert-DateString ([String]$Date, [String[]]$Format)
    {$result = New-Object DateTime    
    try
           { $convertible = [DateTime]::TryParseExact($Date,
                $Format,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$result) 
                
          return $result }
    catch
           {Log "Problem with Converting date, using default value"
         Log $_.Exception.Message
         return Get-Date -Year 2014 -Month 1 -Day 1}} 