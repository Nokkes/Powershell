$arrInts = @(); 48..57 | % {$arrInts+=[char]$_}
$arrAlpabet1 = @(); 65..90  | % {$arrAlpabet1+=[char]$_}
$arrAlpabet2 = @(); 97..122 | % {$arrAlpabet2+=[char]$_}

$arrInts+=$arrAlpabet1+=$arrAlpabet2
1..50 | % {($arrInts | Get-Random -Count 54) -join ''}