$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.Collections.Generic.List[string]]([System.IO.File]::ReadAllLines($path))

# 1. Remove duplicate header line around 707
for ($i = 700; $i -lt 715; $i++) {
    if ($lines[$i] -eq '    Write-Host "EXEMPLOS PRATICOS DE USO:" -ForegroundColor Yellow' -and $lines[$i+1] -like '*COBRINDO TODOS OS PARAMETROS*') {
        $lines.RemoveAt($i)
        Write-Host "Removed duplicate line at $i"
        break
    }
}

# 2. Add '}' after the closing Cyan line
for ($i = 770; $i -lt 785; $i++) {
    if ($lines[$i] -like '*================================================================================" -ForegroundColor Cyan*' -and $lines[$i+1] -notmatch '^\s*\}\s*$') {
        $lines.Insert($i + 1, '}')
        Write-Host "Added closing brace at $($i + 1)"
        break
    }
}

[System.IO.File]::WriteAllLines($path, $lines, [System.Text.UTF8Encoding]::new($true))
Write-Host "File saved successfully."
