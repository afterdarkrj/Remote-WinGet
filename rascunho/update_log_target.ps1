$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.Collections.Generic.List[string]]([System.IO.File]::ReadAllLines($path))

# 1. Find line with '$executionFolder = if ($PSScriptRoot)' and insert $logFile right after the closing brace
for ($i = 2990; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -like '*$executionFolder = if ($PSScriptRoot)*') {
        $lines.Insert($i + 2, '$logFile = Join-Path -Path $executionFolder -ChildPath "output.log"')
        Write-Host "Inserted logFile definition at $($i + 2)"
        break
    }
}

# 2. In foreach ($currentHost in $targetComputers), remove $safeCompName and $logFile = Join-Path ...
for ($i = 3000; $i -lt 3060; $i++) {
    if ($lines[$i] -like '*$safeCompName = *') {
        Write-Host "Removing line: $($lines[$i])"
        $lines.RemoveAt($i)
        Write-Host "Removing line: $($lines[$i])"
        $lines.RemoveAt($i)
        break
    }
}

# 3. Add feedback after Remove-Item -LiteralPath $tempLog
for ($i = 3180; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -like '*Remove-Item -LiteralPath $tempLog*') {
        $lines.Insert($i + 1, '                Write-Host "  -> Log incremental atualizado em: $logFile" -ForegroundColor DarkGray')
        Write-Host "Inserted log feedback at $($i + 1)"
        break
    }
}

# 4. Add consolidated log message at batch end
for ($i = 3190; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -like '*PROCESSAMENTO EM LOTE CONCLU*') {
        $lines.Insert($i + 1, '    Write-Host " Log consolidado gravado em: $logFile" -ForegroundColor DarkGray')
        Write-Host "Inserted batch log notice at $($i + 1)"
        break
    }
}

[System.IO.File]::WriteAllLines($path, $lines, [System.Text.UTF8Encoding]::new($true))
Write-Host "Script successfully updated. Total lines: $($lines.Count)"
