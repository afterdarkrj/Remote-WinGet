$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.IO.File]::ReadAllLines($path)

# Verify line 730 (0-indexed): should be '}'
Write-Host "Line 730: $($lines[730])"
# Verify line 731 (0-indexed): should be '<#'
Write-Host "Line 731: $($lines[731])"
# Verify line 1506 (0-indexed): should be '}'
Write-Host "Line 1506: $($lines[1506])"
# Verify line 1507 (0-indexed): should be empty or comment
Write-Host "Line 1507: $($lines[1507])"
Write-Host "Line 1509: $($lines[1509])"

if ($lines[730] -eq '}' -and $lines[731] -eq '<#' -and $lines[1506] -eq '}') {
    $cleanLines = [System.Collections.Generic.List[string]]::new()
    for ($i = 0; $i -le 730; $i++) {
        $cleanLines.Add($lines[$i])
    }
    for ($i = 1507; $i -lt $lines.Count; $i++) {
        $cleanLines.Add($lines[$i])
    }
    [System.IO.File]::WriteAllLines($path, $cleanLines, [System.Text.UTF8Encoding]::new($true))
    Write-Host "Cleaned successfully. Total lines: $($cleanLines.Count)"
} else {
    Write-Host "Verification failed! Not modifying file."
}
