$errors = $null
$tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile("C:\Temp\Update-StationWinget.ps1", [ref]$tokens, [ref]$errors)
if ($errors.Count -eq 0) {
    Write-Host "PARSER: SUCCESS (0 errors)"
} else {
    Write-Host "PARSER ERRORS:"
    $errors | ForEach-Object { Write-Host "$($_.Extent.StartLineNumber): $($_.Message)" }
}
