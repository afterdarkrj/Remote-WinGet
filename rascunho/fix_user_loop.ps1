$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.IO.File]::ReadAllLines($path)

# Let's find line index where $OnlyList ends:
# "            if ($OnlyList) {"
# ...
# "                continue"
# "            }"
# And the next line is:
# "                    $errSnippet = if ($userUpRes.Output.Length -gt 800) {"

$cutStartIndex = -1
$cutEndIndex = -1

for ($i = 500; $i -lt 620; $i++) {
    if ($lines[$i] -like "*if (`$OnlyList)*") {
        # find the closing brace of if ($OnlyList)
        for ($j = $i; $j -lt $i + 15; $j++) {
            if ($lines[$j] -match '^\s*continue\s*$') {
                $cutStartIndex = $j + 2 # right after closing brace
                break
            }
        }
    }
    if ($lines[$i] -like "*`$errSnippet = if (`$userUpRes.Output.Length*") {
        $cutEndIndex = $i - 1
        break
    }
}

Write-Host "cutStartIndex: $cutStartIndex ($($lines[$cutStartIndex]))"
Write-Host "cutEndIndex: $cutEndIndex ($($lines[$cutEndIndex]))"

$blockToInsert = @(
'            # Executa atualizacao no escopo de usuario',
'            Write-Host "  -> Atualizando [$u]: $($pkg.Name) ($($pkg.Id))..." -ForegroundColor Cyan',
'            $cleanPkgId = $pkg.Id -replace ''[^a-zA-Z0-9_]'',''''',
'            if ($cleanPkgId.Length -gt 20) { $cleanPkgId = $cleanPkgId.Substring(0, 20) }',
'            $userTempUpFile = "C:\Windows\Temp\winget_u_up_${cleanU}_${cleanPkgId}_$PID.txt"',
'            $userUpRes = Invoke-WingetTask -TaskName "W_Up_U_${cleanU}_${cleanPkgId}" `',
'                                          -UserId $u `',
'                                          -IsInteractive $true `',
'                                          -Arguments "upgrade --id `"$($pkg.Id)`" --scope user --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity" `',
'                                          -OutputFile $userTempUpFile `',
'                                          -MaxWaitSeconds $Timeout',
'',
'            $st = switch ($userUpRes.ExitCode) {',
'                0           { "Atualizado" }',
'                3010        { "Requer reboot" }',
'                -1978335189 { "Ja atualizado" }',
'                1460        { "Falhou (Tempo Limite Excedido: ${Timeout}s)" }',
'                default     { "Falhou (Codigo: $($userUpRes.ExitCode))" }',
'            }',
'',
'            if ($userUpRes.ExitCode -eq 0) {',
'                Write-Host "     [OK] Concluido com sucesso." -ForegroundColor Green',
'            } elseif ($userUpRes.ExitCode -eq 3010) {',
'                Write-Host "     [!] Concluido (requer reinicializacao)." -ForegroundColor Yellow',
'            } elseif ($userUpRes.ExitCode -eq -1978335189) {',
'                Write-Host "     [=] Ja esta na versao mais recente." -ForegroundColor Gray',
'            } else {',
'                Write-Host "     [FALHOU] $($st)" -ForegroundColor Red',
'            }',
'',
'            $errSnippet = ""',
'            if ($userUpRes.ExitCode -ne 0 -and $userUpRes.ExitCode -ne 3010 -and $userUpRes.ExitCode -ne -1978335189) {',
'                if ($userUpRes.Output) {'
)

$newLines = @()
$newLines += $lines[0..($cutStartIndex - 1)]
$newLines += $blockToInsert
$newLines += $lines[($cutEndIndex + 1)..($lines.Count - 1)]

[System.IO.File]::WriteAllLines($path, $newLines, [System.Text.UTF8Encoding]::new($true))
Write-Host "Done writing updated file. Total lines: $($newLines.Count)"
