$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.Collections.Generic.List[string]]([System.IO.File]::ReadAllLines($path))

# 1. Update comment header description
for ($i = 10; $i -lt 25; $i++) {
    if ($lines[$i] -like '*COMPORTAMENTO PADRÃO:*') {
        $lines[$i+1] = '    Quando executado sem parâmetros (ou com -Help), o script exibe o manual de ajuda e exemplos.'
        $lines[$i+2] = '    Para consultar pacotes pendentes em modo seguro (sem alterar nada no alvo), basta informar'
        $lines[$i+3] = '    o nome da máquina (ex: .\Update-StationWinget.ps1 -ComputerName "estacao061").'
        Write-Host "Updated comment header at $i"
        break
    }
}

# 2. Update Show-ScriptHelp
for ($i = 650; $i -lt 680; $i++) {
    if ($lines[$i] -like '*COMPORTAMENTO PADRAO*') {
        $lines[$i]   = '    Write-Host "COMPORTAMENTO PADRAO:" -ForegroundColor Yellow'
        $lines[$i+1] = '    Write-Host "  Quando executado sem parametros (ou com -Help), o script exibe este manual completo."'
        $lines[$i+2] = '    Write-Host "  Para consultar pacotes pendentes em modo seguro (sem alterar nada no alvo), basta"'
        $lines[$i+3] = '    Write-Host "  informar o nome da maquina: .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer''"'
        Write-Host "Updated Show-ScriptHelp at $i"
        break
    }
}

# 3. Update the help check to include ($PSBoundParameters.Count -eq 0)
for ($i = 2970; $i -lt 3020; $i++) {
    if ($lines[$i] -like '*if ($Help)*') {
        $lines[$i] = 'if ($Help -or $PSBoundParameters.Count -eq 0) {'
        Write-Host "Updated Help condition at $i"
        break
    }
}

[System.IO.File]::WriteAllLines($path, $lines, [System.Text.UTF8Encoding]::new($true))
Write-Host "Script updated successfully. Total lines: $($lines.Count)"
