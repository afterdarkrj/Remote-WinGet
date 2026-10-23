$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.IO.File]::ReadAllLines($path)

$exampleHeader = @(
'.EXAMPLE',
'    # 1. -ComputerName: Definir computador alvo para consulta',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061"',
'',
'.EXAMPLE',
'    # 2. -ComputerFile: Definir arquivo com lista de maquinas a processar',
'    .\Update-StationWinget.ps1 -ComputerFile "C:\Temp\maquinas.txt"',
'',
'.EXAMPLE',
'    # 3. -Update: Executar atualizacoes de fato de forma silenciosa',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update',
'',
'.EXAMPLE',
'    # 4. -PackageId: Filtrar aplicativo especifico por ID ou wildcard',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -PackageId "VideoLAN.VLC"',
'',
'.EXAMPLE',
'    # 5. -InstallPackage: Instalar um novo aplicativo pelo seu ID do Winget',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -InstallPackage "7zip.7zip"',
'',
'.EXAMPLE',
'    # 6. -ExcludePackages: Ignorar programas especificos durante a atualizacao',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Microsoft.Teams.Classic", "*Dell*")',
'',
'.EXAMPLE',
'    # 7. -IncludeSkipped: Exibir os pacotes ignorados por excecao na tabela de resumo',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Zoom.Zoom") -IncludeSkipped',
'',
'.EXAMPLE',
'    # 8. -TaskTimeoutSeconds: Definir tempo limite customizado em segundos (ex: 1800s)',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -TaskTimeoutSeconds 1800',
'',
'.EXAMPLE',
'    # 9. -Credential: Fornecer credenciais administrativas explicitas',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Credential (Get-Credential) -Update',
'',
'.EXAMPLE',
'    # 10. -PassThru: Retornar objetos no pipeline para exportacao em CSV',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -PassThru | Export-Csv -Path "C:\Temp\pendentes.csv" -NoTypeInformation -Encoding UTF8',
'',
'.EXAMPLE',
'    # 11. -Help: Exibir manual completo de parametros e instrucoes',
'    .\Update-StationWinget.ps1 -Help',
'#>'
)

$helpExamples = @(
'    Write-Host "EXEMPLOS PRATICOS DE USO (1 EXEMPLO PARA CADA PARAMETRO):" -ForegroundColor Yellow',
'    Write-Host ""',
'    Write-Host "  1. Parametro -ComputerName (Definir computador alvo):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  2. Parametro -ComputerFile (Definir arquivo com lista de maquinas):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  3. Parametro -Update (Executar atualizacoes de fato):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  4. Parametro -PackageId (Filtrar aplicativo especifico):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -PackageId ''VideoLAN.VLC''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  5. Parametro -InstallPackage (Instalar novo aplicativo):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -InstallPackage ''7zip.7zip''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  6. Parametro -ExcludePackages (Definir excecoes a ignorar):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -ExcludePackages @(''Microsoft.Teams.Classic'', ''*Dell*'')" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  7. Parametro -IncludeSkipped (Exibir pacotes ignorados na tabela de resumo):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -ExcludePackages @(''Zoom.Zoom'') -IncludeSkipped" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  8. Parametro -TaskTimeoutSeconds (Definir tempo limite customizado em segundos):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -TaskTimeoutSeconds 1800" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  9. Parametro -Credential (Fornecer credenciais administrativas explicitas):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Credential (Get-Credential) -Update" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host " 10. Parametro -PassThru (Exportar objetos para pipeline / CSV):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -PassThru | Export-Csv -Path ''C:\Temp\pendentes.csv'' -NoTypeInformation -Encoding UTF8" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host " 11. Parametro -Help (Exibir manual de instrucoes e parametros):" -ForegroundColor Cyan',
'    Write-Host "     .\Update-StationWinget.ps1 -Help" -ForegroundColor White',
'    Write-Host "================================================================================" -ForegroundColor Cyan',
'}'
)

# Find boundary 1: index of first '.EXAMPLE' up to index of '[CmdletBinding()]'
$exStart = -1
$exEnd = -1
for ($i = 0; $i -lt 150; $i++) {
    if ($lines[$i] -match '^\.EXAMPLE\b' -and $exStart -eq -1) {
        $exStart = $i
    }
    if ($lines[$i] -match '^\[CmdletBinding\(\)\]') {
        $exEnd = $i - 1
        break
    }
}

# Find boundary 2: index of 'Write-Host "EXEMPLOS PRATICOS DE USO' up to the closing '}' of Show-ScriptHelp
$helpStart = -1
$helpEnd = -1
for ($i = 600; $i -lt 900; $i++) {
    if ($lines[$i] -like '*Write-Host "EXEMPLOS PRATICOS DE USO*' -and $helpStart -eq -1) {
        $helpStart = $i
    }
    if ($helpStart -gt -1 -and $lines[$i] -match '^\s*\}\s*$' -and $lines[$i+1] -match '^\s*#\s*=') {
        $helpEnd = $i
        break
    }
}

Write-Host "Boundary 1: $exStart to $exEnd"
Write-Host "Boundary 2: $helpStart to $helpEnd"

$finalLines = [System.Collections.Generic.List[string]]::new()
for ($i = 0; $i -lt $exStart; $i++) { $finalLines.Add($lines[$i]) }
foreach ($l in $exampleHeader) { $finalLines.Add($l) }
for ($i = ($exEnd + 1); $i -lt $helpStart; $i++) { $finalLines.Add($lines[$i]) }
foreach ($l in $helpExamples) { $finalLines.Add($l) }
for ($i = ($helpEnd + 1); $i -lt $lines.Count; $i++) { $finalLines.Add($lines[$i]) }

[System.IO.File]::WriteAllLines($path, $finalLines, [System.Text.UTF8Encoding]::new($true))
Write-Host "Saved successfully. Total lines: $($finalLines.Count)"
