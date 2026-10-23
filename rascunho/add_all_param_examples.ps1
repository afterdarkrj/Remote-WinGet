$path = 'C:\Temp\Update-StationWinget.ps1'
$lines = [System.Collections.Generic.List[string]]([System.IO.File]::ReadAllLines($path))

# 1. Update .EXAMPLE block (between .EXAMPLE at index ~50 and param( at index ~85)
$exampleStart = -1
$exampleEnd = -1

for ($i = 40; $i -lt 100; $i++) {
    if ($lines[$i] -match '^\.EXAMPLE\b' -and $exampleStart -eq -1) {
        $exampleStart = $i
    }
    if ($lines[$i] -match '^param\(') {
        $exampleEnd = $i - 1
        # Also need to make sure the closing '#>' is before param(
        while ($exampleEnd -gt $exampleStart -and $lines[$exampleEnd] -notmatch '^\s*#>') {
            $exampleEnd--
        }
        break
    }
}

Write-Host "Found .EXAMPLE block from line $exampleStart to $exampleEnd"

$newExamplesHeader = @(
'.EXAMPLE',
'    # 1. -ComputerName: Consultar maquina unica em modo seguro',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061"',
'',
'.EXAMPLE',
'    # 2. -ComputerName posicional: Consultar passando diretamente o nome',
'    .\Update-StationWinget.ps1 "estacao061"',
'',
'.EXAMPLE',
'    # 3. -ComputerName com array: Consultar multiplas maquinas',
'    .\Update-StationWinget.ps1 -ComputerName @("estacao061", "estacao062")',
'',
'.EXAMPLE',
'    # 4. -ComputerFile: Processar lista de computadores a partir de arquivo de texto',
'    .\Update-StationWinget.ps1 -ComputerFile "C:\Temp\maquinas.txt"',
'',
'.EXAMPLE',
'    # 5. -Update: Atualizar de fato todos os programas pendentes silenciosamente',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update',
'',
'.EXAMPLE',
'    # 6. -PackageId: Consultar ou atualizar apenas um aplicativo especifico (ID exato ou wildcard)',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -PackageId "VideoLAN.VLC"',
'',
'.EXAMPLE',
'    # 7. -InstallPackage: Instalar um novo aplicativo pelo seu ID do Winget',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -InstallPackage "7zip.7zip"',
'',
'.EXAMPLE',
'    # 8. -ExcludePackages: Atualizar aplicando lista de excecoes',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Microsoft.Teams.Classic", "*Dell*")',
'',
'.EXAMPLE',
'    # 9. -IncludeSkipped: Exibir na tabela final os pacotes ignorados por excecao',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Zoom.Zoom") -IncludeSkipped',
'',
'.EXAMPLE',
'    # 10. -TaskTimeoutSeconds: Atualizar com tempo limite estendido (ex: 1800s / 30 min)',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -TaskTimeoutSeconds 1800',
'',
'.EXAMPLE',
'    # 11. -Credential: Fornecer credenciais administrativas explicitamente',
'    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Credential (Get-Credential) -Update',
'',
'.EXAMPLE',
'    # 12. -PassThru: Exportar objetos para pipeline (CSV ou Out-GridView)',
'    .\Update-StationWinget.ps1 -ComputerFile "C:\Temp\maquinas.txt" -PassThru | Export-Csv -Path "C:\Temp\relatorio.csv" -NoTypeInformation -Encoding UTF8',
'',
'.EXAMPLE',
'    # 13. -Help: Exibir o guia completo de parametros e exemplos',
'    .\Update-StationWinget.ps1 -Help',
'',
'.EXAMPLE',
'    # 14. Mega Exemplo Combinado: Arquivo + Update + Excecoes + Timeout + IncludeSkipped + PassThru',
'    .\Update-StationWinget.ps1 -ComputerFile "C:\Temp\maquinas.txt" -Update -ExcludePackages @("Teams*") -TaskTimeoutSeconds 1200 -IncludeSkipped -PassThru | Export-Csv -Path "C:\Temp\resultado.csv" -NoTypeInformation -Encoding UTF8',
'#>'
)

# Replace lines $exampleStart to $exampleEnd with $newExamplesHeader
$lines.RemoveRange($exampleStart, ($exampleEnd - $exampleStart + 1))
$lines.InsertRange($exampleStart, $newExamplesHeader)

# 2. Update Show-ScriptHelp examples
$helpExStart = -1
$helpExEnd = -1

for ($i = 650; $i -lt 850; $i++) {
    if ($lines[$i] -like '*EXEMPLOS PRATICOS DE USO*') {
        $helpExStart = $i
    }
    if ($helpExStart -gt -1 -and $lines[$i] -match '^}\s*$' -and $lines[$i+1] -match '^\s*#\s*=') {
        $helpExEnd = $i - 1
        break
    }
}

Write-Host "Found Show-ScriptHelp examples block from line $helpExStart to $helpExEnd"

$newHelpExamples = @(
'    Write-Host "EXEMPLOS PRATICOS DE USO (COBRINDO TODOS OS PARAMETROS):" -ForegroundColor Yellow',
'    Write-Host ""',
'    Write-Host "  --- 1. PARAMETRO: -ComputerName (Maquina unica, posicional ou multiplas) ---" -ForegroundColor DarkCyan',
'    Write-Host "  1.1 Consultar maquina unica (modo seguro / somente listagem):" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  1.2 Consultar passando o nome da maquina diretamente por posicao (atalho):" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 ''$TargetComputer''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  1.3 Consultar multiplas maquinas informando uma lista no proprio comando:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName @(''estacao061'', ''estacao062'', ''estacao063'')" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 2. PARAMETRO: -ComputerFile (Lote via arquivo de texto) ---" -ForegroundColor DarkCyan',
'    Write-Host "  2.1 Consultar pendencias em lote lendo lista de maquinas em arquivo .txt:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  2.2 Atualizar em lote todas as maquinas listadas no arquivo:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt'' -Update" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 3. PARAMETRO: -Update (Aplicacao efetiva de atualizacoes) ---" -ForegroundColor DarkCyan',
'    Write-Host "  3.1 Atualizar TODOS os programas pendentes silenciosamente na maquina:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 4. PARAMETRO: -PackageId (Filtrar aplicativo especifico ou wildcard) ---" -ForegroundColor DarkCyan',
'    Write-Host "  4.1 Consultar apenas um aplicativo especifico pelo ID exato:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -PackageId ''VideoLAN.VLC''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  4.2 Atualizar apenas aplicativos que coincidam com um padrao (Wildcard *):" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -PackageId ''*Chrome*''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 5. PARAMETRO: -InstallPackage (Instalar novo aplicativo) ---" -ForegroundColor DarkCyan',
'    Write-Host "  5.1 Instalar um NOVO aplicativo na estacao remota pelo seu ID do Winget:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -InstallPackage ''7zip.7zip''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  5.2 Instalar aplicativo em lote para todas as maquinas de um arquivo:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt'' -InstallPackage ''Git.Git''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 6. PARAMETRO: -ExcludePackages (Lista de excecoes a ignorar) ---" -ForegroundColor DarkCyan',
'    Write-Host "  6.1 Atualizar ignorando programas com problemas conhecidos ou restricoes:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -ExcludePackages @(''Microsoft.Teams.Classic'', ''Oracle.JavaRuntimeEnvironment'', ''*Dell*'')" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 7. PARAMETRO: -IncludeSkipped (Exibir itens ignorados na tabela) ---" -ForegroundColor DarkCyan',
'    Write-Host "  7.1 Atualizar aplicando excecoes e exibindo os pacotes ignorados no resumo final:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -ExcludePackages @(''Zoom.Zoom'') -IncludeSkipped" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 8. PARAMETRO: -TaskTimeoutSeconds (Tempo limite customizado anti-travamento) ---" -ForegroundColor DarkCyan',
'    Write-Host "  8.1 Atualizar programas pesados (ex: Visual Studio / CAD) com timeout de 30 minutos (1800s):" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Update -TaskTimeoutSeconds 1800" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 9. PARAMETRO: -Credential (Autenticacao administrativa explicita) ---" -ForegroundColor DarkCyan',
'    Write-Host "  9.1 Executar solicitando credenciais administrativas explicitas (Get-Credential):" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -Credential (Get-Credential) -Update" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 10. PARAMETRO: -PassThru (Integracao com Pipeline do PowerShell) ---" -ForegroundColor DarkCyan',
'    Write-Host "  10.1 Exportar relatorio tabular de pendencias para arquivo CSV unificado:" -ForegroundColor Cyan',
'    Write-Host "       .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt'' -PassThru | Export-Csv -Path ''C:\Temp\relatorio_geral.csv'' -NoTypeInformation -Encoding UTF8" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  10.2 Exibir resultado em grade interativa do PowerShell (Out-GridView):" -ForegroundColor Cyan',
'    Write-Host "       .\Update-StationWinget.ps1 -ComputerName ''$TargetComputer'' -PassThru | Out-GridView -Title ''Atualizacoes Pendentes''" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- 11. PARAMETRO: -Help (Exibir este manual de ajuda) ---" -ForegroundColor DarkCyan',
'    Write-Host "  11.1 Exibir este manual completo (ou executando o script sem nenhum parametro):" -ForegroundColor Cyan',
'    Write-Host "       .\Update-StationWinget.ps1 -Help" -ForegroundColor White',
'    Write-Host ""',
'    Write-Host "  --- COMBINACAO AVANCADA COMPLETA (Mega Exemplo com multiplos parametros) ---" -ForegroundColor DarkCyan',
'    Write-Host "  12. Lote com arquivo, atualizacao, excecoes, timeout customizado, ignorados visiveis e CSV:" -ForegroundColor Cyan',
'    Write-Host "      .\Update-StationWinget.ps1 -ComputerFile ''C:\Temp\maquinas.txt'' -Update -ExcludePackages @(''Microsoft.Teams*'', ''Veeam*'') -TaskTimeoutSeconds 1200 -IncludeSkipped -PassThru | Export-Csv -Path ''C:\Temp\resultado_lote.csv'' -NoTypeInformation -Encoding UTF8" -ForegroundColor White',
'    Write-Host "================================================================================" -ForegroundColor Cyan'
)

$lines.RemoveRange($helpExStart, ($helpExEnd - $helpExStart + 1))
$lines.InsertRange($helpExStart, $newHelpExamples)

[System.IO.File]::WriteAllLines($path, $lines, [System.Text.UTF8Encoding]::new($true))
Write-Host "Successfully replaced examples in both sections. Total lines: $($lines.Count)"
