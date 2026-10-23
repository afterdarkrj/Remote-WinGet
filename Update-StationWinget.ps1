<#
.SYNOPSIS
    Gerencia, consulta, atualiza e instala programas remotamente via WinRM usando Winget.

.DESCRIPTION
    Este script é executado a partir de uma estação administrativa e conecta na máquina alvo (por padrão estacao061).
    Ele contorna restrições de sessão de rede do WinRM e ConstrainedLanguage mode utilizando tarefas agendadas locais temporárias:
      - Escopo de Máquina: Executa via tarefa sob o contexto NT AUTHORITY\SYSTEM.
      - Escopo de Usuário: Detecta usuários logados ativamente e dispara tarefas interativas ocultas no contexto de cada usuário.
    Garante silêncio absoluto para os usuários (janelas ocultas e parâmetros silenciosos do winget).

    COMPORTAMENTO PADRÃO:
    Quando executado apenas com o nome da máquina (sem o parâmetro -Update e sem -InstallPackage), o script opera em
    modo SOMENTE LISTAGEM, apresentando os pacotes que possuem atualizações pendentes em formato de tabela horizontal
    lado a lado, sem alterar nada.

    Para aplicar atualizações, utilize o parâmetro -Update.
    Para instalar um novo pacote, utilize o parâmetro -InstallPackage.

.PARAMETER ComputerName
    Nome ou IP do computador remoto (Padrão: 'estacao061'). Suporta também uma lista/array de nomes.

.PARAMETER ComputerFile
    Caminho para arquivo de texto (.txt) contendo lista de computadores a serem processados (um por linha).

.PARAMETER Credential
    Credenciais administrativas opcionais para conexão remota.

.PARAMETER Update
    Aplica as atualizações de fato. Se omitido, o script apenas lista os pacotes pendentes (modo seguro de consulta).

.PARAMETER PackageId
    Especifica um único aplicativo para consultar ou atualizar (suporta ID, Nome ou wildcard, ex: 'VideoLAN.VLC', 'Google.Chrome', '*Chrome*').

.PARAMETER InstallPackage
    Instala um pacote específico na máquina de destino pelo seu ID no catálogo Winget (ex: '7zip.7zip', 'Git.Git').

.PARAMETER ExcludePackages
    Lista de IDs ou Nomes de pacotes a ignorar durante a atualização (suporta wildcards, ex: 'Google.Chrome', '*Java*').

.PARAMETER IncludeSkipped
    Se especificado, inclui na tabela final os pacotes ignorados pela lista de exceções.

.PARAMETER PassThru
    Retorna os objetos no pipeline para permitir exportação (ex: Export-Csv) ou manipulação posterior.

.PARAMETER TaskTimeoutSeconds
    Tempo limite em segundos para cada operação de atualização/instalação (Padrão: 900s).

.EXAMPLE
    # 1. -ComputerName: Definir computador alvo para consulta
    .\Update-StationWinget.ps1 -ComputerName "estacao061"

.EXAMPLE
    # 2. -ComputerFile: Definir arquivo com lista de maquinas a processar
    .\Update-StationWinget.ps1 -ComputerFile "C:\Temp\maquinas.txt"

.EXAMPLE
    # 3. -Update: Executar atualizacoes de fato de forma silenciosa
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update

.EXAMPLE
    # 4. -PackageId: Filtrar aplicativo especifico por ID ou wildcard
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -PackageId "VideoLAN.VLC"

.EXAMPLE
    # 5. -InstallPackage: Instalar um novo aplicativo pelo seu ID do Winget
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -InstallPackage "7zip.7zip"

.EXAMPLE
    # 6. -ExcludePackages: Ignorar programas especificos durante a atualizacao
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Microsoft.Teams.Classic", "*Dell*")

.EXAMPLE
    # 7. -IncludeSkipped: Exibir os pacotes ignorados por excecao na tabela de resumo
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -ExcludePackages @("Zoom.Zoom") -IncludeSkipped

.EXAMPLE
    # 8. -TaskTimeoutSeconds: Definir tempo limite customizado em segundos (ex: 1800s)
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Update -TaskTimeoutSeconds 1800

.EXAMPLE
    # 9. -Credential: Fornecer credenciais administrativas explicitas
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -Credential (Get-Credential) -Update

.EXAMPLE
    # 10. -PassThru: Retornar objetos no pipeline para exportacao em CSV
    .\Update-StationWinget.ps1 -ComputerName "estacao061" -PassThru | Export-Csv -Path "C:\Temp\pendentes.csv" -NoTypeInformation -Encoding UTF8

.EXAMPLE
    # 11. -Help: Exibir manual completo de parametros e instrucoes
    .\Update-StationWinget.ps1 -Help
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]]$ComputerName = @("estacao061"),

    [Parameter()]
    [Alias("ComputerList", "ComputersFile", "Lista", "ArquivoMaquinas", "InputFile")]
    [string]$ComputerFile,

    [Parameter()]
    [PSCredential]$Credential,

    [Parameter()]
    [Alias("Upgrade", "Apply")]
    [switch]$Update,

    [Parameter()]
    [Alias("Package", "Id", "App")]
    [string]$PackageId,

    [Parameter()]
    [Alias("Install", "InstallApp")]
    [string]$InstallPackage,

    [Parameter()]
    [string[]]$ExcludePackages = @(),

    [Parameter()]
    [switch]$IncludeSkipped,

    [Parameter()]
    [switch]$PassThru,

    [Parameter()]
    [int]$TaskTimeoutSeconds = 900,

    [Parameter()]
    [Alias("Ajuda", "Examples", "Exemplos")]
    [switch]$Help
)

# Bloco de script que será injetado e executado na máquina de destino
$RemoteScriptBlock = {
    param(
        [string[]]$ExcludeList,
        [int]$Timeout,
        [bool]$OnlyList,
        [string]$TargetPackage,
        [string]$InstallId
    )

    # 1. Localizar o executável do winget em C:\Program Files\WindowsApps
    $appx = Get-ChildItem -Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $appx) {
        return @{
            Results        = @(@{
                Escopo           = "Geral"
                Programa         = "Winget"
                ID               = "Microsoft.DesktopAppInstaller"
                VersaoAtual      = "N/A"
                VersaoDisponivel = "N/A"
                Status           = "Erro: Winget não localizado no sistema de destino."
            })
            MachinePending = 0
            LoggedUsers    = @()
        }
    }
    $wingetPath = $appx.FullName

    # Função interna para executar comandos silenciosamente via Tarefa Agendada local
    # (Essencial para WinRM e para o contexto dos usuários logados)
    function Invoke-WingetTask {
        param(
            [string]$TaskName,
            [string]$UserId,
            [bool]$IsInteractive,
            [string]$Arguments,
            [string]$OutputFile,
            [int]$MaxWaitSeconds
        )

        if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue }

        # Cria um arquivo .cmd temporário para evitar truncamento/conflitos de aspas no cmd.exe /c
        # Sanitiza e limita o tamanho para nunca estourar o limite de caminho (MAX_PATH 260) ou nome de tarefa
        $safeTaskName = $TaskName -replace '[^a-zA-Z0-9_]',''
        if ($safeTaskName.Length -gt 35) {
            $safeTaskName = $safeTaskName.Substring(0, 35)
        }

        # Garante limpeza prévia caso a tarefa anterior com mesmo nome tenha ficado órfã
        $prevTask = Get-ScheduledTask -TaskName $safeTaskName -ErrorAction SilentlyContinue
        if ($prevTask) {
            if ($prevTask.State -eq "Running") {
                Stop-ScheduledTask -TaskName $safeTaskName -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 500
            }
            Unregister-ScheduledTask -TaskName $safeTaskName -Confirm:$false -ErrorAction SilentlyContinue
        }

        $cmdFile = "C:\Windows\Temp\${safeTaskName}_$PID.cmd"
        Set-Content -Path $cmdFile -Value "@echo off`r`n`"$wingetPath`" $Arguments > `"$OutputFile`" 2>&1"
        $vbsFile = $null
        if ($IsInteractive) {
            # Cria um wrapper VBScript executado pelo wscript.exe (subsistema GUI) com WindowStyle=0 (SW_HIDE).
            # Isso impede que o cmd.exe abra ou pisque uma janela de console preta na tela do usuario logado.
            $vbsFile = "C:\Windows\Temp\${safeTaskName}_$PID.vbs"
            $vbsContent = @"
Set WshShell = CreateObject("WScript.Shell")
ret = WshShell.Run("cmd.exe /c """ & "$cmdFile" & """", 0, True)
WScript.Quit ret
"@
            Set-Content -Path $vbsFile -Value $vbsContent -Encoding ASCII

            $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "//B `"$vbsFile`""
            $principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive
        } else {
            $action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$cmdFile`""
            $principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType ServiceAccount -RunLevel Highest
        }


        # Configura tempo limite explícito no próprio agendador do Windows (ExecutionTimeLimit)
        $timeoutSpan = [TimeSpan]::FromSeconds($MaxWaitSeconds + 15)
        $settings = New-ScheduledTaskSettingsSet -Hidden `
                                                -AllowStartIfOnBatteries `
                                                -DontStopIfGoingOnBatteries `
                                                -ExecutionTimeLimit $timeoutSpan
        
        $null = Register-ScheduledTask -TaskName $safeTaskName -Action $action -Principal $principal -Settings $settings -Force
        Start-ScheduledTask -TaskName $safeTaskName

        $waited = 0
        $timedOut = $false
        while ($waited -lt $MaxWaitSeconds) {
            $taskState = (Get-ScheduledTask -TaskName $safeTaskName -ErrorAction SilentlyContinue).State
            if ($taskState -ne "Running") {
                break
            }
            Start-Sleep -Seconds 2
            $waited += 2
        }

        if ($waited -ge $MaxWaitSeconds) {
            $timedOut = $true
        }

        # SE OCORREU TIMEOUT: Interrompe a tarefa e mata a árvore de processos filhos (evita processos zumbis travando a máquina)
        if ($timedOut) {
            Write-Host "     [TIMEOUT] Limite de $MaxWaitSeconds s excedido! Encerrando processos associados..." -ForegroundColor Yellow
            Stop-ScheduledTask -TaskName $safeTaskName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1

            # Mata processos vinculados que possam estar segurando locks no Windows
            $stuckProcs = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object { 
                $_.Name -in @('cmd.exe','winget.exe','wscript.exe','Windows10Upgrade9252.exe','Windows10UpgraderApp.exe','vs_Professional.exe','setup.exe') -and 
                ($_.CommandLine -match $safeTaskName -or $_.CommandLine -match 'winget' -or $_.CommandLine -match 'Windows10Upgrade' -or $_.CommandLine -match 'vs_Professional')
            }
            foreach ($p in $stuckProcs) {
                Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
            }

            Add-Content -Path $OutputFile -Value "`r`n[ERRO] Operação atingiu o tempo limite de $MaxWaitSeconds segundos e foi abortada forçadamente pelo script.`r`n" -ErrorAction SilentlyContinue
        }

        $info = Get-ScheduledTaskInfo -TaskName $safeTaskName -ErrorAction SilentlyContinue
        Unregister-ScheduledTask -TaskName $safeTaskName -Confirm:$false -ErrorAction SilentlyContinue
        Remove-Item $cmdFile -Force -ErrorAction SilentlyContinue
        if ($vbsFile -and (Test-Path $vbsFile)) {
            Remove-Item $vbsFile -Force -ErrorAction SilentlyContinue
        }

        $content = ""
        if (Test-Path $OutputFile) {
            $content = Get-Content $OutputFile -Raw -ErrorAction SilentlyContinue
            Remove-Item $OutputFile -Force -ErrorAction SilentlyContinue
        }

        $exitCode = if ($timedOut) {
            1460
        } else {
            if ($info) {
                $rawRes = $info.LastTaskResult
                if ($rawRes -gt 0x7FFFFFFF) {
                    try {
                        $rawRes = [System.BitConverter]::ToInt32([System.BitConverter]::GetBytes([uint32]$rawRes), 0)
                    } catch { }
                }
                $rawRes
            } else {
                -1
            }
        }

        return @{
            ExitCode = $exitCode
            Output   = $content
            TimedOut = $timedOut
        }
    }

    # Função interna para analisar saída textual de 'winget upgrade' ou 'winget list'
    function Parse-WingetUpgradeOutput {
        param([string]$RawOutput)

        if ([string]::IsNullOrWhiteSpace($RawOutput)) { return @() }

        $lines = $RawOutput -split "`r?`n"
        $headerIndex = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^(Nome|Name)\s+') {
                $headerIndex = $i
                break
            }
        }

        if ($headerIndex -eq -1) { return @() }

        $header = $lines[$headerIndex]
        $idCol = if ($header.IndexOf("Id") -ge 0) { $header.IndexOf("Id") } else { $header.IndexOf("ID") }
        $verCol = if ($header.IndexOf("Version") -ge 0) { $header.IndexOf("Version") } else { $header.IndexOf("Vers") }
        $availCol = if ($header.IndexOf("Available") -ge 0) { $header.IndexOf("Available") } else { $header.IndexOf("Disp") }
        $srcCol = if ($header.IndexOf("Source") -ge 0) { $header.IndexOf("Source") } else { $header.IndexOf("Origem") }

        if ($idCol -lt 0 -or $verCol -lt 0) { return @() }

        $items = @()
        for ($i = $headerIndex + 2; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^\s*\d+\s+(atualiza|upgrade)' -or $line -match '^-+') {
                continue
            }
            if ($line.Length -lt $verCol) { continue }

            $name = $line.Substring(0, $idCol).Trim()

            $lenId = $verCol - $idCol
            if ($line.Length -lt ($idCol + $lenId)) { $lenId = $line.Length - $idCol }
            $id = $line.Substring($idCol, $lenId).Trim()

            $endVer = if ($availCol -gt 0 -and $availCol -gt $verCol) {
                $availCol
            } elseif ($srcCol -gt 0 -and $srcCol -gt $verCol) {
                $srcCol
            } else {
                $line.Length
            }

            $lenVer = $endVer - $verCol
            if ($line.Length -lt ($verCol + $lenVer)) { $lenVer = $line.Length - $verCol }
            $version = $line.Substring($verCol, $lenVer).Trim()

            $avail = ""
            if ($availCol -gt 0 -and $line.Length -ge $availCol) {
                $lenAvail = if ($srcCol -gt 0 -and $line.Length -ge $srcCol) { $srcCol - $availCol } else { $line.Length - $availCol }
                $avail = $line.Substring($availCol, $lenAvail).Trim()
            } else {
                $avail = $version
            }

            if ($id) {
                $items += @{
                    Name      = $name
                    Id        = $id
                    Version   = $version
                    Available = $avail
                }
            }
        }
        return $items
    }

    # Função para verificar se um pacote está na lista de exceções
    function Test-IsExcluded {
        param([string]$PkgId, [string]$PkgName, [string[]]$Exclusions)
        if (-not $Exclusions -or $Exclusions.Count -eq 0) { return $false }
        foreach ($p in $Exclusions) {
            if ($PkgId -like $p -or $PkgName -like $p) { return $true }
        }
        return $false
    }

    # ========================================================
    # 0. MODO INSTALAÇÃO DE NOVO PACOTE ESPECÍFICO
    # ========================================================
    if ($InstallId) {
        $cleanId = $InstallId -replace '[^a-zA-Z0-9_]',''
        if ($cleanId.Length -gt 25) { $cleanId = $cleanId.Substring(0, 25) }
        $installTempFile = "C:\Windows\Temp\winget_inst_${cleanId}_$PID.txt"

        $instRes = Invoke-WingetTask -TaskName "Winget_Inst_${cleanId}" `
                                    -UserId "NT AUTHORITY\SYSTEM" `
                                    -IsInteractive $false `
                                    -Arguments "install --id `"$InstallId`" --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity" `
                                    -OutputFile $installTempFile `
                                    -MaxWaitSeconds $Timeout

        $st = switch ($instRes.ExitCode) {
            0           { "Instalado com sucesso" }
            3010        { "Requer reboot" }
            -1978335189 { "Já instalado" }
            -1978335212 { "Pacote não encontrado" }
            default     {
                if ($instRes.Output -match 'Já instalado|already installed') {
                    "Já instalado"
                } elseif ($instRes.Output -match 'Nenhum pacote|No package found') {
                    "Pacote não encontrado"
                } else {
                    "Falhou (Código: $($instRes.ExitCode))"
                }
            }
        }

        # Consulta detalhes do pacote instalado via 'winget list'
        $listFile = "C:\Windows\Temp\winget_chk_${cleanId}_$PID.txt"
        $chkRes = Invoke-WingetTask -TaskName "Winget_Chk_${cleanId}" `
                                   -UserId "NT AUTHORITY\SYSTEM" `
                                   -IsInteractive $false `
                                   -Arguments "list --id `"$InstallId`" --exact --accept-source-agreements" `
                                   -OutputFile $listFile `
                                   -MaxWaitSeconds 30

        $installedInfo = @(Parse-WingetUpgradeOutput -RawOutput $chkRes.Output)
        $pkgName = $InstallId
        $pkgVer = "N/A"
        if ($installedInfo.Count -gt 0) {
            $pkgName = $installedInfo[0].Name
            $pkgVer = $installedInfo[0].Version
        }

        $errSnippet = ""
        if ($instRes.ExitCode -ne 0 -and $instRes.ExitCode -ne 3010 -and $instRes.ExitCode -ne -1978335189) {
            if ($instRes.Output) {
                $errSnippet = if ($instRes.Output.Length -gt 800) { $instRes.Output.Substring($instRes.Output.Length - 800) } else { $instRes.Output }
            }
        }

        return @{
            Results = @(@{
                Escopo           = "Machine"
                Programa         = $pkgName
                ID               = $InstallId
                VersaoAtual      = "N/A"
                VersaoDisponivel = $pkgVer
                Status           = $st
                ExitCode         = $instRes.ExitCode
                ErrorSnippet     = $errSnippet
            })
            MachinePending = 0
            LoggedUsers    = @()
        }
    }

    $allResults = @()

    # ========================================================
    # 2. ESCOPO DE MÁQUINA
    # ========================================================
    $machineTempQueryFile = "C:\Windows\Temp\winget_m_query_$PID.txt"
    $queryRes = Invoke-WingetTask -TaskName "Winget_Query_Machine" `
                                 -UserId "NT AUTHORITY\SYSTEM" `
                                 -IsInteractive $false `
                                 -Arguments "upgrade --accept-source-agreements" `
                                 -OutputFile $machineTempQueryFile `
                                 -MaxWaitSeconds 90

    $machineUpgrades = @(Parse-WingetUpgradeOutput -RawOutput $queryRes.Output)

    foreach ($pkg in $machineUpgrades) {
        if (-not $pkg -or [string]::IsNullOrWhiteSpace($pkg.Id)) { continue }

        # Se um pacote específico foi solicitado, filtra os demais
        if ($TargetPackage -and ($pkg.Id -notlike $TargetPackage -and $pkg.Name -notlike $TargetPackage)) {
            continue
        }

        $isEx = Test-IsExcluded -PkgId $pkg.Id -PkgName $pkg.Name -Exclusions $ExcludeList
        if ($isEx) {
            $allResults += @{
                Escopo           = "Machine"
                Programa         = $pkg.Name
                ID               = $pkg.Id
                VersaoAtual      = $pkg.Version
                VersaoDisponivel = $pkg.Available
                Status           = "Ignorado (Exceção)"
            }
            continue
        }

        # Se for apenas listagem (sem -Update), não realiza a atualização
        if ($OnlyList) {
            $allResults += @{
                Escopo           = "Machine"
                Programa         = $pkg.Name
                ID               = $pkg.Id
                VersaoAtual      = $pkg.Version
                VersaoDisponivel = $pkg.Available
                Status           = "Pendente"
            }
            continue
        }

        # Filtro Nativo de Seguranca: Assistentes de Atualizacao de SO (Incompativeis com Sessao 0 / Background)
        if ($pkg.Id -match 'Microsoft\.UpdateAssistant' -or $pkg.Id -match 'Microsoft\.WindowsInstallationAssistant') {
            $allResults += @{
                Escopo           = "Machine"
                Programa         = $pkg.Name
                ID               = $pkg.Id
                VersaoAtual      = $pkg.Version
                VersaoDisponivel = $pkg.Available
                Status           = "Ignorado (Assistente de SO - Incompativel com Sessao 0)"
                ExitCode         = $null
                ErrorSnippet     = "Assistentes de migracao de versao do Windows exigem sessao interativa e nao devem ser executados em background via Winget."
            }
            Write-Host "  -> Ignorando assistente de SO incompativel com background: $($pkg.Name)" -ForegroundColor DarkYellow
            continue
        }

        # Executa a atualização silenciosa do pacote
        Write-Host "  -> Atualizando: $($pkg.Name) ($($pkg.Id))..." -ForegroundColor Cyan
        $cleanPkgId = $pkg.Id -replace '[^a-zA-Z0-9_]',''
        if ($cleanPkgId.Length -gt 25) { $cleanPkgId = $cleanPkgId.Substring(0, 25) }
        $machineTempUpFile = "C:\Windows\Temp\winget_m_up_${cleanPkgId}_$PID.txt"
        $upRes = Invoke-WingetTask -TaskName "W_Up_M_${cleanPkgId}" `
                                  -UserId "NT AUTHORITY\SYSTEM" `
                                  -IsInteractive $false `
                                  -Arguments "upgrade --id `"$($pkg.Id)`" --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity" `
                                  -OutputFile $machineTempUpFile `
                                  -MaxWaitSeconds $Timeout

        $st = switch ($upRes.ExitCode) {
            0           { "Atualizado" }
            3010        { "Requer reboot" }
            -1978335189 { "Já atualizado" }
            1460        { "Falhou (Tempo Limite Excedido: ${Timeout}s)" }
            default     { "Falhou (Código: $($upRes.ExitCode))" }
        }

        if ($upRes.ExitCode -eq 0) {
            Write-Host "     [OK] Concluído com sucesso." -ForegroundColor Green
        } elseif ($upRes.ExitCode -eq 3010) {
            Write-Host "     [!] Concluído (requer reinicialização)." -ForegroundColor Yellow
        } elseif ($upRes.ExitCode -eq -1978335189) {
            Write-Host "     [=] Já está na versão mais recente." -ForegroundColor Gray
        } else {
            Write-Host "     [FALHOU] $($st)" -ForegroundColor Red
        }

        $errSnippet = ""
        if ($upRes.ExitCode -ne 0 -and $upRes.ExitCode -ne 3010 -and $upRes.ExitCode -ne -1978335189) {
            if ($upRes.Output) {
                $errSnippet = if ($upRes.Output.Length -gt 800) { $upRes.Output.Substring($upRes.Output.Length - 800) } else { $upRes.Output }
            }
        }

        $allResults += @{
            Escopo           = "Machine"
            Programa         = $pkg.Name
            ID               = $pkg.Id
            VersaoAtual      = $pkg.Version
            VersaoDisponivel = $pkg.Available
            Status           = $st
            ExitCode         = $upRes.ExitCode
            ErrorSnippet     = $errSnippet
        }
    }

    # ========================================================
    # 3. ESCOPO DOS USUÁRIOS LOGADOS
    # ========================================================
    $loggedUsers = @()
    $quserRaw = quser.exe 2>&1
    if ($quserRaw) {
        foreach ($line in ($quserRaw | Select-Object -Skip 1)) {
            $trimmed = $line.ToString().Trim()
            if ($trimmed -match '^>?([a-zA-Z0-9_\.\-]+)\s+') {
                $u = $Matches[1]
                if ($u -notmatch '^(SYSTEM|LOCAL SERVICE|NETWORK SERVICE)$') {
                    $loggedUsers += $u
                }
            }
        }
    }
    $loggedUsers = $loggedUsers | Select-Object -Unique

    foreach ($u in $loggedUsers) {
        $cleanU = $u -replace '[^a-zA-Z0-9_]',''
        if ($cleanU.Length -gt 15) { $cleanU = $cleanU.Substring(0, 15) }
        $userTempQueryFile = "C:\Windows\Temp\winget_u_q_${cleanU}_$PID.txt"
        $userQueryRes = Invoke-WingetTask -TaskName "W_Q_U_${cleanU}" `
                                         -UserId $u `
                                         -IsInteractive $true `
                                         -Arguments "upgrade --scope user --accept-source-agreements" `
                                         -OutputFile $userTempQueryFile `
                                         -MaxWaitSeconds 90

        $userUpgrades = @(Parse-WingetUpgradeOutput -RawOutput $userQueryRes.Output)

        foreach ($pkg in $userUpgrades) {
            if (-not $pkg -or [string]::IsNullOrWhiteSpace($pkg.Id)) { continue }

            # Se um pacote específico foi solicitado, filtra os demais
            if ($TargetPackage -and ($pkg.Id -notlike $TargetPackage -and $pkg.Name -notlike $TargetPackage)) {
                continue
            }

            $isEx = Test-IsExcluded -PkgId $pkg.Id -PkgName $pkg.Name -Exclusions $ExcludeList
            if ($isEx) {
                $allResults += @{
                    Escopo           = "User ($u)"
                    Programa         = $pkg.Name
                    ID               = $pkg.Id
                    VersaoAtual      = $pkg.Version
                    VersaoDisponivel = $pkg.Available
                    Status           = "Ignorado (Exceção)"
                }
                continue
            }

            # Se for apenas listagem (sem -Update), não realiza a atualização
            if ($OnlyList) {
                $allResults += @{
                    Escopo           = "User ($u)"
                    Programa         = $pkg.Name
                    ID               = $pkg.Id
                    VersaoAtual      = $pkg.Version
                    VersaoDisponivel = $pkg.Available
                    Status           = "Pendente"
                }
                continue
            }
            # Executa atualizacao no escopo de usuario
            Write-Host "  -> Atualizando [$u]: $($pkg.Name) ($($pkg.Id))..." -ForegroundColor Cyan
            $cleanPkgId = $pkg.Id -replace '[^a-zA-Z0-9_]',''
            if ($cleanPkgId.Length -gt 20) { $cleanPkgId = $cleanPkgId.Substring(0, 20) }
            $userTempUpFile = "C:\Windows\Temp\winget_u_up_${cleanU}_${cleanPkgId}_$PID.txt"
            $userUpRes = Invoke-WingetTask -TaskName "W_Up_U_${cleanU}_${cleanPkgId}" `
                                          -UserId $u `
                                          -IsInteractive $true `
                                          -Arguments "upgrade --id `"$($pkg.Id)`" --scope user --exact --silent --accept-source-agreements --accept-package-agreements --disable-interactivity" `
                                          -OutputFile $userTempUpFile `
                                          -MaxWaitSeconds $Timeout

            $st = switch ($userUpRes.ExitCode) {
                0           { "Atualizado" }
                3010        { "Requer reboot" }
                -1978335189 { "Ja atualizado" }
                1460        { "Falhou (Tempo Limite Excedido: ${Timeout}s)" }
                default     { "Falhou (Codigo: $($userUpRes.ExitCode))" }
            }

            if ($userUpRes.ExitCode -eq 0) {
                Write-Host "     [OK] Concluido com sucesso." -ForegroundColor Green
            } elseif ($userUpRes.ExitCode -eq 3010) {
                Write-Host "     [!] Concluido (requer reinicializacao)." -ForegroundColor Yellow
            } elseif ($userUpRes.ExitCode -eq -1978335189) {
                Write-Host "     [=] Ja esta na versao mais recente." -ForegroundColor Gray
            } else {
                Write-Host "     [FALHOU] $($st)" -ForegroundColor Red
            }

            $errSnippet = ""
            if ($userUpRes.ExitCode -ne 0 -and $userUpRes.ExitCode -ne 3010 -and $userUpRes.ExitCode -ne -1978335189) {
                if ($userUpRes.Output) {
                    $errSnippet = if ($userUpRes.Output.Length -gt 800) { $userUpRes.Output.Substring($userUpRes.Output.Length - 800) } else { $userUpRes.Output }
                }
            }

            $allResults += @{
                Escopo           = "User ($u)"
                Programa         = $pkg.Name
                ID               = $pkg.Id
                VersaoAtual      = $pkg.Version
                VersaoDisponivel = $pkg.Available
                Status           = $st
                ExitCode         = $userUpRes.ExitCode
                ErrorSnippet     = $errSnippet
            }
        }
    }

    return @{
        Results         = $allResults
        MachinePending  = $machineUpgrades.Count
        LoggedUsers     = $loggedUsers
    }
}

# ========================================================
# FUNÇÕES DE APOIO LOCAL (AJUDA, TABELAS E DIAGNÓSTICO)
# ========================================================

function Show-ScriptHelp {
    param(
        [string]$TargetComputer = "estacao061"
    )

    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host "          MANUAL DE USO E GUIA DE COMANDOS: Update-StationWinget.ps1           " -ForegroundColor Cyan
    Write-Host "================================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRICAO:" -ForegroundColor Yellow
    Write-Host "  Este script gerencia, consulta, atualiza e instala aplicativos remotamente via WinRM"
    Write-Host "  utilizando o Windows Package Manager (Winget). Ele contorna restricoes de rede do WinRM"
    Write-Host "  e o modo ConstrainedLanguage executando atraves de Tarefas Agendadas locais temporarias:"
    Write-Host "    * Escopo de Maquina: Sob o contexto de alta permissao NT AUTHORITY\SYSTEM."
    Write-Host "    * Escopo de Usuario: Detecta usuarios logados e executa tarefas na sessao do usuario."
    Write-Host "    * 100% Silencioso: Sem pop-ups ou janelas visiveis para os usuarios finais.
    * Log Unificado Incremental: Registra automaticamente todas as acoes em 'output.log' na pasta de execucao."
    Write-Host ""
    Write-Host "COMPORTAMENTO PADRAO:" -ForegroundColor Yellow
    Write-Host "  Quando executado sem parametros (ou com -Help), o script exibe este manual completo."
    Write-Host "  Para consultar pacotes pendentes em modo seguro (sem alterar nada no alvo), basta"
    Write-Host "  informar o nome da maquina: .\Update-StationWinget.ps1 -ComputerName '$TargetComputer'"
    Write-Host ""
    Write-Host "PARAMETROS DISPONIVEIS:" -ForegroundColor Yellow
    Write-Host "  -ComputerName <string[]>      Nome ou IP da maquina remota (Padrao: '$TargetComputer'). Aceita multiplas maquinas."
    Write-Host "  -ComputerFile <string>        Arquivo de texto (.txt) com lista de maquinas a processar (uma por linha)."
    Write-Host "  -Update                       Aplica as atualizacoes pendentes de forma silenciosa."
    Write-Host "  -PackageId <string>           Limita a acao a um unico aplicativo (ID, Nome ou wildcard)."
    Write-Host "  -InstallPackage <string>      Instala um novo aplicativo pelo seu ID do Winget (ex: '7zip.7zip')."
    Write-Host "  -ExcludePackages <string[]>   Lista de IDs/Nomes a ignorar nas atualizacoes (suporta wildcards)."
    Write-Host "  -IncludeSkipped               Exibe na tabela final os pacotes que foram ignorados por excecao."
    Write-Host "  -PassThru                     Retorna os objetos no pipeline para permitir exportacao (Export-Csv)."
    Write-Host "  -Credential <PSCredential>    Credencial administrativa explicita para conexao remota."
    Write-Host "  -TaskTimeoutSeconds <int>     Tempo limite por pacote em segundos (Padrao: 900s)."
    Write-Host "  -Help                         Exibe este guia completo de parametros e exemplos."
    Write-Host ""
    Write-Host "EXEMPLOS PRATICOS DE USO (1 EXEMPLO PARA CADA PARAMETRO):" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. Parametro -ComputerName (Definir computador alvo):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer'" -ForegroundColor White
    Write-Host ""
    Write-Host "  2. Parametro -ComputerFile (Definir arquivo com lista de maquinas):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerFile 'C:\Temp\maquinas.txt'" -ForegroundColor White
    Write-Host ""
    Write-Host "  3. Parametro -Update (Executar atualizacoes de fato):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -Update" -ForegroundColor White
    Write-Host ""
    Write-Host "  4. Parametro -PackageId (Filtrar aplicativo especifico):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -PackageId 'VideoLAN.VLC'" -ForegroundColor White
    Write-Host ""
    Write-Host "  5. Parametro -InstallPackage (Instalar novo aplicativo):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -InstallPackage '7zip.7zip'" -ForegroundColor White
    Write-Host ""
    Write-Host "  6. Parametro -ExcludePackages (Definir excecoes a ignorar):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -Update -ExcludePackages @('Microsoft.Teams.Classic', '*Dell*')" -ForegroundColor White
    Write-Host ""
    Write-Host "  7. Parametro -IncludeSkipped (Exibir pacotes ignorados na tabela de resumo):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -Update -ExcludePackages @('Zoom.Zoom') -IncludeSkipped" -ForegroundColor White
    Write-Host ""
    Write-Host "  8. Parametro -TaskTimeoutSeconds (Definir tempo limite customizado em segundos):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -Update -TaskTimeoutSeconds 1800" -ForegroundColor White
    Write-Host ""
    Write-Host "  9. Parametro -Credential (Fornecer credenciais administrativas explicitas):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -Credential (Get-Credential) -Update" -ForegroundColor White
    Write-Host ""
    Write-Host " 10. Parametro -PassThru (Exportar objetos para pipeline / CSV):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -ComputerName '$TargetComputer' -PassThru | Export-Csv -Path 'C:\Temp\pendentes.csv' -NoTypeInformation -Encoding UTF8" -ForegroundColor White
    Write-Host ""
    Write-Host " 11. Parametro -Help (Exibir manual de instrucoes e parametros):" -ForegroundColor Cyan
    Write-Host "     .\Update-StationWinget.ps1 -Help" -ForegroundColor White
    Write-Host "================================================================================" -ForegroundColor Cyan
}


# ==============================================================================
# BASE DE CONHECIMENTO DE DIAGNOSTICO DO WINGET (MAIS DE 200 CASOS CATALOGADOS)
# ==============================================================================
$script:WingetKnowledgeBase = @{
    # --- 1. CODIGOS OFICIAIS DO WINGET CLI (0x8A150001 a 0x8A150110) ---
    '0x8A150001' = @{
        Categoria = 'Erro Interno do WinGet (APPINSTALLER_CLI_ERROR_INTERNAL_ERROR)'
        Causa     = 'O mecanismo interno do cliente WinGet encontrou uma falha de execucao inesperada ou excecao nao tratada.'
        Solucoes  = @(
            'Verificar a integridade do pacote do WinGet executando: winget --version',
            'Atualizar o WinGet manualmente ou reparar o pacote Microsoft.DesktopAppInstaller na estacao.',
            'Tentar a operacao novamente apos reiniciar o servico de tarefas agendadas.'
        )
    }
    '0x8A150002' = @{
        Categoria = 'Argumentos de Linha de Comando Invalidos (APPINSTALLER_CLI_ERROR_INVALID_CL_ARGUMENTS)'
        Causa     = 'Parametros conflitantes ou sintaxe incorreta foram passados para o executavel do WinGet.'
        Solucoes  = @(
            'Verificar os parametros enviados na chamada do winget upgrade ou install.',
            'Testar o comando localmente no terminal da estacao com as mesmas opcoes para verificar a sintaxe.'
        )
    }
    '0x8A150003' = @{
        Categoria = 'Comando Interno Falhou (APPINSTALLER_CLI_ERROR_COMMAND_FAILED)'
        Causa     = 'O comando do WinGet foi interrompido antes de completar sua rotina de execucao.'
        Solucoes  = @(
            'Consultar os logs detalhados do WinGet na pasta %LOCALAPPDATA%\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir.',
            'Tentar executar o comando individualmente para o pacote pelo ID.'
        )
    }
    '0x8A150004' = @{
        Categoria = 'Falha de Manifesto (APPINSTALLER_CLI_ERROR_MANIFEST_FAILED)'
        Causa     = 'O manifesto YAML do pacote no repositorio oficial esta corrompido, incompleto ou com sintaxe invalida.'
        Solucoes  = @(
            'Executar winget source update para sincronizar a versao mais recente do manifesto.',
            'Aguardar atualizacao corretiva do manifesto pela comunidade no repositorio winget-pkgs.',
            'Adicionar o pacote a lista de excecao com -ExcludePackages.'
        )
    }
    '0x8A150005' = @{
        Categoria = 'Operacao Interrompida por Sinal (APPINSTALLER_CLI_ERROR_CTRL_SIGNAL_RECEIVED)'
        Causa     = 'A execucao do WinGet recebeu um sinal de cancelamento (Ctrl+C ou terminacao forcada de processo).'
        Solucoes  = @(
            'Verificar se a tarefa agendada atingiu o tempo limite (Timeout) e foi abortada.',
            'Aumentar o tempo limite atraves do parametro -TaskTimeoutSeconds 1800.'
        )
    }
    '0x8A150006' = @{
        Categoria = 'Falha na Execucao do Instalador (APPINSTALLER_CLI_ERROR_SHELLEXEC_INSTALL_FAILED)'
        Causa     = 'O instalador executado pelo Shell do Windows encerrou com codigo de falha diferente de zero.'
        Solucoes  = @(
            'Verificar o visualizador de eventos do Windows (Event Viewer -> Application) para conferir a causa raiz.',
            'Executar a instalacao interativa na estacao para visualizar mensagens de erro especificas do fabricante.',
            'Ignorar este pacote no lote com -ExcludePackages.'
        )
    }
    '0x8A150007' = @{
        Categoria = 'Versao do Manifesto Nao Suportada (APPINSTALLER_CLI_ERROR_UNSUPPORTED_MANIFESTVERSION)'
        Causa     = 'O manifesto do pacote utiliza uma versao de esquema mais nova do que a suportada pelo cliente WinGet instalado.'
        Solucoes  = @(
            'Atualizar o cliente WinGet (Microsoft.DesktopAppInstaller) na estacao para a versao mais recente.',
            'Verificar atualizacoes pendentes do sistema e da Microsoft Store.'
        )
    }
    '0x8A150008' = @{
        Categoria = 'Falha no Download do Instalador (APPINSTALLER_CLI_ERROR_DOWNLOAD_FAILED)'
        Causa     = 'Nao foi possivel baixar o binario de instalacao do servidor do fornecedor (queda de conexao, timeout ou bloqueio).'
        Solucoes  = @(
            'Verificar a conexao com a internet e resolucao DNS da estacao para a URL do fabricante.',
            'Verificar se firewalls, UTM ou proxy corporativo estao bloqueando o download de arquivos executaveis.',
            'Tentar novamente mais tarde ou excluir o pacote da fila com -ExcludePackages.'
        )
    }
    '0x8A150009' = @{
        Categoria = 'Gravacao em Fonte Somente Leitura (APPINSTALLER_CLI_ERROR_CANNOT_WRITE_TO_UPSTREAM_SOURCE)'
        Causa     = 'Tentativa invalida de modificacao de dados em repositorio upstream somente leitura.'
        Solucoes  = @(
            'Restaurar as fontes padrao do WinGet executando: winget source reset --force.'
        )
    }
    '0x8A15000A' = @{
        Categoria = 'Fonte de Sistema Invalida (APPINSTALLER_CLI_ERROR_INVALID_SYSTEM_FONT)'
        Causa     = 'Configuracao incorreta de fontes de interface ou renderizacao do terminal.'
        Solucoes  = @(
            'Redefinir as preferencias de terminal do Windows para os padroes de fabrica.'
        )
    }
    '0x8A15000B' = @{
        Categoria = 'Falha ao Instalar Fonte de Repositorio (APPINSTALLER_CLI_ERROR_SOURCE_INSTALL_FAILED)'
        Causa     = 'Nao foi possivel adicionar ou inicializar a fonte de pacotes configurada.'
        Solucoes  = @(
            'Verificar se o catalogo local em %LOCALAPPDATA%\Microsoft\WinGet esta corrompido.',
            'Redefinir as fontes executando: winget source reset --force.'
        )
    }
    '0x8A15000C' = @{
        Categoria = 'Falha ao Atualizar Catalogo de Repositorio (APPINSTALLER_CLI_ERROR_SOURCE_UPDATE_FAILED)'
        Causa     = 'A sincronizacao do catalogo de pacotes com os servidores da Microsoft falhou.'
        Solucoes  = @(
            'Verificar conectividade com os servidores da Microsoft (cdn.winget.microsoft.com).',
            'Executar no terminal da estacao: winget source update.'
        )
    }
    '0x8A15000D' = @{
        Categoria = 'Falha ao Desinstalar Fonte (APPINSTALLER_CLI_ERROR_SOURCE_UNINSTALL_FAILED)'
        Causa     = 'Falha ao remover registro de repositorio de pacotes.'
        Solucoes  = @(
            'Executar: winget source reset --force.'
        )
    }
    '0x8A15000E' = @{
        Categoria = 'Falha ao Instalar Extensao (APPINSTALLER_CLI_ERROR_EXTENSION_INSTALL_FAILED)'
        Causa     = 'Nao foi possivel registrar a extensao de comando do WinGet.'
        Solucoes  = @(
            'Verificar compatibilidade do cliente WinGet com extensoes de terceiros.'
        )
    }
    '0x8A15000F' = @{
        Categoria = 'Falha ao Desinstalar Extensao (APPINSTALLER_CLI_ERROR_EXTENSION_UNINSTALL_FAILED)'
        Causa     = 'Falha na desinstalacao de extensao do WinGet.'
        Solucoes  = @(
            'Remover manualmente os arquivos da extensao em %LOCALAPPDATA%\Microsoft\WinGet\Extensions.'
        )
    }
    '0x8A150010' = @{
        Categoria = 'Instalacao Bloqueada por Diretiva (APPINSTALLER_CLI_ERROR_PACKAGES_RESTRICTED)'
        Causa     = 'Diretiva de grupo (GPO) corporativa proibe a instalacao de pacotes por fontes externas.'
        Solucoes  = @(
            'Verificar as politicas de Grupo em: Componentes do Windows -> Gerenciador de Pacotes da Area de Trabalho.',
            'Solicitar inclusao de excecao na GPO corporativa para a estacao.'
        )
    }
    '0x8A150011' = @{
        Categoria = 'Nenhum Pacote Localizado (APPINSTALLER_CLI_ERROR_NO_APPLICATIONS_FOUND)'
        Causa     = 'Nenhum aplicativo correspondeu ao ID ou nome especificado no catalogo de repositorios ativos.'
        Solucoes  = @(
            'Verificar se o ID do pacote esta digitado corretamente.',
            'Executar winget source update para atualizar os indices de busca.'
        )
    }
    '0x8A150012' = @{
        Categoria = 'Nenhuma Fonte de Repositorio Definida (APPINSTALLER_CLI_ERROR_NO_SOURCES_DEFINED)'
        Causa     = 'O cliente WinGet esta configurado sem nenhuma fonte de pacotes ativa.'
        Solucoes  = @(
            'Restaurar as fontes padrao executando no destino: winget source reset --force.',
            'Verificar se o comando winget source list retorna o repositorio winget.'
        )
    }
    '0x8A150013' = @{
        Categoria = 'Multiplos Aplicativos Encontrados (APPINSTALLER_CLI_ERROR_MULTIPLE_APPLICATIONS_FOUND)'
        Causa     = 'O nome ou termo de busca retornou mais de um pacote disponivel, exigindo desambiguacao.'
        Solucoes  = @(
            'Especificar o ID exato com o parametro -PackageId (ex: VideoLAN.VLC em vez de vlc).'
        )
    }
    '0x8A150014' = @{
        Categoria = 'Manifesto Nao Encontrado (APPINSTALLER_CLI_ERROR_NO_MANIFEST_FOUND)'
        Causa     = 'Nao foi possivel localizar os metadados de instalacao para este pacote no catalogo.'
        Solucoes  = @(
            'Atualizar as fontes de pacotes locais com winget source update.',
            'Verificar se o pacote ainda existe no repositorio central.'
        )
    }
    '0x8A150015' = @{
        Categoria = 'Extensao Nao Carregada (APPINSTALLER_CLI_ERROR_EXTENSION_NOT_LOADED)'
        Causa     = 'Uma extensao exigida pelo fluxo de instalacao nao pode ser inicializada.'
        Solucoes  = @(
            'Reinstalar o cliente do WinGet na estacao de destino.'
        )
    }
    '0x8A150016' = @{
        Categoria = 'Nome de Fonte Duplicado (APPINSTALLER_CLI_ERROR_DUPLICATE_SOURCE_NAME)'
        Causa     = 'Conflito na listagem de repositorios do WinGet.'
        Solucoes  = @(
            'Executar: winget source reset --force.'
        )
    }
    '0x8A150017' = @{
        Categoria = 'Tipo de Fonte Invalido (APPINSTALLER_CLI_ERROR_INVALID_SOURCE_TYPE)'
        Causa     = 'Tipo de repositorio desconhecido ou incompativel configurado.'
        Solucoes  = @(
            'Executar: winget source reset --force.'
        )
    }
    '0x8A150018' = @{
        Categoria = 'Pacote e um Bundle Incompativel (APPINSTALLER_CLI_ERROR_PACKAGE_IS_BUNDLE)'
        Causa     = 'O pacote selecionado e um bundle com restricoes de arquitetura.'
        Solucoes  = @(
            'Verificar os requisitos de arquitetura do pacote (x64, ARM64 ou x86).'
        )
    }
    '0x8A150019' = @{
        Categoria = 'Falha em Chamada de API da Fonte (APPINSTALLER_CLI_ERROR_SOURCE_CALL_FAILED)'
        Causa     = 'A comunicacao com o servidor de metadados da fonte falhou durante a requisicao.'
        Solucoes  = @(
            'Verificar estabilidade da conexao de rede da estacao de destino.',
            'Tentar novamente apos alguns minutos.'
        )
    }
    '0x8A15001A' = @{
        Categoria = 'Fonte Nao e Remota (APPINSTALLER_CLI_ERROR_SOURCE_NOT_REMOTE)'
        Causa     = 'Uma operacao que exigia um repositorio online foi invocada em fonte local.'
        Solucoes  = @(
            'Restaurar as fontes padrao com: winget source reset --force.'
        )
    }
    '0x8A15001B' = @{
        Categoria = 'Fonte Nao e Microsoft Store (APPINSTALLER_CLI_ERROR_SOURCE_NOT_STORE)'
        Causa     = 'Invocacao de recurso exclusivo da Microsoft Store em fonte diferente.'
        Solucoes  = @(
            'Garantir que a fonte msstore esteja habilitada caso o pacote dependa da Loja.'
        )
    }
    '0x8A15001C' = @{
        Categoria = 'Busca no Catalogo Central Falhou (APPINSTALLER_CLI_ERROR_NO_APPLICABLE_CORE_SEARCH)'
        Causa     = 'Indice de busca interno do WinGet corrompido ou indisponivel.'
        Solucoes  = @(
            'Recriar o indice executando: winget source update.'
        )
    }
    '0x8A15001D' = @{
        Categoria = 'Localizacao Nao Encontrada na Fonte (APPINSTALLER_CLI_ERROR_LOCATION_NOT_IN_SOURCE)'
        Causa     = 'O endereco de download do binario nao foi localizado na fonte configurada.'
        Solucoes  = @(
            'Atualizar indices com winget source update.',
            'Ignorar pacote no lote com -ExcludePackages.'
        )
    }
    '0x8A15001E' = @{
        Categoria = 'Recurso Experimental Desativado (APPINSTALLER_CLI_ERROR_NO_APPLICABLE_FEATURE)'
        Causa     = 'A operacao solicitada depende de uma funcionalidade experimental desativada.'
        Solucoes  = @(
            'Verificar o arquivo de configuracoes em: winget settings.'
        )
    }
    '0x8A15001F' = @{
        Categoria = 'Hash de Integridade Divergente (APPINSTALLER_CLI_ERROR_INSTALLER_HASH_MISMATCH)'
        Causa     = 'O hash SHA256 do instalador baixado nao confere com o registrado no manifesto oficial. O fornecedor alterou o binario sem atualizar o manifesto.'
        Solucoes  = @(
            'Executar winget source update para obter a versao mais recente do manifesto.',
            'Aguardar que o mantenedor do pacote atualize o hash no repositorio winget-pkgs.',
            'Para nao travar a rotina em lote, ignore este pacote com -ExcludePackages.'
        )
    }
    '0x8A150020' = @{
        Categoria = 'Bloqueio de Seguranca SmartScreen (APPINSTALLER_CLI_ERROR_INSTALLER_SECURITY_CHECK_FAILED)'
        Causa     = 'O Windows Defender SmartScreen bloqueou o instalador por falta de reputacao ou assinatura suspeita.'
        Solucoes  = @(
            'Verificar a integridade e procedencia do software.',
            'Para instalar manualmente, baixe o executavel do site oficial e confirme no prompt.',
            'Adicionar a lista de excecao com -ExcludePackages.'
        )
    }
    '0x8A150021' = @{
        Categoria = 'Divergencia no Tamanho do Download (APPINSTALLER_CLI_ERROR_DOWNLOAD_SIZE_MISMATCH)'
        Causa     = 'O tamanho do arquivo baixado difere do especificado pelo manifesto (download truncado).'
        Solucoes  = @(
            'Verificar estabilidade da conexao de rede.',
            'Executar novamente o processo para efetuar novo download.'
        )
    }
    '0x8A150022' = @{
        Categoria = 'Terminal Nao Suportado (APPINSTALLER_CLI_ERROR_APP_TERMINAL_NOT_SUPPORTED)'
        Causa     = 'O terminal de execucao nao suporta a operacao de console interativa.'
        Solucoes  = @(
            'Utilizar o Windows Terminal ou PowerShell com privilegios administrativos.'
        )
    }
    '0x8A150023' = @{
        Categoria = 'Nenhum Instalador Compativel (APPINSTALLER_CLI_ERROR_NO_APPLICABLE_INSTALLER)'
        Causa     = 'Nenhum instalador do pacote e compativel com a arquitetura do processador (x64/x86/ARM64) ou versao do Windows da maquina.'
        Solucoes  = @(
            'Verificar se o aplicativo possui suporte para a versao do Windows desta estacao.',
            'Adicionar a lista de excecoes permanentes com -ExcludePackages.'
        )
    }
    '0x8A150024' = @{
        Categoria = 'Dependencia Obrigatoria Ausente (APPINSTALLER_CLI_ERROR_INSTALL_MISSING_DEPENDENCY)'
        Causa     = 'O pacote exige pre-requisitos (como .NET Runtime, DirectX ou VCRedist) que nao estao disponiveis.'
        Solucoes  = @(
            'Identificar as dependencias necessarias e instala-las previamente na estacao.',
            'Verificar a instalacao manual dos componentes redistribuiveis da Microsoft.'
        )
    }
    '0x8A150025' = @{
        Categoria = 'Pacote Mutuamente Exclusivo (APPINSTALLER_CLI_ERROR_INSTALL_MUTUALLY_EXCLUSIVE)'
        Causa     = 'Existe um software concorrente ou variante incompativel ja instalada na maquina.'
        Solucoes  = @(
            'Desinstalar a versao conflitante antes de proceder com a nova instalacao.',
            'Adicionar o pacote as excecoes com -ExcludePackages.'
        )
    }
    '0x8A150026' = @{
        Categoria = 'Falha ao Desinstalar Versao Anterior (APPINSTALLER_CLI_ERROR_UNINSTALL_FAILED)'
        Causa     = 'O assistente de desinstalacao da versao antiga falhou durante a rotina de atualizacao.'
        Solucoes  = @(
            'Remover a versao anterior manualmente pelo Painel de Controle -> Programas e Recursos.',
            'Utilizar ferramenta de limpeza de instaladores e tentar novamente.'
        )
    }
    '0x8A150027' = @{
        Categoria = 'Pacote Ja Instalado (APPINSTALLER_CLI_ERROR_PACKAGE_ALREADY_INSTALLED)'
        Causa     = 'A versao exata deste pacote ja se encontra presente no sistema.'
        Solucoes  = @(
            'Nenhuma acao necessaria: o software ja esta em sua versao atual.'
        )
    }
    '0x8A150028' = @{
        Categoria = 'Fonte de Dependencia Indisponivel (APPINSTALLER_CLI_ERROR_DEPENDENCY_SOURCE_UNAVAILABLE)'
        Causa     = 'O repositorio de onde a dependencia do pacote deveria ser baixada esta inacessivel.'
        Solucoes  = @(
            'Verificar conectividade com a internet e fontes com winget source list.',
            'Tentar novamente mais tarde.'
        )
    }
    '0x8A150029' = @{
        Categoria = 'Verificacao de Status Falhou (APPINSTALLER_CLI_ERROR_STATUS_CHECK_FAILED)'
        Causa     = 'A checagem pos-instalacao nao pode confirmar se o aplicativo foi gravado com sucesso no registro.'
        Solucoes  = @(
            'Verificar se o aplicativo esta funcional no menu Iniciar.',
            'Reiniciar a estacao caso haja pendencia de gravacao de arquivos.'
        )
    }
    '0x8A15002A' = @{
        Categoria = 'Recurso Nao Encontrado (APPINSTALLER_CLI_ERROR_NOT_FOUND)'
        Causa     = 'Arquivo, chave de registro ou pasta de destino nao foi localizada durante o fluxo de instalacao.'
        Solucoes  = @(
            'Verificar a integridade do sistema operacional executando: sfc /scannow.'
        )
    }
    '0x8A15002B' = @{
        Categoria = 'Atualizacao Nao Aplicavel (APPINSTALLER_CLI_ERROR_UPDATE_NOT_APPLICABLE)'
        Causa     = 'Nao ha nenhuma atualizacao aplicavel disponivel para este pacote no momento.'
        Solucoes  = @(
            'O aplicativo ja esta na versao mais atual suportada pela maquina.'
        )
    }
    '0x8A15002C' = @{
        Categoria = 'Pacote Nao Instalado (APPINSTALLER_CLI_ERROR_PACKAGE_NOT_INSTALLED)'
        Causa     = 'Tentativa de atualizar um aplicativo que nao esta instalado na maquina.'
        Solucoes  = @(
            'Para instalar o aplicativo pela primeira vez, utilize o parametro -InstallPackage.'
        )
    }
    '0x8A15002D' = @{
        Categoria = 'Versao Fixada por Pinning (APPINSTALLER_CLI_ERROR_PIN_EXISTS)'
        Causa     = 'Uma regra de fixacao de versao (Pin) impede que este aplicativo seja atualizado pelo WinGet.'
        Solucoes  = @(
            'Consultar os pins ativos executando na maquina: winget pin list.',
            'Para liberar a atualizacao do pacote, execute: winget pin remove --id <ID>.'
        )
    }
    '0x8A15002E' = @{
        Categoria = 'Banco de Dados de Pins Inacessivel (APPINSTALLER_CLI_ERROR_PIN_CANNOT_BE_OPENED)'
        Causa     = 'Falha de permissao ou corrupcao no banco de dados de version pinning do WinGet.'
        Solucoes  = @(
            'Verificar permissoes em %LOCALAPPDATA%\Microsoft\WinGet.'
        )
    }
    '0x8A15002F' = @{
        Categoria = 'Indice de Pinning Inacessivel (APPINSTALLER_CLI_ERROR_CANNOT_OPEN_PINNING_INDEX)'
        Causa     = 'Nao foi possivel carregar a tabela de versoes fixadas do WinGet.'
        Solucoes  = @(
            'Redefinir as configuracoes locais do WinGet.'
        )
    }
    '0x8A150030' = @{
        Categoria = 'Pin Nao Encontrado (APPINSTALLER_CLI_ERROR_NO_PIN_FOUND)'
        Causa     = 'Tentativa de desfixar um pacote que nao possui regra de pin ativa.'
        Solucoes  = @(
            'Nenhuma acao necessaria.'
        )
    }
    '0x8A150031' = @{
        Categoria = 'Falha em Ferramenta Externa (APPINSTALLER_CLI_ERROR_EXTERNAL_TOOL_FAILED)'
        Causa     = 'Utilitario auxiliar invocado pelo instalador falhou durante a execucao.'
        Solucoes  = @(
            'Verificar se softwares antivirus ou EDR bloquearam a ferramenta auxiliar na pasta Temp.',
            'Consultar o visualizador de eventos do Windows.'
        )
    }
    '0x8A150032' = @{
        Categoria = 'Falha ao Desinstalar Pacote Portatil (APPINSTALLER_CLI_ERROR_PORTABLE_UNINSTALL_FAILED)'
        Causa     = 'Nao foi possivel remover arquivos ou links simbolicos de aplicativo portatil.'
        Solucoes  = @(
            'Verificar se o executavel portatil esta em execucao e encerrar o processo.'
        )
    }
    '0x8A150033' = @{
        Categoria = 'Falha ao Instalar Pacote Portatil (APPINSTALLER_CLI_ERROR_PORTABLE_INSTALL_FAILED)'
        Causa     = 'Falha ao gravar o executavel na pasta de links do WinGet ou criar atalho no PATH.'
        Solucoes  = @(
            'Verificar espaco livre em disco e permissoes de escrita no perfil de usuario.'
        )
    }
    '0x8A150034' = @{
        Categoria = 'Ciclo de Dependencias Detectado (APPINSTALLER_CLI_ERROR_DEPENDENCY_CYCLE)'
        Causa     = 'Os pacotes dependem circularmente uns dos outros nos manifestos da comunidade.'
        Solucoes  = @(
            'Instalar os pacotes componentes isoladamente.',
            'Reportar o problema de dependencia circular no repositorio winget-pkgs.'
        )
    }
    '0x8A150035' = @{
        Categoria = 'Outra Instalacao WinGet em Execucao (APPINSTALLER_CLI_ERROR_ANOTHER_INSTALL_IN_PROGRESS)'
        Causa     = 'Existe outra instancia ativa do WinGet bloqueando o acesso exclusivo ao banco de dados.'
        Solucoes  = @(
            'Aguardar alguns minutos ate a conclusao do processo anterior.',
            'Verificar se ha processos winget.exe travados e encerra-los se necessario.'
        )
    }
    '0x8A150036' = @{
        Categoria = 'Arquivo Excede Tamanho Permitido (APPINSTALLER_CLI_ERROR_FILE_TOO_LARGE)'
        Causa     = 'O arquivo de download ultrapassa os limites configurados na politica de sistema.'
        Solucoes  = @(
            'Verificar configuracoes de cota ou espaco livre na unidade do sistema.'
        )
    }
    '0x8A150037' = @{
        Categoria = 'Falha na Exportacao do Catalogo (APPINSTALLER_CLI_ERROR_EXPORT_CATALOG_FAILED)'
        Causa     = 'Erro ao gerar arquivo JSON de exportacao de pacotes.'
        Solucoes  = @(
            'Verificar caminho e permissoes de escrita do arquivo de exportacao.'
        )
    }
    '0x8A150038' = @{
        Categoria = 'Instalador Proibe Elevacao Administrativa (APPINSTALLER_CLI_ERROR_INSTALLER_PROHIBITS_ELEVATION)'
        Causa     = 'O instalador foi programado pelo fornecedor para rodar estritamente no contexto de usuario comum e recusa execucao sob SYSTEM ou Administrador.'
        Solucoes  = @(
            'Executar a atualizacao direcionada ao escopo do usuario conectado interativamente.',
            'Garantir que o usuario esteja conectado no momento da execucao.',
            'Ignorar na atualizacao de maquina com -ExcludePackages.'
        )
    }
    '0x8A150039' = @{
        Categoria = 'Pacote Portatil Ja Existe (APPINSTALLER_CLI_ERROR_PORTABLE_PACKAGE_ALREADY_EXISTS)'
        Causa     = 'Conflito de nome com executavel portatil existente na pasta de links.'
        Solucoes  = @(
            'Remover atalho previo na pasta de links do WinGet.'
        )
    }
    '0x8A150040' = @{
        Categoria = 'Instalacao Interrompida (APPINSTALLER_CLI_ERROR_INSTALLER_INTERRUPTED)'
        Causa     = 'O instalador foi cancelado abruptamente durante a execucao de suas rotinas.'
        Solucoes  = @(
            'Tentar a execucao novamente.',
            'Verificar se o processo foi finalizado por timeout ou pelo usuario.'
        )
    }
    '0x8A150100' = @{
        Categoria = 'Validacao de Dependencias Falhou (APPINSTALLER_CLI_ERROR_DEPENDENCIES_VALIDATION_FAILED)'
        Causa     = 'O conjunto de pre-requisitos exigidos pelo pacote nao pode ser validado com sucesso.'
        Solucoes  = @(
            'Instalar previamente os frameworks necessarios (.NET Framework / Visual C++ Redistributable).'
        )
    }
    '0x8A150101' = @{
        Categoria = 'Dependencias Ausentes no Repositorio (APPINSTALLER_CLI_ERROR_DEPENDENCIES_MISSING)'
        Causa     = 'As dependencias declaradas pelo software nao estao disponiveis para download na fonte configurada.'
        Solucoes  = @(
            'Instalar os componentes exigidos diretamente do site oficial do fabricante.',
            'Adicionar o pacote as excecoes com -ExcludePackages.'
        )
    }
    '0x8A150104' = @{
        Categoria = 'Termos da Fonte Nao Aceitos (APPINSTALLER_CLI_ERROR_SOURCE_AGREEMENTS_NOT_ACCEPTED)'
        Causa     = 'A fonte de pacotes requer aceitacao explicita dos termos de servico da Microsoft Store ou repositorio.'
        Solucoes  = @(
            'O script ja envia --accept-source-agreements automaticamente. Se persistir, execute: winget source update --accept-source-agreements.'
        )
    }
    '0x8A150105' = @{
        Categoria = 'Termos do Pacote Nao Aceitos (APPINSTALLER_CLI_ERROR_PACKAGE_AGREEMENTS_NOT_ACCEPTED)'
        Causa     = 'O software requer aceitacao explicita do contrato de licenca (EULA) do fabricante.'
        Solucoes  = @(
            'O script ja envia --accept-package-agreements. Se exigir confirmacao interativa, instale em sessao de usuario.'
        )
    }
    '0x8A150107' = @{
        Categoria = 'Autenticacao de Proxy Necessaria (APPINSTALLER_CLI_ERROR_PROXY_AUTHENTICATION_REQUIRED)'
        Causa     = 'O servidor proxy corporativo da rede bloqueou a conexao exigindo credenciais de autenticacao (HTTP 407).'
        Solucoes  = @(
            'Configurar as credenciais do proxy ou liberar as URLs do WinGet/CDN na politica do gateway de rede.',
            'Verificar as configuracoes de proxy da conta SYSTEM em: netsh winhttp show proxy.'
        )
    }
    '0x8A150110' = @{
        Categoria = 'Auto-Atualizacao do Winget Bloqueada / Dependencia MSIX (0x8A150110)'
        Causa     = 'O pacote Microsoft.AppInstaller e o proprio executavel do Winget. A tentativa de atualizar a si mesmo via linha de comando em background falha com o codigo 0x8A150110 (APPINSTALLER_CLI_ERROR_INSTALL_DEPENDENCIES) porque o binario esta em execucao e o subsistema AppX requer sessao interativa da Microsoft Store.'
        Solucoes  = @(
            'Nao e necessario atualizar manualmente: O Windows Update e a Microsoft Store atualizam o App Installer automaticamente em background.',
            'Adicionar a lista de exclusoes para ignorar a tentativa de auto-atualizacao do Winget: -ExcludePackages @("Microsoft.AppInstaller")',
            'Caso deseje atualizar a versao do Winget manualmente no host, execute no PowerShell como Administrador: Add-AppxPackage -Path "<caminho>\Microsoft.DesktopAppInstaller.msixbundle".'
        )
    }

    # --- 2. CODIGOS DO WINDOWS INSTALLER (MSI / MSIEXEC) ---
    '1601' = @{
        Categoria = 'Servico Windows Installer Inacessivel (MSI 1601)'
        Causa     = 'O servico Windows Installer (msiserver) esta desativado, corrompido ou inacessivel no sistema.'
        Solucoes  = @(
            'Iniciar o servico no PowerShell administrativo da maquina: Start-Service msiserver.',
            'Configurar o servico para inicio automatico/manual: Set-Service msiserver -StartupType Manual.',
            'Verificar a integridade do sistema com: sfc /scannow.'
        )
    }
    '1602' = @{
        Categoria = 'Instalacao Cancelada pelo Usuario (MSI 1602)'
        Causa     = 'A operacao de instalacao foi cancelada manualmente pelo usuario ou por sinal de interrupcao.'
        Solucoes  = @(
            'Executar o instalador silencioso sem intervencao de interface do usuario.',
            'Repetir a operacao certificando-se de que nenhum outro processo encerre o msiexec.'
        )
    }
    '1603' = @{
        Categoria = 'Erro Fatal do Windows Installer (MSI 1603)'
        Causa     = 'O instalador MSI encontrou um erro irrecuperavel durante uma Custom Action, falha ao iniciar servico, falta de permissao em pasta do sistema ou arquivo bloqueado.'
        Solucoes  = @(
            'Testar a instalacao individual do aplicativo para capturar o log detalhado.',
            'Verificar o visualizador de eventos do Windows (Event Viewer -> Application) na maquina remota.',
            'Adicionar o pacote a lista de excecao com -ExcludePackages.'
        )
    }
    '1604' = @{
        Categoria = 'Instalacao Suspensa Incompleta (MSI 1604)'
        Causa     = 'Uma transacao anterior do Windows Installer foi suspensa antes de ser finalizada.'
        Solucoes  = @(
            'Reiniciar a maquina de destino para que o Windows finalize transacoes pendentes de instalacao.',
            'Limpar a chave de registro InProgress em HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer se necessario.'
        )
    }
    '1605' = @{
        Categoria = 'Acao Valida Apenas para Produtos Instalados (MSI 1605)'
        Causa     = 'A acao de desinstalacao ou patch foi invocada para um aplicativo que nao esta registrado no catalogo MSI.'
        Solucoes  = @(
            'Verificar se o aplicativo foi removido previamente por outro metodo.',
            'Reinstalar o aplicativo para reestruturar o registro do componente.'
        )
    }
    '1606' = @{
        Categoria = 'Recurso Nao Registrado (MSI 1606)'
        Causa     = 'ID de recurso interno desconhecido no banco de dados MSI.'
        Solucoes  = @(
            'Baixar novamente o pacote do fornecedor para garantir integridade.'
        )
    }
    '1607' = @{
        Categoria = 'Componente Nao Registrado (MSI 1607)'
        Causa     = 'Identificador de componente do Windows Installer corrompido ou ausente.'
        Solucoes  = @(
            'Executar reparo da instalacao do produto atraves do Painel de Controle.'
        )
    }
    '1608' = @{
        Categoria = 'Propriedade Desconhecida no MSI (MSI 1608)'
        Causa     = 'Propriedade interna do instalador nao declarada no esquema.'
        Solucoes  = @(
            'Entrar em contato com o desenvolvedor do software para correcao do instalador.'
        )
    }
    '1610' = @{
        Categoria = 'Dados de Configuracao Corrompidos (MSI 1610)'
        Causa     = 'As informacoes de registro do produto em HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer estao danificadas.'
        Solucoes  = @(
            'Utilizar a ferramenta oficial da Microsoft Program Install and Uninstall Troubleshooter para reparar o registro.',
            'Reinstalar a versao completa do aplicativo.'
        )
    }
    '1612' = @{
        Categoria = 'Origem de Instalacao Indisponivel (MSI 1612)'
        Causa     = 'O arquivo .MSI original correspondente a versao anteriormente instalada foi removido da pasta de cache C:\Windows\Installer.'
        Solucoes  = @(
            'Baixar o instalador da versao anterior para apontar a origem ou utilizar o utilitario Microsoft Troubleshooter.',
            'Adicionar a lista de excecoes para nao travar o lote com -ExcludePackages.'
        )
    }
    '1613' = @{
        Categoria = 'Versao do Mecanismo MSI Incompativel (MSI 1613)'
        Causa     = 'O pacote MSI requer uma versao mais nova do Windows Installer do que a suportada pelo sistema.'
        Solucoes  = @(
            'Instalar as atualizacoes mais recentes do Windows Update na estacao.'
        )
    }
    '1614' = @{
        Categoria = 'Produto Desinstalado (MSI 1614)'
        Causa     = 'O produto foi desinstalado e nao pode receber acoes de modificacao.'
        Solucoes  = @(
            'Efetuar uma instalacao limpa utilizando -InstallPackage.'
        )
    }
    '1615' = @{
        Categoria = 'Sintaxe SQL Invalida no MSI (MSI 1615)'
        Causa     = 'Erro interno na tabela de consultas do banco de dados do instalador.'
        Solucoes  = @(
            'Obter instalador atualizado junto ao fabricante.'
        )
    }
    '1618' = @{
        Categoria = 'Instalador Ocupado (MSI 1618 - ERROR_INSTALL_ALREADY_RUNNING)'
        Causa     = 'Outro processo de instalacao (msiexec.exe ou Windows Update) ja esta em execucao no sistema bloqueando o mutex global.'
        Solucoes  = @(
            'Aguardar alguns minutos ate a conclusao da instalacao em andamento.',
            'Se o processo estiver travado, encerre o msiexec no PowerShell: Stop-Process -Name msiexec -Force.',
            'Tentar novamente o comando apos liberar o processo.'
        )
    }
    '1619' = @{
        Categoria = 'Pacote MSI Nao Pode Ser Aberto (MSI 1619)'
        Causa     = 'O arquivo de instalacao MSI nao foi localizado no caminho indicado ou a conta nao possui permissao de leitura.'
        Solucoes  = @(
            'Verificar se o arquivo baixado na pasta temporaria esta acessivel pela conta SYSTEM.',
            'Checar se softwares antivirus moveram o arquivo para quarentena.'
        )
    }
    '1620' = @{
        Categoria = 'Pacote de Instalacao Invalido (MSI 1620)'
        Causa     = 'O arquivo .msi esta corrompido ou nao e um pacote de instalacao legitimo.'
        Solucoes  = @(
            'Limpar o cache temporario e realizar novo download.',
            'Verificar se o link do fabricante entrega um instalador valido.'
        )
    }
    '1621' = @{
        Categoria = 'Falha ao Iniciar UI do Instalador (MSI 1621)'
        Causa     = 'Ocorreu um erro ao carregar a interface de usuario do instalador em uma sessao nao interativa.'
        Solucoes  = @(
            'Garantir a flag silenciosa na instalacao (--silent /qn).',
            'Executar a instalacao interativa caso o instalador rejeite modo headless.'
        )
    }
    '1622' = @{
        Categoria = 'Falha ao Abrir Arquivo de Log MSI (MSI 1622)'
        Causa     = 'O msiexec nao conseguiu criar o arquivo de log especificado na pasta temporaria.'
        Solucoes  = @(
            'Verificar permissoes de escrita na pasta C:\Windows\Temp.',
            'Garantir que a particao do sistema nao esteja cheia.'
        )
    }
    '1623' = @{
        Categoria = 'Idioma Nao Suportado pelo Instalador (MSI 1623)'
        Causa     = 'O idioma padrao do sistema operacional nao e aceito pelo pacote MSI.'
        Solucoes  = @(
            'Verificar pacotes multilingues disponibilizados pelo fornecedor.'
        )
    }
    '1624' = @{
        Categoria = 'Falha ao Aplicar Transformacao MST (MSI 1624)'
        Causa     = 'O arquivo de transformacao .mst esta corrompido ou e incompativel com a versao do MSI.'
        Solucoes  = @(
            'Verificar o arquivo de customizacao MST utilizado pelo administrador.'
        )
    }
    '1625' = @{
        Categoria = 'Instalacao Bloqueada por Diretiva de Sistema (MSI 1625)'
        Causa     = 'A instalacao foi proibida por diretivas de grupo (Software Restriction Policies ou AppLocker).'
        Solucoes  = @(
            'Verificar a chave de registro: HKLM:\SOFTWARE\Policies\Microsoft\Windows\Installer -> DisableMSI.',
            'Adicionar regra de liberacao no AppLocker para o certificado do fornecedor.'
        )
    }
    '1626' = @{
        Categoria = 'Funcao Nao Pode Ser Executada (MSI 1626)'
        Causa     = 'Falha de chamada de API interna do Windows Installer.'
        Solucoes  = @(
            'Reiniciar o servico msiserver e tentar novamente.'
        )
    }
    '1627' = @{
        Categoria = 'Funcao Interna Falhou (MSI 1627)'
        Causa     = 'Erro de execucao em rotina interna da tabela do instalador.'
        Solucoes  = @(
            'Testar com versao mais recente do aplicativo.'
        )
    }
    '1628' = @{
        Categoria = 'Tabela Invalida no Banco MSI (MSI 1628)'
        Causa     = 'Estrutura relacional do banco de dados MSI corrompida.'
        Solucoes  = @(
            'Baixar novamente o pacote do fornecedor.'
        )
    }
    '1629' = @{
        Categoria = 'Incompatibilidade de Tipo de Dados (MSI 1629)'
        Causa     = 'Parametro passado com tipo divergente do esperado pelo banco MSI.'
        Solucoes  = @(
            'Verificar argumentos customizados enviados ao instalador.'
        )
    }
    '1630' = @{
        Categoria = 'Tipo de Dados Nao Suportado (MSI 1630)'
        Causa     = 'Estrutura de dados nao suportada pela versao atual do Windows Installer.'
        Solucoes  = @(
            'Atualizar o Windows com os ultimos patches cumulativos.'
        )
    }
    '1631' = @{
        Categoria = 'Falha ao Inicializar Windows Installer (MSI 1631)'
        Causa     = 'O servico Windows Installer nao pode ser iniciado por falta de memoria ou arquivo corrompido.'
        Solucoes  = @(
            'Verificar a integridade dos arquivos do sistema: sfc /scannow.',
            'Reiniciar o computador.'
        )
    }
    '1632' = @{
        Categoria = 'Pasta Temporaria Bloqueada ou Inacessivel (MSI 1632)'
        Causa     = 'A pasta temporaria esta em uma unidade cheia ou sem privilegios de gravacao para o instalador.'
        Solucoes  = @(
            'Verificar espaco livre no disco C:. Liberar espaco se estiver abaixo de 5 GB.',
            'Conferir permissoes de escrita na pasta C:\Windows\Temp.'
        )
    }
    '1633' = @{
        Categoria = 'Plataforma Nao Suportada pelo Pacote (MSI 1633)'
        Causa     = 'O pacote MSI foi compilado para outra arquitetura de processador (ex: x86 em ARM64 ou 64-bit em 32-bit).'
        Solucoes  = @(
            'Baixar a variante compatcompativel com a arquitetura do processador da estacao.',
            'Adicionar a lista de excecao com -ExcludePackages.'
        )
    }
    '1634' = @{
        Categoria = 'Componente Nao Utilizado (MSI 1634)'
        Causa     = 'Acao referenciando componente inativo nesta instalacao.'
        Solucoes  = @(
            'Executar o instalador padrao sem propriedades adicionais.'
        )
    }
    '1635' = @{
        Categoria = 'Pacote de Atualizacao MSP Nao Pode Ser Aberto (MSI 1635)'
        Causa     = 'O arquivo de patch (.msp) nao foi localizado ou esta corrompido.'
        Solucoes  = @(
            'Baixar novamente o patch ou executar instalacao limpa da versao completa.'
        )
    }
    '1636' = @{
        Categoria = 'Pacote de Atualizacao MSP Invalido (MSI 1636)'
        Causa     = 'O arquivo .msp esta danificado ou nao e compativel com o produto instalado.'
        Solucoes  = @(
            'Verificar a versao instalada antes de aplicar o patch.'
        )
    }
    '1637' = @{
        Categoria = 'Patch Nao Suportado pelo Instalador (MSI 1637)'
        Causa     = 'O servico do Windows Installer nao possui suporte para o tipo de patch fornecido.'
        Solucoes  = @(
            'Atualizar o Windows Installer via Windows Update.'
        )
    }
    '1638' = @{
        Categoria = 'Conflito de Versao Instalada (MSI 1638)'
        Causa     = 'Outra versao deste produto ja esta instalada. O instalador MSI nao suporta upgrade direto sobre a versao atual sem desinstalacao previa.'
        Solucoes  = @(
            'Desinstalar a versao legada da estacao atraves do Painel de Controle ou WMI.',
            'Instalar a nova versao limpa com -InstallPackage.',
            'Adicionar a lista de excecao com -ExcludePackages.'
        )
    }
    '1639' = @{
        Categoria = 'Linha de Comando do msiexec Invalida (MSI 1639)'
        Causa     = 'Argumento ou sintaxe de parametro incorreta passada para o msiexec.'
        Solucoes  = @(
            'Verificar aspas e parametros de linha de comando enviados ao instalador.'
        )
    }
    '1640' = @{
        Categoria = 'Instalacao Remota Nao Permitida (MSI 1640)'
        Causa     = 'Instalacoes via Terminal Server / Sessao Remota sao restritas a administradores ou proibidas por politica.'
        Solucoes  = @(
            'Habilitar modo de instalacao no Terminal Server executando: change user /install.',
            'Retornar ao modo de execucao apos a instalacao: change user /execute.'
        )
    }
    '1641' = @{
        Categoria = 'Reinicializacao do Sistema Iniciada (MSI 1641)'
        Causa     = 'A instalacao foi concluida com sucesso e iniciou o reinicio imediato da maquina.'
        Solucoes  = @(
            'Aguardar a estacao reiniciar para que as modificacoes sejam concluidas.'
        )
    }
    '1642' = @{
        Categoria = 'Destino do Patch Nao Encontrado (MSI 1642)'
        Causa     = 'O software a ser atualizado pelo patch MSP nao foi encontrado ou esta em uma versao diferente da exigida.'
        Solucoes  = @(
            'Confirmar se o produto base esta instalado antes de aplicar a atualizacao.'
        )
    }
    '1643' = @{
        Categoria = 'Patch Bloqueado por Diretiva (MSI 1643)'
        Causa     = 'O patch foi bloqueado por politica de restricao de software (SRP/GPO).'
        Solucoes  = @(
            'Liberar a aplicacao de patches na politica corporativa.'
        )
    }
    '1644' = @{
        Categoria = 'Transformacao Rejeitada por Diretiva (MSI 1644)'
        Causa     = 'A transformacao MST nao e permitida pelas regras de seguranca da maquina.'
        Solucoes  = @(
            'Assinar digitalmente o arquivo MST ou revisar as diretivas de restricao.'
        )
    }
    '1648' = @{
        Categoria = 'Instalacao Remota Proibida (MSI 1648)'
        Causa     = 'O pacote proibe instalacao atraves de chamada remota nao interativa.'
        Solucoes  = @(
            'Executar a instalacao localmente na maquina ou via console interativo.'
        )
    }
    '1650' = @{
        Categoria = 'Remocao de Patch Nao Suportada (MSI 1650)'
        Causa     = 'O patch aplicado nao suporta desinstalacao.'
        Solucoes  = @(
            'Reinstalar o produto base caso deseje remover a modificacao.'
        )
    }
    '1654' = @{
        Categoria = 'Condicao de Execucao Nao Atendida (MSI 1654)'
        Causa     = 'Uma condicao de pre-requisito (LaunchCondition) definida pelo desenvolvedor no MSI falhou.'
        Solucoes  = @(
            'Verificar os requisitos de sistema (versao minima de build do Windows, memoria ou componentes).'
        )
    }
    '1714' = @{
        Categoria = 'Falha ao Remover Versao Anterior (MSI 1714)'
        Causa     = 'A versao mais antiga deste produto nao pode ser desinstalada durante o upgrade por ausencia do MSI original.'
        Solucoes  = @(
            'Utilizar o Microsoft Program Install and Uninstall Troubleshooter para limpar registros antigos.',
            'Ignorar este pacote no lote com -ExcludePackages.'
        )
    }
    '3010' = @{
        Categoria = 'Reinicializacao Necessaria (MSI 3010 - ERROR_SUCCESS_REBOOT_REQUIRED)'
        Causa     = 'A instalacao foi realizada com sucesso, mas arquivos protegidos foram agendados para substituicao na proxima reinicializacao.'
        Solucoes  = @(
            'Agendar o reinicio da estacao fora do horario comercial para concluir a aplicacao das alteracoes.',
            'O aplicativo ja se encontra instalado e funcionara plenamente apos o reboot.'
        )
    }

    # --- 3. CODIGOS DE SISTEMA WIN32 E SISTEMA OPERACIONAL ---
    '1' = @{
        Categoria = 'Funcao Incorreta (Win32 1 - ERROR_INVALID_FUNCTION)'
        Causa     = 'Operacao invalida solicitada ao sistema pelo executavel do instalador.'
        Solucoes  = @(
            'Verificar compatibilidade do software com a edicao do sistema operacional.'
        )
    }
    '2' = @{
        Categoria = 'Arquivo Nao Encontrado (Win32 2 - ERROR_FILE_NOT_FOUND)'
        Causa     = 'O sistema nao encontrou o arquivo executavel ou instalador no caminho especificado.'
        Solucoes  = @(
            'Verificar se o arquivo baixado foi removido por antivirus ou expirou na pasta Temp.'
        )
    }
    '3' = @{
        Categoria = 'Caminho Nao Encontrado (Win32 3 - ERROR_PATH_NOT_FOUND)'
        Causa     = 'A pasta de destino da instalacao ou diretorio temporario nao existe.'
        Solucoes  = @(
            'Garantir a existencia das pastas C:\Windows\Temp e do diretorio Program Files.'
        )
    }
    '4' = @{
        Categoria = 'Muitos Arquivos Abertos (Win32 4 - ERROR_TOO_MANY_OPEN_FILES)'
        Causa     = 'Esgotamento de descritores de arquivos abertos simultaneamente.'
        Solucoes  = @(
            'Reiniciar a maquina para liberar handles de arquivos pendentes.'
        )
    }
    '5' = @{
        Categoria = 'Acesso Negado / Permissao Insuficiente (Win32 5 - ERROR_ACCESS_DENIED)'
        Causa     = 'A conta em execucao nao possui privilegios de gravacao na pasta de instalacao, registro ou servico.'
        Solucoes  = @(
            'Executar o script com credenciais de Administrador do Dominio ou Administrador Local.',
            'Garantir que a tarefa agendada esteja rodando com o maior nivel de privilegio (HighestRunLevel).'
        )
    }
    '8' = @{
        Categoria = 'Memoria Insuficiente (Win32 8 - ERROR_NOT_ENOUGH_MEMORY)'
        Causa     = 'O sistema nao dispoe de memoria RAM livre suficiente para carregar o instalador.'
        Solucoes  = @(
            'Fechar aplicativos pesados em execucao na maquina de destino ou expandir o arquivo de paginacao.'
        )
    }
    '13' = @{
        Categoria = 'Dados Invalidos (Win32 13 - ERROR_INVALID_DATA)'
        Causa     = 'Estrutura de dados ou arquivo corrompido durante a descompactacao.'
        Solucoes  = @(
            'Baixar novamente o pacote limpando a pasta temporaria.'
        )
    }
    '14' = @{
        Categoria = 'Armazenamento Insuficiente (Win32 14 - ERROR_OUTOFMEMORY)'
        Causa     = 'Esgotamento de pool de memoria do kernel para alocacao de buffers.'
        Solucoes  = @(
            'Reiniciar a maquina para desfragmentar o pool de recursos de sistema.'
        )
    }
    '15' = @{
        Categoria = 'Unidade Invalida (Win32 15 - ERROR_INVALID_DRIVE)'
        Causa     = 'A letra da unidade de destino especificada para instalacao nao existe.'
        Solucoes  = @(
            'Verificar se o aplicativo esta configurado para instalar em uma unidade de rede nao mapeada.'
        )
    }
    '19' = @{
        Categoria = 'Midia Protegida Contra Gravacao (Win32 19 - ERROR_WRITE_PROTECT)'
        Causa     = 'O disco de destino esta em modo somente leitura ou protegido por politica.'
        Solucoes  = @(
            'Verificar atributos do volume com diskpart.'
        )
    }
    '21' = @{
        Categoria = 'Dispositivo Nao Pronto (Win32 21 - ERROR_NOT_READY)'
        Causa     = 'A unidade de armazenamento demorou a responder a chamada do sistema operacional.'
        Solucoes  = @(
            'Verificar saude do disco rigido/SSD com chkdsk.'
        )
    }
    '32' = @{
        Categoria = 'Violacao de Compartilhamento de Arquivo (Win32 32 - ERROR_SHARING_VIOLATION)'
        Causa     = 'O instalador tentou sobrescrever um arquivo executavel ou DLL que esta em uso por outro processo ativo.'
        Solucoes  = @(
            'Encerrar os processos em execucao relacionados ao aplicativo antes de prosseguir.',
            'Agendar a atualizacao fora do horario de trabalho ou reiniciar o computador.'
        )
    }
    '33' = @{
        Categoria = 'Violacao de Bloqueio de Arquivo (Win32 33 - ERROR_LOCK_VIOLATION)'
        Causa     = 'Outro processo bloqueou um trecho do arquivo necessario para a instalacao.'
        Solucoes  = @(
            'Identificar o processo bloqueador utilizando a ferramenta Handle ou Process Explorer.'
        )
    }
    '50' = @{
        Categoria = 'Requisicao de Rede Nao Suportada (Win32 50 - ERROR_NOT_SUPPORTED)'
        Causa     = 'O sistema operacional de destino nao implementa o recurso de rede solicitado.'
        Solucoes  = @(
            'Verificar configuracoes de protocolo SMB e firewall.'
        )
    }
    '53' = @{
        Categoria = 'Caminho de Rede Nao Encontrado (Win32 53 - ERROR_BAD_NETPATH)'
        Causa     = 'O recurso de rede ou compartilhamento UNC de origem nao pode ser localizado.'
        Solucoes  = @(
            'Verificar conectividade com o servidor de arquivos e resolucao DNS.'
        )
    }
    '67' = @{
        Categoria = 'Nome de Rede Nao Encontrado (Win32 67 - ERROR_BAD_NET_NAME)'
        Causa     = 'O servidor remoto nao reconheceu o nome do compartilhamento de rede.'
        Solucoes  = @(
            'Verificar o nome do compartilhamento no servidor de distribuicao.'
        )
    }
    '80' = @{
        Categoria = 'Arquivo Ja Existe (Win32 80 - ERROR_FILE_EXISTS)'
        Causa     = 'Um arquivo critico ja existe e o instalador nao tem permissao para sobrescrever.'
        Solucoes  = @(
            'Renomear ou mover o arquivo conflitante antes de atualizar.'
        )
    }
    '87' = @{
        Categoria = 'Parametro Incorreto (Win32 87 - ERROR_INVALID_PARAMETER)'
        Causa     = 'Sintaxe incorreta de argumento passada na chamada do executavel.'
        Solucoes  = @(
            'Revisar os parametros de instalacao do pacote.'
        )
    }
    '112' = @{
        Categoria = 'Espaco Insuficiente em Disco (Win32 112 - ERROR_DISK_FULL)'
        Causa     = 'O disco rigido onde esta localizada a pasta temporaria ou o diretorio de destino esta completamente cheio.'
        Solucoes  = @(
            'Executar limpeza de disco (cleanmgr) na estacao para liberar arquivos temporarios.',
            'Garantir pelo menos 5 GB a 10 GB livres na unidade C: para instalacoes em lote.'
        )
    }
    '123' = @{
        Categoria = 'Sintaxe de Nome de Arquivo Invalida (Win32 123 - ERROR_INVALID_NAME)'
        Causa     = 'Caracteres ilegais foram detectados no nome do caminho ou arquivo.'
        Solucoes  = @(
            'Evitar o uso de caracteres especiais ou caminhos excessivamente longos.'
        )
    }
    '145' = @{
        Categoria = 'Diretorio Nao Vazio (Win32 145 - ERROR_DIR_NOT_EMPTY)'
        Causa     = 'Falha ao remover pasta de instalacao anterior porque continha arquivos residuais de usuario.'
        Solucoes  = @(
            'Limpar arquivos residuais na pasta de instalacao anterior antes de reinstalar.'
        )
    }
    '183' = @{
        Categoria = 'Arquivo ou Chave Ja Existe (Win32 183 - ERROR_ALREADY_EXISTS)'
        Causa     = 'Nao foi possivel criar o objeto porque ele ja existe no registro ou sistema de arquivos.'
        Solucoes  = @(
            'Executar o instalador com parametro de reparo ou desinstalar previamente.'
        )
    }
    '259' = @{
        Categoria = 'Nenhum Item Adicional (Win32 259 - ERROR_NO_MORE_ITEMS)'
        Causa     = 'Fim de enumeracao de dados de instalacao.'
        Solucoes  = @(
            'Operacao concluida normalmente.'
        )
    }
    '1058' = @{
        Categoria = 'Servico Desativado (Win32 1058 - ERROR_SERVICE_DISABLED)'
        Causa     = 'O servico de sistema exigido pelo software esta com o tipo de inicializacao Desativado.'
        Solucoes  = @(
            'Alterar o tipo de inicializacao do servico afetado para Manual ou Automatico e iniciar o servico.'
        )
    }
    '1060' = @{
        Categoria = 'Servico Nao Existe (Win32 1060 - ERROR_SERVICE_DOES_NOT_EXIST)'
        Causa     = 'O instalador tentou configurar ou iniciar um servico que nao se encontra registrado no Windows.'
        Solucoes  = @(
            'Verificar se o aplicativo requer um servico de base previamente instalado.'
        )
    }
    '1062' = @{
        Categoria = 'Servico Nao Foi Iniciado (Win32 1062 - ERROR_SERVICE_NOT_ACTIVE)'
        Causa     = 'Falha ao enviar comando de controle a um servico inativo.'
        Solucoes  = @(
            'Iniciar manualmente o servico atraves do console services.msc.'
        )
    }
    '1223' = @{
        Categoria = 'Operacao Cancelada pelo Usuario (Win32 1223 - ERROR_CANCELLED)'
        Causa     = 'O assistente de instalacao foi fechado pelo operador da maquina.'
        Solucoes  = @(
            'Garantir a execucao com argumentos silenciosos (/S, /qn, /silent).'
        )
    }
    '1260' = @{
        Categoria = 'Programa Bloqueado por Diretiva de Grupo (Win32 1260 - ERROR_KM_DRIVER_BLOCKED)'
        Causa     = 'O programa foi bloqueado por uma politica de restricao de software (GPO / AppLocker / SmartScreen).'
        Solucoes  = @(
            'Verificar as regras de Restricao de Software em secpol.msc na maquina de destino.',
            'Adicionar uma regra de excecao para o hash ou certificado do instalador.'
        )
    }
    '1450' = @{
        Categoria = 'Recursos de Sistema Insuficientes (Win32 1450 - ERROR_NO_SYSTEM_RESOURCES)'
        Causa     = 'Esgotamento de recursos de paginacao de pool nao paginavel do sistema operacional.'
        Solucoes  = @(
            'Reiniciar o computador para liberar recursos de kernel.'
        )
    }
    '1460' = @{
        Categoria = 'Tempo Limite de Operacao Expirado (Win32 1460 - ERROR_TIMEOUT)'
        Causa     = 'A chamada do instalador demorou mais tempo para responder do que o limite permitido.'
        Solucoes  = @(
            'Aumentar o tempo limite atraves do parametro -TaskTimeoutSeconds 1800.'
        )
    }

    # --- 4. CODIGOS HRESULT, REDE, SSL E CERTIFICADOS ---
    '0x80004005' = @{
        Categoria = 'Erro Generico do Windows (E_FAIL - 0x80004005)'
        Causa     = 'Falha nao especificada retornada por componente COM ou API nativa do Windows durante a instalacao.'
        Solucoes  = @(
            'Verificar os logs de instalacao do aplicativo para detalhamento da falha.',
            'Executar sfc /scannow para validar arquivos protegidos do Windows.'
        )
    }
    '0x80004002' = @{
        Categoria = 'Interface COM Nao Suportada (E_NOINTERFACE - 0x80004002)'
        Causa     = 'O instalador tentou interagir com uma versao incompativel de DLL do sistema.'
        Solucoes  = @(
            'Atualizar as bibliotecas C++ Runtime e .NET Framework da estacao.'
        )
    }
    '0x80004004' = @{
        Categoria = 'Operacao Abortada (E_ABORT - 0x80004004)'
        Causa     = 'A rotina de execucao foi finalizada antes de concluir todas as etapas.'
        Solucoes  = @(
            'Executar novamente a instalacao individualmente.'
        )
    }
    '0x80040154' = @{
        Categoria = 'Classe COM Nao Registrada (REGDB_E_CLASSNOTREG - 0x80040154)'
        Causa     = 'Um componente COM necessario para o funcionamento do instalador nao esta registrado no sistema.'
        Solucoes  = @(
            'Registrar as DLLs necessarias com regsvr32 ou reinstalar o pacote redistribuivel da aplicacao.'
        )
    }
    '0x80070005' = @{
        Categoria = 'Acesso Negado (E_ACCESSDENIED - 0x80070005)'
        Causa     = 'Permissoes insuficientes na conta de execucao para modificar pastas de sistema ou chaves HKLM.'
        Solucoes  = @(
            'Executar como Administrador ou fornecer credenciais administrativas com -Credential.'
        )
    }
    '0x8007000E' = @{
        Categoria = 'Memoria do Sistema Esgotada (E_OUTOFMEMORY - 0x8007000E)'
        Causa     = 'Falha ao alocar memoria RAM durante o processo de descompactacao.'
        Solucoes  = @(
            'Fechar programas que consomem muita memoria e tentar novamente.'
        )
    }
    '0x800704EC' = @{
        Categoria = 'Bloqueado por Diretiva de Seguranca (0x800704EC)'
        Causa     = 'Este programa foi bloqueado por uma politica de grupo da empresa (AppLocker ou SRP).'
        Solucoes  = @(
            'Verificar as diretivas corporativas em: gpedit.msc -> Politicas de Controle de Aplicativos.',
            'Solicitar inclusao do aplicativo nas regras de liberacao do AppLocker.'
        )
    }
    '0x80070643' = @{
        Categoria = 'Erro Fatal de Instalacao (0x80070643)'
        Causa     = 'Codigo HRESULT classico que encapsula erro fatal de instaladores MSI ou atualizacoes do Windows.'
        Solucoes  = @(
            'Verificar se o aplicativo ja possui uma versao mais nova ou corrompida instalada.',
            'Consultar o visualizador de eventos do Windows (Event Viewer) para a causa primaria.',
            'Adicionar a lista de excecao com -ExcludePackages.'
        )
    }
    '0x80070652' = @{
        Categoria = 'Outra Instalacao em Andamento (0x80070652)'
        Causa     = 'Mutex de instalacao bloqueado por outro processo em execucao.'
        Solucoes  = @(
            'Aguardar a conclusao do processo anterior ou finalizar instancias orfas de msiexec.exe.'
        )
    }
    '0x80072EE2' = @{
        Categoria = 'Tempo Limite de Conexao a Internet (0x80072EE2)'
        Causa     = 'A requisicao para o servidor de download expirou sem resposta (Timeout de rede).'
        Solucoes  = @(
            'Verificar se ha lentidao ou instabilidade no link de internet da estacao.',
            'Checar se ha filtros de conteudo web bloqueando a requisicao.'
        )
    }
    '0x80072EE7' = @{
        Categoria = 'Nome de Servidor Nao Resolvido via DNS (0x80072EE7)'
        Causa     = 'A estacao remota nao conseguiu resolver o dominio do fornecedor no servidor DNS configurado.'
        Solucoes  = @(
            'Verificar as configuracoes de DNS da placa de rede da estacao de destino.',
            'Limpar o cache DNS executando: Clear-DnsClientCache.',
            'Checar se a maquina possui saida direta para a internet.'
        )
    }
    '0x80072EFD' = @{
        Categoria = 'Falha ao Conectar ao Servidor Web (0x80072EFD)'
        Causa     = 'Nao foi possivel abrir a porta TCP 80/443 com o servidor do fornecedor (Firewall/Roteador bloqueando).'
        Solucoes  = @(
            'Verificar regras de saida de firewall corporativo para conexoes HTTPS.'
        )
    }
    '0x80072F0C' = @{
        Categoria = 'Certificado de Cliente SSL Necessario (0x80072F0C)'
        Causa     = 'O servidor web requer autenticacao mutua TLS (mTLS) com certificado de cliente.'
        Solucoes  = @(
            'Verificar regras de proxy reverso da rede corporativa.'
        )
    }
    '0x80072F78' = @{
        Categoria = 'Resposta HTTP Invalida do Servidor (0x80072F78)'
        Causa     = 'O servidor web do fornecedor enviou uma resposta malformada ou conexao foi resetada por proxy.'
        Solucoes  = @(
            'Verificar inspecao profunda de pacotes SSL em firewalls corporativos.'
        )
    }
    '0x80072F8F' = @{
        Categoria = 'Falha de Certificado SSL / Relogio Incorreto (0x80072F8F)'
        Causa     = 'A data e hora da estacao estao divergentes do servidor, invalidando o certificado TLS/SSL.'
        Solucoes  = @(
            'Sincronizar o relogio da maquina com o servidor de dominio NTP: w32tm /resync.',
            'Verificar fuso horario correto no Windows.'
        )
    }
    '0x80190190' = @{
        Categoria = 'Erro HTTP 400 - Requisicao Invalida (0x80190190)'
        Causa     = 'O servidor de download rejeitou a requisicao por cabecalhos ou formato invalido.'
        Solucoes  = @(
            'Atualizar o cliente WinGet para que envie requisicoes compativeis com a CDN do fabricante.'
        )
    }
    '0x80190191' = @{
        Categoria = 'Erro HTTP 401 - Nao Autorizado (0x80190191)'
        Causa     = 'O download exige credenciais de autenticacao no servidor web.'
        Solucoes  = @(
            'Verificar se o aplicativo requer licenca ou login previo para download.'
        )
    }
    '0x80190193' = @{
        Categoria = 'Erro HTTP 403 - Acesso Proibido (0x80190193)'
        Causa     = 'O servidor do fornecedor bloqueou o acesso ao arquivo (geobloqueio, hotlinking ou IP restrito).'
        Solucoes  = @(
            'Verificar se o IP de saida da empresa esta bloqueado na CDN do fabricante.',
            'Efetuar o download manualmente ou ignorar no lote com -ExcludePackages.'
        )
    }
    '0x80190194' = @{
        Categoria = 'Erro HTTP 404 - Arquivo Nao Encontrado (0x80190194)'
        Causa     = 'O arquivo de download foi removido do servidor pelo fornecedor (link quebrado ou versao aposentada).'
        Solucoes  = @(
            'Aguardar atualizacao do manifesto pela comunidade no repositorio winget-pkgs.',
            'Adicionar a lista de excecao com -ExcludePackages.'
        )
    }
    '0x80190197' = @{
        Categoria = 'Erro HTTP 407 - Autenticacao no Proxy Necessaria (0x80190197)'
        Causa     = 'O servidor Proxy da rede exige login e senha para autorizar o download.'
        Solucoes  = @(
            'Configurar autenticacao no proxy corporativo para a conta do computador.'
        )
    }
    '0x801901F4' = @{
        Categoria = 'Erro HTTP 500 - Erro Interno do Servidor (0x801901F4)'
        Causa     = 'O servidor web do desenvolvedor apresentou falha interna ao atender a requisicao.'
        Solucoes  = @(
            'Tentar novamente mais tarde quando os servidores do fabricante forem restabelecidos.'
        )
    }
    '0x801901F6' = @{
        Categoria = 'Erro HTTP 502 - Bad Gateway (0x801901F6)'
        Causa     = 'Falha de comunicacao entre servidores proxy ou CDN intermediarios.'
        Solucoes  = @(
            'Tentar novamente apos alguns minutos.'
        )
    }
    '0x801901F7' = @{
        Categoria = 'Erro HTTP 503 - Servico Indisponivel (0x801901F7)'
        Causa     = 'Servidor do fabricante temporariamente sobrecarregado ou em manutencao.'
        Solucoes  = @(
            'Tentar novamente mais tarde.'
        )
    }
    '0x801901F8' = @{
        Categoria = 'Erro HTTP 504 - Gateway Timeout (0x801901F8)'
        Causa     = 'O servidor intermediario nao obteve resposta no tempo limite.'
        Solucoes  = @(
            'Repetir a operacao em outro horario.'
        )
    }
    '0x800B0100' = @{
        Categoria = 'Assinatura Digital Ausente no Instalador (0x800B0100)'
        Causa     = 'O arquivo baixado nao contem uma assinatura digital Authenticode reconhecida.'
        Solucoes  = @(
            'Verificar a procedencia do pacote e baixar a versao oficial assinada.'
        )
    }
    '0x800B0101' = @{
        Categoria = 'Certificado Digital Expirado (0x800B0101)'
        Causa     = 'O certificado com o qual o fornecedor assinou o executavel esta com data de validade vencida.'
        Solucoes  = @(
            'Verificar com o desenvolvedor do software se ha uma versao com certificado renovado.'
        )
    }
    '0x800B0109' = @{
        Categoria = 'Certificado Raiz Nao Confiavel (0x800B0109)'
        Causa     = 'A cadeia de certificacao do instalador termina em uma autoridade raiz nao instalada no repositorio de certificados do Windows.'
        Solucoes  = @(
            'Atualizar a lista de certificados raiz confiaveis do Windows via Windows Update.',
            'Instalar o certificado raiz da Autoridade Certificadora na maquina de destino.'
        )
    }
    '0x800B010A' = @{
        Categoria = 'Cadeia de Certificacao Incompleta (0x800B010A)'
        Causa     = 'Nao foi possivel construir a cadeia de certificados ate uma raiz confiavel (falta certificado intermediario).'
        Solucoes  = @(
            'Instalar os certificados intermediarios necessarios no repositorio do computador.'
        )
    }
    '0x800B010E' = @{
        Categoria = 'Certificado Digital Revogado (0x800B010E)'
        Causa     = 'O certificado do fornecedor foi revogado pela Autoridade Certificadora por motivos de seguranca.'
        Solucoes  = @(
            'Nao instalar este executavel: risco de seguranca por certificado comprometido.',
            'Adicionar a lista de exclusoes permanentes com -ExcludePackages.'
        )
    }
    '0x800B0111' = @{
        Categoria = 'Certificado Explicitamente Nao Confiavel (0x800B0111)'
        Causa     = 'O certificado do instalador foi adicionado a pasta de Certificados Nao Confiaveis da maquina por GPO.'
        Solucoes  = @(
            'Verificar as politicas de seguranca da organizacao em certlm.msc.'
        )
    }
    '0x80092004' = @{
        Categoria = 'Objeto Criptografico Nao Encontrado (0x80092004)'
        Causa     = 'Assinatura ou chave publica nao pode ser lida no cabecalho do arquivo.'
        Solucoes  = @(
            'Verificar se o arquivo baixado foi corrompido durante o download.'
        )
    }
    '0x80240017' = @{
        Categoria = 'Atualizacao Nao Aplicavel (0x80240017)'
        Causa     = 'A atualizacao nao se aplica a edicao ou build atual do Windows.'
        Solucoes  = @(
            'Ignorar este componente na estacao.'
        )
    }
    '0x80240020' = @{
        Categoria = 'Operacao Exige Interacao do Usuario (0x80240020)'
        Causa     = 'O pacote nao pode ser instalado de forma nao interativa e exige resposta na tela.'
        Solucoes  = @(
            'Executar a instalacao logado na sessao de usuario.'
        )
    }
    '0x80244018' = @{
        Categoria = 'Download Bloqueado por Proxy (0x80244018)'
        Causa     = 'O gateway corporativo retornou status HTTP 403 Forbidden para o servico de atualizacao.'
        Solucoes  = @(
            'Liberar dominios da Microsoft e de CDNs autorizadas na lista de permissoes do firewall.'
        )
    }
    '0x80244019' = @{
        Categoria = 'Arquivo de Catalogo Nao Encontrado (0x80244019)'
        Causa     = 'Erro 404 retornado pelo servidor de atualizacao do Windows.'
        Solucoes  = @(
            'Limpar o cache do SoftwareDistribution se afetar atualizacoes cumulativas.'
        )
    }
    '0x8024402C' = @{
        Categoria = 'DNS do Servidor de Atualizacao Nao Resolvido (0x8024402C)'
        Causa     = 'Falha de resolucao de nomes WinHTTP para o servidor de catalogo.'
        Solucoes  = @(
            'Verificar os servidores DNS configurados na placa de rede da maquina.'
        )
    }

    # --- 5. CODIGOS DO SUBSISTEMA APPX / MSIX E MICROSOFT STORE ---
    '0x80073CF0' = @{
        Categoria = 'Falha no Pacote AppX (0x80073CF0)'
        Causa     = 'O pacote MSIX nao pode ser aberto pelo subsistema de implantacao AppX.'
        Solucoes  = @(
            'Verificar se o arquivo .msixbundle baixado esta integro.',
            'Tentar reinstalar o pacote individualmente.'
        )
    }
    '0x80073CF1' = @{
        Categoria = 'Pacote AppX Ja Existe (0x80073CF1)'
        Causa     = 'O aplicativo com a mesma identidade ja esta registrado para outro usuario.'
        Solucoes  = @(
            'Registrar o pacote para o usuario atual utilizando Add-AppxPackage.'
        )
    }
    '0x80073CF2' = @{
        Categoria = 'Repositorio AppX Corrompido (0x80073CF2)'
        Causa     = 'O banco de dados de registro de aplicativos AppX do Windows esta danificado.'
        Solucoes  = @(
            'Executar em PowerShell como Administrador: Dism /Online /Cleanup-Image /RestoreHealth.',
            'Reiniciar a estacao de trabalho.'
        )
    }
    '0x80073CF3' = @{
        Categoria = 'Falha ao Resolver Dependencias AppX (0x80073CF3)'
        Causa     = 'O pacote MSIX depende de outros pacotes de framework (ex: VCLibs ou Microsoft.UI.Xaml) que nao estao presentes.'
        Solucoes  = @(
            'Instalar os pacotes de framework pre-requisito antes do pacote principal.',
            'Garantir acesso aos servidores da Microsoft Store.'
        )
    }
    '0x80073CF4' = @{
        Categoria = 'Espaco Insuficiente para Pacote AppX (0x80073CF4)'
        Causa     = 'Nao ha espaco em disco suficiente no volume selecionado para descompactar o pacote MSIX.'
        Solucoes  = @(
            'Liberar espaco na particao do sistema operacional.'
        )
    }
    '0x80073CF6' = @{
        Categoria = 'Falha de Registro do Pacote AppX (0x80073CF6)'
        Causa     = 'O manifesto do pacote falhou ao registrar servicos, manipuladores de protocolo ou extensoes no Windows.'
        Solucoes  = @(
            'Consultar o visualizador de eventos em: Microsoft-Windows-AppXDeployment-Server/Operational.',
            'Reiniciar o servico AppXSvc ou a estacao de trabalho.'
        )
    }
    '0x80073CF9' = @{
        Categoria = 'Implantacao AppX Bloqueada por Diretiva (0x80073CF9)'
        Causa     = 'A instalacao de aplicativos empacotados foi bloqueada por politica de grupo corporativa (GPO).'
        Solucoes  = @(
            'Verificar a diretiva: Permitir a instalacao de todos os aplicativos confiaveis (Sideloading).',
            'Ajustar regras corporativas em GPO.'
        )
    }
    '0x80073CFA' = @{
        Categoria = 'Modo de Desenvolvedor Necessario (0x80073CFA)'
        Causa     = 'O pacote nao possui assinatura comercial da Microsoft Store e exige modo desenvolvedor para ser instalado.'
        Solucoes  = @(
            'Habilitar o Modo de Desenvolvedor nas configuracoes do Windows caso seja um pacote interno corporativo.'
        )
    }
    '0x80073CFB' = @{
        Categoria = 'Falha na Validacao de Integridade AppX (0x80073CFB)'
        Causa     = 'O mapa de blocos (AppxBlockMap.xml) nao confere com os arquivos de conteudo do pacote.'
        Solucoes  = @(
            'Baixar novamente o arquivo .msixbundle para obter uma copia integra.'
        )
    }
    '0x80073CFC' = @{
        Categoria = 'Roaming Nao Suportado (0x80073CFC)'
        Causa     = 'O aplicativo nao oferece suporte para sincronizacao em perfil de usuario movel (Roaming Profile).'
        Solucoes  = @(
            'Instalar o aplicativo em perfil local.'
        )
    }
    '0x80073CFD' = @{
        Categoria = 'Servico Firewall do Windows Desativado (0x80073CFD)'
        Causa     = 'O servico Windows Firewall (mpssvc) esta parado ou desativado. O subsistema AppX exige o servico ativo para registrar pacotes.'
        Solucoes  = @(
            'Iniciar e habilitar o servico Windows Firewall: Set-Service mpssvc -StartupType Automatic; Start-Service mpssvc.',
            'O subsistema UWP/AppX do Windows bloqueia qualquer instalacao de aplicativo se o firewall estiver desativado.'
        )
    }
    '0x80073CFE' = @{
        Categoria = 'Falha ao Mover Pacote AppX (0x80073CFE)'
        Causa     = 'Erro ao realocar arquivos entre volumes de armazenamento.'
        Solucoes  = @(
            'Manter o local de instalacao padrao no drive C:.'
        )
    }
    '0x80073D00' = @{
        Categoria = 'Volume de Destino Nao Vazio (0x80073D00)'
        Causa     = 'A unidade designada para gravacao de novos apps possui arquivos residuais incompativeis.'
        Solucoes  = @(
            'Configurar o armazenamento de novos aplicativos na unidade padrao em Configuracoes -> Sistema -> Armazenamento.'
        )
    }
    '0x80073D01' = @{
        Categoria = 'Pacote AppX Nao Confiavel (0x80073D01)'
        Causa     = 'O certificado de assinatura do pacote nao esta na lista de autoridades confiaveis da maquina.'
        Solucoes  = @(
            'Importar o certificado publico do pacote para o repositorio Pessoas Confiaveis da maquina.'
        )
    }
    '0x80073D02' = @{
        Categoria = 'Implantacao Abortada por Logoff de Usuario (0x80073D02)'
        Causa     = 'O usuario desconectou da sessao durante o registro do aplicativo no perfil.'
        Solucoes  = @(
            'Executar a instalacao com o usuario conectado ou sob o escopo de maquina (SYSTEM).'
        )
    }
    '0x80073D05' = @{
        Categoria = 'Falha ao Limpar Armazenamento Anterior (0x80073D05)'
        Causa     = 'O Windows nao conseguiu deletar a pasta de dados da versao anterior em %LOCALAPPDATA%\Packages por arquivo bloqueado.'
        Solucoes  = @(
            'Reiniciar a estacao para desbloquear os arquivos residuais da aplicacao.',
            'Tentar novamente a instalacao.'
        )
    }
    '0x80073D19' = @{
        Categoria = 'Pacote Principal Necessario (0x80073D19)'
        Causa     = 'Tentativa de instalar um pacote opcional ou pacote de idiomas sem que o pacote principal esteja instalado.'
        Solucoes  = @(
            'Instalar previamente o pacote principal do aplicativo.'
        )
    }
    '0x80073D23' = @{
        Categoria = 'Identidade de Conta SYSTEM Nao Permitida (0x80073D23)'
        Causa     = 'O pacote MSIX proibe expressamente registro sob a conta de servico NT AUTHORITY\SYSTEM e exige uma conta de usuario real.'
        Solucoes  = @(
            'Instalar o pacote na sessao interativa do usuario logado na estacao.',
            'Adicionar a lista de excecao de maquina com -ExcludePackages.'
        )
    }
    '0x80080204' = @{
        Categoria = 'Manifesto AppX Invalido (0x80080204)'
        Causa     = 'O arquivo AppxManifest.xml dentro do pacote contem tags ou esquemas XML corrompidos.'
        Solucoes  = @(
            'Aguardar versao corrigida disponibilizada pelo desenvolvedor.'
        )
    }
    '0x80080207' = @{
        Categoria = 'Conteudo do Pacote Corrompido (0x80080207)'
        Causa     = 'Arquivos de payload do pacote MSIX sofreram corrupcao de transferencia.'
        Solucoes  = @(
            'Excluir o arquivo baixado e efetuar novo download.'
        )
    }

    # --- 6. INSTALADORES DE TERCEIROS (INNO SETUP, NSIS, INSTALLSHIELD, WIX BURN) ---
    'Inno_1' = @{
        Categoria = 'Falha de Inicializacao Inno Setup (Code 1)'
        Causa     = 'O instalador Inno Setup falhou durante a rotina inicial de validacao de ambiente ou pre-requisitos.'
        Solucoes  = @(
            'Verificar compatibilidade do software com a maquina.',
            'Executar interativamente na estacao para identificar a mensagem do assistente.'
        )
    }
    'Inno_2' = @{
        Categoria = 'Instalacao Inno Setup Cancelada (Code 2)'
        Causa     = 'A instalacao foi cancelada pelo usuario ou processo externo.'
        Solucoes  = @(
            'Garantir parametros silenciosos (/VERYSILENT /SUPPRESSMSGBOXES /NORESTART).'
        )
    }
    'Inno_3' = @{
        Categoria = 'Falha Fatal Inno Setup (Code 3)'
        Causa     = 'Erro grave durante a extracao de arquivos do arquivo compactado Inno Setup.'
        Solucoes  = @(
            'Verificar espaco livre em disco e integridade do arquivo baixado.'
        )
    }
    'Inno_4' = @{
        Categoria = 'Inno Setup Abortado por Pre-Requisito (Code 4)'
        Causa     = 'O instalador foi abortado devido a uma condicao obrigatoria de pre-requisito nao atendida.'
        Solucoes  = @(
            'Instalar os componentes obrigatorios antes de rodar o pacote.'
        )
    }
    'Inno_5' = @{
        Categoria = 'Erro Critico Inno Setup (Code 5)'
        Causa     = 'Excecao fatal nao tratada pelo executavel do instalador Inno Setup.'
        Solucoes  = @(
            'Reportar o problema ao desenvolvedor ou testar versao mais recente.'
        )
    }
    'NSIS_1' = @{
        Categoria = 'Instalacao NSIS Cancelada ou Invalida (Code 1)'
        Causa     = 'O instalador NSIS (Nullsoft) encerrou devido a argumento incorreto ou cancelamento.'
        Solucoes  = @(
            'Verificar se o argumento de silencio (/S) esta em maiusculas (o NSIS diferencia maiusculas de minusculas).'
        )
    }
    'NSIS_2' = @{
        Categoria = 'Falha de Extracao de Arquivos NSIS (Code 2)'
        Causa     = 'Erro fatal de CRC ao descompactar arquivos do instalador NSIS.'
        Solucoes  = @(
            'O arquivo baixado esta corrompido. Efetuar novo download.',
            'Verificar se o disco rigido possui espaco livre.'
        )
    }
    'IS_-1' = @{
        Categoria = 'Falha Geral InstallShield (Code -1)'
        Causa     = 'Falha generica de inicializacao do mecanismo InstallShield.'
        Solucoes  = @(
            'Verificar se o mecanismo InstallShield (ISScript.msi) precisa ser atualizado.'
        )
    }
    'IS_-2' = @{
        Categoria = 'Arquivo de Resposta Invalido InstallShield (Code -2)'
        Causa     = 'O arquivo de resposta (.iss) para instalacao silenciosa nao foi localizado ou tem sintaxe invalida.'
        Solucoes  = @(
            'Gerar novo arquivo de resposta executando: setup.exe -r.'
        )
    }
    'IS_-3' = @{
        Categoria = 'Argumento Invalido InstallShield (Code -3)'
        Causa     = 'Parametro incorreto passado para a linha de comando do InstallShield.'
        Solucoes  = @(
            'Revisar os argumentos passados na chamada do instalador.'
        )
    }
    'IS_-4' = @{
        Categoria = 'Memoria Insuficiente InstallShield (Code -4)'
        Causa     = 'O mecanismo InstallShield nao dispoe de memoria livre suficiente.'
        Solucoes  = @(
            'Reiniciar a estacao para liberar recursos.'
        )
    }
    'IS_-5' = @{
        Categoria = 'Arquivo Ja Existe InstallShield (Code -5)'
        Causa     = 'O arquivo de destino ja existe e o instalador nao tem autorizacao para substituir.'
        Solucoes  = @(
            'Desinstalar a versao anterior antes de prosseguir.'
        )
    }
    'IS_-6' = @{
        Categoria = 'Falha ao Gravar Arquivo de Resposta (Code -6)'
        Causa     = 'Permissoes insuficientes na pasta de gravacao de logs do InstallShield.'
        Solucoes  = @(
            'Conferir permissoes de escrita na pasta temporaria.'
        )
    }
    '0x800701BE' = @{
        Categoria = 'Falha de Validacao do Bundle WiX Burn (0x800701BE)'
        Causa     = 'O bundle WiX Burn detectou discrepancia de hash ou assinatura em um de seus pacotes internos.'
        Solucoes  = @(
            'Baixar novamente o pacote completo do fornecedor.'
        )
    }
    '0x800701BF' = @{
        Categoria = 'Pre-Requisito com Falha no Bundle WiX Burn (0x800701BF)'
        Causa     = 'Um dos pacotes de pre-requisito inclusos no bundle falhou na instalacao.'
        Solucoes  = @(
            'Consultar os logs individuais do WiX Burn na pasta C:\Windows\Temp ou %TEMP%.'
        )
    }
}


function Get-PackageDiagnosis {
    param(
        [PSCustomObject]$Item,
        [string]$ComputerName = "estacao"
    )

    $id   = if ($Item.ID) { "$($Item.ID)" } else { "" }
    $prog = if ($Item.Programa) { "$($Item.Programa)" } else { "" }
    $code = if ($Item.ExitCode -ne $null) { "$($Item.ExitCode)".Trim() } else { "" }
    $err  = if ($Item.ErrorSnippet) { "$($Item.ErrorSnippet)" } else { "" }
    $stat = if ($Item.Status) { "$($Item.Status)" } else { "" }

    # ==============================================================================
    # A. DIAGNOSTICO ESPECIFICO DE SOFTWARES CONHECIDOS (28 APLICATIVOS POPULARES)
    # ==============================================================================

    # 1. Ausencia do Winget no Sistema (Windows Server / LTSC / Sem Microsoft Store)
    if ($stat -match 'não localizado|nao localizado' -or ($id -match 'DesktopAppInstaller' -and $stat -match 'Erro')) {
        return @{
            Categoria        = 'Winget / AppInstaller Nao Instalado no Sistema'
            Causa            = 'O executavel do Winget (Microsoft.DesktopAppInstaller) nao foi encontrado no sistema. Isso ocorre tipicamente em servidores Windows Server (2019/2022/2025) ou edicoes Windows LTSC/Enterprise que nao incluem a Microsoft Store pre-instalada.'
            Solucoes         = @(
                'Instalar o Winget manualmente no host: Baixar o pacote Microsoft.DesktopAppInstaller (.msixbundle) e suas dependencias (Microsoft.VCLibs e Microsoft.UI.Xaml) do repositorio oficial no GitHub (https://github.com/microsoft/winget-cli/releases).',
                'Executar no PowerShell do destino como Administrador: Add-AppxPackage -Path "<caminho>\Microsoft.DesktopAppInstaller.msixbundle" -DependencyPath @("<caminho>\VCLibs.appx", "<caminho>\UI.Xaml.appx")',
                'Verificar se o host e um servidor (ex: Windows Server) antes de inclui-lo na lista de atualizacao de estacoes clientes.'
            )
            ExclusaoSugerida = 'Microsoft.DesktopAppInstaller'
        }
    }

    # 2. Auto-Atualizacao do Winget (Microsoft.AppInstaller / 0x8A150110 / 2316632336)
    if ($id -in @('Microsoft.AppInstaller', 'Microsoft.DesktopAppInstaller') -or ($code -in @('2316632336', '0x8A150110') -and ($id -match 'AppInstaller' -or $prog -match 'DesktopAppInstaller'))) {
        return @{
            Categoria        = 'Auto-Atualizacao do Winget Bloqueada / Dependencia MSIX (0x8A150110)'
            Causa            = 'O pacote Microsoft.AppInstaller e o proprio executavel do Winget. A tentativa do Winget de atualizar a si mesmo via linha de comando em background falha com o codigo 0x8A150110 (APPINSTALLER_CLI_ERROR_INSTALL_DEPENDENCIES) porque o binario esta em execucao e o subsistema AppX requer sessao interativa da Microsoft Store.'
            Solucoes         = @(
                'Nao e necessario atualizar manualmente: O Windows Update e a Microsoft Store atualizam o App Installer automaticamente em background durante a manutencao.',
                'Adicionar a lista de exclusoes para ignorar a tentativa de auto-atualizacao do Winget: -ExcludePackages @("Microsoft.AppInstaller")',
                'Caso deseje atualizar a versao do Winget manualmente no host, execute no PowerShell como Administrador: Add-AppxPackage -Path "<caminho>\Microsoft.DesktopAppInstaller.msixbundle"'
            )
            ExclusaoSugerida = 'Microsoft.AppInstaller'
        }
    }

    # 3. Dell Command | Update (ExitCode 4, 1603, 1714, 1612, 2316632070)
    if ($id -match 'Dell\.CommandUpdate' -or $prog -match 'Dell Command' -or ($code -in @('4','1603','1714','1612','2316632070') -and ($id -match 'Dell' -or $prog -match 'Dell'))) {
        return @{
            Categoria        = 'Corrupcao de Cache MSI / Conflito de Versao Anterior'
            Causa            = 'O instalador da nova versao tenta desinstalar a versao anterior, mas o arquivo MSI original correspondente no cache do Windows (C:\Windows\Installer) foi removido ou corrompido, gerando falha no msiexec (codigos 4, 1603 ou 1714).'
            Solucoes         = @(
                'Remover a versao antiga manualmente pelo Painel de Controle -> Programas e Recursos ou utilizando a ferramenta Microsoft Program Install and Uninstall Troubleshooter.',
                'Ou baixar o executavel da Dell diretamente do site de suporte e rodar como administrador para recriar o cache.',
                'Para prosseguir com a atualizacao em lote dos demais aplicativos sem travar, execute com: -ExcludePackages @("Dell.CommandUpdate*")'
            )
            ExclusaoSugerida = 'Dell.CommandUpdate*'
        }
    }

    # 4. Zoom Workplace / Zoom Client (ExitCode 1603, 1714, 1612)
    if ($id -match 'Zoom\.Zoom' -or $prog -match 'Zoom' -or ($code -in @('1603','1714','1612') -and ($id -match 'Zoom' -or $prog -match 'Zoom'))) {
        return @{
            Categoria        = 'Incompatibilidade de Atualizacao MSI / Chaves de Registro Orfas'
            Causa            = 'O instalador MSI do Zoom falhou ao tentar atualizar a versao anterior legada por ausencia do pacote MSI original em C:\Windows\Installer (Erro MSI 1603 / 1714).'
            Solucoes         = @(
                'Executar a ferramenta oficial de limpeza da Zoom (CleanZoom.exe) na estacao para remover completamente a instalacao antiga e seus registros residuais.',
                'Apos a limpeza, instalar a versao atual limpa executando: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -InstallPackage "Zoom.Zoom"',
                'Para prosseguir com a atualizacao em lote dos demais aplicativos, ignore este pacote com: -ExcludePackages @("Zoom.Zoom")'
            )
            ExclusaoSugerida = 'Zoom.Zoom'
        }
    }

    # 5. Microsoft Teams Classic (404 / 0x80190194 / 2149122452 / CDN removido)
    if ($id -match 'Microsoft\.Teams\.Classic' -or $err -match '0x80190194|404|Not Found' -or $code -eq '2149122452') {
        return @{
            Categoria        = 'Aplicativo Descontinuado / URL de Download Inexistente (HTTP 404)'
            Causa            = 'O Microsoft Teams Classic foi oficialmente descontinuado e aposentado pela Microsoft em 2024. A Microsoft removeu os binarios de download de sua CDN, tornando impossivel atualiza-lo via Winget.'
            Solucoes         = @(
                'Desinstalar o Teams Classic da maquina de destino.',
                'Instalar a nova geracao do Teams executando: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -InstallPackage "Microsoft.Teams"',
                'Adicionar o Teams Classic permanentemente na lista de exclusao: -ExcludePackages @("Microsoft.Teams.Classic")'
            )
            ExclusaoSugerida = 'Microsoft.Teams.Classic'
        }
    }

    # 6. Veeam Agent for Microsoft Windows (v3 para v13 nao suportado diretamente)
    if ($id -match 'Veeam\.VeeamAgent' -or $prog -match 'Veeam Agent') {
        return @{
            Categoria        = 'Upgrade Direto Nao Suportado pelo Fornecedor'
            Causa            = 'A arquitetura do instalador do Veeam Agent v3.0 e legada e o instalador da v13 nao aceita atualizacao direta in-place, abortando o assistente de instalacao.'
            Solucoes         = @(
                'Desinstalar a versao legada v3.0.2 pelo Painel de Controle ou WMI.',
                'Instalar a nova versao limpa via: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -InstallPackage "Veeam.VeeamAgent"',
                'Para nao interromper o fluxo em lote dos demais programas: -ExcludePackages @("Veeam.VeeamAgent*")'
            )
            ExclusaoSugerida = 'Veeam.VeeamAgent*'
        }
    }

    # 7. SafeSign 64-bits / Softwares de Certificado Digital (ExitCode 2)
    if ($id -match 'SafeSign' -or $prog -match 'SafeSign' -or ($id -match 'AETEurope' -or $prog -match 'AETEurope')) {
        return @{
            Categoria        = 'Falha de Descompactacao de Arquivo (Archive Shell COM em Sessao 0)'
            Causa            = 'O instalador do SafeSign e distribuido em formato ZIP ou CAB e tenta invocar objetos Shell COM ou assistente interativo do Explorer, o que e bloqueado na Sessao 0 (Sessao de Servico do SYSTEM).'
            Solucoes         = @(
                'Para softwares de certificado digital (ex: SafeSign), instalar/atualizar diretamente com o executavel oficial fornecido pela Autoridade Certificadora.',
                'Para pacotes afetados por essa limitacao do Winget em background, execute a atualizacao em sessao de usuario interativa ou adicione a exclusao: -ExcludePackages @("' + $id + '")'
            )
            ExclusaoSugerida = $id
        }
    }

    # 8. Microsoft.UI.Xaml / Microsoft.VCLibs (Pacotes AppX do Sistema)
    if ($id -match 'Microsoft\.UI\.Xaml' -or $id -match 'Microsoft\.VCLibs') {
        return @{
            Categoria        = 'Pacote de Infraestrutura AppX / WinUI do Windows'
            Causa            = 'Estes pacotes sao bibliotecas de runtime internas compartilhadas do subsistema AppX/MSIX do Windows.'
            Solucoes         = @(
                'Este pacote e gerenciado e atualizado de forma automatica pelo Windows Update e Microsoft Store.',
                'Nao requer atualizacao forcada via Winget. Adicione a lista de exclusao: -ExcludePackages @("Microsoft.UI.Xaml*", "Microsoft.VCLibs*")'
            )
            ExclusaoSugerida = $id
        }
    }

    # 9. Discord (Discord.Discord)
    if ($id -match 'Discord\.Discord' -or $prog -match 'Discord') {
        return @{
            Categoria        = 'Instalador de Escopo de Usuario (%LocalAppData%)'
            Causa            = 'O Discord foi desenvolvido para ser instalado estritamente no perfil do usuario (%LocalAppData%\Discord). O instalador recusa execucao sob a conta SYSTEM e falha em tarefas agendadas de maquina.'
            Solucoes         = @(
                'Executar a atualizacao direcionada a sessao do usuario conectado: O script ja detecta o usuario e dispara em sua sessao.',
                'Garantir que o usuario esteja conectado interativamente no momento da execucao.',
                'Adicionar a lista de exclusao de maquina: -ExcludePackages @("Discord.Discord")'
            )
            ExclusaoSugerida = 'Discord.Discord'
        }
    }

    # 10. Slack (SlackTechnologies.Slack)
    if ($id -match 'SlackTechnologies\.Slack' -or $prog -match 'Slack') {
        return @{
            Categoria        = 'Conflito de Instalacao por Usuario vs por Maquina'
            Causa            = 'O Slack possui duas variantes: a versao MSI (por maquina em Program Files) e a versao por usuario (%LocalAppData%). A tentativa de atualizar uma variante com o instalador da outra causa erro de registro.'
            Solucoes         = @(
                'Padronizar o parque com a versao Slack MSI para Maquina (Slack Machine-Wide Installer).',
                'Desinstalar versoes locais instaladas individualmente no perfil dos usuarios.'
            )
            ExclusaoSugerida = 'SlackTechnologies.Slack'
        }
    }

    # 11. Docker Desktop (Docker.DockerDesktop)
    if ($id -match 'Docker\.DockerDesktop' -or $prog -match 'Docker Desktop') {
        return @{
            Categoria        = 'Pre-Requisito de Virtualizacao / WSL2 Nao Atendido'
            Causa            = 'O Docker Desktop exige a presenca ativa do subsistema WSL2 (Windows Subsystem for Linux) ou recursos de Hyper-V habilitados no Windows.'
            Solucoes         = @(
                'Habilitar o WSL2 e VirtualMachinePlatform executando no PowerShell: wsl --install --no-distribution.',
                'Verificar se a virtualizacao de CPU (VT-x / AMD-V) esta ativada na BIOS da maquina.',
                'Ignorar temporariamente no lote: -ExcludePackages @("Docker.DockerDesktop")'
            )
            ExclusaoSugerida = 'Docker.DockerDesktop'
        }
    }

    # 12. Oracle VM VirtualBox (Oracle.VirtualBox)
    if ($id -match 'Oracle\.VirtualBox' -or $prog -match 'VirtualBox') {
        return @{
            Categoria        = 'Instalacao de Driver de Kernel Bloqueada em Background'
            Causa            = 'O instalador do VirtualBox necessita registrar drivers de filtro de rede e virtualizacao NDIS6. Em sessoes silenciosas sem interacao, o Windows pode bloquear a instalacao de drivers nao certificados pela politica corporativa.'
            Solucoes         = @(
                'Instalar o certificado digital da Oracle no repositorio de Editores Confiaveis antes de atualizar.',
                'Executar a instalacao interativamente ou ignorar no lote: -ExcludePackages @("Oracle.VirtualBox")'
            )
            ExclusaoSugerida = 'Oracle.VirtualBox'
        }
    }

    # 13. Java Runtime / JDK (Oracle.JavaRuntimeEnvironment / Oracle.JDK)
    if ($id -match 'Oracle\.JavaRuntimeEnvironment' -or $id -match 'Oracle\.JDK' -or $prog -match 'Java(TM) SE') {
        return @{
            Categoria        = 'Processo Java Bloqueando Arquivos Binarios'
            Causa            = 'Instancias ativas de java.exe, javaw.exe ou servicos baseados em Java impedem a sobrescrita das DLLs em C:\Program Files\Java.'
            Solucoes         = @(
                'Encerrar aplicacoes Java em execucao na estacao antes de proceder com o upgrade: Stop-Process -Name java, javaw -Force -ErrorAction SilentlyContinue.',
                'Agendar a atualizacao do Java fora do horario de expediente.'
            )
            ExclusaoSugerida = $id
        }
    }

    # 14. Node.js (OpenJS.NodeJS)
    if ($id -match 'OpenJS\.NodeJS' -or $prog -match 'Node.js') {
        return @{
            Categoria        = 'Processo Node.js em Execucao'
            Causa            = 'Processos node.exe em segundo plano (daemons, ferramentas de build ou scripts) bloqueiam a substituicao do binario principal.'
            Solucoes         = @(
                'Finalizar processos node.exe ativos na estacao remota.',
                'Executar novamente a atualizacao: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -Update -PackageId "OpenJS.NodeJS"'
            )
            ExclusaoSugerida = 'OpenJS.NodeJS'
        }
    }

    # 15. Git for Windows (Git.Git)
    if ($id -match 'Git\.Git' -or $prog -match 'Git') {
        return @{
            Categoria        = 'Binarios do Git Bloqueados por IDEs em Uso'
            Causa            = 'Editores como VS Code, Visual Studio ou terminais abertos mantem git.exe, ssh.exe ou bash.exe abertos em segundo plano.'
            Solucoes         = @(
                'Solicitar ao desenvolvedor que feche o VS Code ou finalizar os processos do Git antes de atualizar.',
                'Adicionar a lista de excecoes do lote caso nao queira interromper o trabalho: -ExcludePackages @("Git.Git")'
            )
            ExclusaoSugerida = 'Git.Git'
        }
    }

    # 16. Visual Studio Code (Microsoft.VisualStudioCode)
    if ($id -match 'Microsoft\.VisualStudioCode' -or $prog -match 'Visual Studio Code') {
        return @{
            Categoria        = 'Conflito entre Versao User Installer e System Installer'
            Causa            = 'O VS Code possui um instalador para o perfil do usuario (%LocalAppData%) e outro para a maquina inteira (Program Files). Tentar atualizar a versao de usuario atraves da conta SYSTEM gera falha.'
            Solucoes         = @(
                'Instalar a versao correta: Microsoft.VisualStudioCode.System para escopo de maquina global.',
                'Para usuarios comuns, a atualizacao do VS Code ocorre de forma automatica e integrada no proprio editor.'
            )
            ExclusaoSugerida = 'Microsoft.VisualStudioCode'
        }
    }

    # 17. Google Chrome / Brave Browser
    if ($id -match 'Google\.Chrome' -or $id -match 'Brave\.Brave' -or $prog -match 'Google Chrome|Brave Browser') {
        return @{
            Categoria        = 'Navegador Aberto Bloqueando DLLs Principais'
            Causa            = 'O navegador esta em execucao com multiplas abas ou processos de extensao, bloqueando a substituicao dos arquivos chrome.dll ou brave.dll.'
            Solucoes         = @(
                'O Winget agenda a substituicao para a proxima reinicializacao do navegador na maioria dos casos.',
                'Para forcar o encerramento previo: Stop-Process -Name chrome, brave -Force -ErrorAction SilentlyContinue.'
            )
            ExclusaoSugerida = $id
        }
    }

    # 18. Mozilla Firefox (Mozilla.Firefox)
    if ($id -match 'Mozilla\.Firefox' -or $prog -match 'Firefox') {
        return @{
            Categoria        = 'Processo do Firefox Ativo Bloqueando Atualizacao'
            Causa            = 'O processo firefox.exe esta ativo na memoria, impedindo a sobregravacao do executavel sem reinicializacao do navegador.'
            Solucoes         = @(
                'Encerrar o Firefox ou aguardar o usuario fechar a aplicacao para efetivar o update.'
            )
            ExclusaoSugerida = 'Mozilla.Firefox'
        }
    }

    # 19. Adobe Acrobat Reader (Adobe.Acrobat.Reader.64-bit)
    if ($id -match 'Adobe\.Acrobat\.Reader' -or $prog -match 'Adobe Acrobat Reader') {
        return @{
            Categoria        = 'Conflito com Servico Adobe ARM / Processos AcroRd32 Ativos'
            Causa            = 'O servico AdobeARMservice ou processos leitores de PDF abertos na sessao bloqueiam a execucao silenciosa do instalador.'
            Solucoes         = @(
                'Finalizar processos: Stop-Process -Name AcroRd32, Acrobat -Force -ErrorAction SilentlyContinue.',
                'Utilizar a ferramenta Adobe Customization Wizard caso haja politicas corporativas de PDF.'
            )
            ExclusaoSugerida = 'Adobe.Acrobat.Reader*'
        }
    }

    # 20. AnyDesk / TeamViewer / RustDesk
    if ($id -match 'AnyDesk\.AnyDesk' -or $id -match 'TeamViewer\.TeamViewer' -or $id -match 'RustDesk\.RustDesk') {
        return @{
            Categoria        = 'Servico de Acesso Remoto Bloqueado em Execucao'
            Causa            = 'O servico de suporte remoto esta rodando em segundo plano e bloqueia a substituicao de seus binarios executaveis sem parada temporaria do servico.'
            Solucoes         = @(
                'Parar temporariamente o servico antes de atualizar (ex: Stop-Service AnyDesk).',
                'Executar a atualizacao individualmente ou adicionar a lista de excecoes: -ExcludePackages @("' + $id + '")'
            )
            ExclusaoSugerida = $id
        }
    }

    # 21. Spotify (Spotify.Spotify)
    if ($id -match 'Spotify\.Spotify' -or $prog -match 'Spotify') {
        return @{
            Categoria        = 'Instalador Exclusivo de Usuario'
            Causa            = 'O instalador do Spotify roda exclusivamente na pasta %AppData% do usuario e bloqueia qualquer execucao sob privilegios elevados de SYSTEM.'
            Solucoes         = @(
                'Executar a atualizacao direcionada a sessao do usuario logado.',
                'Adicionar a lista de excecao de maquina: -ExcludePackages @("Spotify.Spotify")'
            )
            ExclusaoSugerida = 'Spotify.Spotify'
        }
    }

    # 22. Notepad++ (Notepad++.Notepad++)
    if ($id -match 'Notepad\+\+\.Notepad\+\+' -or $prog -match 'Notepad\+\+') {
        return @{
            Categoria        = 'Instancia do Notepad++ Aberta'
            Causa            = 'O processo notepad++.exe esta em execucao na sessao do usuario, bloqueando a atualizacao silenciosa.'
            Solucoes         = @(
                'Encerrar o Notepad++ na maquina de destino antes de atualizar: Stop-Process -Name notepad++ -Force -ErrorAction SilentlyContinue.'
            )
            ExclusaoSugerida = 'Notepad++.Notepad++'
        }
    }

    # 23. Wireshark (WiresharkFoundation.Wireshark)
    if ($id -match 'WiresharkFoundation\.Wireshark' -or $prog -match 'Wireshark') {
        return @{
            Categoria        = 'Driver Npcap Bloqueado em Captura de Pacotes'
            Causa            = 'O driver de captura de pacotes Npcap / WinPcap esta em uso ou exige confirmacao interativa de licenca.'
            Solucoes         = @(
                'Garantir que nenhuma captura de rede esteja em andamento.',
                'Instalar previamente o pacote Npcap OEM ou silencioso.'
            )
            ExclusaoSugerida = 'WiresharkFoundation.Wireshark'
        }
    }

    # 24. Postman (Postman.Postman)
    if ($id -match 'Postman\.Postman' -or $prog -match 'Postman') {
        return @{
            Categoria        = 'Aplicacao Electron Restrita a Perfil de Usuario'
            Causa            = 'O Postman e instalado por padrao na pasta AppData do usuario e seu auto-updater interno gerencia as atualizacoes.'
            Solucoes         = @(
                'Permitir que o Postman se auto-atualize pela interface do usuario ou disparar no escopo de usuario conectado.'
            )
            ExclusaoSugerida = 'Postman.Postman'
        }
    }

    # 25. Microsoft Power BI Desktop (Microsoft.PowerBI)
    if ($id -match 'Microsoft\.PowerBI' -or $prog -match 'Power BI Desktop') {
        return @{
            Categoria        = 'Dependencia de WebView2 / Instancia em Uso'
            Causa            = 'O Power BI Desktop exige o Microsoft Edge WebView2 Runtime e fecha conexoes de modelo de dados caso haja relatorios abertos.'
            Solucoes         = @(
                'Garantir que o WebView2 Runtime esteja instalado na maquina.',
                'Encerrar o executavel PBIDesktop.exe antes de atualizar.'
            )
            ExclusaoSugerida = 'Microsoft.PowerBI'
        }
    }

    # 26. DBeaver (dbeaver.dbeaver)
    if ($id -match 'dbeaver\.dbeaver' -or $prog -match 'DBeaver') {
        return @{
            Categoria        = 'Conexoes de Banco de Dados Ativas no DBeaver'
            Causa            = 'O processo dbeaver.exe bloqueia a atualizacao se houver conexoes JDBC abertas na sessao.'
            Solucoes         = @(
                'Encerrar o processo dbeaver.exe na estacao antes de proceder com o upgrade.'
            )
            ExclusaoSugerida = 'dbeaver.dbeaver'
        }
    }

    # 27. Python 3 (Python.Python.3.*)
    if ($id -match 'Python\.Python\.3' -or $prog -match 'Python 3') {
        return @{
            Categoria        = 'Conflito de Versoes Minoritarias de Python'
            Causa            = 'O instalador do Python trata cada versao menor (ex: 3.11, 3.12, 3.13) como produtos independentes e nao como atualizacao in-place.'
            Solucoes         = @(
                'Se deseja a versao mais nova, use -InstallPackage para adicionar a nova versao em paralelo.',
                'Remover a versao minoritaria anterior pelo Painel de Controle caso nao precise de retrocompatibilidade.'
            )
            ExclusaoSugerida = 'Python.Python*'
        }
    }

    # 28. Windows 10 & 11 Update Assistants (Microsoft.UpdateAssistant / Microsoft.WindowsInstallationAssistant)
    if ($id -match 'Microsoft\.UpdateAssistant' -or $id -match 'Microsoft\.WindowsInstallationAssistant' -or $prog -match 'Update Assistant|Installation Assistant') {
        return @{
            Categoria        = 'Assistente de Atualizacao de Sistema Operacional (Incompativel com Winget Background)'
            Causa            = 'Estes pacotes sao ferramentas interativas de atualizacao de versao principal do Windows. Em maquinas ja rodando Windows 11 atualizado, o executavel nao realiza atualizacao de software comum e trava na Sessao 0 (SYSTEM) aguardando interacao ou verificacao de compatibilidade de SO.'
            Solucoes         = @(
                'Adicionar a lista de exclusao de atualizacao de aplicativos em lote: -ExcludePackages @("Microsoft.UpdateAssistant*", "Microsoft.WindowsInstallationAssistant*")',
                'Para atualizacoes de seguranca e builds do Windows, utilize os canais oficiais: Windows Update, WSUS ou Intune.',
                'Caso o assistente nao seja mais necessario na estacao, desinstale-o pelo Painel de Controle -> Programas e Recursos.'
            )
            ExclusaoSugerida = 'Microsoft.*Assistant*'
        }
    }

    # 29. Microsoft Visual Studio (Microsoft.VisualStudio.*)
    if ($id -match 'Microsoft\.VisualStudio\.' -or ($prog -match 'Visual Studio' -and $prog -notmatch 'Code')) {
        return @{
            Categoria        = 'Suite Pesada com Timeout Elevado / Bloqueio de Instalador (Visual Studio 2022)'
            Causa            = 'O instalador do Visual Studio 2022 (vs_Professional.exe / setup.exe) baixa e aplica pacotes pesados (dezenas de GB) de compiladores e SDKs. A instalacao em segundo plano costuma exceder o tempo limite padrao (900s) e retem o mutex global de instaladores do Windows, bloqueando os demais pacotes.'
            Solucoes         = @(
                'Para atualizar o Visual Studio isoladamente aumentando o tempo limite: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -Update -PackageId "' + $id + '" -TaskTimeoutSeconds 3600',
                'Para atualizacoes normais de aplicativos em lote, ignore o Visual Studio: -ExcludePackages @("Microsoft.VisualStudio*")',
                'Ou atualize diretamente pelo Visual Studio Installer interativo na maquina do desenvolvedor.'
            )
            ExclusaoSugerida = 'Microsoft.VisualStudio*'
        }
    }

    # 30. Sped Fiscal ICMS IPI / PVA Receita Federal (ReceitaFederaldoBrasil.Sped*)
    if ($id -match 'ReceitaFederaldoBrasil\.Sped' -or $prog -match 'Sped Fiscal|Sped Contribuicoes') {
        return @{
            Categoria        = 'Aplicativo Java PVA da Receita Federal (Instalador Customizado)'
            Causa            = 'Os Programas Validadores e Assinadores (PVA) da Receita Federal exigem runtime Java especifico e possuem instalador que solicita confirmacao de diretorio (%ProgramFiles%\Programas SPED), rejeitando atualizacao silenciosa em Sessao 0.'
            Solucoes         = @(
                'Executar a atualizacao em sessao de usuario ou instalar a nova versao do PVA com o instalador oficial baixado do portal SPED da Receita Federal.',
                'Adicionar a lista de excecoes: -ExcludePackages @("ReceitaFederaldoBrasil.Sped*")'
            )
            ExclusaoSugerida = 'ReceitaFederaldoBrasil.Sped*'
        }
    }

    # ==============================================================================
    # B. BUSCA DIRETA NO BANCO DE CONHECIMENTO DE CODIGOS (200+ ENTRADAS INDEXADAS)
    # ==============================================================================

    # Normaliza o codigo para busca em multiplos formatos (Decimal, Hex, Signed Int)
    $lookupKeys = @()
    if ($code) {
        $lookupKeys += $code
        $lookupKeys += $code.ToUpper()
        $lookupKeys += $code.ToLower()

        $parsedNum = $null
        if ([int64]::TryParse($code, [ref]$parsedNum)) {
            if ($parsedNum -lt 0) {
                $uintVal = [uint32]$parsedNum
                $lookupKeys += "$uintVal"
                $lookupKeys += ("0x{0:X8}" -f $uintVal)
                $lookupKeys += ("0x{0:x8}" -f $uintVal)
            } else {
                $lookupKeys += ("0x{0:X8}" -f $parsedNum)
                $lookupKeys += ("0x{0:x8}" -f $parsedNum)
                $lookupKeys += ("0x{0:X}" -f $parsedNum)
                $lookupKeys += ("0x{0:x}" -f $parsedNum)
                if ($parsedNum -gt 2147483647 -and $parsedNum -le 4294967295) {
                    $intSigned = [int32][uint32]$parsedNum
                    $lookupKeys += "$intSigned"
                }
            }
        }
    }

    # Verifica se ha ocorrencia direta de codigo HRESULT ou Hex no snippet de erro
    if ($err -match '(0x[0-9a-fA-F]{8})') {
        $hexMatch = $matches[1]
        $lookupKeys += $hexMatch.ToUpper()
        $lookupKeys += $hexMatch.ToLower()
    }

    foreach ($k in $lookupKeys) {
        if ($script:WingetKnowledgeBase -and $script:WingetKnowledgeBase.ContainsKey($k)) {
            $entry = $script:WingetKnowledgeBase[$k]
            $sug = if ($entry.ExclusaoSugerida) { $entry.ExclusaoSugerida } else { $id }
            return @{
                Categoria        = $entry.Categoria
                Causa            = $entry.Causa
                Solucoes         = $entry.Solucoes
                ExclusaoSugerida = $sug
            }
        }
    }

    # ==============================================================================
    # C. PADROES TEXTUAIS E HEURISTICAS DE FALHA (FALLBACKS INTELIGENTES)
    # ==============================================================================
    if ($err -match 'HASH_MISMATCH|hash' -or $stat -match 'hash') {
        return @{
            Categoria        = 'Hash de Integridade Divergente (INSTALLER_HASH_MISMATCH)'
            Causa            = 'O hash SHA256 do instalador baixado nao confere com o manifesto oficial. O fornecedor publicou nova versao no mesmo link sem atualizar o repositorio do Winget.'
            Solucoes         = @(
                'Aguardar atualizacao do manifesto oficial no repositorio do Winget pela comunidade.',
                'Ou efetuar o download e a atualizacao manual diretamente do site do desenvolvedor.',
                'Para prosseguir com os demais aplicativos: -ExcludePackages @("' + $id + '")'
            )
            ExclusaoSugerida = $id
        }
    }

    if ($err -match '0x80072ee|0x8019019|download|connection|timeout' -or $stat -match 'download|conexão|conexao') {
        return @{
            Categoria        = 'Falha de Conectividade / Download Interrompido'
            Causa            = 'Nao foi possivel completar o download do pacote. Conexao interrompida, servidor offline ou bloqueio por firewall/proxy corporativo.'
            Solucoes         = @(
                'Verificar se a estacao de destino possui acesso a internet e resolucao de nomes DNS para o dominio do aplicativo.',
                'Checar regras de proxy ou inspecao SSL em gateways corporativos.',
                'Tentar novamente mais tarde ou ignorar na fila com: -ExcludePackages @("' + $id + '")'
            )
            ExclusaoSugerida = $id
        }
    }

    if ($err -match 'reboot|3010' -or $stat -match 'reboot') {
        return @{
            Categoria        = 'Reinicializacao do Sistema Necessaria (MSI 3010)'
            Causa            = 'A instalacao foi concluida, porem arquivos protegidos de sistema foram enfileirados para substituicao na proxima reinicializacao.'
            Solucoes         = @(
                'Agendar o reinicio da maquina fora do horario de trabalho para efetivar a nova versao.',
                'O aplicativo ja se encontra operacional para a maioria das funcoes.'
            )
            ExclusaoSugerida = $id
        }
    }

    # ==============================================================================
    # D. ERRO GENERICO / NAO CATALOGADO (FALLBACK)
    # ==============================================================================
    $exitMsg = if ($code) { "Codigo de saida retornado pelo instalador: $code." } else { "Falha reportada pelo gerenciador Winget." }
    return @{
        Categoria        = 'Erro na Execucao do Instalador'
        Causa            = "$exitMsg O instalador do fornecedor retornou codigo de erro durante a execucao silenciosa em segundo plano."
        Solucoes         = @(
            'Testar a atualizacao deste aplicativo isoladamente: .\Update-StationWinget.ps1 -ComputerName "' + $ComputerName + '" -Update -PackageId "' + $id + '"',
            'Verificar os logs de instalacao no destino em: %LOCALAPPDATA%\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir',
            'Adicionar a lista de excecoes para prosseguir com os demais pacotes: -ExcludePackages @("' + $id + '")'
        )
        ExclusaoSugerida = $id
    }
}


function Show-ErrorAnalysis {
    param(
        [array]$FailedItems,
        [string]$ComputerName = "estacao"
    )

    if (-not $FailedItems -or $FailedItems.Count -eq 0) { return }

    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Red
    Write-Host "            DIAGNOSTICO AUTOMATICO DE ERROS E SOLUCOES SUGERIDAS               " -ForegroundColor Red
    Write-Host "================================================================================" -ForegroundColor Red
    Write-Host "Foram identificadas falhas em $($FailedItems.Count) pacote(s) durante o processo em $ComputerName." -ForegroundColor Yellow
    Write-Host "Abaixo apresentamos a analise de causa raiz e as acoes recomendadas para cada um:" -ForegroundColor Gray
    Write-Host ""

    $suggestedExcludes = @()
    $idx = 1
    foreach ($item in $FailedItems) {
        $diag = Get-PackageDiagnosis -Item $item -ComputerName $ComputerName
        if ($diag.ExclusaoSugerida) {
            $suggestedExcludes += $diag.ExclusaoSugerida
        }

        Write-Host "[$idx] $($item.Programa) (ID: $($item.ID))" -ForegroundColor Cyan
        Write-Host "    Escopo:        $($item.Escopo) | Status: $($item.Status)" -ForegroundColor Gray
        if ($item.ExitCode -ne $null -and "$($item.ExitCode)" -ne "") {
            Write-Host "    Codigo Saida:  $($item.ExitCode)" -ForegroundColor DarkGray
        }
        Write-Host "    Diagnostico:   $($diag.Categoria)" -ForegroundColor Yellow
        Write-Host "    Causa Raiz:    $($diag.Causa)" -ForegroundColor White
        Write-Host "    Solucao Proposta:" -ForegroundColor Green
        foreach ($sol in $diag.Solucoes) {
            Write-Host "      * $sol" -ForegroundColor Gray
        }
        Write-Host ""
        $idx++
    }

    $uniqueExcludes = $suggestedExcludes | Select-Object -Unique
    if ($uniqueExcludes -and $uniqueExcludes.Count -gt 0) {
        Write-Host "--------------------------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "COMANDO SUGERIDO PARA RETENTAR O RESTANTE IGNORANDO OS PACOTES COM FALHA:" -ForegroundColor Yellow
        $quotedList = ($uniqueExcludes | ForEach-Object { "`"$_`"" }) -join ", "
        Write-Host ".\Update-StationWinget.ps1 -ComputerName `"$ComputerName`" -Update -ExcludePackages @($quotedList)" -ForegroundColor White
    }
    Write-Host "================================================================================" -ForegroundColor Red
}

function Show-FormattedTable {
    param(
        [array]$Items,
        [string]$Mode = "List"
    )

    if (-not $Items -or $Items.Count -eq 0) { return }

    # Detecta a largura real do terminal do usuário (padrão 110 se não detectado)
    $termWidth = 110
    try {
        if ($Host.UI.RawUI.WindowSize.Width -gt 60) {
            $termWidth = $Host.UI.RawUI.WindowSize.Width
        }
    } catch {}

    $colVer1 = switch ($Mode) {
        "Install" { "Versao Anterior" }
        "Update"  { "Versao Anterior" }
        default   { "Versao Atual" }
    }
    $colVer2 = switch ($Mode) {
        "Install" { "Versao Instalada" }
        "Update"  { "Versao Instalada" }
        default   { "Disponivel" }
    }

    $colList = @(
        @{ Header = "Escopo";   Prop = "Escopo";           Min = 8;  Max = 12 },
        @{ Header = "Programa"; Prop = "Programa";         Min = 18; Max = 36 },
        @{ Header = "ID";       Prop = "ID";               Min = 14; Max = 30 },
        @{ Header = $colVer1;   Prop = "VersaoAtual";      Min = 10; Max = 16 },
        @{ Header = $colVer2;   Prop = "VersaoDisponivel"; Min = 10; Max = 16 },
        @{ Header = "Status";   Prop = "Status";           Min = 8;  Max = 18 }
    )

    # Calcula a largura de cada coluna baseando-se no conteúdo
    foreach ($col in $colList) {
        $maxDataLen = $col.Header.Length
        foreach ($item in $Items) {
            $val = "$($item.$($col.Prop))"
            if ($val.Length -gt $maxDataLen) { $maxDataLen = $val.Length }
        }
        $col.Width = [Math]::Min($col.Max, [Math]::Max($col.Min, $maxDataLen))
    }

    # Garante que a linha total nunca ultrapasse a largura da tela do usuário
    $availableWidth = $termWidth - 2
    $totalWidth = 0
    foreach ($c in $colList) { $totalWidth += $c.Width + 2 }

    if ($totalWidth -gt $availableWidth) {
        $overflow = $totalWidth - $availableWidth
        $reduceProg = [Math]::Min($overflow, $colList[1].Width - $colList[1].Min)
        $colList[1].Width -= $reduceProg
        $overflow -= $reduceProg
        if ($overflow -gt 0) {
            $reduceId = [Math]::Min($overflow, $colList[2].Width - $colList[2].Min)
            $colList[2].Width -= $reduceId
        }
    }

    # Cabeçalho
    $headerLine = ""
    $sepLine = ""
    for ($i = 0; $i -lt $colList.Count; $i++) {
        $c = $colList[$i]
        $hText = $c.Header
        if ($hText.Length -gt $c.Width) { $hText = $hText.Substring(0, $c.Width) }
        $headerLine += $hText.PadRight($c.Width)
        $sepLine += ("-" * $c.Width)
        if ($i -lt ($colList.Count - 1)) {
            $headerLine += "  "
            $sepLine += "  "
        }
    }

    Write-Host $headerLine -ForegroundColor Yellow
    Write-Host $sepLine -ForegroundColor DarkGray

    # Linhas de dados (cada programa ocupa estritamente 1 linha sem quebrar)
    foreach ($item in $Items) {
        $rowLine = ""
        for ($i = 0; $i -lt $colList.Count; $i++) {
            $c = $colList[$i]
            $val = "$($item.$($c.Prop))"
            if ($val.Length -gt $c.Width) {
                if ($c.Width -gt 3) {
                    $val = $val.Substring(0, $c.Width - 3) + "..."
                } else {
                    $val = $val.Substring(0, $c.Width)
                }
            }
            $rowLine += $val.PadRight($c.Width)
            if ($i -lt ($colList.Count - 1)) {
                $rowLine += "  "
            }
        }
        Write-Host $rowLine
    }
}

# ========================================================
# EXECUÇÃO DO ORQUESTRADOR LOCAL / REMOTO
# ========================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Resolução da lista de computadores a processar (via arquivo ou parâmetro)
$targetComputers = @()
if ($ComputerFile) {
    if (-not (Test-Path -Path $ComputerFile)) {
        Write-Error "Arquivo de lista de computadores não encontrado: '$ComputerFile'"
        return
    }
    $rawLines = Get-Content -Path $ComputerFile -ErrorAction Stop
    foreach ($line in $rawLines) {
        $trimmed = $line.Trim()
        if ($trimmed -and -not $trimmed.StartsWith("#") -and -not $trimmed.StartsWith(";")) {
            $targetComputers += $trimmed
        }
    }
    if ($targetComputers.Count -eq 0) {
        Write-Error "O arquivo '$ComputerFile' não contém nenhum nome de computador válido."
        return
    }
} else {
    $targetComputers = @($ComputerName)
}

if ($Help -or $PSBoundParameters.Count -eq 0) {
    Show-ScriptHelp -TargetComputer ($targetComputers[0])
    return
}

$executionFolder = (Get-Location).Path
if (-not $executionFolder) {
    $executionFolder = if ($PSScriptRoot) { $PSScriptRoot } else { "." }
}
$logFile = Join-Path -Path $executionFolder -ChildPath "output.log"

$isUpdate = $Update.IsPresent
$isListOnly = (-not $isUpdate) -and (-not $InstallPackage)

$allBatchResults = @()
$totalComputers  = $targetComputers.Count
$compIdx         = 1

if ($totalComputers -gt 1) {
    Write-Host "================================================================================" -ForegroundColor Magenta
    Write-Host " INICIANDO PROCESSAMENTO EM LOTE DE $totalComputers MÁQUINA(S) " -ForegroundColor Magenta
    if ($ComputerFile) {
        Write-Host " Origem da lista: $ComputerFile" -ForegroundColor DarkGray
    }
    Write-Host "================================================================================" -ForegroundColor Magenta
}

foreach ($currentHost in $targetComputers) {
    $startTime = Get-Date
    $tempLog = [System.IO.Path]::GetTempFileName()

    try {
        Start-Transcript -Path $tempLog -Force | Out-Null
    } catch {}

    try {
        if ($totalComputers -gt 1) {
            Write-Host ""
            Write-Host "################################################################################" -ForegroundColor Cyan
            Write-Host " [$compIdx/$totalComputers] PROCESSANDO ESTAÇÃO: $currentHost " -ForegroundColor Cyan
            Write-Host "################################################################################" -ForegroundColor Cyan
        }

        Write-Host "==========================================================" -ForegroundColor Cyan
        if ($InstallPackage) {
            Write-Host "Iniciando INSTALAÇÃO remota via Winget em: $currentHost" -ForegroundColor Cyan
            Write-Host "Pacote alvo para instalação: $InstallPackage" -ForegroundColor Yellow
        } elseif ($isUpdate) {
            Write-Host "Iniciando ATUALIZAÇÃO remota via Winget em: $currentHost" -ForegroundColor Cyan
            if ($PackageId) {
                Write-Host "Alvo específico selecionado: $PackageId" -ForegroundColor Yellow
            }
        } else {
            Write-Host "Verificação remota via Winget em: $currentHost [MODO CONSULTA / SOMENTE LISTAGEM]" -ForegroundColor Cyan
            if ($PackageId) {
                Write-Host "Filtro de aplicativo ativo: $PackageId" -ForegroundColor Yellow
            }
            Write-Host "(Dica: Execute com o parâmetro -Help para ver opções de uso e exemplos práticos)" -ForegroundColor DarkGray
        }
        if ($ExcludePackages.Count -gt 0) {
            Write-Host "Exceções configuradas: $($ExcludePackages -join ', ')" -ForegroundColor Yellow
        }
        Write-Host "==========================================================" -ForegroundColor Cyan

        $invokeParams = @{
            ComputerName = $currentHost
            ScriptBlock  = $RemoteScriptBlock
            ArgumentList = @($ExcludePackages, $TaskTimeoutSeconds, [bool]$isListOnly, [string]$PackageId, [string]$InstallPackage)
        }
        if ($Credential) {
            $invokeParams["Credential"] = $Credential
        }

        $remoteData = $null
        try {
            $remoteData = Invoke-Command @invokeParams -ErrorAction Stop
        } catch {
            Write-Error "Erro ao conectar ou executar em $currentHost : $_"
            $allBatchResults += [PSCustomObject]@{
                Computador       = $currentHost
                Escopo           = "Geral"
                Programa         = "Conexão Remota"
                ID               = "WinRM"
                VersaoAtual      = "N/A"
                VersaoDisponivel = "N/A"
                VersaoAnterior   = "N/A"
                VersaoInstalada  = "N/A"
                Status           = "Erro de Conexão"
                ExitCode         = -1
                ErrorSnippet     = "$_"
            }
            continue
        }

        # Informações de diagnóstico retornadas da estação
        $rawResults     = $remoteData.Results
        $loggedUsers    = $remoteData.LoggedUsers
        $machinePending = $remoteData.MachinePending

        if (-not $InstallPackage) {
            Write-Host "[+] Escopo de Máquina: $machinePending pacote(s) com atualização disponível no total." -ForegroundColor Gray
            if ($loggedUsers -and $loggedUsers.Count -gt 0) {
                Write-Host "[+] Usuário(s) logado(s) detectado(s): $($loggedUsers -join ', ')" -ForegroundColor Gray
            } else {
                Write-Host "[+] Nenhum usuário conectado no momento." -ForegroundColor Gray
            }
        }

        # Processamento e montagem dos objetos
        $formattedList = @()
        if ($rawResults) {
            foreach ($r in $rawResults) {
                $formattedList += [PSCustomObject]@{
                    Computador       = $currentHost
                    Escopo           = $r.Escopo
                    Programa         = $r.Programa
                    ID               = $r.ID
                    VersaoAtual      = $r.VersaoAtual
                    VersaoDisponivel = $r.VersaoDisponivel
                    VersaoAnterior   = $r.VersaoAtual
                    VersaoInstalada  = $r.VersaoDisponivel
                    Status           = $r.Status
                    ExitCode         = $r.ExitCode
                    ErrorSnippet     = $r.ErrorSnippet
                }
            }
        }

        # Filtro de exibição conforme opção IncludeSkipped
        $reportItems = @(
            if (-not $IncludeSkipped) {
                $formattedList | Where-Object { $_.Status -notmatch '^Ignorado' }
            } else {
                $formattedList
            }
        )

        Write-Host ""
        if ($reportItems.Count -gt 0) {
            if ($InstallPackage) {
                Write-Host "Resultado da Instalação em $currentHost :" -ForegroundColor Green
                Show-FormattedTable -Items $reportItems -Mode "Install"
            } elseif (-not $isUpdate) {
                Write-Host "Pacotes Pendentes de Atualização em $currentHost :" -ForegroundColor Cyan
                Show-FormattedTable -Items $reportItems -Mode "List"
            } else {
                Write-Host "Resumo de Programas Atualizados em $currentHost :" -ForegroundColor Green
                Show-FormattedTable -Items $reportItems -Mode "Update"
            }
        } else {
            if ($PackageId) {
                Write-Host "O aplicativo '$PackageId' não possui atualizações pendentes ou já está na versão mais recente em $currentHost." -ForegroundColor Green
            } elseif ($formattedList.Count -gt 0) {
                Write-Host "Nenhum pacote pendente de ação (todos os pacotes encontrados estão na lista de exceções)." -ForegroundColor Yellow
            } else {
                Write-Host "Tudo em dia! Nenhuma atualização pendente encontrada para a máquina ou usuários logados em $currentHost." -ForegroundColor Green
            }
        }

        # Se houver itens com falha (em modo Update ou Install), exibe a análise diagnóstica automática
        $failedItems = @($formattedList | Where-Object { $_.Status -match '^(Falhou|Erro)' })
        if ($failedItems.Count -gt 0) {
            Show-ErrorAnalysis -FailedItems $failedItems -ComputerName $currentHost
        }

        $allBatchResults += $formattedList
    }
    finally {
        try {
            Stop-Transcript | Out-Null
            if (Test-Path -LiteralPath $tempLog) {
                $rawLogLines = Get-Content -LiteralPath $tempLog
                $inHeader = $false
                $cleanLogLines = [System.Collections.Generic.List[string]]::new()

                $endTime = Get-Date
                $duration = New-TimeSpan -Start $startTime -End $endTime
                $durStr = "{0:D2}m {1:D2}s" -f [int]$duration.TotalMinutes, $duration.Seconds

                # Cabeçalho com data/hora de início e computador alvo
                $cleanLogLines.Add("================================================================================")
                $cleanLogLines.Add("EXECUÇÃO INICIADA EM: $($startTime.ToString('dd/MM/yyyy HH:mm:ss')) | ALVO: $currentHost")
                $cleanLogLines.Add("================================================================================")

                foreach ($line in $rawLogLines) {
                    if ($line -match '^\*{6,}') {
                        $inHeader = -not $inHeader
                        continue
                    }
                    if ($inHeader) { continue }
                    $cleanLogLines.Add(($line -replace '^INFO:\s', ''))
                }

                # Rodapé com data/hora de conclusão e tempo total decorrido
                $cleanLogLines.Add("--------------------------------------------------------------------------------")
                $cleanLogLines.Add("EXECUÇÃO CONCLUÍDA EM: $($endTime.ToString('dd/MM/yyyy HH:mm:ss')) | DURAÇÃO: $durStr")
                $cleanLogLines.Add("================================================================================")
                $cleanLogLines.Add("")

                [System.IO.File]::AppendAllLines($logFile, $cleanLogLines, [System.Text.UTF8Encoding]::new($true))
                Remove-Item -LiteralPath $tempLog -Force -ErrorAction SilentlyContinue
                Write-Host "  -> Log incremental atualizado em: $logFile" -ForegroundColor DarkGray
            }
        } catch {}
    }

    $compIdx++
}

if ($totalComputers -gt 1) {
    Write-Host ""
    Write-Host "================================================================================" -ForegroundColor Magenta
    Write-Host " PROCESSAMENTO EM LOTE CONCLUÍDO ($totalComputers MÁQUINAS PROCESSADAS) " -ForegroundColor Magenta
    Write-Host " Log consolidado gravado em: $logFile" -ForegroundColor DarkGray
    Write-Host "================================================================================" -ForegroundColor Magenta
}

if ($PassThru) {
    return $allBatchResults
}
