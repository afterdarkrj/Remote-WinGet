<#
.SYNOPSIS
    Compilador Nativo: Transforma Update-StationWinget.ps1 em Update-StationWinget.exe
.DESCRIPTION
    Utiliza exclusivamente o compilador C# nativo do Windows (csc.exe) presente em qualquer
    estacao Windows 10/11 ou Windows Server, sem necessidade de internet, PSGallery ou modulos adicionais.
#>
param(
    [string]$SourceScript = (Join-Path $PSScriptRoot "Update-StationWinget.ps1"),
    [string]$OutputFile   = (Join-Path $PSScriptRoot "Update-StationWinget.exe"),
    [string]$Version      = "2.1.0.0"
)

$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"

if (-not (Test-Path $csc)) {
    Write-Error "Compilador csc.exe nao encontrado em: $csc"
    return
}

if (-not (Test-Path $SourceScript)) {
    Write-Error "Script fonte nao encontrado em: $SourceScript"
    return
}

Write-Host "Iniciando compilacao nativa de $SourceScript..." -ForegroundColor Cyan

$tempCs = [System.IO.Path]::GetTempFileName() + ".cs"

$csharpCode = @"
using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Text;

[assembly: AssemblyTitle("Update-StationWinget CLI")]
[assembly: AssemblyDescription("Orquestrador Remoto de Atualizacoes Winget")]
[assembly: AssemblyCompany("TI Corporativa")]
[assembly: AssemblyProduct("Update-StationWinget")]
[assembly: AssemblyVersion("$Version")]
[assembly: AssemblyFileVersion("$Version")]

namespace UpdateStationWingetLauncher
{
    class Program
    {
        static int Main(string[] args)
        {
            string tempScript = null;
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                string scriptContent;
                using (var stream = assembly.GetManifestResourceStream("Script.ps1"))
                {
                    if (stream == null)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Erro: Recurso de script interno nao encontrado.");
                        Console.ResetColor();
                        return 1;
                    }
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        scriptContent = reader.ReadToEnd();
                    }
                }

                string tempDir = Path.GetTempPath();
                string tempFileName = "USW_" + Guid.NewGuid().ToString("N") + ".ps1";
                tempScript = Path.Combine(tempDir, tempFileName);

                File.WriteAllText(tempScript, scriptContent, new UTF8Encoding(true));

                var sbArgs = new StringBuilder();
                sbArgs.Append("-NoProfile -ExecutionPolicy Bypass -File \"");
                sbArgs.Append(tempScript);
                sbArgs.Append("\"");

                foreach (var arg in args)
                {
                    sbArgs.Append(" ").Append(EscapeArgument(arg));
                }

                var psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = sbArgs.ToString(),
                    UseShellExecute = false,
                    WorkingDirectory = Environment.CurrentDirectory
                };

                Console.CancelKeyPress += (sender, e) =>
                {
                    e.Cancel = true;
                };

                using (var process = Process.Start(psi))
                {
                    process.WaitForExit();
                    return process.ExitCode;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Erro inesperado na execucao: " + ex.Message);
                Console.ResetColor();
                return 1;
            }
            finally
            {
                if (tempScript != null && File.Exists(tempScript))
                {
                    try { File.Delete(tempScript); } catch { }
                }
            }
        }

        private static string EscapeArgument(string arg)
        {
            if (string.IsNullOrEmpty(arg))
                return "\"\"";

            if (arg.IndexOfAny(new char[] { ' ', '\t', '\n', '\v', '\"' }) == -1)
                return arg;

            var sb = new StringBuilder();
            sb.Append('\"');
            for (int i = 0; i < arg.Length; i++)
            {
                int backslashCount = 0;
                while (i < arg.Length && arg[i] == '\\')
                {
                    backslashCount++;
                    i++;
                }

                if (i == arg.Length)
                {
                    sb.Append('\\', backslashCount * 2);
                    break;
                }
                else if (arg[i] == '\"')
                {
                    sb.Append('\\', backslashCount * 2 + 1);
                    sb.Append('\"');
                }
                else
                {
                    sb.Append('\\', backslashCount);
                    sb.Append(arg[i]);
                }
            }
            sb.Append('\"');
            return sb.ToString();
        }
    }
}
"@

try {
    [System.IO.File]::WriteAllText($tempCs, $csharpCode, [System.Text.Encoding]::UTF8)

    $argsCsc = @(
        "/target:exe",
        "/optimize+",
        "/out:$OutputFile",
        "/resource:$SourceScript,Script.ps1",
        $tempCs
    )

    $proc = Start-Process -FilePath $csc -ArgumentList $argsCsc -NoNewWindow -PassThru -Wait

    if ($proc.ExitCode -eq 0 -and (Test-Path $OutputFile)) {
        $item = Get-Item $OutputFile
        Write-Host "Executavel gerado com sucesso!" -ForegroundColor Green
        Write-Host "Arquivo: $($item.FullName) ($([math]::Round($item.Length / 1KB, 1)) KB)" -ForegroundColor Green
    } else {
        Write-Error "Falha na compilacao com codigo de erro $($proc.ExitCode)."
    }
} finally {
    if (Test-Path $tempCs) {
        Remove-Item -Path $tempCs -Force -ErrorAction SilentlyContinue
    }
}
