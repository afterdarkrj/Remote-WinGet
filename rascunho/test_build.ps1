$csc = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$sourceScript = "C:\Temp\Update-StationWinget.ps1"
$outputExe = "C:\Users\sjjunior\.gemini\antigravity-cli\brain\d2a2258a-99ba-4459-b3bc-c663030a31ec\scratch\Test-Launcher.exe"
$csSource = "C:\Users\sjjunior\.gemini\antigravity-cli\brain\d2a2258a-99ba-4459-b3bc-c663030a31ec\scratch\Launcher.cs"

$csharpCode = @'
using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Text;

[assembly: AssemblyTitle("Update-StationWinget CLI")]
[assembly: AssemblyDescription("Orquestrador Remoto de Atualizacoes Winget")]
[assembly: AssemblyCompany("TI Corporativa")]
[assembly: AssemblyProduct("Update-StationWinget")]
[assembly: AssemblyVersion("2.1.0.0")]
[assembly: AssemblyFileVersion("2.1.0.0")]

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
'@

[System.IO.File]::WriteAllText($csSource, $csharpCode, [System.Text.Encoding]::UTF8)

# Compile with csc.exe
$argsCsc = @(
    "/target:exe",
    "/optimize+",
    "/out:$outputExe",
    "/resource:$sourceScript,Script.ps1",
    $csSource
)

Write-Host "Compiling test executable..."
$proc = Start-Process -FilePath $csc -ArgumentList $argsCsc -NoNewWindow -PassThru -Wait

if ($proc.ExitCode -eq 0 -and (Test-Path $outputExe)) {
    $item = Get-Item $outputExe
    Write-Host "COMPILATION SUCCESSFUL!" -ForegroundColor Green
    Write-Host "Exe size: $($item.Length) bytes ($([math]::Round($item.Length / 1KB, 2)) KB)"
} else {
    Write-Host "COMPILATION FAILED with code $($proc.ExitCode)" -ForegroundColor Red
}
