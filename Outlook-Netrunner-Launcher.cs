using System;
using System.Diagnostics;
using System.IO;

/// <summary>
/// Outlook Netrunner Launcher
/// Busca y ejecuta el script PowerShell en la carpeta del EXE.
/// Compatible con PowerShell 5.1+ (Windows 7+).
/// </summary>
class Program
{
    static int Main(string[] args)
    {
        // Encontrar el script en la carpeta del EXE
        string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        string exeDir = Path.GetDirectoryName(exePath);
        string scriptPath = Path.Combine(exeDir, "Diagnose-OutlookPerf.ps1");

        // Validar que el script existe
        if (!File.Exists(scriptPath))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine("ERROR: No se encontro Diagnose-OutlookPerf.ps1 en la carpeta del EXE.");
            Console.Error.WriteLine("Ruta esperada: " + scriptPath);
            Console.ResetColor();
            return 1;
        }

        try
        {
            // Construir argumentos para PowerShell (forzar UTF-8)
            string argsStr = string.Join(" ", EscapeArgs(args));
            string psArgs = "-NoProfile -ExecutionPolicy Bypass -InputFormat None -File \"" + scriptPath + "\" " + argsStr;

            // Ejecutar PowerShell
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = psArgs,
                UseShellExecute = false,
                CreateNoWindow = false
            };

            using (Process proc = Process.Start(psi))
            {
                proc.WaitForExit();
                return proc.ExitCode;
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine("ERROR: " + ex.Message);
            Console.ResetColor();
            return 1;
        }
    }

    static string[] EscapeArgs(string[] args)
    {
        if (args == null || args.Length == 0) return args;

        string[] escaped = new string[args.Length];
        for (int i = 0; i < args.Length; i++)
        {
            // Si el argumento contiene espacios, envolver en comillas
            if (args[i].Contains(" "))
            {
                escaped[i] = "\"" + args[i] + "\"";
            }
            else
            {
                escaped[i] = args[i];
            }
        }
        return escaped;
    }
}
