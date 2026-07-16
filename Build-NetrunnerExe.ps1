#requires -Version 5.1
<#
.SYNOPSIS
    Compila Outlook Netrunner en un EXE portátil.

.DESCRIPTION
    Crea un ejecutable (launcher) que busca y ejecuta el script PS en su carpeta.
    - Portable: funciona en cualquier carpeta
    - Requiere PowerShell 5.1+ en el target
    - Genera un EXE de ~15 KB
    - El script Diagnose-OutlookPerf.ps1 debe estar en la misma carpeta

.PARAMETER OutputPath
    Ruta del EXE de salida. Default: .\Outlook-Netrunner.exe

.EXAMPLE
    .\Build-NetrunnerExe.ps1
    Genera Outlook-Netrunner.exe en el directorio actual.
#>
param(
    [string]$OutputPath = ".\Outlook-Netrunner.exe"
)

$ErrorActionPreference = "Stop"

Write-Host "Outlook Netrunner :: Build Script (PowerShell 5.1+ compatible)`n" -ForegroundColor Cyan

# --- Verificar prerequisitos ---
Write-Host "[~] Verificando requisitos..."
$launcherFile = Join-Path $PSScriptRoot "Outlook-Netrunner-Launcher.cs"
$scriptFile = Join-Path $PSScriptRoot "Diagnose-OutlookPerf.ps1"

if (-not (Test-Path $launcherFile)) {
    Write-Host "[x] ERROR: No se encontro Outlook-Netrunner-Launcher.cs" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $scriptFile)) {
    Write-Host "[x] ADVERTENCIA: No se encontro Diagnose-OutlookPerf.ps1 en la carpeta" -ForegroundColor Yellow
    Write-Host "   El script debe estar en la misma carpeta que el EXE para ejecutarse." -ForegroundColor Gray
}

Write-Host "[+] Archivos listos" -ForegroundColor Green

# --- Encontrar compilador C# ---
Write-Host "[~] Buscando compilador C#..."
$cscPath = $null

# Intentar primero con .NET Framework 4.x
$cscFramework = @("C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
                  "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe") |
    Where-Object { Test-Path $_ } | Select-Object -First 1

# Si no, buscar en Visual Studio
if (-not $cscFramework) {
    $vsVersions = @(2022, 2019, 2017)
    foreach ($vsVer in $vsVersions) {
        $vsPath = "C:\Program Files\Microsoft Visual Studio\$vsVer\*\MSBuild\Current\Bin\Roslyn\csc.exe"
        $found = @(Get-Item $vsPath -ErrorAction SilentlyContinue) | Select-Object -First 1
        if ($found) {
            $cscFramework = $found.FullName
            break
        }
    }
}

if (-not $cscFramework) {
    Write-Host "[x] ERROR: No se encontro el compilador C# (csc.exe)" -ForegroundColor Red
    Write-Host "   Instala: .NET Framework 4.5+ o Visual Studio" -ForegroundColor Gray
    exit 1
}

Write-Host "[+] Compilador encontrado: $cscFramework" -ForegroundColor Green

# --- Compilar C# ---
Write-Host "[~] Compilando launcher C#..."

$outputExe = Resolve-Path $OutputPath -ErrorAction SilentlyContinue
if (-not $outputExe) {
    $outputExe = $OutputPath
} else {
    $outputExe = $outputExe.Path
}

$compileOutput = & $cscFramework "/out:$outputExe" "/target:exe" "/optimize" $launcherFile 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "[x] ERROR: La compilacion fallo (exit code: $LASTEXITCODE)" -ForegroundColor Red
    $compileOutput | ForEach-Object { Write-Host "    $_" -ForegroundColor Gray }
    exit 1
}

Write-Host "[+] Compilacion exitosa" -ForegroundColor Green

# --- Verificar salida ---
if (-not (Test-Path $outputExe)) {
    Write-Host "[x] ERROR: El EXE no fue creado." -ForegroundColor Red
    exit 1
}

$exeSize = [Math]::Round((Get-Item $outputExe).Length / 1KB, 1)
Write-Host ""
Write-Host "✓ EXITO" -ForegroundColor Green
Write-Host "  Launcher generado: $outputExe"
Write-Host "  Tamano          : $exeSize KB"
Write-Host ""
Write-Host "Distribucion:" -ForegroundColor Cyan
Write-Host "  1. Copia Outlook-Netrunner.exe"
Write-Host "  2. Copia Diagnose-OutlookPerf.ps1 a la misma carpeta"
Write-Host "  3. Ejecuta: Outlook-Netrunner.exe"
Write-Host ""
Write-Host "Uso:" -ForegroundColor Cyan
Write-Host "  Outlook-Netrunner.exe"
Write-Host "  Outlook-Netrunner.exe -Mailbox usuario@dominio.com"
Write-Host "  Outlook-Netrunner.exe -LocalOnly"
Write-Host "  Outlook-Netrunner.exe -Mailbox usuario@dominio.com -Auto -ReportPath .\informe.html"
Write-Host ""
Write-Host "Requisitos en el equipo target:" -ForegroundColor Cyan
Write-Host "  • PowerShell 5.1 o superior (Windows 7 SP1+ / Server 2008 R2+)"
Write-Host "  • Acceso a Exchange Online (para diagnostico remoto)"
Write-Host ""
