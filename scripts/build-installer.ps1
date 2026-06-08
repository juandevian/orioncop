[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [string]$Platform = "x64",
    [string]$MsBuildPath = "",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

function Get-IsccPath {
    $candidates = @(
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $cmd = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw "No se encontró ISCC.exe (Inno Setup). Instala Inno Setup 6 para generar el instalador."
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $repoRoot

if (-not $SkipBuild) {
    Write-Host "==> Ejecutando build nativo de OrionPlus (default)..."
    $buildScript = Join-Path $repoRoot "scripts\build.ps1"
    & $buildScript -Configuration $Configuration -Platform $Platform -MsBuildPath $MsBuildPath
    if ($LASTEXITCODE -ne 0) {
        throw "Falló scripts\build.ps1"
    }
}

$issFile = Join-Path $repoRoot "scripts\installer\OrionPlusInstaller.iss"
if (-not (Test-Path $issFile)) {
    throw "No se encontró el script Inno: $issFile"
}

$iscc = Get-IsccPath
Write-Host "==> Compilando instalador con Inno Setup..."
& $iscc $issFile
if ($LASTEXITCODE -ne 0) {
    throw "Falló la compilación del instalador Inno Setup."
}

$installerPath = Join-Path $repoRoot "build\installer\qa\OrionPlus-QA-Setup.exe"
if (-not (Test-Path $installerPath)) {
    throw "No se generó el instalador esperado: $installerPath"
}

$file = Get-Item $installerPath
Write-Host ("Instalador generado: {0} ({1} bytes, {2})" -f $file.FullName, $file.Length, $file.LastWriteTime)
