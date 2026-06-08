[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [string]$Platform = "x64",
    [string]$SolutionPath = ".\OrionCop.Net.sln",
    [string]$MsBuildPath = "",
    [switch]$SkipRestore
)

$ErrorActionPreference = "Stop"
$nativeOutputDir = "build\releases\orion-plus"

function Get-MSBuildPath {
    $candidates = @(
        "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe",
        "C:\Program Files\Microsoft Visual Studio\18\Insiders\MSBuild\Current\bin\MSBuild.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    $cmd = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw "No se encontró MSBuild. Instala Visual Studio Build Tools o especifica -MsBuildPath."
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $repoRoot

if (-not (Test-Path $SolutionPath)) {
    throw "No se encontró la solución: $SolutionPath"
}

if (-not $SkipRestore) {
    Write-Host "==> Restaurando paquetes NuGet..."
    & nuget restore $SolutionPath
    if ($LASTEXITCODE -ne 0) {
        throw "Falló nuget restore para $SolutionPath"
    }
}

$resolvedMsBuildPath = $MsBuildPath
if ([string]::IsNullOrWhiteSpace($resolvedMsBuildPath)) {
    $resolvedMsBuildPath = Get-MSBuildPath
}

if (-not (Test-Path $resolvedMsBuildPath)) {
    throw "No existe MSBuild en: $resolvedMsBuildPath"
}

Write-Host "==> Compilando OrionCop ($Configuration|$Platform)..."
& $resolvedMsBuildPath $SolutionPath "/p:Configuration=$Configuration" "/p:Platform=$Platform" "/m" "/nologo"
if ($LASTEXITCODE -ne 0) {
    throw "Falló la compilación con MSBuild."
}

$orionPlusPath = Join-Path $repoRoot ($nativeOutputDir + "\OrionPlus.exe")
if (-not (Test-Path $orionPlusPath)) {
    throw "La compilación terminó, pero no existe OrionPlus.exe en: $orionPlusPath"
}

$artifact = Get-Item $orionPlusPath

$requiredRuntimeFiles = @(
    "OrionCopIU.exe.config",
    "OrionCopL.dll",
    "OriIntCon.dll",
    "RepOriCop.dll",
    "PanL.dll",
    "PanDat.dll",
    "OriWin.dll",
    "WinCom.dll"
)

$missing = @()
foreach ($file in $requiredRuntimeFiles) {
    $path = Join-Path $repoRoot ($nativeOutputDir + "\" + $file)
    if (-not (Test-Path $path)) {
        $missing += $file
    }
}

if ($missing.Count -gt 0) {
    throw ("Faltan archivos runtime en " + $nativeOutputDir + ": " + ($missing -join ", "))
}

Write-Host "==> Build completado."
Write-Host ("OrionPlus.exe: {0} ({1} bytes, {2})" -f $artifact.FullName, $artifact.Length, $artifact.LastWriteTime)
