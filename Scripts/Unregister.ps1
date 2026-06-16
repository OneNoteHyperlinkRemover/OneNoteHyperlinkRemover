#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Unregister OneNoteHyperlinkRemover COM add-in.
#>

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir

$assemblyPath = Join-Path $projectDir "bin\Release\OneNoteHyperlinkRemover.dll"
if (-not (Test-Path $assemblyPath)) {
    $assemblyPath = Join-Path $projectDir "bin\Debug\OneNoteHyperlinkRemover.dll"
}

$regasmPath = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if (-not (Test-Path $regasmPath)) {
    $regasmPath = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
}

if (Test-Path $assemblyPath) {
    Write-Host "Unregistering COM component..." -ForegroundColor Cyan
    & $regasmPath /unregister $assemblyPath
}

$addinRegPath = "HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn"
if (Test-Path $addinRegPath) {
    Write-Host "Removing OneNote add-in registry..." -ForegroundColor Cyan
    Remove-Item -Path $addinRegPath -Force
}

Write-Host ""
Write-Host "Unregistration successful!" -ForegroundColor Green
Write-Host "Please restart OneNote to complete uninstall." -ForegroundColor Yellow
