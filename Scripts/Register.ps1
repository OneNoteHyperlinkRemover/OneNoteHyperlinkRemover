#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Register OneNoteHyperlinkRemover COM add-in.
.DESCRIPTION
    Uses regasm to register the assembly as a COM component and adds OneNote add-in registry entry.
    Requires Administrator privileges.
#>

param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$assemblyPath = Join-Path $projectDir "bin\$Configuration\OneNoteHyperlinkRemover.dll"
$regasmPath = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

if (-not (Test-Path $assemblyPath)) {
    Write-Error "Assembly not found: $assemblyPath`nPlease build the project first."
    exit 1
}

if (-not (Test-Path $regasmPath)) {
    $regasmPath = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    if (-not (Test-Path $regasmPath)) {
        Write-Error "RegAsm.exe not found. Please install .NET Framework 4.x."
        exit 1
    }
}

Write-Host "Registering COM component..." -ForegroundColor Cyan
& $regasmPath /codebase $assemblyPath
if ($LASTEXITCODE -ne 0) {
    Write-Error "RegAsm registration failed."
    exit 1
}

$addinRegPath = "HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn"

Write-Host "Setting OneNote add-in registry..." -ForegroundColor Cyan
if (-not (Test-Path $addinRegPath)) {
    New-Item -Path $addinRegPath -Force | Out-Null
}

Set-ItemProperty -Path $addinRegPath -Name "Description" -Value "Remove auto-converted URL hyperlinks" -Type String
Set-ItemProperty -Path $addinRegPath -Name "FriendlyName" -Value "OneNote Hyperlink Remover" -Type String
Set-ItemProperty -Path $addinRegPath -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host ""
Write-Host "Registration successful!" -ForegroundColor Green
Write-Host "Please restart OneNote to load the add-in." -ForegroundColor Yellow
Write-Host ""
Write-Host "To uninstall, run: .\Unregister.ps1" -ForegroundColor Gray
