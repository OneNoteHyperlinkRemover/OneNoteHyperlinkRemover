#Requires -RunAsAdministrator
<#
.SYNOPSIS
    注销 OneNoteHyperlinkRemover COM 加载项。
#>

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# 尝试 Release 和 Debug 路径
$assemblyPath = Join-Path $scriptDir "bin\Release\OneNoteHyperlinkRemover.dll"
if (-not (Test-Path $assemblyPath)) {
    $assemblyPath = Join-Path $scriptDir "bin\Debug\OneNoteHyperlinkRemover.dll"
}

$regasmPath = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
if (-not (Test-Path $regasmPath)) {
    $regasmPath = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
}

if (Test-Path $assemblyPath) {
    Write-Host "注销 COM 组件..." -ForegroundColor Cyan
    & $regasmPath /unregister $assemblyPath
}

# 删除注册表项
$addinRegPath = "HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn"
if (Test-Path $addinRegPath) {
    Write-Host "删除注册表项..." -ForegroundColor Cyan
    Remove-Item -Path $addinRegPath -Force
}

Write-Host ""
Write-Host "注销成功！" -ForegroundColor Green
Write-Host "请重启 OneNote 以完成卸载。" -ForegroundColor Yellow
