#Requires -RunAsAdministrator
<#
.SYNOPSIS
    注册 OneNoteHyperlinkRemover COM 加载项。
.DESCRIPTION
    使用 regasm 注册程序集为 COM 组件，并在注册表中添加 OneNote 加载项条目。
    需要以管理员权限运行。
#>

param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$assemblyPath = Join-Path $projectDir "bin\$Configuration\OneNoteHyperlinkRemover.dll"
$regasmPath = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"

# 检查程序集是否存在
if (-not (Test-Path $assemblyPath)) {
    Write-Error "找不到程序集: $assemblyPath`n请先编译项目。"
    exit 1
}

# 检查 regasm 是否存在
if (-not (Test-Path $regasmPath)) {
    # 尝试 32 位路径
    $regasmPath = "$env:windir\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    if (-not (Test-Path $regasmPath)) {
        Write-Error "找不到 RegAsm.exe，请确保安装了 .NET Framework 4.x。"
        exit 1
    }
}

Write-Host "注册 COM 组件..." -ForegroundColor Cyan
& $regasmPath /codebase $assemblyPath
if ($LASTEXITCODE -ne 0) {
    Write-Error "RegAsm 注册失败。"
    exit 1
}

# 设置 OneNote 加载项注册表项
$addinRegPath = "HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn"

Write-Host "设置 OneNote 加载项注册表..." -ForegroundColor Cyan
if (-not (Test-Path $addinRegPath)) {
    New-Item -Path $addinRegPath -Force | Out-Null
}

Set-ItemProperty -Path $addinRegPath -Name "Description" -Value "移除自动超链接转换" -Type String
Set-ItemProperty -Path $addinRegPath -Name "FriendlyName" -Value "超链接移除工具" -Type String
Set-ItemProperty -Path $addinRegPath -Name "LoadBehavior" -Value 3 -Type DWord

Write-Host ""
Write-Host "注册成功！" -ForegroundColor Green
Write-Host "请重启 OneNote 以加载插件。" -ForegroundColor Yellow
Write-Host ""
Write-Host "如需卸载，请运行: .\Unregister.ps1" -ForegroundColor Gray
