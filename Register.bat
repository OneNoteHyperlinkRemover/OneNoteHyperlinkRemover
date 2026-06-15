@echo off
:: 注册 OneNoteHyperlinkRemover COM 加载项（简易版）
:: 需要以管理员权限运行

echo 注册 OneNoteHyperlinkRemover...

set ASSEMBLY=%~dp0bin\Release\OneNoteHyperlinkRemover.dll
set REGASM=%windir%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

if not exist "%ASSEMBLY%" (
    echo 错误: 找不到程序集 %ASSEMBLY%
    echo 请先编译项目。
    pause
    exit /b 1
)

"%REGASM%" /codebase "%ASSEMBLY%"
if errorlevel 1 (
    echo 错误: RegAsm 注册失败
    pause
    exit /b 1
)

:: 添加 OneNote 加载项注册表项
reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v Description /t REG_SZ /d "移除自动超链接转换" /f
reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v FriendlyName /t REG_SZ /d "超链接移除工具" /f
reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo 注册成功！请重启 OneNote 以加载插件。
pause
