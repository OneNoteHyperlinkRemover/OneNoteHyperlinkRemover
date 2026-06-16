@echo off
:: 注销 OneNoteHyperlinkRemover COM 加载项
:: 需要以管理员权限运行

echo 注销 OneNoteHyperlinkRemover...

set ASSEMBLY=%~dp0..\bin\Release\OneNoteHyperlinkRemover.dll
set REGASM=%windir%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

if not exist "%ASSEMBLY%" (
    set ASSEMBLY=%~dp0..\bin\Debug\OneNoteHyperlinkRemover.dll
)

if exist "%ASSEMBLY%" (
    "%REGASM%" /unregister "%ASSEMBLY%"
)

:: 删除注册表项
reg delete "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /f

echo.
echo 注销成功！请重启 OneNote 以完成卸载。
pause
