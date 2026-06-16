@echo off
:: Unregister OneNoteHyperlinkRemover COM add-in
:: Requires Administrator privileges

echo Unregistering OneNoteHyperlinkRemover...

set ASSEMBLY=%~dp0..\bin\Release\OneNoteHyperlinkRemover.dll
set REGASM=%windir%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

if not exist "%ASSEMBLY%" (
    set ASSEMBLY=%~dp0..\bin\Debug\OneNoteHyperlinkRemover.dll
)

if exist "%ASSEMBLY%" (
    "%REGASM%" /unregister "%ASSEMBLY%"
)

reg delete "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /f

echo.
echo Unregistration successful! Please restart OneNote.
pause
