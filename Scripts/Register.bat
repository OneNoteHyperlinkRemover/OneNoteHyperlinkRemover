@echo off
:: Register OneNoteHyperlinkRemover COM add-in (simple version)
:: Requires Administrator privileges

echo Registering OneNoteHyperlinkRemover...

set ASSEMBLY=%~dp0..\bin\Release\OneNoteHyperlinkRemover.dll
set REGASM=%windir%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe

if not exist "%ASSEMBLY%" (
    echo Error: Assembly not found: %ASSEMBLY%
    echo Please build the project first.
    pause
    exit /b 1
)

"%REGASM%" /codebase "%ASSEMBLY%"
if errorlevel 1 (
    echo Error: RegAsm registration failed
    pause
    exit /b 1
)

reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v Description /t REG_SZ /d "Remove auto-converted URL hyperlinks" /f
reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v FriendlyName /t REG_SZ /d "OneNote Hyperlink Remover" /f
reg add "HKCU\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" /v LoadBehavior /t REG_DWORD /d 3 /f

echo.
echo Registration successful! Please restart OneNote.
pause
