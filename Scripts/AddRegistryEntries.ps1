#Requires -RunAsAdministrator
# Add missing COM registry entries to match OneMore's structure

$clsid = "HKLM:\SOFTWARE\Classes\CLSID\{b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e}"
$progId = "HKLM:\SOFTWARE\Classes\OneNoteHyperlinkRemover.AddIn"

# Add Programmable key
New-Item -Path "$clsid\Programmable" -Force | Out-Null
Write-Host "Added: Programmable"

# Add AppID
Set-ItemProperty $clsid -Name "AppID" -Value "{b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e}"
Write-Host "Added: AppID"

# Add VersionIndependentProgID
New-Item -Path "$clsid\VersionIndependentProgID" -Force | Out-Null
Set-ItemProperty "$clsid\VersionIndependentProgID" -Name "(default)" -Value "OneNoteHyperlinkRemover.AddIn"
Write-Host "Added: VersionIndependentProgID"

# Add CurVer on ProgId
New-Item -Path "$progId\CurVer" -Force | Out-Null
Set-ItemProperty "$progId\CurVer" -Name "(default)" -Value "OneNoteHyperlinkRemover.AddIn"
Write-Host "Added: CurVer"

# Ensure LoadBehavior is 3
Set-ItemProperty "HKCU:\Software\Microsoft\Office\OneNote\Addins\OneNoteHyperlinkRemover.AddIn" -Name "LoadBehavior" -Value 3 -Type DWord
Write-Host "LoadBehavior = 3"

Write-Host "`nDone. Restart OneNote to test."
