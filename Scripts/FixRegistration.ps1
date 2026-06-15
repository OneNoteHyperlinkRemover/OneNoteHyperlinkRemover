#Requires -RunAsAdministrator
# Fix COM registration to match OneMore's structure exactly

$ourClsid = "{b7a3d2e1-4f6c-4a8b-9e1d-3c5f7a9b2d4e}"
$progId = "OneNoteHyperlinkRemover.AddIn"
$asmPath = "D:\work\OneNote add-ins\OneNote_add-ins_com\bin\Release\OneNoteHyperlinkRemover.dll"

# First, unregister old entries
$regasm = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
& $regasm /unregister $asmPath 2>&1 | Out-Null

# Clean up old entries
Remove-Item "HKLM:\SOFTWARE\Classes\$progId" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKLM:\SOFTWARE\Classes\CLSID\$ourClsid" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Cleaned up old entries"

# === ProgId ===
New-Item -Path "HKLM:\SOFTWARE\Classes\$progId" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\$progId" -Name "(default)" -Value "$progId"
New-Item -Path "HKLM:\SOFTWARE\Classes\$progId\CLSID" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\$progId\CLSID" -Name "(default)" -Value $ourClsid
New-Item -Path "HKLM:\SOFTWARE\Classes\$progId\CurVer" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\$progId\CurVer" -Name "(default)" -Value "$progId.1"
Write-Host "Registered ProgId"

# === Versioned ProgId ===
New-Item -Path "HKLM:\SOFTWARE\Classes\$progId.1" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\$progId.1" -Name "(default)" -Value "Addin class"
New-Item -Path "HKLM:\SOFTWARE\Classes\$progId.1\CLSID" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\$progId.1\CLSID" -Name "(default)" -Value $ourClsid
Write-Host "Registered versioned ProgId"

# === CLSID ===
$clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$ourClsid"
New-Item -Path $clsidPath -Force | Out-Null
Set-ItemProperty $clsidPath -Name "(default)" -Value $progId
Set-ItemProperty $clsidPath -Name "AppID" -Value $ourClsid

# Implemented Categories
New-Item -Path "$clsidPath\Implemented Categories" -Force | Out-Null
New-Item -Path "$clsidPath\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}" -Force | Out-Null

# InprocServer32
New-Item -Path "$clsidPath\InprocServer32" -Force | Out-Null
Set-ItemProperty "$clsidPath\InprocServer32" -Name "(default)" -Value "mscoree.dll"
Set-ItemProperty "$clsidPath\InprocServer32" -Name "ThreadingModel" -Value "Both"
Set-ItemProperty "$clsidPath\InprocServer32" -Name "CodeBase" -Value "file:///$($asmPath.Replace('\','/'))"
Set-ItemProperty "$clsidPath\InprocServer32" -Name "Class" -Value $progId
Set-ItemProperty "$clsidPath\InprocServer32" -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty "$clsidPath\InprocServer32" -Name "Assembly" -Value "OneNoteHyperlinkRemover, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"

# Versioned InprocServer32
New-Item -Path "$clsidPath\InprocServer32\1.0.0.0" -Force | Out-Null
Set-ItemProperty "$clsidPath\InprocServer32\1.0.0.0" -Name "Assembly" -Value "OneNoteHyperlinkRemover, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty "$clsidPath\InprocServer32\1.0.0.0" -Name "CodeBase" -Value "file:///$($asmPath.Replace('\','/'))"
Set-ItemProperty "$clsidPath\InprocServer32\1.0.0.0" -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty "$clsidPath\InprocServer32\1.0.0.0" -Name "Class" -Value $progId

# ProgId
New-Item -Path "$clsidPath\ProgId" -Force | Out-Null
Set-ItemProperty "$clsidPath\ProgId" -Name "(default)" -Value $progId

# Programmable
New-Item -Path "$clsidPath\Programmable" -Force | Out-Null

# TypeLib (use CLSID as TypeLib ID - same as OneMore)
New-Item -Path "$clsidPath\TypeLib" -Force | Out-Null
Set-ItemProperty "$clsidPath\TypeLib" -Name "(default)" -Value $ourClsid

# VersionIndependentProgID
New-Item -Path "$clsidPath\VersionIndependentProgID" -Force | Out-Null
Set-ItemProperty "$clsidPath\VersionIndependentProgID" -Name "(default)" -Value $progId

Write-Host "Registered CLSID"

# === AppID ===
New-Item -Path "HKLM:\SOFTWARE\Classes\AppID\$ourClsid" -Force | Out-Null
Set-ItemProperty "HKLM:\SOFTWARE\Classes\AppID\$ourClsid" -Name "DllSurrogate" -Value ""
Write-Host "Registered AppID"

# === HKCU entries (same as OneMore) ===
New-Item -Path "HKCU:\SOFTWARE\Classes\AppID\$ourClsid" -Force | Out-Null
Set-ItemProperty "HKCU:\SOFTWARE\Classes\AppID\$ourClsid" -Name "DllSurrogate" -Value ""

# OneNote Addins key (both HKLM and HKCU)
foreach ($root in @("HKLM:\SOFTWARE\Microsoft\Office\OneNote\Addins", "HKCU:\SOFTWARE\Microsoft\Office\OneNote\Addins")) {
    New-Item -Path "$root\$progId" -Force | Out-Null
    Set-ItemProperty "$root\$progId" -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty "$root\$progId" -Name "Description" -Value "Remove auto-hyperlinks"
    Set-ItemProperty "$root\$progId" -Name "FriendlyName" -Value "HyperlinkRemover"
}
Write-Host "Registered OneNote Addins (HKLM + HKCU)"

Write-Host "`nDone! Restart OneNote to test."
