# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OneNote COM Add-in that removes automatic URL-to-hyperlink conversion. OneNote auto-converts typed/pasted URLs into clickable hyperlinks with no disable setting; this add-in scans page content and strips auto-converted `<a>` tags, restoring plain text. UI supports Chinese and English (auto-detected from system locale).

## Technology

This is a **COM Add-in** (not VSTO) — OneNote does not support VSTO project templates. The add-in implements `IDTExtensibility2` and `IRibbonExtensibility` directly, targeting .NET Framework 4.8 (C# 9.0). OneNote page content is accessed via `Microsoft.Office.Interop.OneNote.IApplication` COM API and manipulated as XML using LINQ to XML.

Reference project: [OneMore](https://github.com/stevencohn/OneMore) — a large-scale OneNote COM add-in in C#.

## Commands

- **Build**: `msbuild OneNoteHyperlinkRemover.sln /p:Configuration=Release` (or open in Visual Studio 2022)
- **Generate icons**: `powershell -ExecutionPolicy Bypass -File Scripts\GenerateIcons.ps1`
- **Register add-in** (admin PowerShell): `.\Scripts\Register.ps1 -Configuration Release`
- **Unregister add-in** (admin PowerShell): `.\Scripts\Unregister.ps1`
- **View logs** (DEBUG only): `%LOCALAPPDATA%\OneNoteHyperlinkRemover\addin.log`

No test framework, linter, or formatter is configured.

## Architecture

**Key files**:

| File | Role |
|------|------|
| `AddIn.cs` | Entry point — `IDTExtensibility2` lifecycle + `IRibbonExtensibility` Ribbon callbacks |
| `OneNoteHelper.cs` | Wraps `IApplication` COM object; manages page read/write with explicit `Marshal.ReleaseComObject` |
| `HyperlinkRemover.cs` | Core logic: parse page XML → find auto-converted `<a>` in CDATA → replace with plain text |
| `ClipboardMonitor.cs` | Background clipboard watcher that strips zero-width spaces from copied text |
| `Strings.cs` | Bilingual string dictionary (zh/en), auto-detected from `CultureInfo.CurrentUICulture` |
| `Logger.cs` | Simple file logger (DEBUG builds only, compiled out in Release via `[Conditional("DEBUG")]`) |
| `Ribbon\Ribbon.xml` | Ribbon UI definition — all labels use `getLabel` callbacks for i18n |
| `Properties\AssemblyInfo.cs` | Assembly metadata and COM visibility |

**Page XML structure** (OneNote namespace `one:`):
```
one:Page > one:Outline > one:OEChildren > one:OE > one:T (CDATA containing HTML <a> tags)
```

**Hyperlink detection**: Regex matches `<a href="...">text</a>` inside CDATA sections of `<T>` elements. Auto-converted links are identified when `href` equals the display text, or the display text starts with `http`.

**Anti-re-linking**: After stripping a hyperlink, `BreakUrlPattern` inserts zero-width spaces at `://` and `www.` to prevent OneNote from auto-converting the URL again. `ClipboardMonitor` cleans these characters when text is copied.

**COM lifecycle**: `OnConnection` deliberately does NOT store the `IApplication` reference (to avoid preventing OneNote shutdown — lesson from OneMore). Each operation creates a fresh COM object via `OneNoteHelper` and releases it immediately.

**Bilingual**: All Ribbon labels are resolved at runtime via `getLabel`/`getScreentip`/`getSupertip` callbacks that read from `Strings.cs`. Language is detected once from `CultureInfo.CurrentUICulture` (Chinese if `zh`, otherwise English).

## Development Workflow

1. Build in Visual Studio 2022 (requires "Office/SharePoint development" workload)
2. Run `Scripts\Register.ps1` as admin to register the COM component and add registry entry
3. Restart OneNote — find the add-in group on the Home tab
4. For debugging: attach VS to `ONENOTE.EXE` process
5. After code changes: rebuild → re-register → restart OneNote

## Key Constraints

- Cannot intercept/prevent auto-hyperlink conversion in real-time; only post-hoc removal
- OneNote COM API does not support undo — users should back up content before bulk removal
- Must target .NET Framework 4.x (not .NET Core/5+)
- COM registration requires admin privileges
- `IApplication` COM references must be released explicitly to avoid blocking OneNote shutdown
