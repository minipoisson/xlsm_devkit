# xlsm_devkit

[日本語版 README はこちら](README.ja.md)

A development toolkit for Excel VBA projects that provides module import/export with encoding conversion and sheet map export.  
Designed to support xlsm development workflows that combine the VBE (Visual Basic Editor) with external editors such as VS Code.

## Features

| Macro | Description |
| :--- | :--- |
| `ExportAllModules` | Exports all VBA modules to `src/*.bas` (ANSI → UTF-8 conversion) |
| `ImportAllModules` | Imports `src/*.bas` into the VBA project (UTF-8 → ANSI conversion). Skips `xlsm_devkit` itself |
| `ExportAllSheetMapsToMD` | Exports cell values, shapes, formulas, and styles of all sheets to `sheet/*.md` |

## Usage

### Initial setup (add to a new workbook)

1. Open your target `.xlsm` workbook and press `Alt + F11` to open the VBE.
2. In Project Explorer, right-click the target VBA project and import `xlsm_devkit.bas`.
3. Enable "Trust access to the VBA project object model" in Excel settings.
4. Run `ExportAllModules` once and confirm that `src/` is created next to the workbook.
5. Edit files in `src/*.bas` with VS Code, then run `ImportAllModules` while `xlsm_devkit` is selected in the VBE.

### Exporting modules

1. Run the `ExportAllModules` macro in Excel.
2. Each module is written as a `.bas` file (UTF-8) to a `src/` folder next to the workbook.

### Importing modules

> **Important prerequisite**
>
> `ImportAllModules` uses `Application.VBE.ActiveCodePane` to obtain its own module name so that it can skip itself during import.  
> Therefore, you must **open the VBE and select the `xlsm_devkit` module** before running the macro.  
> If a different module or window is active, the self-skip logic will not work correctly.

1. Open the VBE and select the `xlsm_devkit` module in the Project Explorer.
2. Run the `ImportAllModules` macro.
3. Files from `src/*.bas` are loaded into the project. Existing modules have their code replaced; new modules are added.

### Exporting sheet maps

1. Run the `ExportAllSheetMapsToMD` macro.
2. A Markdown file for each sheet is written to a `sheet/` folder next to the workbook.

## Prerequisites

### Trust access to the VBA project object model

Enable the following setting in Excel:

```
File → Options → Trust Center → Trust Center Settings
  → Macro Settings → Check "Trust access to the VBA project object model"
```

Both export and import will fail if this setting is disabled.

### Character encoding

VBA's `VBComponents.Export` / `VBComponents.Import` always use the system ANSI code page (e.g. Shift_JIS on Japanese Windows) when reading and writing files.  
This module uses ADODB.Stream and the Win32 API `GetACP()` to keep files on disk in UTF-8 while transparently converting to and from ANSI for the VBE.

## File layout

```
<workbook folder>/
  src/          # Exported .bas files (UTF-8)
  sheet/        # Exported sheet map .md files (UTF-8)
```

## Limitations

- `xlsm_devkit` itself is never imported by `ImportAllModules` — a running module cannot delete or overwrite itself. To update `xlsm_devkit`, paste the new code manually in the VBE or handle it as a separate import step.
- Form modules (`Type = 3`) are excluded from export.
- Requires Windows and Microsoft Excel with VBA.

## Tested environment

- Windows
- Verified: Microsoft Excel 2010 or later (32-bit and 64-bit)

## Version compatibility

- 32-bit Excel: Intended to work on Excel 2007 or later (VBA6 branch included; not yet tested)
- 64-bit Excel: Excel 2010 or later (`VBA7` / `PtrSafe` required)
- This project depends on Windows Excel because it uses `GetACP` and `VBProject` automation
