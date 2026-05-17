# xlsm_devkit

[日本語版 README はこちら](README.ja.md)

A development toolkit for Excel VBA projects that provides module/form import/export with encoding conversion and sheet map export.  
Designed to support xlsm development workflows that combine the VBE (Visual Basic Editor) with external editors such as VS Code.

## Core module

Import `xlsm_devkit.bas` alone to get the following macros:

| Macro | Description |
| :--- | :--- |
| `ExportAllModulesFormsSheetMaps` | Exports all VBA modules and forms to `src/`, and all sheet maps to `sheet/` |
| `ImportAllModulesFormsSheetMaps` | Imports all VBA modules and forms from `src/`, and all sheet maps from `sheet/` |
| `CallExportAllComponents` | Exports all VBA modules and forms to `src/` |
| `CallImportAllComponents` | Imports all VBA modules and forms from `src/` |
| `CallExportAllSheetMapsToMD` | Exports cell values, shapes, formulas, and styles of all sheets to `sheet/*.md` |
| `CallImportAllSheetMapsFromMD` | Restores cell values, formulas, styles, named ranges, data-validation lists, and merged regions from `sheet/*.md` |

Modules (`.bas`) and forms (`.frm`/`.frx`) are handled together by the same export/import operation.  
`xlsm_devkit` itself is never imported — a running module cannot overwrite itself.

## Optional features

These features require additional files to be imported alongside `xlsm_devkit.bas`.

### InsertDelete

Inserts or deletes rows/columns, saves before/after sheet maps as Markdown files, and generates an AI prompt for updating affected VBA references.

Required files (all must be imported into the same VBA project):

| File | Role |
| :--- | :--- |
| `devkit_InsertDelete.bas` | Feature logic |
| `devkit_frmInsertDelete.frm` + `devkit_frmInsertDelete.frx` | Setup dialog |
| `devkit_frmInstruction.frm` + `devkit_frmInstruction.frx` | Result/import dialog (shared with Move) |

Entry point: `ShowInsertDeleteForm`

### Move

Records a cell-range cut-and-paste operation via Excel's macro recorder, captures before/after sheet maps, and generates an AI prompt for updating affected VBA references.

Required files (all must be imported into the same VBA project):

| File | Role |
| :--- | :--- |
| `devkit_Move.bas` | Feature logic |
| `devkit_frmMoveSetup.frm` + `devkit_frmMoveSetup.frx` | Setup dialog |
| `devkit_frmMoveWait.frm` + `devkit_frmMoveWait.frx` | Recording dialog |
| `devkit_frmInstruction.frm` + `devkit_frmInstruction.frx` | Result/import dialog (shared with InsertDelete) |

Entry point: `ShowMoveSetupForm`

## Usage

### Initial setup (add to a new workbook)

1. Open your target `.xlsm` workbook and press `Alt + F11` to open the VBE.
2. In Project Explorer, right-click the target VBA project and import `xlsm_devkit.bas`.
3. Enable "Trust access to the VBA project object model" in Excel settings.
4. Run `CallExportAllComponents` once and confirm that `src/` is created next to the workbook.
5. Edit files in `src/` with VS Code, then run `CallImportAllComponents`.

### Exporting modules and forms

1. Run the `CallExportAllComponents` macro in Excel.
2. Each module is written as a `.bas` file (UTF-8, BOM-less) and each form as a `.frm` file (UTF-8, BOM-less) to a `src/` folder next to the workbook. If a form has binary resources, VBE also writes a companion `.frx` file.

### Importing modules and forms

1. Run the `CallImportAllComponents` macro.
2. Files from `src/` are loaded into the project. Existing modules and forms have their code updated; new modules and forms are added. The companion `*.frx` in the same folder is picked up automatically.

### Exporting sheet maps

1. Run the `CallExportAllSheetMapsToMD` macro.
2. A Markdown file for each sheet is written to a `sheet/` folder next to the workbook.

### Importing sheet maps

1. Run the `CallImportAllSheetMapsFromMD` macro.
2. Each `sheet/*.md` file is read and applied to the corresponding sheet: cell values, formulas, background/foreground colors, font sizes, data-validation lists, named ranges, and merged regions are all restored.

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
This module uses ADODB.Stream and the Win32 API `GetACP()` to keep files on disk in BOM-less UTF-8 while transparently converting to and from ANSI for the VBE.

## File layout

```
<workbook folder>/
  src/          # Exported .bas and .frm files (UTF-8, BOM-less) + companion .frx binaries
  sheet/        # Exported sheet map .md files (UTF-8, BOM-less)
```

## Limitations

- `xlsm_devkit` itself is never imported by `CallImportAllComponents` — a running module cannot delete or overwrite itself. To update `xlsm_devkit`, paste the new code manually in the VBE.
- Requires Windows and Microsoft Excel with VBA.

## Tested environment

- Windows
- Verified: Microsoft Excel 2010 or later (32-bit and 64-bit)

## Version compatibility

- 32-bit Excel: Intended to work on Excel 2007 or later (VBA6 branch included; not yet tested)
- 64-bit Excel: Excel 2010 or later (`VBA7` / `PtrSafe` required)
- This project depends on Windows Excel because it uses `GetACP` and `VBProject` automation
