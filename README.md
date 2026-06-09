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
| `CallInitDevMode` | Creates `DEV_<name>.xlsm` from the current workbook and imports all `devkit_*` files from `src/` into it |
| `CallSaveAsRelease` | Saves a production copy (strips `DEV_` prefix, removes all devkit modules); call from a `DEV_` workbook |

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

### Launcher

A central dialog that exposes all import/export operations in one place and lets users switch the UI language. Buttons for InsertDelete and Move also appear when those modules are loaded.

Required files (all must be imported into the same VBA project):

| File | Role |
| :--- | :--- |
| `devkit_Launch.bas` | Feature logic |
| `devkit_frmLauncher.frm` + `devkit_frmLauncher.frx` | Launcher dialog |

Entry point: `ShowLauncherForm`

### Internationalization

UI text in all optional-feature dialogs is localized via INI-based language files. The i18n functions (`t()`, `Fmt()`, `SetLang()`, `GetLangCode()`) are built into `xlsm_devkit.bas` — no additional VBA module is required. Place the `lang/` folder next to the workbook to enable localization.

27 languages are bundled: Arabic, Bengali, English, Spanish, Persian, French, German, Hindi, Indonesian, Japanese, Javanese, Korean, Malay, Marathi, Portuguese, Punjabi, Russian, Swahili, Tamil, Telugu, Thai, Turkish, Ukrainian, Urdu, Vietnamese, Simplified Chinese, and Traditional Chinese.

The active language is auto-detected from Windows. It can be overridden via the Launcher's language selector or programmatically:

```vba
SetLang "ja"   ' switch to Japanese
SetLang ""     ' revert to system auto-detection
```

## Usage

### Initial setup (add to a new workbook)

1. Open your target `.xlsm` workbook and press `Alt + F11` to open the VBE.
2. In Project Explorer, right-click the target VBA project and import `xlsm_devkit.bas`.
3. Enable "Trust access to the VBA project object model" in Excel settings.
4. Run `ExportAllModulesFormsSheetMaps` once and confirm that `src/` and `sheet/` are created next to the workbook.
5. Edit files in `src/` with VS Code, then run `ImportAllModulesFormsSheetMaps`.

### DEV / release workflow

When you want to keep the user-facing workbook free from devkit modules, use the `DEV_` naming convention:

**Start development (`CallInitDevMode`)**

1. Open the production workbook (e.g. `MyTool.xlsm`) and import `xlsm_devkit.bas` manually (same as the initial setup step above).
2. Place the optional `devkit_*.bas`, `devkit_*.frm`, and `devkit_*.frx` files from the xlsm_devkit release into `src/` next to the workbook (only the ones you need).
3. Run `CallInitDevMode` from the Macro dialog (`Alt+F8`). It creates `DEV_MyTool.xlsm` in the same folder and imports all `devkit_*` files from `src/` into the copy.
4. Close `MyTool.xlsm` **without saving** — this keeps the production file free from devkit modules.
5. Open `DEV_MyTool.xlsm` and develop there.

**Release (`CallSaveAsRelease`)**

When you are ready to ship:

1. From `DEV_MyTool.xlsm`, run `CallSaveAsRelease` from the Macro dialog (`Alt+F8`), or use the **Save as Release** button in the Launcher (enabled only when the workbook name starts with `DEV_`).
2. It saves a copy named `MyTool.xlsm` (strips `DEV_` prefix) and removes all `xlsm_devkit` and `devkit_*` modules from that copy.
3. The `DEV_MyTool.xlsm` workbook is unchanged — continue developing from it.

### Exporting

Run `ExportAllModulesFormsSheetMaps` to export all modules, forms, and sheet maps at once.

- Each module → `src/*.bas` (UTF-8, BOM-less), each form → `src/*.frm` (UTF-8, BOM-less), binary resources → `src/*.frx`.
- Each sheet → `sheet/*.md` (UTF-8, BOM-less).

To export only modules and forms (without sheet maps), run `CallExportAllComponents`.  
To export only sheet maps, run `CallExportAllSheetMapsToMD`.

### Importing

Run `ImportAllModulesFormsSheetMaps` to import all modules, forms, and sheet maps at once.

- Files from `src/` are loaded into the project. Existing modules/forms are updated; new ones are added. Companion `*.frx` files are picked up automatically.
- Each `sheet/*.md` is applied to the corresponding sheet: cell values, formulas, background/foreground colors, font sizes, data-validation lists, named ranges, and merged regions are all restored.

To import only modules and forms (without sheet maps), run `CallImportAllComponents`.  
To import only sheet maps, run `CallImportAllSheetMapsFromMD`.

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

## Upgrading

When you replace `xlsm_devkit.bas` with a newer version, re-export all sheet maps before importing them again:

1. Import the new `xlsm_devkit.bas` into the VBE.
2. Run `CallExportAllSheetMapsToMD` to re-generate `sheet/*.md` in the current format.
3. Commit the refreshed `sheet/*.md` files.

Sheet map files generated by an older version may not import correctly with a newer version because the sheet map format can change between releases.

## Limitations

- `xlsm_devkit` itself is never imported by `ImportAllModulesFormsSheetMaps` or `CallImportAllComponents` — a running module cannot delete or overwrite itself. To update `xlsm_devkit`, paste the new code manually in the VBE.
- When `SKIP_DEVKIT_MODULES = True` (the default): `devkit_*` optional modules and forms are skipped during both import and export, and `xlsm_devkit` itself is also skipped during export. Set `SKIP_DEVKIT_MODULES = False` inside `xlsm_devkit.bas` when developing the optional devkit modules.
- While a Move capture is in progress, `devkit_Move` cannot be reimported — it is on the active call stack, and reimporting would reset the VBA runtime and crash Excel.
- Requires Windows and Microsoft Excel with VBA.

## Tested environment

- Windows
- Verified: Microsoft Excel 2010 or later (32-bit and 64-bit)

## Version compatibility

- 32-bit Excel: Intended to work on Excel 2007 or later (VBA6 branch included; not yet tested)
- 64-bit Excel: Excel 2010 or later (`VBA7` / `PtrSafe` required)
- This project depends on Windows Excel because it uses `GetACP` and `VBProject` automation
