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
| `CallInitDevMode` | Creates `DEV_<name>.xlsm` from the current workbook and imports all `devkit_*` files from `src/` into it. If the source workbook is `.xlsx`, the DEV_ copy is automatically created as `.xlsm`. |
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

### Testing harness

Treat a workbook like testable code. The optional `devkit_Test.bas` module plus a
PowerShell runner answer a high-value question: **no matter what boundary, blank,
invalid, or random value is entered into any input cell, does the result sheet stay
free of Excel errors (`#DIV/0!`, `#N/A`, ...)?**

- Standard Windows PowerShell only — no Python, no external modules.
- Specs are JSON (`tests/workbook.meta.json`, `tests/no_error.test.json`); results are
  written to `test-results/latest/result.{json,md}`.
- Sheets can be referenced by tab name (`sheet`) or VBA `code_name`; tests can seed cells
  before the run via Markdown `fixtures`.
- Assertions go beyond `no_error`: fixed **scenario** tests and generative **property**
  tests can check `equals` / `blank` / `not_blank` / `not_contains` / `within_range` /
  `matches_regex` / `all_blank_or_hidden`, and failing property cases are automatically
  shrunk to a minimal counterexample. (`matches_regex` uses the bundled pure-VBA
  `devkit_Regex.bas`.)
- The original workbook is never modified (every run uses a work copy under the workbook
  folder, so it inherits the workbook's Excel trusted location).

```powershell
.\tools\New-SampleTestWorkbook.ps1                                   # build a sample (dev only)
.\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm
```

See **[TESTING.md](TESTING.md)** for the full API, JSON schema, and runbook.

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

> **Working with `.xlsx` source files:** `CallInitDevMode` also supports `.xlsx` workbooks. The DEV_ copy is always created as `.xlsm` (macro-enabled format), so all devkit functionality is available regardless of the source format. After development, `CallSaveAsRelease` produces a `MyTool.xlsm` release file; use **File → Save As** if you need a `.xlsx` output instead.

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
- Each `sheet/*.md` is applied to the corresponding sheet: cell values, formulas, background/foreground colors, font sizes, number formats, alignment, cell protection, data-validation lists, named ranges, and merged regions are all restored.
- Sheet maps whose Markdown file has not changed since the last export or import are skipped automatically. File size and modification time are compared against a sync cache (`sheet/xlsm_devkit_sync.tsv`); only changed files are processed, so a full re-import of an unmodified workbook completes in near-instant time.

To import only modules and forms (without sheet maps), run `CallImportAllComponents`.  
To import only sheet maps, run `CallImportAllSheetMapsFromMD`.

### Selective import (`xlsm_devkit.ini`)

By default every token in a sheet map is applied on import. To disable specific categories, place an `xlsm_devkit.ini` file next to the workbook with an `[import]` section:

```ini
[import]
; 1 = apply on import (default if key or file is absent), 0 = skip
; Exception: unlocked defaults to 0 (opt-in) -- applying Locked=False to unintended
; cells can weaken sheet protection; enable explicitly to restore lock state.
; Works with both merge=0 and merge=1 (lazy-unmerge handles slave cells).
value=1
formula=1
numfmt=1
unlocked=0
list=1
bg=0
fg=0
font_size=0
bold=0
italic=0
strike=0
wrap=0
halign=0
valign=0
merge=0
hidden_rows=0
hidden_cols=0
shapes=0
```

The most impactful setting is `merge=0`: it skips `ws.Cells.UnMerge`, a blocking COM call that can take 8--15 minutes on sheets with thousands of merged regions. With `merge=0` the merged-cell structure is left intact and only cell values/formulas are updated.

## Sheet Map Format

Each sheet is exported to `sheet/<CodeName>.md`. The file begins with a metadata
header, followed by a Markdown table, and optionally a Shapes section.

### File header

```
# Sheet Configuration
- VBA CodeName: Sheet1
- Excel UI Name: Summary
- Hidden Rows: 3, 5-7
- Hidden Columns: B, D-F
```

`Hidden Rows` and `Hidden Columns` appear only when the sheet contains hidden rows or columns.

### Cell table

```markdown
| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | Hello | - | Bold |
| B2 | rate | 0.05 | `=A1*0.05` | FontSize:14; NumFmt:0.00% |
| C3 | - | !merged_left | - | - |
```

| Column | Content |
| :--- | :--- |
| `Address` | Cell address (e.g. `A1`, `B12`) |
| `Name` | Named range bound to this cell; `-` if none. Sheet-scoped names include the worksheet prefix (`Sheet1!myRange`). |
| `Value / Label` | Display value (formula result when a formula is present). For merge-slave cells, a merge marker (see below). |
| `Formula` | The formula wrapped in backticks (`` `=A1*B1` ``); `-` if not a formula cell. |
| `Style` | Semicolon-separated style tokens; `-` if none. |

Cells with no value, no formula, and no background color are omitted from the export.

### Style tokens

| Token | Format | Exported when | Import flag | When absent |
| :--- | :--- | :--- | :--- | :--- |
| `BG:#rrggbb` | 6-digit hex | background color set | `bg=1` | unchanged |
| `FG:#rrggbb` | 6-digit hex | font color set | `fg=1` | unchanged |
| `FontSize:<n>` | integer pt | differs from standard font size | `font_size=1` | unchanged |
| `Bold` | flag | `Font.Bold = True` | `bold=1` | reset to `False` |
| `Italic` | flag | `Font.Italic = True` | `italic=1` | reset to `False` |
| `Strike` | flag | `Font.Strikethrough = True` | `strike=1` | reset to `False` |
| `Wrap` | flag | `WrapText = True` | `wrap=1` | reset to `False` |
| `Unlocked` | flag | `Locked = False` | `unlocked=1` | reset to `True` (Locked) |
| `NumFmt:<format>` | Excel format string | not `"General"` | `numfmt=1` | unchanged |
| `List:<formula>` | e.g. `=$A$1:$A$10` | validation list present | `list=1` | unchanged |
| `HAlign:<value>` | see below | never (import-only) | `halign=1` | unchanged |
| `VAlign:<value>` | see below | never (import-only) | `valign=1` | unchanged |

`HAlign` values: `General`, `Left`, `Center`, `Right`, `Fill`, `Justify`, `CenterAcrossSelection`, `Distributed` (case-insensitive).  
`VAlign` values: `Top`, `Center`, `Bottom`, `Justify`, `Distributed` (case-insensitive).

**Boolean-token reset behaviour:** When a boolean import flag is enabled (`bold=1`, etc.) and the token is absent from a cell's style string, the attribute is reset to its default value on import. This mirrors export semantics -- only non-default values are written.

### Merge-slave markers

Slave cells of a merged region appear in the cell table with a marker in the `Value / Label` column:

| Marker | Meaning |
| :--- | :--- |
| `!merged_left` | slave cell in the same row as the master (extends right) |
| `!merged_up` | slave cell in the same column as the master (extends down) |
| `!merged_ul` | slave cell right of and below the master |

Slave rows are skipped during import; merge geometry is rebuilt in a separate pass (`merge=1` required).

### Escape sequences

The following characters are escaped in cell values and style token values:

| Escape | Character |
| :--- | :--- |
| `\\` | literal backslash |
| `\n` | newline (CR, LF, or CRLF) |
| `\v` | vertical tab (Chr 11) |
| `\|` | pipe (Markdown table delimiter) |
| `\;` | semicolon (style token delimiter -- style values only) |

In Shapes fields the literal string `-` (the sentinel for "empty") is written as `\-`.

### Shapes section

When the sheet contains shapes (text boxes, buttons, etc.) a `## Shapes` section is appended after the cell table:

```markdown
## Shapes

| Address | Name | Label | Formula | OnAction | Style |
| :--- | :--- | :--- | :--- | :--- | :--- |
| A1 | Button 1 | Click me | - | Module1.DoAction | BG:#4472C4; FG:#FFFFFF |
```

| Column | Content |
| :--- | :--- |
| `Address` | Top-left cell under the shape |
| `Name` | Shape's internal name |
| `Label` | Visible text (`-` if none; `\-` if the text is literally `-`) |
| `Formula` | Linked cell formula in backticks; `-` if none |
| `OnAction` | Assigned macro name; `-` if none |
| `Style` | `BG:`, `FG:`, and `FontSize:` tokens for the shape |

Import flag: `shapes=1`.

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
  src/                     # Exported .bas and .frm files (UTF-8, BOM-less) + companion .frx binaries
  sheet/                   # Exported sheet map .md files (UTF-8, BOM-less)
  xlsm_devkit.ini          # Optional: selective import settings (see Selective import)
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
