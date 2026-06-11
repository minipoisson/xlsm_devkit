# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added
- `InitDevMode` (.xlsx support): when the source workbook is `.xlsx`, the DEV_ copy is now created as `.xlsm` (macro-enabled format) instead of `.xlsx`. A temporary `.xlsx` file is used internally during the copy+import step and deleted after `SaveAs` completes. This allows `CallInitDevMode` to be used on `.xlsx` workbooks that contain complex formulas or data, while still injecting the devkit VBA components correctly.

## [1.6.0] - 2026-06-09

### Added
- Export progress: `Application.StatusBar` shows "Exporting sheet map N/N: `<sheet>`" during `ExportAllSheetMapsToMD`; status bar is restored in a `lblFin` handler even if an export error occurs.
- `NumFmt:<format>` style token: exports the cell's number format (when it is not `General`) and re-applies it on import (roundtrip).
- `Unlocked` style token: exports `rng.Locked = False` and re-applies on import (roundtrip).
- `HAlign:<value>` / `VAlign:<value>` import-only style tokens: apply horizontal or vertical alignment when written manually in a sheet map (`general`, `left`, `center`, `right`, `fill`, `justify`, `distributed`, and numeric `XlHAlign` / `XlVAlign` values are accepted; not exported).
- `EscapeStyleValue` / `UnescapeStyleValue`: new private helpers; extend `EscapeCellValue` / `UnescapeCellValue` by also escaping `;` as `\;`, allowing number-format strings that contain semicolons (e.g. `#,##0.00;[Red]-#,##0.00`) to round-trip correctly.
- `\-` escape sequence (`\-` decodes to `-`) added to both `UnescapeCellValue` and `UnescapeStyleValue`; shape export now emits `\-` instead of the `-` sentinel when a shape label or OnAction string is the literal character `"-"`.
- `ParseStyleTokens`: new private helper; splits style strings on unescaped `;` (respects `\;`).
- `DevkitStringBuilder` UDT with `StringBuilderInit` / `StringBuilderAppend` / `StringBuilderToString`; used in `GenerateSheetMapMarkdown` to build output in O(n) rather than O(n^2).
- `IMPORT_DIAGNOSTICS_ENABLED` constant (default `False`): when set `True`, `StartImportDiagnosticLog` / `LogImportDiagnostic` append diagnostic events to `devkit_import_diagnostics.log` (FSO TextStream, one open/close per call for crash safety) in the workbook folder.
- `AbortIfProtectedImportSheets`: pre-import guard; shows a message and aborts if any target sheet is protected.
- `ApplyShapeMapMarkdown`: new private sub; reads the `## Shapes` section of a sheet map and applies label, formula, OnAction, and style back onto the worksheet (updates existing shapes by name, or creates a rectangle placeholder for shapes not found).

### Changed
- `ApplyCellStyle`: rewritten to use `ParseStyleTokens` (escape-aware) instead of `Split(styleStr, "; ")`; adds `NumFmt`, `Unlocked`, `HAlign`, `VAlign` handling; errors during token application are now logged via `LogImportDiagnostic` instead of silently swallowed.
- `GenerateSheetMapMarkdown`: refactored to use `DevkitStringBuilder`; style-building extracted into `BuildCellStyle`; shape labels, OnAction strings, and formula strings are now escaped with `EscapeCellValue`.
- `ParseMDTableRow`: rewritten to use `TrimMDTableField` (strips exactly one leading and one trailing space) instead of `Trim` per field; preserves leading/trailing whitespace in cell values.
- `ImportAllSheetMapsFromMD`: now calls `AbortIfProtectedImportSheets` before starting; `ImportAllModulesFormsSheetMaps` also calls it in advance.

### Fixed
- Export: cells containing error values (`#DIV/0!`, `#REF!`, etc.) now export `rng.Text` as a fallback instead of raising a runtime error.
- `ExportAllSheetMapsToMD`: added `lblErr` / `lblFin` error handler that restores `Application.StatusBar` and `DisplayStatusBar` even if an error occurs mid-export.
- `ReconstructMerges`, `ApplyHiddenRows`, `ApplyHiddenCols`: added per-operation `On Error Resume Next` guards with `LogImportDiagnostic` logging so individual cell/row/column failures no longer abort the entire import.
- `ApplySheetMapMarkdown`: removed a redundant `On Error Resume Next` immediately before `On Error GoTo 0` in the pre-clear block.

## [1.5.0] - 2026-06-08

### Added
- `InitDevMode`: creates a `DEV_<name>.xlsm` copy of the current workbook and
  imports all `devkit_*` modules/forms found in `src/` into it, enabling a clean
  dev/release separation workflow.
- `SaveAsRelease`: strips all `xlsm_devkit` and `devkit_*` VBA components and saves
  a clean production copy (removes `DEV_` prefix from filename). Call from the DEV_
  workbook.
- `CallInitDevMode` / `CallSaveAsRelease`: public entry-point wrappers callable from
  the Macro dialog (`Alt+F8`) or a Quick Access Toolbar button.
- `btnSaveAsRelease` button added to `devkit_frmLauncher`; enabled only when the
  workbook name starts with `DEV_`.
- `btn_save_as_release` i18n key added to all 27 language files.
- `ImportComponentIntoProject`: private helper that imports a UTF-8 source file into
  any target VBProject (used by `InitDevMode` for cross-workbook injection).
- `IsDevkitComponent`: private helper that identifies devkit module names (used by
  `SaveAsRelease` for stripping).
- DEV/release workflow documented in README (English and Japanese).
- 11 new i18n keys (`init_*`, `release_*`) added to all 27 language files.

### Fixed
- `InitDevMode` / `SaveAsRelease`: replaced `AutomationSecurity = msoAutomationSecurityForceDisable`
  with `EnableEvents = False` when opening workbooks, preventing VBProject lockout that caused
  silent termination with no error raised.

## [1.4.3] - 2026-06-06

### Added
- `msg.sheet_map_import_error` key added to all 27 language files; displayed when
  sheet map import fails mid-way.

### Changed
- `StripAttributeLines`: rewrote from O(n²) string concatenation loop to O(n)
  array-join to avoid quadratic slowdown on large modules.
- `LongArrayToRangeStr` / `ColNumArrayToLetterRangeStr` consolidated into
  `NumArrayToRangeStr` + `FormatNumRange` to eliminate duplicated range-formatting
  logic.
- `NormalizeStartUpPosition`: parameter types widened from `Integer` to `Long`.

### Fixed
- `ImportAllSheetMapsFromMD`: `Application.ScreenUpdating`, `Calculation`, and
  `EnableEvents` are now restored in an error handler so they are never left
  disabled if the import fails mid-way.
- `xlsm_devkit`: added `IsError` guard before reading cell values to prevent
  runtime errors on cells containing error values (e.g. `#REF!`, `#VALUE!`).
- `devkit_frmInsertDelete`: added `m_syncLock` to `txtCount_Change` to prevent
  mutual re-entry with `spnCount_Change` (spinbutton ↔ textbox sync loop).
- `devkit_frmInstruction`: removed redundant `If MsgBox(...) = vbOK` wrapper in
  `btnImport_Click`; fixed parenthesized `MsgBox()` call to statement form.
- `devkit_frmLauncher`: `cboLang_Change` now calls `SetLang ""` uniformly for the
  auto-detect entry (index 0) instead of branching on `idx = 0`.
- `devkit_InsertDelete` / `devkit_Move`: renamed local variable `t` to `txt` in
  `BuildInstructionText` / `BuildMoveInstructionText` to avoid shadowing the
  global `t()` i18n function.

## [1.4.2] - 2026-05-28

### Fixed
- `devkit_frmInstruction`: renamed the `lblInstruction` label control in the form
  binary (was stored as `Label1`), fixing a runtime error introduced in v1.4.1
  where `UserForm_Initialize` sets `lblInstruction.Caption` but the control did
  not exist under that name.

## [1.4.1] - 2026-05-20

### Added
- i18n for `devkit_frmInstruction`: `lbl_instruction` and `lbl_import_note` labels now
  localized across all 27 language files (`[frmInstruction]` section).
- i18n for `devkit_frmInsertDelete`: all labels, buttons, and action strings localized
  across 27 languages using the existing `T()` / `Fmt()` API.
- i18n for `devkit_frmMoveSetup`: 10 keys (`caption`, `frame_src_ws`,
  `lbl_code_name_src`, `lbl_sheet_name_src`, `frame_dst_ws`, `lbl_code_name_dst`,
  `lbl_sheet_name_dst`, `btn_ok`, `btn_cancel`, `discard_confirm`) added across
  all 27 language files.
- `msg.optional_module_missing` key added to all 27 language files; displayed when
  a launcher button is pressed but the corresponding optional module is not loaded.

### Changed
- `GetLangCode` now caches the OS language detection result in `m_LangCodeDetected`
  to avoid repeated `GetLocaleInfoA` calls within a single session.

### Fixed
- Move: `modDeleteNote` instruction text rewritten with `VS Code:` / `VBE:` action
  prefixes so that VS Code AI correctly deletes `src\<module>.bas` rather than
  misinterpreting the VBE-deletion instruction as applying to source files.
- `SaveAsUTF8` (used by all export paths — `.bas`, `.frm`, `.md`) now normalizes
  trailing newlines: strips extra blank lines appended by Excel's VBA exporter and
  ensures each file ends with exactly one CRLF.

## [1.4.0] - 2026-05-18

### Added
- i18n support: INI-based language files under `lang/`; `T(key)` for localized
  string lookup and `Fmt(template, args…)` for `{0}`/`{1}` placeholder
  substitution. Built into `xlsm_devkit.bas` — no additional module required.
- `SetLang(code)` / `GetLangCode()` — set or query the active language at
  runtime; pass `""` to revert to auto-detection.
- `DetectSystemLang()` — resolves the Windows locale to a language code; checks
  `lang\<lang>-<region>.ini` (e.g. `zh-CN.ini`) first via
  `LOCALE_SISO3166CTRYNAME`, then `lang\<lang>.ini`, then falls back to `"en"`.
- `devkit_frmLauncher` redesigned as a full-featured launcher: exposes all
  import/export operations (all modules + forms + sheet maps; modules + forms
  only; sheet maps only) and a language-selector drop-down. InsertDelete and
  Move buttons appear automatically when those optional modules are loaded.
- 27 languages bundled in `lang/`: Arabic (`ar`), Bengali (`bn`), English
  (`en`), Spanish (`es`), Persian (`fa`), French (`fr`), German (`de`),
  Hindi (`hi`), Indonesian (`id`), Japanese (`ja`), Javanese (`jv`),
  Korean (`ko`), Malay (`ms`), Marathi (`mr`), Portuguese (`pt`),
  Punjabi (`pa`), Russian (`ru`), Swahili (`sw`), Tamil (`ta`), Telugu (`te`),
  Thai (`th`), Turkish (`tr`), Ukrainian (`uk`), Urdu (`ur`),
  Vietnamese (`vi`), Simplified Chinese (`zh-CN`), Traditional Chinese
  (`zh-TW`).

## [1.3.0] - 2026-05-17

### Added
- `ExportAllForms` — exports all UserForms (`comp.Type = 3`) to `src/*.frm` (ANSI → UTF-8); companion `*.frx` binary files are written as-is by VBE
- `ImportAllForms` — imports `src/*.frm` into the VBA project (UTF-8 → ANSI), replacing both the designer and code by removing and re-importing each form; the `*.frx` binary in the same folder is picked up automatically by VBE
- `callExportAllForms` / `callImportAllForms` — public entry-point wrappers
- `devkit_InsertDelete` — optional module: insert/delete rows or columns and generate AI instruction prompts showing before/after sheet maps
- `devkit_Move` — optional module: record a cell-range move operation via Excel macro recorder and generate AI instruction prompts
- `ExportAllModulesFormsSheetMaps` / `ImportAllModulesFormsSheetMaps` — unified entry points that export/import modules, forms, and sheet maps in one call
- `SKIP_DEVKIT_MODULES` constant in `xlsm_devkit.bas` — set `True` (default) to skip `devkit_*` modules on import; set `False` when developing the devkit optional modules themselves

### Changed
- Replaced per-file overwrite confirmation dialogs with a single start-of-operation confirmation
- Optional feature modules renamed to `devkit_` prefix (`devkit_InsertDelete`, `devkit_Move`, etc.)
- `callExportAllComponents` / `callImportAllComponents` consolidated; new unified wrappers `ExportAllModulesFormsSheetMaps` / `ImportAllModulesFormsSheetMaps` added

## [1.1.0] - 2026-05-12

### Added
- `ImportAllSheetMapsFromMD` — restores cell values, formulas, styles, named ranges, data-validation lists, and merged regions from `sheet/*.md` back into the workbook
- `callImportAllSheetMapsFromMD` — public entry-point wrapper for `ImportAllSheetMapsFromMD`

### Changed
- Replaced all legacy VBA file I/O functions (`Dir`, `Kill`, `FileCopy`, `Name`, `MkDir`) with `Scripting.FileSystemObject` throughout `ExportAllModules`, `ImportAllModules`, `ReplaceExistingComponentCodeFromFile`, and `ImportNewComponentFromFile`. This fixes silent failures on UNC paths (`\\server\share\...`) and eliminates the global-state hazard of the `Dir` loop.
- `call*` wrapper subs are now explicitly declared `Public`.

### Fixed
- `SaveAsUTF8` and `ConvertEncoding` now write BOM-less UTF-8 by switching to a binary stream and skipping the 3-byte BOM that `ADODB.Stream` prepends. Exported `src/*.bas` and `sheet/*.md` files are now recognised as plain UTF-8 (not "UTF-8 with BOM") by VS Code and other tools.
- `ImportAllModules` now checks that the `src/` folder exists before iterating its contents and shows a clear message if it is missing, preventing a runtime error from `fso.GetFolder`.

### Removed
- Private helper `FileExists` — callers now use `fso.FileExists` directly.

## [1.0.1] - 2026-05-07

### Fixed
- `ImportAllModules` now identifies itself via the module-level constant `MODULE_NAME = "xlsm_devkit"` instead of `Application.VBE.ActiveCodePane.CodeModule.Name`. The previous approach failed to protect the module from being overwritten when the macro was invoked from another module or via a shortcut key.

## [1.0.0] - 2026-05-05

### Added
- `ExportAllModules` — exports all VBA modules to `src/*.bas` (ANSI → UTF-8)
- `ImportAllModules` — imports `src/*.bas` into the VBA project (UTF-8 → ANSI), skipping `xlsm_devkit` itself
- `ExportAllSheetMapsToMD` — exports cell values, shapes, formulas, and styles to `sheet/*.md`
- VBA7 / 32-bit compatibility via `#If VBA7` and `PtrSafe`
