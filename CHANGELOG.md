# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

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
