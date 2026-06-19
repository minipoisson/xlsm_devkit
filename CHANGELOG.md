# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.13.0] - 2026-06-19

### Added
- Testing harness (Phase 1 MVP): new optional module `src/devkit_Test.bas` plus a
  PowerShell runner that treats a workbook like testable code. It generates boundary
  and random inputs into a temp copy and asserts that the result sheet stays free of
  Excel errors (`no_error` property check). All VBA API is non-interactive and uses
  JSON in/out (`DevkitTestPing`, `DevkitResolveInputs`, `DevkitApplyInputs`,
  `DevkitCalculateFullRebuild`, `DevkitAssertNoErrors`, `DevkitWriteResult`).
  Standard Windows PowerShell only -- no Python or external modules.
  - `tools/Invoke-XlsmDevkitTest.ps1` -- orchestrator (work copy, COM, generation, cleanup)
  - `tools/New-SampleTestWorkbook.ps1` -- builds a sample workbook for verification
  - `examples/test-harness/tests/*.json` + `tests/fixtures/*.md` -- meta/role, `no_error`
    test, and fixture templates
  - `TESTING.md` -- harness reference; README (EN/JA) sections added
  - The module carries the `devkit_` prefix, so it is auto-skipped on import/export and
    stripped from released workbooks by existing machinery.
- Testing harness: sheet references now accept `code_name` (VBA CodeName) in addition to
  `sheet` (tab display name), disambiguating sheets whose tab names collide. The CodeName
  takes priority when both are present; `DevkitResolveInputs` echoes both.
- Testing harness: test fixtures (`DevkitApplyFixture`). A test spec may seed cells with
  values/formulas before the run via a `fixtures` block referencing a Markdown sheet-map
  table (`md_file` or inline `md`). Application is non-destructive (only listed cells are
  written; no ClearContents/ClearFormats). Reuses the core Markdown helpers.
- Testing harness: the work copy is placed under the workbook folder (`test-results/work`)
  instead of the system temp folder, so it inherits the workbook's Excel trusted location
  (the system temp folder is usually not trusted, which blocks macros). The runner also
  lowers `Application.AutomationSecurity` for the explicit local test copy.

### Changed
- `xlsm_devkit.bas`: `ParseMDTableRow` and `UnescapeCellValue` are now `Public` so optional
  modules (e.g. `devkit_Test`) can reuse Markdown table parsing / cell-value unescaping.
  No behavioural change to existing import/export.

## [1.12.0] - 2026-06-14

### Fixed
- `ApplyCellStyle`: boolean-attribute reset (`Bold`, `Italic`, `Strike`, `Wrap`,
  `Unlocked`) now runs for cells with empty style strings when `merge=0`. Previously
  the function exited before the reset block for empty-style cells, leaving existing
  Excel formatting unchanged when a token was removed from the Markdown. The reset
  now uses read-before-write: each attribute is read first and the write is skipped
  when the value already equals the desired default, minimising unnecessary COM writes.

## [1.11.0] - 2026-06-14

### Fixed
- `ApplyCellStyle`: when a boolean-attribute import flag is enabled (`bold=1`, `italic=1`,
  `strike=1`, `wrap=1`, `unlocked=1`), the corresponding attribute is now reset to its
  default value before token processing. Previously, if the token was absent from a cell's
  style string, the existing Excel value was left unchanged, causing drift when `merge=0`
  (since `ClearFormats` is skipped in that mode). Absence of a token is now treated as
  "restore default" for all enabled flags, consistent with export semantics (non-defaults
  only are emitted). For `merge=1` the reset is redundant but harmless -- `ClearFormats`
  already performs it.
- `xlsm_devkit.ini` / README: corrected the `unlocked=0` comment which incorrectly stated
  "enable only together with merge=1". The `rng.MergeArea.Locked` fix and lazy-unmerge
  introduced in v1.10.0 make `unlocked=1` + `merge=0` fully safe. The default remains `0`
  as an opt-in safeguard against accidentally weakening sheet protection.

## [1.10.0] - 2026-06-14

### Added
- Selective import write-back: place `xlsm_devkit.ini` in the workbook folder with an `[import]` section to control which style categories are applied on import. Each key (`value`, `formula`, `numfmt`, `unlocked`, `list`, `bg`, `fg`, `font_size`, `bold`, `italic`, `strike`, `wrap`, `halign`, `valign`, `merge`, `hidden_rows`, `hidden_cols`, `shapes`) accepts `0` or `1`. Missing keys and a missing file both default to `1` for backward compatibility, except `unlocked` which defaults to `0` (see **Changed** below). Setting `merge=0` skips `ws.Cells.UnMerge` entirely, eliminating the 8--15 minute bottleneck on sheets with thousands of merged regions.
- `DisabledImportSettingsStr`: new private helper; returns a space-separated list of `ImportSettings` flags currently set to `False` (e.g. `"bg fg font_size ..."`), or `"(all enabled)"` when all flags are `True`. Called at import start and logged via `LogImportDiagnostic` so diagnostic logs show which categories are inactive.

### Performance
- `ExportAllSheetMapsToMD` now sets `ScreenUpdating = False`, `Calculation = xlManual`, and `EnableEvents = False` during the export loop (same as import). This prevents screen repainting between cell reads, making export significantly faster for large sheets and ensuring the status bar progress message is visible.
- `ImportAllSheetMapsFromMD` skips sheets whose Markdown file has not changed since the last export or import. After each export/import the file size and modification time of every `*.md` file are written to `sheet/xlsm_devkit_sync.tsv`. On the next import run each file is compared against this cache; unchanged files are skipped entirely, reducing a full re-import of unmodified sheets to a near-instant check.

### Changed
- `ImportUnlocked` now defaults to `False` (previously `True` via the all-enabled backward-compatibility default). Setting `rng.Locked = False` on slave cells of still-merged ranges fails with Error 1004 when `merge=0` (UnMerge skipped). To apply `Unlocked` during import, explicitly set `unlocked=1` in `xlsm_devkit.ini`; this works correctly only when `merge=1` is also set so cells are individually addressable after UnMerge.

### Fixed
- `ExportAllSheetMapsToMD` and `ImportAllSheetMapsFromMD` now call `DoEvents` inside the per-sheet loop so that `Application.StatusBar` progress messages are repainted in real time. Previously the status bar was updated only after the loop completed, making progress invisible on large workbooks.
- `ApplySheetMapMarkdown` no longer raises `RPC_E_DISCONNECTED` on sheets with large merged regions after `ws.Cells.UnMerge` completes. A `DoEvents` call immediately after `UnMerge` flushes the COM message queue; an additional `DoEvents` every 1000 rows in Pass 1 prevents queue starvation on very large sheets.
- `ApplyShapeFields` no longer logs Err 438 for Picture-type shapes. Picture objects do not have a text frame; the label assignment is now guarded by `shp.HasTextFrame` and silently skipped instead of attempting the write and logging a spurious error.
- `ApplyCellStyle` no longer emits `WARN unknown style token` for recognized tokens that are disabled in `xlsm_devkit.ini` (e.g. `bg=0` suppresses `BG:#rrggbb` tokens that were incorrectly falling through to the WARN branch). Token name recognition is now separate from the cfg-flag check; known tokens with disabled flags are silently skipped. Only genuinely unknown tokens still produce a WARN.
- `ApplyCellStyle`: `BOLD`, `ITALIC`, `STRIKE`, `WRAP`, and `UNLOCKED` token application is now guarded with `On Error Resume Next` / `On Error GoTo 0` and `LogImportDiagnostic` error logging, consistent with `BG`, `FG`, `FONTSIZE`, `NUMFMT`, `HALIGN`, and `VALIGN`. Previously, a runtime error in any of these five branches (e.g. Err 1004 from `rng.Locked = False` on a slave cell of a merged range) would abort the entire import.
- `DisabledImportSettingsStr`: fixed a compile error -- `ReDim Preserve` on a fixed-size array declared as `Dim parts(17) As String` is not permitted; changed to a dynamic array with an initial `ReDim`.
- `ApplyCellStyle` -- `UNLOCKED` token: changed `rng.Locked = False` to
  `rng.MergeArea.Locked = False`. Excel raises Err 1004 when setting `.Locked` on any single-cell
  reference that belongs to a merged area (master or slave); using `MergeArea` addresses the
  entire merged region and succeeds regardless of whether `rng` is the master, a slave, or an
  individual cell.
- `ApplyCellStyle`: when `rng` is a slave cell of a currently-merged region (can occur when
  `merge=0` leaves cells merged but the sheet map treats those cells as individual), the merge
  area is temporarily unmerged, all style tokens are applied, then the area is re-merged.
  This is a defensive preamble that protects all tokens; the `UNLOCKED` fix above makes it
  specifically unnecessary for the Locked property, but it remains for other potential failures.

## [1.9.0] - 2026-06-11

### Performance
- `ApplySheetMapMarkdown` no longer freezes on sheets with thousands of merge slave cells (e.g. Sheet4 with 7357 lines and ~6000 slaves). Two additional COM-call bottlenecks eliminated:
  1. Pass 1 slave handling called `ws.Range(addr)` + `ws.Cells(...).Address(...)` per slave cell (~3 COM calls × ~6000 slaves = ~18 000 calls). Now uses `AddrRowNum`/`AddrColNum`/`CellAddrStr` pure string parsing -- zero COM calls per slave unless a named range must be applied.
  2. `ReconstructMerges` called `ws.Range` on every slave entry to read `.Row`/`.Column` (~18 000 COM calls for ~6000 slaves). Now uses the same string parsers and makes only one `ws.Range(...).Merge` call per distinct merge group.
- New private helper: `CellAddrStr(r, c)` -- inverse of `AddrRowNum`/`AddrColNum`; converts row/column numbers back to an A1-style address string without touching the worksheet object model.

## [1.8.0] - 2026-06-11

### Added
- Import progress: `Application.StatusBar` shows "Importing sheet map N/N: `<sheet>`" during `ImportAllSheetMapsFromMD`; status bar is saved and restored even if an error occurs mid-import.

### Fixed
- `ApplySheetMapMarkdown` no longer freezes or disconnects COM on sheets with large merged regions. Two root causes were eliminated:
  1. The per-cell cleanup loop built an `Application.Union` object incrementally, causing O(n^2) slowdown and eventual COM disconnect (`RPC_E_DISCONNECTED`) on sheets with many cells outside the Markdown.
  2. `ApplyCellStyle` was called for every merge slave cell individually (~350+ calls on large sheets), even though slave styles are irrelevant after `ReconstructMerges` (master style propagates to the whole merged region).
- Cleanup is now 3 batch operations on the Markdown bounding box (`ClearContents` / `ClearFormats` / `Validation.Delete`). Merge slaves are recorded for Pass 2 only, with no per-cell style application.
- Pass 0 bounding box calculation no longer makes COM calls (`ws.Range(addr)` per address); replaced with pure string parsing helpers `AddrRowNum` / `AddrColNum`.

## [1.7.0] - 2026-06-11

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
