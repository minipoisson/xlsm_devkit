# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

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
