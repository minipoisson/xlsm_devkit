# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.0.1] - 2026-05-07

### Fixed
- `ImportAllModules` now identifies itself via the module-level constant `MODULE_NAME = "xlsm_devkit"` instead of `Application.VBE.ActiveCodePane.CodeModule.Name`. The previous approach failed to protect the module from being overwritten when the macro was invoked from another module or via a shortcut key.

## [1.0.0] - 2026-05-05

### Added
- `ExportAllModules` — exports all VBA modules to `src/*.bas` (ANSI → UTF-8)
- `ImportAllModules` — imports `src/*.bas` into the VBA project (UTF-8 → ANSI), skipping `xlsm_devkit` itself
- `ExportAllSheetMapsToMD` — exports cell values, shapes, formulas, and styles to `sheet/*.md`
- VBA7 / 32-bit compatibility via `#If VBA7` and `PtrSafe`
