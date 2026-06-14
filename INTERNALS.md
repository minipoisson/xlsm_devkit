# INTERNALS.md

Architecture and implementation notes for contributors and future maintainers.
For user-facing documentation see README / README.ja. For contribution workflow see CONTRIBUTING.md (TBD).

---

## Module Layout

```
src/xlsm_devkit.bas        # Core module -- the release asset
src/devkit_*.bas/.frm/.frx # Optional modules (UI forms, helpers)
lang/                      # i18n INI files (27 languages)
sheet/                     # Sheet map *.md + sync.tsv (git-ignored)
xlsm_devkit.ini            # ImportSettings sample
```

**Git-ignored source files** (workbook-local, never committed):
`src/Module*.bas`, `src/Sheet*.bas`, `src/ThisWorkbook.bas`, `sheet/`

**Binary files**: `src/*.frx` must never be edited by hand.
If an unintended Excel operation mutates them, restore with `git restore src/*.frx`.

---

## ImportSettings / xlsm_devkit.ini

Place `xlsm_devkit.ini` in the same folder as the workbook to override import defaults.

```ini
[import]
; code defaults: unlocked=0 is the only exception, all others default to 1
value=1  formula=1  numfmt=1  unlocked=0  list=1
bg=0  fg=0  font_size=0  bold=0  italic=0  strike=0  wrap=0
halign=0  valign=0
merge=0      ; skip UnMerge + ClearFormats + ReconstructMerges (biggest speedup)
hidden_rows=0  hidden_cols=0  shapes=0
```

`LoadImportSettings(wbFolder)` is called once at the top of `ImportAllSheetMapsFromMD`.

**`merge=0` implication**: `ClearFormats` is also skipped -- clearing formats would reset
the merge structure that `UnMerge` was skipped to preserve. These two are always coupled.

**`unlocked=0` default rationale**: Defaulting to opt-in prevents accidentally weakening
sheet protection. `unlocked=1` + `merge=0` is fully safe since v1.10.0 (`rng.MergeArea.Locked`
fix and lazy-unmerge).

---

## `ApplyCellStyle` Processing Order (as of v1.12.0)

```
ApplyCellStyle(rng, styleStr, cfg)
```

1. **Variable declarations + ParseStyleTokens guard**
   If `styleStr` is `""` or `"-"`, skip `ParseStyleTokens` (no tokens to parse).

2. **Lazy-unmerge preamble**
   If `rng` is a merge slave cell, call `rng.MergeArea.UnMerge` and record
   `savedMergeArea` for re-merge at the epilogue.

3. **Boolean reset (read-before-write)**
   For each boolean attribute (`Bold`, `Italic`, `Strike`, `Wrap`, `Unlocked`),
   if the corresponding `cfg.Import*` flag is `True`, read the current Excel value
   and write the default only when it differs.
   This block runs for **every cell including empty-style cells** to prevent drift
   when `merge=0` (where `ClearFormats` is skipped).

4. **Empty-style early exit**
   If `styleStr = ""` or `"-"`, `GoTo lblRemerge` (skip token loop).

5. **Token loop**
   Apply each style token: `Bold=True`, `MergeArea.Locked=False`, color values, etc.

6. **`lblRemerge:` re-merge epilogue**
   If step 2 unmerged the cell, call `savedMergeArea.Merge`.

**Why steps 3 and 4 are in this order**: The reset must run even when there are no tokens
(empty style = "all tokens removed"). Placing the guard before the reset would skip the
reset for cells whose style string was cleared.

---

## Import Pipeline: `ApplySheetMapMarkdown`

### Pass 0 -- Bounding box calculation (zero COM calls)

Pure-string parsing using `AddrRowNum` / `AddrColNum` / `CellAddrStr` helpers.
Computes the bounding box of all addressed cells without touching the Excel object model.

### Pre-clear

| `merge` setting | Actions |
|---|---|
| `merge=1` | `ws.Cells.UnMerge` + `DoEvents` → bounding-box `ClearContents` + `ClearFormats` + `Validation.Delete` |
| `merge=0` | `ClearContents` only (no `ClearFormats`, no `UnMerge`) |

### Pass 1 -- Value / formula / style application

Iterates parsed rows and calls `ApplyCellStyle` per cell.
`DoEvents` is called every 1000 rows to keep the status bar responsive.

Slave cells (Value column = `!merged_left`, `!merged_up`, `!merged_ul`) are excluded
in Pass 1 via `cValue` check and never reach `ApplyCellStyle`. This behaviour is
unchanged since v1.8.0.

### Pass 2 -- ReconstructMerges (`merge=1` only)

Rebuilds merge regions from the slave-cell markers recorded during Pass 1.
Skipped entirely when `merge=0`.

---

## Sheet Map Markdown Format (brief)

Full specification: `## Sheet Map Format` section in README / README.ja (added in commit `254ae69`).

Key points:
- File location: `sheet/<codeName>.md`
- Table header: `| Address | Name | Value / Label | Formula | Style |`
- Style tokens: `BG:#rrggbb`, `FG:#rrggbb`, `FontSize:<n>`, `Bold`, `Italic`, `Strike`,
  `Wrap`, `Unlocked`, `NumFmt:<fmt>`, `List:<fml>`, `HAlign:<v>` (import-only),
  `VAlign:<v>` (import-only)
- Slave markers (Value column): `!merged_left`, `!merged_up`, `!merged_ul`

---

## Diagnostics

`IMPORT_DIAGNOSTICS_ENABLED` constant at line 32 of `xlsm_devkit.bas`.

- Production value: `False`
- Enable only in a separate test project; never commit `True` to this repository.

---

## Known Failed Approaches

These are recorded to prevent revisiting the same dead ends.

| Approach | Why it fails | Correct alternative |
|---|---|---|
| `rng.Locked = False` on a merged cell via individual reference | Raises Err 1004 -- Excel does not allow per-cell Locked on a member of a merged region | `rng.MergeArea.Locked = False` |
| `ImportUnlocked` defaulting to `True` | With `merge=0`, `MergeArea.Locked` is called on every cell, causing Err 1004 bursts on sheets with merges | Default to `False` (opt-in) |
| Token-name check AND `cfg` flag in the same `And` condition inside `ApplyCellStyle` | When `cfg.ImportBold = False`, a `Bold` token falls through to the `Else` branch and is logged as WARN instead of being silently skipped | Separate the cfg-flag guard from the token dispatcher |
| `ClearFormats` when `merge=0` | Resets the merge structure that `UnMerge` was skipped to preserve -- the two operations are contradictory | When `merge=0`, skip both `UnMerge` and `ClearFormats` |
| `Application.Union` accumulated in a loop for merge reconstruction | O(n^2) COM calls; triggers `RPC_E_DISCONNECTED` on large sheets | Use a 3-call bounding-box approach instead |
| `ws.Range(addr)` in Pass 0 / Pass 1 / `ReconstructMerges` | Each call is a COM round-trip; scales poorly | Pure-string address helpers (`AddrRowNum`, `AddrColNum`, `CellAddrStr`) -- zero COM in Pass 0 |
| Legacy VBA file functions (`Dir`, `Kill`, `Open`, `FreeFile`, `MkDir`, etc.) | Unreliable on UNC paths (`\\server\share\...`); `Dir` also holds module-level global state that corrupts nested loops | `Scripting.FileSystemObject` (late-bound via `CreateObject`) for all file I/O |
