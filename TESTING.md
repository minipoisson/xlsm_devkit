# Testing harness (Phase 1 MVP)

Treat an Excel workbook like testable code. The Phase 1 MVP answers one high-value
question for real-world workbooks:

> **No matter what boundary, blank, invalid, or random value is entered into any input
> cell, does the result sheet stay free of Excel errors (`#DIV/0!`, `#N/A`, ...)?**

You do not need to define the *correct* result to get value from this -- the `no_error`
property check is useful even before any expected values exist.

This is an **optional** feature: the `devkit_Test.bas` module carries the `devkit_`
prefix, so `xlsm_devkit` skips it on import/export and strips it from released
workbooks automatically (no core changes required).

## Architecture

```
AI / CLI agent
   |  authors / edits meta.json + *.test.json, reads result.md
   v
tools/Invoke-XlsmDevkitTest.ps1   (standard Windows PowerShell; no Python, no modules)
   |  temp-copy the workbook, start one Excel COM instance,
   |  generate boundary + random input cases,
   |  call the Devkit* API via Application.Run (JSON in / JSON out)
   v
devkit_Test.bas   (inside the workbook; non-interactive)
   v
test-results/latest/result.json  (machine-readable)
test-results/latest/result.md    (human / AI readable)
```

Two rules keep the harness predictable:

1. Every operation on Excel goes through a fixed public API. No ad-hoc VBA is injected.
2. PowerShell is the orchestrator only; what to write and what to check lives in the
   VBA runtime plus the JSON specs.

## Public VBA API (`devkit_Test.bas`)

All functions are non-interactive and take/return a JSON string.

| Function | Purpose |
| :--- | :--- |
| `DevkitTestPing()` | Connectivity, module version, workbook name |
| `DevkitResolveInputs(metaJson)` | Resolve `role=input` ranges; auto-detect unlocked cells on protected sheets |
| `DevkitApplyInputs(inputsJson)` | Write values into target cells (with a `blank`/`text`/`number` type hint) |
| `DevkitCalculateFullRebuild()` | `Application.CalculateFullRebuild` |
| `DevkitAssertNoErrors(targetJson)` | Scan target ranges / `used_range` for Excel error values |
| `DevkitApplyFixture(fixtureJson)` | Seed cells with values/formulas from a Markdown table before the run (non-destructive) |
| `DevkitWriteResult(resultJson)` | Write `result.json` + `result.md` to a given output folder |

## Sheet references (`sheet` vs `code_name`)

Anywhere a sheet is referenced (meta sheet definitions, `expect` targets, fixtures), you
may identify it by either:

- `"sheet": "Result"` — the tab display name, or
- `"code_name": "Sheet2"` — the VBA CodeName.

Use `code_name` when tab names can collide or change. When both are present the CodeName
wins. `DevkitResolveInputs` returns both for each input so downstream calls are unambiguous.

## File layout (in a real workbook folder)

```
<workbook folder>/
  YourTool.xlsm                      # must contain devkit_Test.bas
  tests/
    workbook.meta.json               # cell role definitions
    no_error.test.json               # the property test
    fixtures/
      seed.md                        # optional fixture sheet-map tables
  test-results/
    latest/
      result.json
      result.md
    work/                            # transient work copy of the workbook (auto-cleaned)
  tools/
    Invoke-XlsmDevkitTest.ps1
```

`test-results/` is generated output (gitignored). `tests/*` are authored specs.

## `workbook.meta.json`

Defines per-sheet roles. Input cells can be auto-detected (unlocked cells on
protected sheets) so you usually do not enumerate them by hand.

```json
{
  "version": 1,
  "input_detection": { "unlocked_cells_on_protected_sheets": true },
  "sheets": {
    "Input":  { "role": "user_input_sheet" },
    "Result": { "role": "output_sheet",
                "outputs": [ { "used_range": true, "assertions": ["no_error"] } ] }
  }
}
```

## `no_error.test.json`

```json
{
  "tests": [
    {
      "name": "No Excel error on the Result sheet for any input",
      "type": "property",
      "fixtures": [
        { "sheet": "Result", "md_file": "tests/fixtures/seed.md" }
      ],
      "inputs": [ { "role": "input" } ],
      "generate": { "strategy": "boundary_and_random", "cases": 50 },
      "expect": [ { "sheet": "Result", "used_range": true, "assert": "no_error" } ]
    }
  ]
}
```

Boundary values exercised include: blank, `0`, `-1`, a large number, a decimal, a
long string, a numeric-looking string, a date-like string, and whitespace; random
values fill the remaining cases. Detected Excel errors:
`#DIV/0! #N/A #VALUE! #REF! #NAME? #NUM! #NULL! #SPILL! #CALC!`.

## Fixtures (seed cells before the run)

A test may declare `fixtures` that set specific cells **once, before the case loop** —
useful for static context a formula depends on (lookup tables, mode flags, parameters).
Each fixture references a sheet (`sheet` or `code_name`) and supplies a Markdown
sheet-map table, either inline (`md`) or from a file (`md_file`, relative to the workbook
folder). Only the listed cells are written — it is **non-destructive** (no clearing of
surrounding cells, formats, or validation).

```markdown
| Address | Name | Value | Formula | Style |
|---------|------|-------|---------|-------|
| H1      |      | 5     |         |       |
| H2      |      |       | =H1*2   |       |
```

A formula (column 4) takes priority over a literal value (column 3), matching the core
sheet-map import. The Style column is ignored by fixtures (value/formula only). You can
paste a slice of an existing `sheet/*.md` straight into a fixture file.

## Running

```powershell
# 1. (dev only) build a sample workbook with protected inputs + error-prone formulas
.\tools\New-SampleTestWorkbook.ps1

# 2. run the harness
.\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm
```

Useful switches: `-Cases <n>` overrides the case count, `-Seed <n>` makes generation
reproducible, `-MetaPath` / `-TestPath` point at non-default spec locations.

The runner exits `1` when any case fails (so it slots into CI), and always writes
`result.md` with a failure table (case, sheet, address, error, formula, inputs).

### Requirements

- Excel must be installed (the runner drives it through COM).
- The workbook must contain `devkit_Test.bas`. Fixtures additionally require the core
  `xlsm_devkit.bas` module in the same workbook (it provides the reused `ParseMDTableRow`
  / `UnescapeCellValue` helpers) -- normal dev workbooks already have it.
- Building the sample workbook requires *Trust access to the VBA project object model*
  (Excel Options > Trust Center > Macro Settings).
- The original workbook is never modified -- every run operates on a **work copy** placed
  under `test-results/work/` next to the workbook. This keeps the copy inside the
  workbook's Excel **trusted location** so macros run; the system temp folder is usually
  not trusted. Ensure the trusted location is set to **include subfolders**.

## Roadmap (Phases 1-6)

The harness is built in phases; each phase is independently useful. The JSON specs and
the VBA public API are stable across phases, so later phases extend rather than replace.

- **Phase 1 (MVP) -- implemented.** Cell role definitions (`workbook.meta.json`) plus the
  `no_error` property check. This release also adds `code_name` sheet references and
  Markdown fixtures.
- **Phase 2 -- planned.** Fixed-scenario tests (`scenario` type) and the assertions
  `blank` / `not_blank` / `equals` / `not_contains` (e.g. "for male patients, the
  obstetrics result cell must be blank").
- **Phase 3 -- planned.** Full property-based testing: `within_range`, `matches_regex`,
  `all_blank_or_hidden`, plus failing-case shrinking.
- **Phase 4 -- planned.** Snapshot regression using the existing sheet-map export
  (`CallExportAllSheetMapsToMD`): `same_as_snapshot`, `changed_only_in_allowed_ranges`,
  combined with Git diff to judge whether a change moved expected values.
- **Phase 5 -- planned.** AI test-generation: emit a prompt/runbook (like InsertDelete /
  Move) that hands `sheet/*.md` + meta to an AI to propose `*.test.yaml`/`.json` cases.
- **Phase 6 -- planned.** Python / pytest interop. The specs are language-neutral, so a
  pytest adapter only needs to call the same public API via `Application.Run` (pywin32);
  PowerShell remains the primary, zero-install driver.
