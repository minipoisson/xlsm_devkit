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
| `DevkitWriteResult(resultJson)` | Write `result.json` + `result.md` to a given output folder |

## File layout (in a real workbook folder)

```
<workbook folder>/
  YourTool.xlsm                      # must contain devkit_Test.bas
  tests/
    workbook.meta.json               # cell role definitions
    no_error.test.json               # the property test
  test-results/
    latest/
      result.json
      result.md
  tools/
    Invoke-XlsmDevkitTest.ps1
```

`test-results/` is generated output (gitignored). `tests/*.json` are authored specs.

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
- The workbook must contain `devkit_Test.bas`.
- Building the sample workbook requires *Trust access to the VBA project object model*
  (Excel Options > Trust Center > Macro Settings).
- The original workbook is never modified -- every run operates on a temp copy.

## Roadmap (beyond Phase 1)

`scenario` tests + `blank`/`equals`/`not_blank` assertions (Phase 2) -> richer
property-based testing (Phase 3) -> snapshot regression via the existing sheet-map
export (Phase 4) -> AI test-generation prompts (Phase 5) -> optional Python/pytest
drivers reusing the same JSON specs (Phase 6).
