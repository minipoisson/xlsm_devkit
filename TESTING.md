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
| `DevkitAssertExpectations(json)` | Check fixed-scenario expectations per cell (`blank` / `not_blank` / `equals` / `not_contains`) |
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

## Fixed-scenario tests (`scenario` type)

Alongside the generative `property` / `no_error` check, a test may be a **fixed
scenario**: apply a known set of inputs once, recalculate, then assert specific cells
against expected outcomes. Use it for concrete business rules ("for a male patient the
obstetrics result cell must be blank"; "quantity=2, price=50 -> total=100").

Set `"type": "scenario"`. Instead of `generate`, list explicit `inputs` (each a cell
reference plus `value` and optional `type` hint, same as `DevkitApplyInputs`), and give
each `expect` entry an `assert` and, where relevant, a `value`:

```json
{
  "tests": [
    {
      "name": "Quantity=2, price=50 yields expected metrics",
      "type": "scenario",
      "fixtures": [ { "sheet": "Result", "md_file": "tests/fixtures/seed.md" } ],
      "inputs": [
        { "sheet": "Input", "address": "B2", "value": "2",  "type": "number" },
        { "sheet": "Input", "address": "B3", "value": "50", "type": "number" }
      ],
      "expect": [
        { "sheet": "Result", "address": "B2", "assert": "equals",       "value": "25" },
        { "sheet": "Result", "address": "B4", "assert": "equals",       "value": "100" },
        { "sheet": "Result", "address": "B3", "assert": "not_blank" },
        { "sheet": "Result", "address": "B6", "assert": "not_contains", "value": "#" }
      ]
    }
  ]
}
```

Assertions:

| `assert` | Extra fields | Passes when |
| :--- | :--- | :--- |
| `blank` | -- | the cell has no content (empty, or a formula returning `""`) |
| `not_blank` | -- | the cell has any content (an error value counts as content) |
| `equals` | `value` | the cell's **displayed text** equals `value` (exact, case-sensitive) |
| `not_contains` | `value` | the cell's **displayed text** does not contain `value` (substring) |
| `within_range` | `min` / `max` | the cell's **displayed text**, read as a number, is within `[min, max]` (**inclusive**) |
| `matches_regex` | `pattern`, `ignore_case`, `multiline` | the cell's **displayed text** matches the regex `pattern` |
| `all_blank_or_hidden` | -- | every cell in the (possibly multi-cell) target is blank or sits in a hidden row/column |

The Phase 3 asserts (`within_range`, `matches_regex`, `all_blank_or_hidden`) work in both
`scenario` and `property` tests. Notes:

- **`within_range`** takes the number from the **displayed** text (`Range.Text`), not the
  raw cell value. A cell showing `25.0` (a rounded display of `24.99`) therefore counts as
  `>= 25` -- what the user sees. At least one of `min` / `max` is required; a missing bound
  is unbounded on that side. Non-numeric displays fail.
- **`matches_regex`** uses the vendored, pure-VBA **`devkit_Regex`** engine
  ([sihlfall/vba-regex](https://github.com/sihlfall/vba-regex), MIT), which implements a
  **JavaScript-flavoured** dialect (not .NET / not VBScript). `ignore_case` and `multiline`
  are optional booleans; the match succeeds when the pattern is found anywhere in the text
  (anchor with `^`/`$` for a full match). The target workbook must contain
  `src/devkit_Regex.bas` (see Requirements).
- **`all_blank_or_hidden`** accepts a multi-cell `address` (e.g. `A1:D50`); it fails on the
  first cell that is both visible and non-blank.

`equals` and `not_contains` compare against the cell's *displayed* text
(`Range.Text`) — what a user sees after number formatting — with no numeric tolerance.
Write `value` to match the formatted display (e.g. a currency cell showing `$25.00`).
A scenario counts as one "case"; a failure lists the assertion mismatch
(`expected "…" got "…"`) in the result's failure table.

## Running

```powershell
# 1. (dev only) build a sample workbook with protected inputs + error-prone formulas
.\tools\New-SampleTestWorkbook.ps1

# 2. run the harness (default: no_error property check)
.\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm

# run a fixed-scenario spec instead (-TestPath is resolved from the current directory)
.\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm -TestPath .\examples\test-harness\tests\scenario.test.json
```

Useful switches: `-Cases <n>` overrides the case count, `-Seed <n>` makes generation
reproducible, `-MetaPath` / `-TestPath` point at non-default spec locations,
`-MaxShrinkSteps <n>` caps shrink re-runs (default 200), `-NoShrink` disables shrinking.

The runner exits `1` when any case fails (so it slots into CI), and always writes
`result.md` with a failure table (case, sheet, address, error, formula, inputs).

### Failing-case shrinking (property tests)

When a generated `property` case fails, the runner **minimises** the counterexample before
recording it: greedily, one input at a time, it tries neutralising each input to blank and
then `0`, keeping the change only while the **same** failing cell + assertion still
reproduces. Iteration continues to a fixpoint, so `result.md` reports the smallest inputs
that still break the property (marked `(shrunk)`), making the cause easier to see. Shrink
re-runs share one budget across the whole run (`-MaxShrinkSteps`, default 200) so a run
with many failures cannot explode Excel COM cost; `-NoShrink` turns it off. Shrinking
applies to `property` tests only -- `scenario` inputs are fixed by definition.

### Requirements

- Excel must be installed (the runner drives it through COM).
- The workbook must contain `devkit_Test.bas`. Fixtures additionally require the core
  `xlsm_devkit.bas` module in the same workbook (it provides the reused `ParseMDTableRow`
  / `UnescapeCellValue` helpers) -- normal dev workbooks already have it. The
  `matches_regex` assertion additionally requires `src/devkit_Regex.bas` (the vendored
  pure-VBA engine); import it alongside `devkit_Test.bas` if you use `matches_regex`.
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
- **Phase 2 -- implemented.** Fixed-scenario tests (`scenario` type) and the assertions
  `blank` / `not_blank` / `equals` / `not_contains` (e.g. "for male patients, the
  obstetrics result cell must be blank"). See "Fixed-scenario tests" above.
- **Phase 3 -- implemented.** Full property-based testing: the assertions `within_range`,
  `matches_regex` (JS-flavoured, via the vendored MIT `devkit_Regex` engine) and
  `all_blank_or_hidden`, usable in both `scenario` and `property` tests, plus failing-case
  shrinking (`-MaxShrinkSteps` / `-NoShrink`). See the assertions table and "Failing-case
  shrinking" above.
- **Phase 4 -- planned.** Snapshot regression using the existing sheet-map export
  (`CallExportAllSheetMapsToMD`): `same_as_snapshot`, `changed_only_in_allowed_ranges`,
  combined with Git diff to judge whether a change moved expected values.
- **Phase 5 -- planned.** AI test-generation: emit a prompt/runbook (like InsertDelete /
  Move) that hands `sheet/*.md` + meta to an AI to propose `*.test.yaml`/`.json` cases.
- **Phase 6 -- planned.** Python / pytest interop. The specs are language-neutral, so a
  pytest adapter only needs to call the same public API via `Application.Run` (pywin32);
  PowerShell remains the primary, zero-install driver.
