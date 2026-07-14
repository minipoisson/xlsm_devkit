<#
.SYNOPSIS
    Build the sample test workbook used to develop / verify the xlsm_devkit test harness.

.DESCRIPTION
    Constructs examples\test-harness\sample.xlsm from scratch via Excel COM:
      - Sheet "Input"  : protected, with B2 (quantity) and B3 (unit price) UNLOCKED.
      - Sheet "Result" : formulas that reference Input, including two that can error:
            B2 = Input!B3 / Input!B2          -> #DIV/0! when quantity is 0 / blank
            B3 = VLOOKUP(Input!B2, table, ..) -> #N/A   when quantity is not 1/2/3
      - Imports src\xlsm_devkit.bas, src\devkit_Test.bas and src\devkit_Regex.bas into the
        workbook's VBProject so the runner can call the Devkit* API via Application.Run.
        (devkit_Test reuses Public helpers from xlsm_devkit.bas; matches_regex uses
        devkit_Regex, so all three modules must be present.)
      - Extra Phase 3 outputs: B8 = a "Qnnn" code string (matches_regex target), the
        Total B4 (within_range target), a naturally-blank block J1:J3, and the hidden
        lookup columns E:F (all_blank_or_hidden hidden-branch demo). A spare unlocked
        Input!B4 gives shrinking an input it can neutralise.

    ASCII-only sheet names (the role mechanism is language-independent), which keeps
    PowerShell free of encoding pitfalls.

    Requires "Trust access to the VBA project object model" to be enabled
    (Excel Options > Trust Center > Macro Settings) so the module can be imported.

    The produced .xlsm is gitignored (a binary build artifact).

.EXAMPLE
    .\tools\New-SampleTestWorkbook.ps1
#>
[CmdletBinding()]
param(
    [string]$OutPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
# devkit_Test reuses Public helpers from xlsm_devkit.bas (ParseMDTableRow,
# UnescapeCellValue), so the core module must be present in the workbook too.
$coreSrc = Join-Path $repoRoot 'src\xlsm_devkit.bas'
$moduleSrc = Join-Path $repoRoot 'src\devkit_Test.bas'
$regexSrc = Join-Path $repoRoot 'src\devkit_Regex.bas'
if (-not (Test-Path -LiteralPath $coreSrc)) { throw "xlsm_devkit.bas not found: $coreSrc" }
if (-not (Test-Path -LiteralPath $moduleSrc)) { throw "devkit_Test.bas not found: $moduleSrc" }
if (-not (Test-Path -LiteralPath $regexSrc)) { throw "devkit_Regex.bas not found: $regexSrc" }

if ([string]::IsNullOrWhiteSpace($OutPath)) {
    $OutPath = Join-Path $repoRoot 'examples\test-harness\sample.xlsm'
}
$outDir = Split-Path -Parent $OutPath
$null = New-Item -ItemType Directory -Force -Path $outDir

function Remove-ComObject($obj) {
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

$xl = $null
$wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false

    $wb = $xl.Workbooks.Add()
    # Reduce to a single sheet, then build the two we need.
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

    $wsIn = $wb.Worksheets.Item(1)
    $wsIn.Name = 'Input'
    $wsRes = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wsIn)
    $wsRes.Name = 'Result'

    # --- Input sheet ---------------------------------------------------------
    $wsIn.Range('A1').Value2 = 'Field'
    $wsIn.Range('B1').Value2 = 'Value'
    $wsIn.Range('A2').Value2 = 'Quantity'
    $wsIn.Range('A3').Value2 = 'Unit Price'
    $wsIn.Range('A4').Value2 = 'Spare (unused input)'
    $wsIn.Range('B2').Value2 = 1
    $wsIn.Range('B3').Value2 = 100
    # Unlock the input cells (including a spare B4 that no result formula reads, so
    # shrinking has an input it can neutralise), then protect the sheet.
    $wsIn.Range('B2:B4').Locked = $false
    $wsIn.Protect()

    # --- Result sheet --------------------------------------------------------
    $wsRes.Range('A1').Value2 = 'Metric'
    $wsRes.Range('B1').Value2 = 'Value'
    $wsRes.Range('A2').Value2 = 'Price per unit (div)'
    $wsRes.Range('A3').Value2 = 'Tier lookup'
    $wsRes.Range('A4').Value2 = 'Total'
    $wsRes.Range('A6').Value2 = 'Fixture-dependent (1/H1)'
    $wsRes.Range('B2').Formula = '=Input!B3/Input!B2'
    $wsRes.Range('B3').Formula = '=VLOOKUP(Input!B2,$E$2:$F$4,2,FALSE)'
    $wsRes.Range('B4').Formula = '=Input!B2*Input!B3'
    # B6 errors (#DIV/0!) unless a fixture seeds H1 before the run -- demonstrates DevkitApplyFixture.
    $wsRes.Range('B6').Formula = '=1/Result!H1'
    # Small lookup table for the VLOOKUP (tiers 1/2/3 only -> other quantities give #N/A).
    $wsRes.Range('E2').Value2 = 1; $wsRes.Range('F2').Value2 = 10
    $wsRes.Range('E3').Value2 = 2; $wsRes.Range('F3').Value2 = 20
    $wsRes.Range('E4').Value2 = 3; $wsRes.Range('F4').Value2 = 30

    # --- Phase 3 assertion targets ------------------------------------------
    # matches_regex: a "Qnnn" code string derived from the quantity.
    $wsRes.Range('A8').Value2 = 'Code'
    $wsRes.Range('B8').Formula = '="Q"&TEXT(Input!B2,"000")'
    # all_blank_or_hidden: hide the lookup columns E:F so a non-blank but hidden range
    # still passes; J1:J3 is left naturally blank for the blank branch.
    $wsRes.Columns('E:F').Hidden = $true

    # --- Import the core + test + regex modules into the VBProject -----------
    # Core first: devkit_Test calls its Public ParseMDTableRow / UnescapeCellValue.
    # devkit_Regex provides the matches_regex engine used by devkit_Test.
    try {
        $null = $wb.VBProject.VBComponents.Import($coreSrc)
        $null = $wb.VBProject.VBComponents.Import($regexSrc)
        $null = $wb.VBProject.VBComponents.Import($moduleSrc)
    } catch {
        throw "Could not import modules into the VBProject. Enable " +
              "'Trust access to the VBA project object model' in Excel Trust Center. " +
              "Underlying error: $($_.Exception.Message)"
    }

    # --- Save as macro-enabled workbook (52 = xlOpenXMLWorkbookMacroEnabled) --
    if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }
    $wb.SaveAs($OutPath, 52)
    Write-Host "Sample workbook written: $OutPath"
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    if ($null -ne $xl) { try { $xl.Quit() } catch { } }
    Remove-ComObject $wb
    Remove-ComObject $xl
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
