<#
.SYNOPSIS
    Build the sample test workbook used to develop / verify the xlsm_devkit test harness.

.DESCRIPTION
    Constructs examples\test-harness\sample.xlsm from scratch via Excel COM:
      - Sheet "Input"  : protected, with B2 (quantity) and B3 (unit price) UNLOCKED.
      - Sheet "Result" : formulas that reference Input, including two that can error:
            B2 = Input!B3 / Input!B2          -> #DIV/0! when quantity is 0 / blank
            B3 = VLOOKUP(Input!B2, table, ..) -> #N/A   when quantity is not 1/2/3
      - Imports src\devkit_Test.bas into the workbook's VBProject so the runner can
        call the Devkit* API via Application.Run.

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
$moduleSrc = Join-Path $repoRoot 'src\devkit_Test.bas'
if (-not (Test-Path -LiteralPath $moduleSrc)) { throw "devkit_Test.bas not found: $moduleSrc" }

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
    $wsIn.Range('B2').Value2 = 1
    $wsIn.Range('B3').Value2 = 100
    # Unlock the two input cells, then protect the sheet.
    $wsIn.Range('B2:B3').Locked = $false
    $wsIn.Protect()

    # --- Result sheet --------------------------------------------------------
    $wsRes.Range('A1').Value2 = 'Metric'
    $wsRes.Range('B1').Value2 = 'Value'
    $wsRes.Range('A2').Value2 = 'Price per unit (div)'
    $wsRes.Range('A3').Value2 = 'Tier lookup'
    $wsRes.Range('A4').Value2 = 'Total'
    $wsRes.Range('B2').Formula = '=Input!B3/Input!B2'
    $wsRes.Range('B3').Formula = '=VLOOKUP(Input!B2,$E$2:$F$4,2,FALSE)'
    $wsRes.Range('B4').Formula = '=Input!B2*Input!B3'
    # Small lookup table for the VLOOKUP (tiers 1/2/3 only -> other quantities give #N/A).
    $wsRes.Range('E2').Value2 = 1; $wsRes.Range('F2').Value2 = 10
    $wsRes.Range('E3').Value2 = 2; $wsRes.Range('F3').Value2 = 20
    $wsRes.Range('E4').Value2 = 3; $wsRes.Range('F4').Value2 = 30

    # --- Import the test module into the VBProject ---------------------------
    try {
        $null = $wb.VBProject.VBComponents.Import($moduleSrc)
    } catch {
        throw "Could not import devkit_Test.bas into the VBProject. Enable " +
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
