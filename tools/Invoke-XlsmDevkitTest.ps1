<#
.SYNOPSIS
    Run xlsm_devkit Phase 1 MVP tests (no_error property check) against a workbook.

.DESCRIPTION
    Standard Windows PowerShell only -- no Python, no external modules.
    The script:
      1. Copies the target workbook to a temp file (the original is never tested).
      2. Opens one Excel COM instance (Visible=$false, DisplayAlerts=$false, UpdateLinks=0).
      3. Reads tests/workbook.meta.json and the test spec (no_error.test.json).
      4. Resolves input cells via DevkitResolveInputs (unlocked cells on protected sheets).
      5. Generates boundary + random input cases and, per case, calls
         DevkitApplyInputs -> DevkitCalculateFullRebuild -> DevkitAssertNoErrors.
      6. Writes result.json / result.md next to the ORIGINAL workbook via DevkitWriteResult.

    All VBA API calls pass and return JSON strings (Application.Run).

.PARAMETER Workbook
    Path to the .xlsm/.xlsb workbook to test. The workbook must contain devkit_Test.bas.

.PARAMETER MetaPath
    Path to workbook.meta.json. Default: <workbook folder>\tests\workbook.meta.json

.PARAMETER TestPath
    Path to the test spec JSON. Default: <workbook folder>\tests\no_error.test.json

.PARAMETER Cases
    Override the number of generated cases (otherwise taken from the test spec).

.PARAMETER Seed
    Random seed for reproducible generation.

.PARAMETER AllowDestructive
    Gate for destructive operations. The no_error check is read-centric and does not
    need it; the switch exists so future destructive test types must opt in explicitly.

.EXAMPLE
    .\tools\Invoke-XlsmDevkitTest.ps1 -Workbook .\examples\test-harness\sample.xlsm
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Workbook,
    [string]$MetaPath,
    [string]$TestPath,
    [int]$Cases = 0,
    [int]$Seed = 0,
    [switch]$AllowDestructive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Resolve-Default([string]$value, [string]$fallback) {
    if ([string]::IsNullOrWhiteSpace($value)) { return $fallback } else { return $value }
}

# Safe optional-property access for PSCustomObjects produced by ConvertFrom-Json.
function Get-Prop($obj, [string]$name, $default = $null) {
    if ($null -eq $obj) { return $default }
    $p = $obj.PSObject.Properties[$name]
    if ($null -eq $p) { return $default }
    return $p.Value
}

# Build the boundary value set: each entry is @{ value=...; type=...; label=... }
function Get-BoundarySet {
    @(
        @{ value = '';          type = 'blank';  label = '(blank)' },
        @{ value = '0';         type = 'number'; label = '0' },
        @{ value = '-1';        type = 'number'; label = '-1' },
        @{ value = '99999999';  type = 'number'; label = '99999999' },
        @{ value = '1.5';       type = 'number'; label = '1.5' },
        @{ value = ('A' * 255); type = 'text';   label = 'long-string(255)' },
        @{ value = '123';       type = 'text';   label = 'numeric-string' },
        @{ value = '2099-12-31'; type = 'text';  label = 'date-like' },
        @{ value = '   ';       type = 'text';   label = 'whitespace' }
    )
}

function Get-RandomValue([System.Random]$rng) {
    switch ($rng.Next(0, 3)) {
        0 { return @{ value = ([string]($rng.Next(-100000, 100001)));        type = 'number' } }
        1 { return @{ value = ([string]([math]::Round($rng.NextDouble() * 1000, 4))); type = 'number' } }
        default { return @{ value = ('X' + $rng.Next(0, 9999));               type = 'text' } }
    }
}

# Format a case's inputs for the result.md "Inputs" column.
function Format-InputsSummary($caseInputs) {
    ($caseInputs | ForEach-Object {
        $loc = if ($_.address) { $_.address } else { $_.range }
        "$($_.sheet)!$loc=$($_.value) [$($_.type)]"
    }) -join '; '
}

function Remove-ComObject($obj) {
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# ---------------------------------------------------------------------------
# Resolve paths
# ---------------------------------------------------------------------------

$wbPath = (Resolve-Path -LiteralPath $Workbook).Path
$wbFolder = Split-Path -Parent $wbPath
$wbName = Split-Path -Leaf $wbPath

$MetaPath = Resolve-Default $MetaPath (Join-Path $wbFolder 'tests\workbook.meta.json')
$TestPath = Resolve-Default $TestPath (Join-Path $wbFolder 'tests\no_error.test.json')

if (-not (Test-Path -LiteralPath $MetaPath)) { throw "meta.json not found: $MetaPath" }
if (-not (Test-Path -LiteralPath $TestPath)) { throw "test spec not found: $TestPath" }

$metaJson = Get-Content -LiteralPath $MetaPath -Raw -Encoding UTF8
$testSpec = Get-Content -LiteralPath $TestPath -Raw -Encoding UTF8 | ConvertFrom-Json

$outputDir = Join-Path $wbFolder 'test-results\latest'
$lockPath  = Join-Path $wbFolder 'xlsm_devkit_test.lock'

# ---------------------------------------------------------------------------
# Single-instance lock
# ---------------------------------------------------------------------------

if (Test-Path -LiteralPath $lockPath) {
    throw "Another test run appears active (lock file present): $lockPath. Delete it if stale."
}
$null = New-Item -ItemType Directory -Force -Path $outputDir
Set-Content -LiteralPath $lockPath -Value ([string]$PID) -Encoding ASCII

# ---------------------------------------------------------------------------
# Work copy (never test the original)
# ---------------------------------------------------------------------------

$ext = [System.IO.Path]::GetExtension($wbPath)
$workDir = Join-Path $wbFolder 'test-results\work'
$null = New-Item -ItemType Directory -Force -Path $workDir
$tmpPath = Join-Path $workDir ("xlsmdevkit_" + [System.Guid]::NewGuid().ToString('N') + $ext)
Copy-Item -LiteralPath $wbPath -Destination $tmpPath -Force

$xl = $null
$wb = $null
$exitCode = 0
$previousAutomationSecurity = $null

try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    try {
        $previousAutomationSecurity = $xl.AutomationSecurity
        # 1 = msoAutomationSecurityLow. The runner can only call Devkit* APIs if
        # macros are enabled for this explicit, local test-workbook copy.
        $xl.AutomationSecurity = 1
    } catch {
        Write-Warning "Could not set Excel AutomationSecurity. If DevkitTestPing cannot run, add this workbook folder to Trusted Locations or enable macros for automation."
    }

    # Open(Filename, UpdateLinks=0, ReadOnly=$false)
    $wb = $xl.Workbooks.Open($tmpPath, 0, $false)
    $runPrefix = "'" + $wb.Name + "'!"
    if ($null -ne $previousAutomationSecurity) {
        try { $xl.AutomationSecurity = $previousAutomationSecurity } catch { }
    }

    # --- 1. Ping -------------------------------------------------------------
    try {
        $ping = $xl.Run($runPrefix + 'DevkitTestPing') | ConvertFrom-Json
    } catch {
        throw "Could not run DevkitTestPing in the workbook copy: $tmpPath. " +
              "Check that src/devkit_Test.bas was imported into the workbook and that Excel allows macros for this run. " +
              "Underlying error: $($_.Exception.Message)"
    }
    if (-not $ping.ok) { throw "DevkitTestPing failed: $($ping.error)" }
    Write-Host "Connected: devkit_Test $($ping.version) on $($ping.workbook)"

    # --- 2. Resolve input targets -------------------------------------------
    $resolved = $xl.Run($runPrefix + 'DevkitResolveInputs', $metaJson) | ConvertFrom-Json
    if ($null -ne ($resolved.PSObject.Properties['error'])) { throw "DevkitResolveInputs failed: $($resolved.error)" }
    # Normalize to plain hashtables so optional address/range are always present (as $null).
    $normTargets = @(@(Get-Prop $resolved 'inputs') | ForEach-Object {
        @{ sheet = $_.sheet; address = (Get-Prop $_ 'address'); range = (Get-Prop $_ 'range') }
    })
    if ($normTargets.Count -eq 0) { throw "No input cells resolved. Check meta.json roles / unlocked cells." }
    Write-Host "Resolved $($normTargets.Count) input target(s)."

    # --- 3. Build assertion targets from the test spec expect blocks --------
    $assertTargets = New-Object System.Collections.ArrayList
    $caseCount = $Cases
    foreach ($test in @(Get-Prop $testSpec 'tests')) {
        $gen = Get-Prop $test 'generate'
        $genCases = Get-Prop $gen 'cases'
        if ($caseCount -le 0 -and $genCases) { $caseCount = [int]$genCases }
        foreach ($exp in @(Get-Prop $test 'expect')) {
            $t = @{ sheet = $exp.sheet }
            $expUsed = Get-Prop $exp 'used_range'
            $expRange = Get-Prop $exp 'range'
            if ($expUsed) { $t.used_range = $true }
            elseif ($expRange) { $t.range = $expRange }
            else { $t.used_range = $true }
            [void]$assertTargets.Add($t)
        }
    }
    if ($caseCount -le 0) { $caseCount = 50 }
    $assertJson = ConvertTo-Json @{ targets = @($assertTargets) } -Depth 6 -Compress

    # --- 4. Generate cases (boundary sweep, then random) --------------------
    if ($Seed -eq 0) { $rng = New-Object System.Random } else { $rng = New-Object System.Random($Seed) }
    $boundary = Get-BoundarySet

    $caseList = New-Object System.Collections.ArrayList
    # Phase A: boundary sweep -- every target gets boundary[j] (guarantees 0, blank, etc.)
    for ($j = 0; $j -lt $boundary.Count -and $caseList.Count -lt $caseCount; $j++) {
        $ci = @()
        foreach ($tg in $normTargets) {
            $ci += @{ sheet = $tg.sheet; address = $tg.address; range = $tg.range;
                      value = $boundary[$j].value; type = $boundary[$j].type }
        }
        [void]$caseList.Add($ci)
    }
    # Phase B: random per-target
    while ($caseList.Count -lt $caseCount) {
        $ci = @()
        foreach ($tg in $normTargets) {
            if ($rng.Next(0, 100) -lt 30) { $pick = $boundary[$rng.Next(0, $boundary.Count)] }
            else { $pick = Get-RandomValue $rng }
            $ci += @{ sheet = $tg.sheet; address = $tg.address; range = $tg.range;
                      value = $pick.value; type = $pick.type }
        }
        [void]$caseList.Add($ci)
    }

    # --- 5. Run the cases ----------------------------------------------------
    $failures = New-Object System.Collections.ArrayList
    $failedCases = 0
    $caseNo = 0
    foreach ($ci in $caseList) {
        $caseNo++
        $applyInputs = @($ci | ForEach-Object {
            $o = @{ sheet = $_.sheet; value = $_.value; type = $_.type }
            if ($_.address) { $o.address = $_.address }
            if ($_.range)   { $o.range = $_.range }
            $o
        })
        $applyJson = ConvertTo-Json @{ inputs = $applyInputs } -Depth 6 -Compress

        $applyRes = $xl.Run($runPrefix + 'DevkitApplyInputs', $applyJson) | ConvertFrom-Json
        if ($null -ne ($applyRes.PSObject.Properties['error'])) { throw "DevkitApplyInputs failed: $($applyRes.error)" }

        $null = $xl.Run($runPrefix + 'DevkitCalculateFullRebuild') | ConvertFrom-Json

        $assertRes = $xl.Run($runPrefix + 'DevkitAssertNoErrors', $assertJson) | ConvertFrom-Json
        if ($null -ne ($assertRes.PSObject.Properties['error'])) { throw "DevkitAssertNoErrors failed: $($assertRes.error)" }

        if ([int]$assertRes.failCount -gt 0) {
            $failedCases++
            $summary = Format-InputsSummary $ci
            foreach ($f in @($assertRes.failures)) {
                if ($failures.Count -lt 200) {
                    [void]$failures.Add(@{
                        case = $caseNo; inputs = $summary;
                        sheet = $f.sheet; address = $f.address;
                        error = $f.error; formula = $f.formula
                    })
                }
            }
        }
    }

    # --- 6. Write result -----------------------------------------------------
    $passed = ($failedCases -eq 0)
    $payload = @{
        outputDir   = $outputDir
        workbook    = $wbName
        generatedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        summary     = @{ cases = $caseList.Count; failedCases = $failedCases; passed = $passed }
        failures    = @($failures)
    }
    $payloadJson = ConvertTo-Json $payload -Depth 8 -Compress
    $writeRes = $xl.Run($runPrefix + 'DevkitWriteResult', $payloadJson) | ConvertFrom-Json
    if (-not $writeRes.ok) { throw "DevkitWriteResult failed: $($writeRes.error)" }

    Write-Host ""
    Write-Host "Cases run : $($caseList.Count)"
    Write-Host "Failed    : $failedCases"
    Write-Host "Result    : $(if ($passed) {'PASS'} else {'FAIL'})"
    Write-Host "result.md : $($writeRes.md)"
    if (-not $passed) { $exitCode = 1 }
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }
    if ($null -ne $xl -and $null -ne $previousAutomationSecurity) {
        try { $xl.AutomationSecurity = $previousAutomationSecurity } catch { }
    }
    if ($null -ne $xl) { try { $xl.Quit() } catch { } }
    Remove-ComObject $wb
    Remove-ComObject $xl
    $wb = $null; $xl = $null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tmpPath) { Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $lockPath) { Remove-Item -LiteralPath $lockPath -Force -ErrorAction SilentlyContinue }
}

exit $exitCode
