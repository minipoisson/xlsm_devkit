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
        $sheetLabel = if ($_.code_name) { $_.code_name } else { $_.sheet }
        "$sheetLabel!$loc=$($_.value) [$($_.type)]"
    }) -join '; '
}

function Remove-ComObject($obj) {
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# Invokes a test's fixtures (seed cells) once, before the test runs. Reuses the
# same shape as the original Phase 1 fixture step; [string]-casts Get-Content -Raw
# so PSPath/PSProvider note properties are not serialized into the JSON.
function Invoke-TestFixtures($xl, [string]$prefix, $test, [string]$wbFolder) {
    $fixtureList = New-Object System.Collections.ArrayList
    foreach ($fx in @(Get-Prop $test 'fixtures')) {
        $fxObj = @{}
        $fxSheet = Get-Prop $fx 'sheet'
        $fxCode  = Get-Prop $fx 'code_name'
        if ($fxCode)  { $fxObj.code_name = $fxCode }
        if ($fxSheet) { $fxObj.sheet = $fxSheet }
        $fxMd = Get-Prop $fx 'md'
        $fxMdFile = Get-Prop $fx 'md_file'
        if ($fxMdFile) {
            $fxPath = if ([System.IO.Path]::IsPathRooted($fxMdFile)) { $fxMdFile } else { Join-Path $wbFolder $fxMdFile }
            if (-not (Test-Path -LiteralPath $fxPath)) { throw "fixture md_file not found: $fxPath" }
            $fxMd = [string](Get-Content -LiteralPath $fxPath -Raw -Encoding UTF8)
        }
        if ($fxMd) { $fxObj.md = $fxMd; [void]$fixtureList.Add($fxObj) }
    }
    if ($fixtureList.Count -gt 0) {
        $fixtureJson = ConvertTo-Json @{ fixtures = @($fixtureList) } -Depth 6 -Compress
        $fxRes = $xl.Run($prefix + 'DevkitApplyFixture', $fixtureJson) | ConvertFrom-Json
        if ($null -ne ($fxRes.PSObject.Properties['error'])) { throw "DevkitApplyFixture failed: $($fxRes.error)" }
        Write-Host "  fixtures: applied $($fxRes.applied) cell(s)."
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

    # --- 2. Resolve input targets (needed only by property-type tests) ------
    $resolved = $xl.Run($runPrefix + 'DevkitResolveInputs', $metaJson) | ConvertFrom-Json
    if ($null -ne ($resolved.PSObject.Properties['error'])) { throw "DevkitResolveInputs failed: $($resolved.error)" }
    # Normalize to plain hashtables so optional code_name/address/range are always present (as $null).
    $normTargets = @(@(Get-Prop $resolved 'inputs') | ForEach-Object {
        @{ sheet = $_.sheet; code_name = (Get-Prop $_ 'code_name'); address = (Get-Prop $_ 'address'); range = (Get-Prop $_ 'range') }
    })
    Write-Host "Resolved $($normTargets.Count) input target(s)."

    # --- 3. Run each test, dispatching on its "type" (default: property) -----
    if ($Seed -eq 0) { $rng = New-Object System.Random } else { $rng = New-Object System.Random($Seed) }
    $boundary = Get-BoundarySet

    $failures = New-Object System.Collections.ArrayList
    $totalCases = 0
    $failedCases = 0

    foreach ($test in @(Get-Prop $testSpec 'tests')) {
        $testType = Resolve-Default (Get-Prop $test 'type') 'property'
        $testName = Resolve-Default (Get-Prop $test 'name') $testType
        Write-Host "Test: $testName [$testType]"

        # Seed cells for this test (non-destructive), before running it.
        Invoke-TestFixtures $xl $runPrefix $test $wbFolder

        if ($testType -eq 'scenario') {
            # --- scenario: apply fixed inputs once, then check expectations ---
            $scInputs = @(@(Get-Prop $test 'inputs') | ForEach-Object {
                @{ sheet = (Get-Prop $_ 'sheet'); code_name = (Get-Prop $_ 'code_name');
                   address = (Get-Prop $_ 'address'); range = (Get-Prop $_ 'range');
                   value = (Get-Prop $_ 'value'); type = (Get-Prop $_ 'type') }
            })
            $applyInputs = @($scInputs | ForEach-Object {
                $o = @{ value = $_.value; type = $_.type }
                if ($_.code_name) { $o.code_name = $_.code_name }
                if ($_.sheet)     { $o.sheet = $_.sheet }
                if ($_.address)   { $o.address = $_.address }
                if ($_.range)     { $o.range = $_.range }
                $o
            })
            $applyJson = ConvertTo-Json @{ inputs = @($applyInputs) } -Depth 6 -Compress
            $applyRes = $xl.Run($runPrefix + 'DevkitApplyInputs', $applyJson) | ConvertFrom-Json
            if ($null -ne ($applyRes.PSObject.Properties['error'])) { throw "DevkitApplyInputs failed: $($applyRes.error)" }

            $null = $xl.Run($runPrefix + 'DevkitCalculateFullRebuild') | ConvertFrom-Json

            $expList = @(@(Get-Prop $test 'expect') | ForEach-Object {
                $e = @{}
                $eCode = Get-Prop $_ 'code_name'; $eSheet = Get-Prop $_ 'sheet'
                if ($eCode)  { $e.code_name = $eCode }
                if ($eSheet) { $e.sheet = $eSheet }
                $eAddr = Get-Prop $_ 'address'; if ($eAddr) { $e.address = $eAddr }
                $eAssert = Get-Prop $_ 'assert'; if ($eAssert) { $e.assert = $eAssert }
                # value is compared as displayed text, so serialize it as a string.
                if ($null -ne $_.PSObject.Properties['value']) { $e.value = [string](Get-Prop $_ 'value') }
                $e
            })
            $expJson = ConvertTo-Json @{ expectations = @($expList) } -Depth 6 -Compress
            $assertRes = $xl.Run($runPrefix + 'DevkitAssertExpectations', $expJson) | ConvertFrom-Json
            if ($null -ne ($assertRes.PSObject.Properties['error'])) { throw "DevkitAssertExpectations failed: $($assertRes.error)" }

            $totalCases++
            if ([int]$assertRes.failCount -gt 0) {
                $failedCases++
                $summary = Format-InputsSummary $scInputs
                foreach ($f in @($assertRes.failures)) {
                    if ($failures.Count -lt 200) {
                        [void]$failures.Add(@{
                            case = $testName; inputs = $summary;
                            sheet = $f.sheet; address = $f.address;
                            error = $f.error; formula = $f.formula
                        })
                    }
                }
            }
        }
        else {
            # --- property (default): generate inputs, assert no Excel errors --
            if ($normTargets.Count -eq 0) {
                throw "No input cells resolved for property test '$testName'. Check meta.json roles / unlocked cells."
            }

            # Assertion targets from this test's expect blocks (no_error over ranges).
            $assertTargets = New-Object System.Collections.ArrayList
            foreach ($exp in @(Get-Prop $test 'expect')) {
                $t = @{}
                $expSheet = Get-Prop $exp 'sheet'
                $expCode  = Get-Prop $exp 'code_name'
                if ($expCode)  { $t.code_name = $expCode }
                if ($expSheet) { $t.sheet = $expSheet }
                $expUsed = Get-Prop $exp 'used_range'
                $expRange = Get-Prop $exp 'range'
                if ($expUsed) { $t.used_range = $true }
                elseif ($expRange) { $t.range = $expRange }
                else { $t.used_range = $true }
                [void]$assertTargets.Add($t)
            }
            $assertJson = ConvertTo-Json @{ targets = @($assertTargets) } -Depth 6 -Compress

            # Case count: -Cases override, else generate.cases, else 50.
            $caseCount = $Cases
            if ($caseCount -le 0) {
                $genCases = Get-Prop (Get-Prop $test 'generate') 'cases'
                if ($genCases) { $caseCount = [int]$genCases }
            }
            if ($caseCount -le 0) { $caseCount = 50 }

            # Generate cases: boundary sweep, then random per-target.
            $caseList = New-Object System.Collections.ArrayList
            for ($j = 0; $j -lt $boundary.Count -and $caseList.Count -lt $caseCount; $j++) {
                $ci = @()
                foreach ($tg in $normTargets) {
                    $ci += @{ sheet = $tg.sheet; code_name = $tg.code_name; address = $tg.address; range = $tg.range;
                              value = $boundary[$j].value; type = $boundary[$j].type }
                }
                [void]$caseList.Add($ci)
            }
            while ($caseList.Count -lt $caseCount) {
                $ci = @()
                foreach ($tg in $normTargets) {
                    if ($rng.Next(0, 100) -lt 30) { $pick = $boundary[$rng.Next(0, $boundary.Count)] }
                    else { $pick = Get-RandomValue $rng }
                    $ci += @{ sheet = $tg.sheet; code_name = $tg.code_name; address = $tg.address; range = $tg.range;
                              value = $pick.value; type = $pick.type }
                }
                [void]$caseList.Add($ci)
            }

            # Run the cases.
            $caseNo = 0
            foreach ($ci in $caseList) {
                $caseNo++
                $applyInputs = @($ci | ForEach-Object {
                    $o = @{ value = $_.value; type = $_.type }
                    if ($_.code_name) { $o.code_name = $_.code_name }
                    if ($_.sheet)     { $o.sheet = $_.sheet }
                    if ($_.address)   { $o.address = $_.address }
                    if ($_.range)     { $o.range = $_.range }
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
            $totalCases += $caseList.Count
        }
    }

    # --- 4. Write result -----------------------------------------------------
    $passed = ($failedCases -eq 0)
    $payload = @{
        outputDir   = $outputDir
        workbook    = $wbName
        generatedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        summary     = @{ cases = $totalCases; failedCases = $failedCases; passed = $passed }
        failures    = @($failures)
    }
    $payloadJson = ConvertTo-Json $payload -Depth 8 -Compress
    $writeRes = $xl.Run($runPrefix + 'DevkitWriteResult', $payloadJson) | ConvertFrom-Json
    if (-not $writeRes.ok) { throw "DevkitWriteResult failed: $($writeRes.error)" }

    Write-Host ""
    Write-Host "Cases run : $totalCases"
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
