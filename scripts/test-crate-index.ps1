# Offline tests for the crates.io sparse-index helpers in CrateIndex.psm1.
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:failures = 0
$module = Join-Path $PSScriptRoot 'CrateIndex.psm1'
Import-Module $module -Force

function Assert-Equal {
    param([string]$Name, $Actual, $Expected)
    if ($Actual -ceq $Expected) {
        Write-Host "PASS: $Name" -ForegroundColor Green
    } else {
        $script:failures++
        Write-Host "FAIL: $Name (expected '$Expected', got '$Actual')" -ForegroundColor Red
    }
}

Assert-Equal -Name 'index URL for uiautomation' `
    -Actual (Get-CrateIndexUrl 'uiautomation') `
    -Expected 'https://index.crates.io/ui/au/uiautomation'

Assert-Equal -Name 'index URL for uiautomation_derive' `
    -Actual (Get-CrateIndexUrl 'uiautomation_derive') `
    -Expected 'https://index.crates.io/ui/au/uiautomation_derive'

# Sparse index payload: newline-delimited JSON lines (text/plain), may include
# CRLF, blank lines and a trailing newline.
$fakeIndex = "{`"vers`":`"1.0.0`"}`r`n{`"vers`":`"0.7.10`"}`r`n`r`n"
$versions = ConvertFrom-CrateIndexText -Text $fakeIndex
Assert-Equal -Name 'parsed entry count' -Actual $versions.Count -Expected 2
Assert-Equal -Name 'parsed first version' -Actual $versions[0] -Expected '1.0.0'
Assert-Equal -Name 'parsed second version' -Actual $versions[1] -Expected '0.7.10'

$empty = ConvertFrom-CrateIndexText -Text ''
Assert-Equal -Name 'empty payload yields empty list' -Actual @($empty).Count -Expected 0

if ($script:failures -gt 0) {
    Write-Host "$script:failures test case(s) failed" -ForegroundColor Red
    exit 1
}
Write-Host 'All crate-index helper tests passed' -ForegroundColor Green
# Explicit success code so the result never depends on a stale $LASTEXITCODE.
exit 0
