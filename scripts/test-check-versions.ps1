# Offline tests for check-versions.ps1: build minimal workspace fixtures in a
# temp directory and assert exit codes and message fragments.
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$sut = Join-Path $scriptDir 'check-versions.ps1'
$script:failures = 0

function New-Fixture {
    param(
        [string]$Require,           # derive version requirement written into the main manifest; '___MISSING___' omits the dependency
        [string]$MainVersion = '0.25.2',
        [string]$DeriveVersion = '0.7.10'
    )
    $dir = Join-Path ([System.IO.Path]::GetTempPath()) ('cver-' + [System.Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path "$dir/crates/uiautomation/src" -Force | Out-Null
    New-Item -ItemType Directory -Path "$dir/crates/uiautomation_derive/src" -Force | Out-Null

    Set-Content -Encoding UTF8 -Path "$dir/Cargo.toml" -Value @'
[workspace]
members = ["crates/uiautomation", "crates/uiautomation_derive"]
resolver = "2"
'@

    $depLine = if ($Require -eq '___MISSING___') {
        ''
    } else {
        "uiautomation_derive = { version = `"$Require`", path = `"../uiautomation_derive`" }"
    }
    Set-Content -Encoding UTF8 -Path "$dir/crates/uiautomation/Cargo.toml" -Value @"
[package]
name = "uiautomation"
version = "$MainVersion"
edition = "2024"

[dependencies]
$depLine
"@
    Set-Content -Encoding UTF8 -Path "$dir/crates/uiautomation_derive/Cargo.toml" -Value @"
[package]
name = "uiautomation_derive"
version = "$DeriveVersion"
edition = "2024"

[lib]
proc-macro = true
"@
    Set-Content -Encoding UTF8 -Path "$dir/crates/uiautomation/src/lib.rs" -Value ''
    Set-Content -Encoding UTF8 -Path "$dir/crates/uiautomation_derive/src/lib.rs" -Value ''
    return $dir
}

function Invoke-Check {
    param([string]$Root, [string]$Tag)
    $argList = @('-NoProfile', '-File', $sut, '-Root', $Root)
    if ($Tag) { $argList += @('-ExpectedTag', $Tag) }
    $output = & pwsh @argList 2>&1
    return [pscustomobject]@{ Code = $LASTEXITCODE; Output = ($output -join "`n") }
}

function Assert-Case {
    param(
        [string]$Name,
        [string]$Require,
        [string]$Tag,
        [int]$ExpectedCode,
        [string]$OutputMustContain
    )
    $fixture = New-Fixture -Require $Require
    try {
        $r = Invoke-Check -Root $fixture -Tag $Tag
        $ok = ($r.Code -eq $ExpectedCode) -and ($r.Output -like "*$OutputMustContain*")
        if ($ok) {
            Write-Host "PASS: $Name" -ForegroundColor Green
        } else {
            $script:failures++
            Write-Host "FAIL: $Name (expected exit code $ExpectedCode and output containing '$OutputMustContain', got $($r.Code))" -ForegroundColor Red
            Write-Host $r.Output
        }
    } finally {
        Remove-Item -Recurse -Force $fixture
    }
}

Assert-Case -Name 'loose requirement 0.7.9 is rejected' -Require '0.7.9' -ExpectedCode 1 -OutputMustContain 'version check failed'
Assert-Case -Name 'caret requirement ^0.7.10 is rejected' -Require '^0.7.10' -ExpectedCode 1 -OutputMustContain '=0.7.10'
Assert-Case -Name 'pinned old version =0.7.9 is rejected' -Require '=0.7.9' -ExpectedCode 1 -OutputMustContain 'version check failed'
Assert-Case -Name 'pinned current version =0.7.10 passes' -Require '=0.7.10' -ExpectedCode 0 -OutputMustContain 'version check passed'
Assert-Case -Name 'missing dependency is rejected' -Require '___MISSING___' -ExpectedCode 1 -OutputMustContain 'missing its dependency on uiautomation_derive'
Assert-Case -Name 'matching tag v0.25.2 passes' -Require '=0.7.10' -Tag 'v0.25.2' -ExpectedCode 0 -OutputMustContain 'tag v0.25.2 matches'
Assert-Case -Name 'wrong tag v0.26.0 is rejected' -Require '=0.7.10' -Tag 'v0.26.0' -ExpectedCode 1 -OutputMustContain 'tag must be'

if ($script:failures -gt 0) {
    Write-Host "$script:failures test case(s) failed" -ForegroundColor Red
    exit 1
}
Write-Host 'All test cases passed' -ForegroundColor Green
