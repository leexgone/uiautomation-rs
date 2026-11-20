# check-versions.ps1 的离线测试：在临时目录构造最小 workspace 夹具，断言退出码。
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = $PSScriptRoot
$sut = Join-Path $scriptDir 'check-versions.ps1'
$script:failures = 0

function New-Fixture {
    param(
        [string]$Require,           # 写进主包清单的 derive 版本要求；'___MISSING___' 表示不写该依赖
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
            Write-Host "FAIL: $Name (期望退出码 $ExpectedCode 且包含 '$OutputMustContain'，实际 $($r.Code))" -ForegroundColor Red
            Write-Host $r.Output
        }
    } finally {
        Remove-Item -Recurse -Force $fixture
    }
}

Assert-Case -Name '旧版本要求 0.7.9 判失败' -Require '0.7.9' -ExpectedCode 1 -OutputMustContain '版本校验失败'
Assert-Case -Name 'caret 要求 ^0.7.10 判失败' -Require '^0.7.10' -ExpectedCode 1 -OutputMustContain '=0.7.10'
Assert-Case -Name '精确旧版本 =0.7.9 判失败' -Require '=0.7.9' -ExpectedCode 1 -OutputMustContain '版本校验失败'
Assert-Case -Name '精确当前版本 =0.7.10 通过' -Require '=0.7.10' -ExpectedCode 0 -OutputMustContain '版本校验通过'
Assert-Case -Name '依赖缺失判失败' -Require '___MISSING___' -ExpectedCode 1 -OutputMustContain '缺少对 uiautomation_derive 的依赖'
Assert-Case -Name '正确 tag v0.25.2 通过' -Require '=0.7.10' -Tag 'v0.25.2' -ExpectedCode 0 -OutputMustContain 'tag v0.25.2 一致'
Assert-Case -Name '错误 tag v0.26.0 判失败' -Require '=0.7.10' -Tag 'v0.26.0' -ExpectedCode 1 -OutputMustContain 'tag 必须为'

if ($script:failures -gt 0) {
    Write-Host "共有 $script:failures 个用例失败" -ForegroundColor Red
    exit 1
}
Write-Host '全部用例通过' -ForegroundColor Green
