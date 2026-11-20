# 校验 uiautomation 与 uiautomation_derive 的版本一致性。
# - 默认（CI）：主包对 derive 的版本要求必须是 "=<derive 当前版本>"（精确锁死）。
# - -ExpectedTag <tag>：额外校验 tag 必须等于 "v<主包版本>"（发版时使用）。
# - -Root <path>：仓库根目录，默认取脚本所在目录的上级（测试可传入临时夹具）。
[CmdletBinding()]
param(
    [string]$ExpectedTag,
    [string]$Root = (Split-Path $PSScriptRoot -Parent)
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Write-Fail {
    param([string]$Message)
    Write-Host "[版本校验失败] $Message" -ForegroundColor Red
    exit 1
}

$manifestPath = Join-Path $Root 'Cargo.toml'
if (-not (Test-Path $manifestPath)) {
    Write-Fail "在 '$Root' 下找不到 Cargo.toml，请通过 -Root 指定仓库根目录。"
}

$metadataJson = & cargo metadata --format-version 1 --no-deps --manifest-path $manifestPath
if ($LASTEXITCODE -ne 0 -or -not $metadataJson) {
    Write-Fail "cargo metadata 执行失败，请确认 Rust 工具链可用且清单有效：$manifestPath"
}
$metadata = $metadataJson | ConvertFrom-Json

function Find-WorkspaceCrate {
    param([string]$Name)
    foreach ($pkg in $metadata.packages) {
        if ($pkg.name -ne $Name) { continue }
        $normalized = ($pkg.manifest_path -replace '\\', '/')
        if ($normalized -like "*/crates/$Name/Cargo.toml") {
            return $pkg
        }
    }
    return $null
}

$main = Find-WorkspaceCrate -Name 'uiautomation'
if (-not $main) { Write-Fail 'metadata 中找不到工作区 crate uiautomation。' }
$derive = Find-WorkspaceCrate -Name 'uiautomation_derive'
if (-not $derive) { Write-Fail 'metadata 中找不到工作区 crate uiautomation_derive。' }

$dep = $main.dependencies | Where-Object { $_.name -eq 'uiautomation_derive' } | Select-Object -First 1
if (-not $dep) {
    Write-Fail 'uiautomation 的 Cargo.toml 中缺少对 uiautomation_derive 的依赖声明。'
}

$expectedReq = "=$($derive.version)"
if ($dep.req -ne $expectedReq) {
    Write-Fail @"
uiautomation 对 uiautomation_derive 的版本要求为 '$($dep.req)'，但工作区中 uiautomation_derive 的版本是 '$($derive.version)'。
两者必须精确锁死，请将 crates/uiautomation/Cargo.toml 中该依赖改为：
  uiautomation_derive = { version = "$expectedReq", path = "../uiautomation_derive" }
"@
}

if ($ExpectedTag) {
    # 注意：PowerShell 变量大小写不敏感，局部变量不能命名为 $expectedTag，
    # 否则会覆盖参数 $ExpectedTag。
    $correctTag = "v$($main.version)"
    if ($ExpectedTag -ne $correctTag) {
        Write-Fail "tag 为 '$ExpectedTag'，但 crates/uiautomation 的版本是 '$($main.version)'；tag 必须为 '$correctTag'，请先修正版本号再重新打 tag。"
    }
}

$tagNote = ''
if ($ExpectedTag) { $tagNote = "；tag $ExpectedTag 一致" }
Write-Host "[版本校验通过] uiautomation $($main.version) 精确依赖 uiautomation_derive $expectedReq$tagNote。" -ForegroundColor Green
exit 0
