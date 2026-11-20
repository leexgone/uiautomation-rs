# 幂等发布一个 crate 到 crates.io：
# 1. 从 cargo metadata 读取工作区中该 crate 的版本；
# 2. 查询 crates.io sparse 索引，版本已存在则跳过（tag 重跑安全）；
# 3. 不存在则先 cargo publish --dry-run，再正式 cargo publish；
# 4. 发布后轮询索引，确认新版本可见（为后续主包发布消除传播延迟）。
# 凭证：cargo 自动读取环境变量 CARGO_REGISTRY_TOKEN，不要通过命令行传 token。
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('uiautomation', 'uiautomation_derive')]
    [string]$Crate
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot -Parent
$workspaceManifest = Join-Path $root 'Cargo.toml'
Set-Location $root

$metadataJson = & cargo metadata --format-version 1 --no-deps --manifest-path $workspaceManifest
if ($LASTEXITCODE -ne 0 -or -not $metadataJson) { throw 'cargo metadata 执行失败' }
$pkg = ($metadataJson | ConvertFrom-Json).packages |
    Where-Object { $_.name -eq $Crate } |
    Select-Object -First 1
if (-not $pkg) { throw "metadata 中找不到 crate $Crate" }
$version = $pkg.version

# crates.io 索引路径规则：
# 1 字符 -> 1/<name>；2 字符 -> 2/<name>；3 字符 -> 3/<首字符>/<name>；
# 更长 -> <前2字符>/<第3-4字符>/<name>
function Get-CrateIndexUrl {
    param([string]$Name)
    $lower = $Name.ToLowerInvariant()
    if ($lower.Length -eq 1) { return "https://index.crates.io/1/$lower" }
    if ($lower.Length -eq 2) { return "https://index.crates.io/2/$lower" }
    if ($lower.Length -eq 3) { return "https://index.crates.io/3/$($lower[0])/$lower" }
    return "https://index.crates.io/$($lower.Substring(0, 2))/$($lower.Substring(2, 2))/$lower"
}

function Get-PublishedVersions {
    param([string]$Url)
    $versions = @()
    try {
        # sparse 索引返回换行分隔的 JSON 文本行（Content-Type 为 text/plain），
        # 必须先取原文再按行拆分；直接 Invoke-RestMethod/整体 ConvertFrom-Json 会失败。
        $text = (Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Get).Content
        foreach ($line in ($text -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $versions += ($line | ConvertFrom-Json).vers
            }
        }
    } catch {
        # 404 表示该 crate 在索引中尚无任何记录（两个目标 crate 不会出现，但按通用规则处理）
        $status = 0
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        if ($status -ne 404) { throw }
    }
    return $versions
}

$url = Get-CrateIndexUrl -Name $Crate
if ((Get-PublishedVersions -Url $url) -contains $version) {
    Write-Host "[$Crate $version] crates.io 上已存在，跳过发布。" -ForegroundColor Yellow
    exit 0
}

Write-Host "[$Crate $version] 未发布，先执行 dry-run ..." -ForegroundColor Cyan
cargo publish -p $Crate --dry-run
if ($LASTEXITCODE -ne 0) { throw "[$Crate] dry-run 失败" }

Write-Host "[$Crate $version] 正式发布 ..." -ForegroundColor Cyan
# 注意：cargo publish 不支持 --all-features；全 feature 编译由发版 workflow 中
# `cargo check --workspace --all-targets --all-features` 闸门保证，publish 自身只验证默认 feature。
cargo publish -p $Crate
if ($LASTEXITCODE -ne 0) { throw "[$Crate] publish 失败" }

$deadline = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    if ((Get-PublishedVersions -Url $url) -contains $version) {
        Write-Host "[$Crate $version] 已进入 crates.io 索引。" -ForegroundColor Green
        exit 0
    }
    Write-Host '索引尚未出现新版本，继续等待 ...'
}
throw "[$Crate $version] 发布后 120 秒内未在索引中出现；可直接重跑该 tag（流程幂等）。"
