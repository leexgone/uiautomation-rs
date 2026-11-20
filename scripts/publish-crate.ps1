# Idempotently publish one crate to crates.io:
# 1. Read the workspace crate version from cargo metadata;
# 2. Query the crates.io sparse index; skip when the version already exists (safe tag re-runs);
# 3. Otherwise run cargo publish --dry-run, then cargo publish;
# 4. Poll the index until the new version is visible (removes propagation delay before the main crate publishes).
# Credentials: cargo reads the CARGO_REGISTRY_TOKEN environment variable; never pass the token on the command line.
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
if ($LASTEXITCODE -ne 0 -or -not $metadataJson) { throw 'cargo metadata failed' }
$pkg = ($metadataJson | ConvertFrom-Json).packages |
    Where-Object { $_.name -eq $Crate } |
    Select-Object -First 1
if (-not $pkg) { throw "crate $Crate not found in metadata" }
$version = $pkg.version

# crates.io index path rules:
# 1 char  -> 1/<name>; 2 chars -> 2/<name>; 3 chars -> 3/<first-char>/<name>;
# longer  -> <first-2-chars>/<chars-3-4>/<name>
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
        # The sparse index returns newline-delimited JSON lines (Content-Type text/plain),
        # so fetch raw text and split per line; Invoke-RestMethod / a single ConvertFrom-Json would fail.
        $text = (Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Get).Content
        foreach ($line in ($text -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $versions += ($line | ConvertFrom-Json).vers
            }
        }
    } catch {
        # 404 means the crate has no index entry yet (not expected for either target crate, but handled generically).
        $status = 0
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        if ($status -ne 404) { throw }
    }
    return $versions
}

$url = Get-CrateIndexUrl -Name $Crate
if ((Get-PublishedVersions -Url $url) -contains $version) {
    Write-Host "[$Crate $version] already published on crates.io, skipping." -ForegroundColor Yellow
    exit 0
}

Write-Host "[$Crate $version] not published yet, running dry-run ..." -ForegroundColor Cyan
cargo publish -p $Crate --dry-run
if ($LASTEXITCODE -ne 0) { throw "[$Crate] dry-run failed" }

Write-Host "[$Crate $version] publishing ..." -ForegroundColor Cyan
# cargo publish has no --all-features; the full-feature build is guaranteed by the release
# workflow's `cargo check --workspace --all-targets --all-features` gate; publish itself validates default features.
cargo publish -p $Crate
if ($LASTEXITCODE -ne 0) { throw "[$Crate] publish failed" }

$deadline = (Get-Date).AddSeconds(120)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    if ((Get-PublishedVersions -Url $url) -contains $version) {
        Write-Host "[$Crate $version] is now visible in the crates.io index." -ForegroundColor Green
        exit 0
    }
    Write-Host 'new version not yet in the index, waiting ...'
}
throw "[$Crate $version] did not appear in the index within 120s after publishing; re-run the tag (the flow is idempotent)."
