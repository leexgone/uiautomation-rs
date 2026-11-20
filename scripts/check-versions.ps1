# Validate version consistency between uiautomation and uiautomation_derive.
# - Default (CI): the main crate's requirement on the derive crate must be
#   "=<current derive version>" (exact pin).
# - -ExpectedTag <tag>: additionally require the tag to equal "v<main crate version>" (used by release).
# - -Root <path>: repository root; defaults to the parent of this script's directory (tests pass a fixture).
[CmdletBinding()]
param(
    [string]$ExpectedTag,
    [string]$Root = (Split-Path $PSScriptRoot -Parent)
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Write-Fail {
    param([string]$Message)
    Write-Host "[version check failed] $Message" -ForegroundColor Red
    exit 1
}

$manifestPath = Join-Path $Root 'Cargo.toml'
if (-not (Test-Path $manifestPath)) {
    Write-Fail "Cannot find Cargo.toml under '$Root'; pass the repository root with -Root."
}

try {
    $metadataJson = & cargo metadata --format-version 1 --no-deps --manifest-path $manifestPath
    if ($LASTEXITCODE -ne 0 -or -not $metadataJson) { throw 'cargo metadata exited unsuccessfully' }
} catch {
    Write-Fail "cargo metadata failed; make sure the Rust toolchain is installed and the manifest is valid ($manifestPath): $_"
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
if (-not $main) { Write-Fail 'workspace crate uiautomation not found in metadata.' }
$derive = Find-WorkspaceCrate -Name 'uiautomation_derive'
if (-not $derive) { Write-Fail 'workspace crate uiautomation_derive not found in metadata.' }

$dep = $main.dependencies | Where-Object { $_.name -eq 'uiautomation_derive' } | Select-Object -First 1
if (-not $dep) {
    Write-Fail 'uiautomation is missing its dependency on uiautomation_derive in Cargo.toml.'
}

$expectedReq = "=$($derive.version)"
if ($dep.req -ne $expectedReq) {
    Write-Fail @"
uiautomation requires uiautomation_derive as '$($dep.req)', but the workspace version of uiautomation_derive is '$($derive.version)'.
The two must be pinned exactly. Change the dependency in crates/uiautomation/Cargo.toml to:
  uiautomation_derive = { version = "$expectedReq", path = "../uiautomation_derive" }
"@
}

if ($ExpectedTag) {
    # PowerShell variable names are case-insensitive; do not name this $expectedTag,
    # it would shadow the -ExpectedTag parameter.
    $correctTag = "v$($main.version)"
    if ($ExpectedTag -ne $correctTag) {
        Write-Fail "tag is '$ExpectedTag', but crates/uiautomation version is '$($main.version)'; tag must be '$correctTag'. Fix the version before tagging again."
    }
}

$tagNote = ''
if ($ExpectedTag) { $tagNote = "; tag $ExpectedTag matches" }
Write-Host "[version check passed] uiautomation $($main.version) pins uiautomation_derive $expectedReq$tagNote." -ForegroundColor Green
exit 0
