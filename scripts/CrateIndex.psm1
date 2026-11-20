# Helpers for the crates.io sparse index, shared by publish-crate.ps1 and its tests.

# Index URL for a crate name. Both published crates here have more than 3 chars
# (uiautomation=12, uiautomation_derive=17), so only the general rule is needed:
# <first 2 chars>/<chars 3-4>/<lowercase name>. Underscores stay as-is.
function Get-CrateIndexUrl {
    param([Parameter(Mandatory = $true)][string]$Name)

    $lower = $Name.ToLowerInvariant()
    return "https://index.crates.io/$($lower.Substring(0, 2))/$($lower.Substring(2, 2))/$lower"
}

# Parse a sparse-index payload (newline-delimited JSON, Content-Type text/plain)
# into the list of published version strings.
function ConvertFrom-CrateIndexText {
    param([AllowEmptyString()][string]$Text)

    $versions = @()
    if ([string]::IsNullOrWhiteSpace($Text)) { return $versions }
    foreach ($line in ($Text -split "`r?`n")) {
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            $versions += ($line | ConvertFrom-Json).vers
        }
    }
    return $versions
}

# Fetch the published versions for a crate. 404 (no index entry yet) yields an
# empty list; any other transport error is propagated.
function Get-PublishedVersions {
    param([Parameter(Mandatory = $true)][string]$Url)

    try {
        $text = (Invoke-WebRequest -Uri $Url -UseBasicParsing -Method Get -TimeoutSec 15).Content
        return ConvertFrom-CrateIndexText -Text $text
    } catch {
        # Connection-level exceptions (DNS/TLS/timeout) have no Response property;
        # read the property defensively so StrictMode cannot turn this into a
        # PropertyNotFoundException that hides the real error.
        $responseProp = $_.Exception.PSObject.Properties['Response']
        $status = 0
        if ($responseProp -and $responseProp.Value) {
            $status = [int]$responseProp.Value.StatusCode
        }
        if ($status -eq 404) { return @() }
        throw
    }
}

Export-ModuleMember -Function Get-CrateIndexUrl, ConvertFrom-CrateIndexText, Get-PublishedVersions
