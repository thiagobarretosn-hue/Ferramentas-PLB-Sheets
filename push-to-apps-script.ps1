<#
.SYNOPSIS
    Push local files to a Google Apps Script project via REST API.

.PARAMETER ScriptId
    Target project ID. Omit to use .clasp.json in the same folder.

.PARAMETER Reauth
    Re-authenticate with Google (runs clasp login, then exits).
    Use this when the token expires or to switch accounts.

.PARAMETER AuthFile
    Path to credentials file. Default: ~/.clasprc.json
    To use a different Google account, save its credentials to a
    separate file and pass it here.

.EXAMPLE
    # Push using .clasp.json in current folder
    .\push-to-apps-script.ps1

    # Push to a specific project ID
    .\push-to-apps-script.ps1 -ScriptId "1diHfk63bDsQh3nmumhq9ANi4Jmo8NxDBgTd0keGx98Tg37jKm29ie1tn"

    # Use a different Google account
    .\push-to-apps-script.ps1 -ScriptId "1abc..." -AuthFile "$env:USERPROFILE\.clasprc-work.json"

    # Re-authenticate (token expired or switching accounts)
    .\push-to-apps-script.ps1 -Reauth
    # After -Reauth, save a copy for a second account:
    # Copy-Item "$env:USERPROFILE\.clasprc.json" "$env:USERPROFILE\.clasprc-work.json"
    # Then re-run clasp login to restore your main account
#>
param(
    [string]$ScriptId = "",
    [switch]$Reauth,
    [string]$AuthFile = "$env:USERPROFILE\.clasprc.json"
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot

# ── Re-auth ──────────────────────────────────────────────────────────────────
if ($Reauth) {
    Write-Host "Opening Google login..." -ForegroundColor Cyan
    clasp login
    Write-Host ""
    Write-Host "Authenticated. To save credentials for a second account:" -ForegroundColor Green
    Write-Host "  Copy-Item `"$env:USERPROFILE\.clasprc.json`" `"$env:USERPROFILE\.clasprc-ACCOUNTNAME.json`"" -ForegroundColor Gray
    Write-Host "  Then run clasp login again to restore your main account." -ForegroundColor Gray
    exit
}

# ── Resolve ScriptId ─────────────────────────────────────────────────────────
if (-not $ScriptId) {
    $claspJsonPath = Join-Path $root ".clasp.json"
    if (Test-Path $claspJsonPath) {
        $ScriptId = (Get-Content $claspJsonPath -Raw | ConvertFrom-Json).scriptId
        Write-Host "ScriptId: $ScriptId  (from .clasp.json)" -ForegroundColor DarkGray
    } else {
        Write-Error "Provide -ScriptId or place .clasp.json in the same folder as this script."
    }
}

# ── Load credentials ─────────────────────────────────────────────────────────
if (-not (Test-Path $AuthFile)) {
    Write-Error "No credentials at '$AuthFile'. Run with -Reauth to authenticate."
}

$clasprc = Get-Content $AuthFile -Raw | ConvertFrom-Json

# Support clasp v2 (.token) and v3 (.tokens.default)
$creds = if ($clasprc.PSObject.Properties['tokens']) { $clasprc.tokens.default }
         else                                          { $clasprc.token }

if (-not $creds) {
    Write-Error "Could not read credentials from '$AuthFile'. Run with -Reauth."
}

# ── Refresh access token ──────────────────────────────────────────────────────
Write-Host "Refreshing token..." -ForegroundColor Cyan
$tok = Invoke-RestMethod "https://oauth2.googleapis.com/token" -Method Post -Body @{
    client_id     = $creds.client_id
    client_secret = $creds.client_secret
    refresh_token = $creds.refresh_token
    grant_type    = "refresh_token"
}
$headers = @{ Authorization = "Bearer $($tok.access_token)" }
Write-Host "Token OK" -ForegroundColor Green

# ── Ignored dirs and files (matches .claspignore) ────────────────────────────
$ignoreDirs  = @('.git', '.claude', 'docs', 'node_modules')
$ignoreFiles = @('CLAUDE.md', 'README.md', 'Código.js')

# ── Collect files ─────────────────────────────────────────────────────────────
$allFiles = [System.Collections.Generic.List[PSCustomObject]]::new()

# Manifest must come first (Apps Script API requirement)
$manifestPath = Join-Path $root "appsscript.json"
if (Test-Path $manifestPath) {
    $allFiles.Add([PSCustomObject]@{
        name   = "appsscript"
        type   = "JSON"
        source = (Get-Content $manifestPath -Raw -Encoding UTF8)
    })
}

# Scan all .gs and .html files, excluding ignored paths
Get-ChildItem -Path $root -Recurse -File |
    Where-Object {
        $rel    = $_.FullName.Substring($root.Length + 1)
        $topDir = $rel.Split([IO.Path]::DirectorySeparatorChar)[0]
        $ext    = $_.Extension.ToLower()

        if ($topDir -in $ignoreDirs)      { return $false }
        if ($_.Name -in $ignoreFiles)     { return $false }
        if ($_.Name -eq "appsscript.json"){ return $false }  # already added
        if ($ext -notin @('.gs', '.html')) { return $false }
        return $true
    } |
    Sort-Object FullName |
    ForEach-Object {
        $rel = $_.FullName.Substring($root.Length + 1) -replace '\\', '/'
        $src = Get-Content $_.FullName -Raw -Encoding UTF8

        if ($_.Extension.ToLower() -eq '.gs') {
            $allFiles.Add([PSCustomObject]@{
                name   = $rel -replace '\.gs$', ''   # e.g. "BOM" or "lib/Shared/Config"
                type   = "SERVER_JS"
                source = $src
            })
        } else {
            # HTML files MUST keep .html in name (Apps Script API requirement)
            $allFiles.Add([PSCustomObject]@{
                name   = $rel                          # e.g. "BomSidebar.html"
                type   = "HTML"
                source = $src
            })
        }
    }

Write-Host ""
Write-Host "Files to push ($($allFiles.Count)):" -ForegroundColor Cyan
$allFiles | ForEach-Object {
    Write-Host ("  {0,-12} {1}" -f $_.type, $_.name) -ForegroundColor Gray
}

# ── Push ──────────────────────────────────────────────────────────────────────
$bodyBytes = [System.Text.Encoding]::UTF8.GetBytes(
    (@{ files = $allFiles } | ConvertTo-Json -Depth 10)
)

Write-Host ""
Write-Host "Pushing to project $ScriptId ..." -ForegroundColor Cyan

$result = Invoke-RestMethod `
    -Uri     "https://script.googleapis.com/v1/projects/$ScriptId/content" `
    -Method  Put `
    -Headers $headers `
    -ContentType "application/json; charset=utf-8" `
    -Body    $bodyBytes

Write-Host "Done — $($result.files.Count) files in project." -ForegroundColor Green
