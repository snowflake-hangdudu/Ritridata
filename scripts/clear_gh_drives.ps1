param(
    [switch]$NoPrompt
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ClearGScript = Join-Path $ScriptDir "clear_g_drive.ps1"
$ClearHScript = Join-Path $ScriptDir "clear_h_drive.ps1"

if (-not (Test-Path -LiteralPath $ClearGScript)) {
    Write-Error "Script not found: $ClearGScript"
    exit 1
}

if (-not (Test-Path -LiteralPath $ClearHScript)) {
    Write-Error "Script not found: $ClearHScript"
    exit 1
}

Write-Host "This script will run in order:"
Write-Host "  1) Clear drive G:"
Write-Host "  2) Clear drive H:"
Write-Host ""

if (-not $NoPrompt) {
    $token = Read-Host "Type CLEAR-GH to continue"
    if ($token -ne "CLEAR-GH") {
        Write-Host "Confirmation token mismatch. Cancelled."
        exit 1
    }
}

Write-Host ">>> Start clearing G:"
if ($NoPrompt) {
    & $ClearGScript -NoPrompt
}
else {
    & $ClearGScript
}

Write-Host ""
Write-Host ">>> Start clearing H:"
if ($NoPrompt) {
    & $ClearHScript -NoPrompt
}
else {
    & $ClearHScript
}

Write-Host ""
Write-Host "G/H clear process completed."
