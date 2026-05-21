param(
    [switch]$NoPrompt
)

$ErrorActionPreference = "Stop"
$DriveRoot = "H:\"
$ConfirmToken = "CLEAR-H"

if (-not (Test-Path -LiteralPath $DriveRoot)) {
    Write-Error "Drive not found: $DriveRoot"
    exit 1
}

Write-Host "Target drive: $DriveRoot"
Write-Host "Delete all files and folders under $DriveRoot (no formatting)."

if (-not $NoPrompt) {
    $inputToken = Read-Host "Type $ConfirmToken to continue"
    if ($inputToken -ne $ConfirmToken) {
        Write-Host "Confirmation token mismatch. Cancelled."
        exit 1
    }
}

$items = Get-ChildItem -LiteralPath $DriveRoot -Force -ErrorAction SilentlyContinue
if (-not $items) {
    Write-Host "$DriveRoot is already empty."
    exit 0
}

foreach ($item in $items) {
    try {
        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
        Write-Host "Deleted: $($item.FullName)"
    }
    catch {
        Write-Warning "Delete failed: $($item.FullName) - $($_.Exception.Message)"
    }
}

Write-Host "Clear completed: $DriveRoot"
