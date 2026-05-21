param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$SkipNames = @("System Volume Information")

function Clear-DriveContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveRoot
    )

    if (-not (Test-Path -LiteralPath $DriveRoot)) {
        Write-Warning "Drive not found: $DriveRoot"
        return
    }

    Write-Host "=== Clearing drive: $DriveRoot ==="
    $items = Get-ChildItem -LiteralPath $DriveRoot -Force -ErrorAction SilentlyContinue
    if (-not $items) {
        Write-Host "Already empty: $DriveRoot"
        return
    }

    foreach ($item in $items) {
        if ($SkipNames -contains $item.Name) {
            Write-Host "Skipped system folder: $($item.FullName)"
            continue
        }

        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Deleted: $($item.FullName)"
        }
        catch {
            Write-Warning "Delete failed: $($item.FullName) - $($_.Exception.Message)"
        }
    }
}

foreach ($drive in $Drives) {
    Clear-DriveContent -DriveRoot $drive
}

Write-Host ""
Write-Host "Done. Normal delete completed for G/H."
