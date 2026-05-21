param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$SkipNames = @("System Volume Information")
$DiskFullHResult = -2147024784

function Invoke-QuickFreeSpaceWipe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveRoot
    )

    $tempFile = Join-Path $DriveRoot ".__fast_wipe.tmp"
    $bufferSize = 8MB
    $buffer = New-Object byte[] $bufferSize
    [long]$writtenBytes = 0
    $stream = $null

    Write-Host "Start quick free-space wipe on $DriveRoot (single pass) ..."

    try {
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }

        $stream = [System.IO.File]::Open(
            $tempFile,
            [System.IO.FileMode]::CreateNew,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None
        )

        while ($true) {
            try {
                $stream.Write($buffer, 0, $buffer.Length)
                $writtenBytes += $buffer.Length
                if (($writtenBytes % 536870912) -eq 0) {
                    $writtenGb = [Math]::Round($writtenBytes / 1GB, 2)
                    Write-Host "  Wiped ~${writtenGb} GB free space on $DriveRoot"
                }
            }
            catch [System.IO.IOException] {
                if ($_.Exception.HResult -eq $DiskFullHResult) {
                    break
                }
                throw
            }
        }

        $stream.Flush($true)
        $stream.Dispose()
        $stream = $null

        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction Stop
        }

        Write-Host "Quick free-space wipe completed: $DriveRoot"
        return $true
    }
    catch {
        Write-Warning "Quick wipe failed on $DriveRoot - $($_.Exception.Message)"
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
        if (Test-Path -LiteralPath $tempFile) {
            Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

function Clear-DriveContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DriveRoot
    )

    if (-not (Test-Path -LiteralPath $DriveRoot)) {
        Write-Warning "Drive not found: $DriveRoot"
        return $false
    }

    Write-Host "=== Clearing drive: $DriveRoot ==="
    $items = Get-ChildItem -LiteralPath $DriveRoot -Force -ErrorAction SilentlyContinue
    if ($items) {
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
    else {
        Write-Host "Already empty: $DriveRoot"
    }

    return (Invoke-QuickFreeSpaceWipe -DriveRoot $DriveRoot)
}

$allOk = $true
foreach ($drive in $Drives) {
    $ok = Clear-DriveContent -DriveRoot $drive
    if (-not $ok) {
        $allOk = $false
    }
    Write-Host ""
}

if ($allOk) {
    Write-Host "Done. Fast secure delete (single-pass free-space wipe) completed for G/H."
}
else {
    Write-Warning "Completed with warnings. Check messages above."
}
