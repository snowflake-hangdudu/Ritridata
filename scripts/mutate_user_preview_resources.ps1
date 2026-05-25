param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$RootFolderName = "scan-allfiles-test"

Add-Type -AssemblyName Microsoft.VisualBasic

function Send-ToRecycleBin {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return $false
    }
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    )
    return $true
}

foreach ($drive in $Drives) {
    if (-not (Test-Path -LiteralPath $drive)) {
        Write-Warning "Drive not found, skipped: $drive"
        continue
    }

    $root = Join-Path $drive $RootFolderName
    $sourceDir = Join-Path $root "source"
    if (-not (Test-Path -LiteralPath $sourceDir -PathType Container)) {
        Write-Warning "Source not found, run stage script first: $sourceDir"
        continue
    }

    Write-Host "Mutating resources on $root"
    $files = Get-ChildItem -LiteralPath $sourceDir -File -Force | Sort-Object Name
    if (-not $files) {
        Write-Warning "No files in source folder: $sourceDir"
        continue
    }

    # Recycle first 2 files
    $recycled = @()
    $files | Select-Object -First 2 | ForEach-Object {
        if (Send-ToRecycleBin -Path $_.FullName) {
            $recycled += $_.Name
            Write-Host "Moved to recycle bin: $($_.FullName)"
        }
    }

    # Permanently delete next 2 files
    $deleted = @()
    $files | Select-Object -Skip 2 -First 2 | ForEach-Object {
        if (Test-Path -LiteralPath $_.FullName) {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            $deleted += $_.Name
            Write-Host "Permanently deleted: $($_.FullName)"
        }
    }

    # Keep one existing control file untouched
    $liveFile = Join-Path $root "live_existing_keep.txt"
    if (Test-Path -LiteralPath $liveFile) {
        Write-Host "Kept existing control file: $liveFile"
    }

    $note = Join-Path $root "mutation_note.txt"
    @(
        "Mutation time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Recycled files: $($recycled -join ', ')"
        "Permanent deleted files: $($deleted -join ', ')"
        "Existing control file kept: live_existing_keep.txt (should not appear in results)"
    ) | Set-Content -LiteralPath $note -Encoding UTF8

    Write-Host "Wrote mutation note: $note"
    Write-Host ""
}

Write-Host "Mutation completed."
