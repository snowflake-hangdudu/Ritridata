param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$RootFolderName = "scan-allfiles-test"

Add-Type -AssemblyName Microsoft.VisualBasic

function Send-ToRecycleBin {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $Path,
        [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
        [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    )
}

function Remove-Permanent {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

foreach ($drive in $Drives) {
    if (-not (Test-Path -LiteralPath $drive)) {
        Write-Warning "Drive not found, skipped: $drive"
        continue
    }

    $root = Join-Path $drive $RootFolderName
    if (-not (Test-Path -LiteralPath $root)) {
        Write-Warning "Test root not found, run generator first: $root"
        continue
    }

    Write-Host "Mutating resources on $root"

    # 1) Send selected files to recycle bin
    $toRecycle = @(
        (Join-Path $root "docs\allfiles_doc_01.txt"),
        (Join-Path $root "images\allfiles_img_01.jpg")
    )
    foreach ($f in $toRecycle) {
        if (Test-Path -LiteralPath $f) {
            Send-ToRecycleBin -Path $f
            Write-Host "Moved to recycle bin: $f"
        }
    }

    # 2) Permanent delete selected files
    $toPermanentDelete = @(
        (Join-Path $root "videos\allfiles_vid_01.mp4"),
        (Join-Path $root "archives\allfiles_zip_01.zip")
    )
    foreach ($f in $toPermanentDelete) {
        if (Test-Path -LiteralPath $f) {
            Remove-Permanent -Path $f
            Write-Host "Permanently deleted: $f"
        }
    }

    # 3) Keep hidden sample as hidden and update timestamp
    $hiddenFile = Join-Path $root "hidden\allfiles_hidden.txt"
    if (Test-Path -LiteralPath $hiddenFile) {
        (Get-Item -LiteralPath $hiddenFile -Force).Attributes = "Hidden"
        (Get-Item -LiteralPath $hiddenFile -Force).LastWriteTime = (Get-Date).AddDays(-7)
        Write-Host "Refreshed hidden file metadata: $hiddenFile"
    }

    # 4) Keep one existing file for exclusion check
    $liveFile = Join-Path $root "live_existing_keep.txt"
    "This file should remain existing and should NOT appear in lost-file scan results." |
        Set-Content -LiteralPath $liveFile -Encoding UTF8
    Write-Host "Kept one existing control file: $liveFile"

    # 5) Write scenario note
    $note = Join-Path $root "mutation_note.txt"
    @(
        "Mutation time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        "Recycled: docs/allfiles_doc_01.txt, images/allfiles_img_01.jpg"
        "Permanent deleted: videos/allfiles_vid_01.mp4, archives/allfiles_zip_01.zip"
        "Hidden sample refreshed: hidden/allfiles_hidden.txt"
        "Existing control file kept: live_existing_keep.txt (should not appear in results)"
    ) | Set-Content -LiteralPath $note -Encoding UTF8
    Write-Host "Wrote mutation note: $note"
    Write-Host ""
}

Write-Host "Done. Mutation scenarios prepared."
