param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$RootFolderName = "scan-allfiles-test"

function Ensure-Dir {
    param([string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Container) {
        return
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
    New-Item -Path $Path -ItemType Directory -Force | Out-Null
}

function Write-TextFile {
    param(
        [string]$Path,
        [string]$Content
    )
    if (Test-Path -LiteralPath $Path) {
        try {
            (Get-Item -LiteralPath $Path -Force).Attributes = "Normal"
        }
        catch {
            Write-Warning "Failed to normalize attributes before write: $Path"
        }
    }
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
}

function Write-RandomBinary {
    param(
        [string]$Path,
        [int]$SizeKB
    )
    $bytes = New-Object byte[] ($SizeKB * 1024)
    (New-Object System.Random).NextBytes($bytes)
    [System.IO.File]::WriteAllBytes($Path, $bytes)
}

function Set-RandomLastWriteTime {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Warning "Skip timestamp update, path not found: $Path"
        return
    }
    $daysBack = Get-Random -Minimum 1 -Maximum 45
    $minutesBack = Get-Random -Minimum 1 -Maximum 1440
    $time = (Get-Date).AddDays(-$daysBack).AddMinutes(-$minutesBack)
    (Get-Item -LiteralPath $Path -Force).LastWriteTime = $time
}

foreach ($drive in $Drives) {
    if (-not (Test-Path -LiteralPath $drive)) {
        Write-Warning "Drive not found, skipped: $drive"
        continue
    }

    $root = Join-Path $drive $RootFolderName
    $docs = Join-Path $root "docs"
    $images = Join-Path $root "images"
    $videos = Join-Path $root "videos"
    $archives = Join-Path $root "archives"
    $hidden = Join-Path $root "hidden"
    $deep = Join-Path $root "deep\nested\level1\level2\level3\level4"
    $zero = Join-Path $root "zero-byte"
    $dupA = Join-Path $root "dup\a"
    $dupB = Join-Path $root "dup\b"

    Ensure-Dir $root
    Ensure-Dir $docs
    Ensure-Dir $images
    Ensure-Dir $videos
    Ensure-Dir $archives
    Ensure-Dir $hidden
    Ensure-Dir $deep
    Ensure-Dir $zero
    Ensure-Dir $dupA
    Ensure-Dir $dupB

    Write-Host "Generating resources on $root"

    # docs
    1..5 | ForEach-Object {
        $file = Join-Path $docs ("allfiles_doc_{0:00}.txt" -f $_)
        Write-TextFile -Path $file -Content "all files test doc $_ on $drive"
        Set-RandomLastWriteTime -Path $file
    }

    # images/videos (binary placeholders)
    1..3 | ForEach-Object {
        $img = Join-Path $images ("allfiles_img_{0:00}.jpg" -f $_)
        Write-RandomBinary -Path $img -SizeKB 64
        Set-RandomLastWriteTime -Path $img
    }
    1..2 | ForEach-Object {
        $vid = Join-Path $videos ("allfiles_vid_{0:00}.mp4" -f $_)
        Write-RandomBinary -Path $vid -SizeKB 256
        Set-RandomLastWriteTime -Path $vid
    }

    # archives
    1..2 | ForEach-Object {
        $zip = Join-Path $archives ("allfiles_zip_{0:00}.zip" -f $_)
        Write-RandomBinary -Path $zip -SizeKB 128
        Set-RandomLastWriteTime -Path $zip
    }

    # hidden file/folder
    Ensure-Dir $hidden
    $hiddenFile = Join-Path $hidden "allfiles_hidden.txt"
    Write-TextFile -Path $hiddenFile -Content "hidden file for all files test"
    (Get-Item -LiteralPath $hidden -Force).Attributes = "Directory,Hidden"
    (Get-Item -LiteralPath $hiddenFile -Force).Attributes = "Hidden"
    Set-RandomLastWriteTime -Path $hiddenFile

    # deep path file
    $deepFile = Join-Path $deep "allfiles_deep_path.log"
    Write-TextFile -Path $deepFile -Content "deep path file"
    Set-RandomLastWriteTime -Path $deepFile

    # zero byte file
    $zeroFile = Join-Path $zero "allfiles_zero.dat"
    New-Item -Path $zeroFile -ItemType File -Force | Out-Null
    Set-RandomLastWriteTime -Path $zeroFile

    # duplicate name in different dirs
    $dupName = "allfiles_same_name.doc"
    $dupFileA = Join-Path $dupA $dupName
    $dupFileB = Join-Path $dupB $dupName
    Write-TextFile -Path $dupFileA -Content "same name file A"
    Write-TextFile -Path $dupFileB -Content "same name file B"
    Set-RandomLastWriteTime -Path $dupFileA
    Set-RandomLastWriteTime -Path $dupFileB

    # existing control file (must stay on disk for exclusion check)
    $liveFile = Join-Path $root "live_existing_keep.txt"
    Write-TextFile -Path $liveFile -Content "existing control file: should NOT appear in lost-file scan results"
    Set-RandomLastWriteTime -Path $liveFile

    # manifest
    $manifest = Join-Path $root "manifest.csv"
    Get-ChildItem -LiteralPath $root -Recurse -Force -File |
        Select-Object FullName, Length, LastWriteTime, Attributes |
        Export-Csv -LiteralPath $manifest -NoTypeInformation -Encoding UTF8

    Write-Host "Generated manifest: $manifest"
    Write-Host ""
}

Write-Host "Done. Resource generation completed."
