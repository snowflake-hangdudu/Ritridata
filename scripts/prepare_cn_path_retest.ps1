param()

$ErrorActionPreference = "Stop"
$Drive = "G:\"
$Root = Join-Path $Drive "cn-path-retest"

Add-Type -AssemblyName Microsoft.VisualBasic

function Ensure-Dir {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Write-TextFile {
    param(
        [string]$Path,
        [string]$Content
    )
    [System.IO.File]::WriteAllText($Path, $Content, [System.Text.Encoding]::UTF8)
}

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

if (-not (Test-Path -LiteralPath $Drive)) {
    Write-Error "Drive not found: $Drive"
    exit 1
}

# Clean old retest data
if (Test-Path -LiteralPath $Root) {
    Remove-Item -LiteralPath $Root -Recurse -Force -ErrorAction SilentlyContinue
}

Ensure-Dir $Root

$groupA = Join-Path $Root "group-A-en-path"
$groupB = Join-Path $Root "group-B-en-path"
$groupC = Join-Path $Root "组C-中文路径"
$groupD = Join-Path $Root "组D-中文路径"

Ensure-Dir $groupA
Ensure-Dir $groupB
Ensure-Dir $groupC
Ensure-Dir $groupD

$files = @(
    @{ Group = "A"; Path = (Join-Path $groupA "en_name_01.txt"); Type = "recycle" },
    @{ Group = "A"; Path = (Join-Path $groupA "en_name_02.txt"); Type = "permanent" },
    @{ Group = "B"; Path = (Join-Path $groupB "中文文件名_01.txt"); Type = "recycle" },
    @{ Group = "B"; Path = (Join-Path $groupB "中文文件名_02.txt"); Type = "permanent" },
    @{ Group = "C"; Path = (Join-Path $groupC "en_name_01.txt"); Type = "recycle" },
    @{ Group = "C"; Path = (Join-Path $groupC "en_name_02.txt"); Type = "permanent" },
    @{ Group = "D"; Path = (Join-Path $groupD "中文文件名_01.txt"); Type = "recycle" },
    @{ Group = "D"; Path = (Join-Path $groupD "中文文件名_02.txt"); Type = "permanent" }
)

foreach ($f in $files) {
    Write-TextFile -Path $f.Path -Content ("group={0}; delete_type={1}; created={2}" -f $f.Group, $f.Type, (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
}

$recycled = @()
$permanent = @()

foreach ($f in $files) {
    if ($f.Type -eq "recycle") {
        if (Send-ToRecycleBin -Path $f.Path) {
            $recycled += $f.Path
        }
    }
    else {
        if (Test-Path -LiteralPath $f.Path) {
            Remove-Item -LiteralPath $f.Path -Force -ErrorAction SilentlyContinue
            $permanent += $f.Path
        }
    }
}

# Keep one existing control file for exclusion check
$liveFile = Join-Path $Root "live_existing_keep.txt"
Write-TextFile -Path $liveFile -Content "existing control file: should NOT appear in lost-file scan results"

$note = Join-Path $Root "mutation_cn_path_note.txt"
@(
    "Case: CN_PATH_RETEST"
    "Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    "Recycled files:"
    ($recycled | ForEach-Object { "  - $_" })
    "Permanent deleted files:"
    ($permanent | ForEach-Object { "  - $_" })
    "Existing keep file:"
    "  - $liveFile"
) | Set-Content -LiteralPath $note -Encoding UTF8

Write-Host "Prepared CN path retest data on $Root"
Write-Host "Recycled: $($recycled.Count) ; Permanent deleted: $($permanent.Count)"
Write-Host "Note: $note"
