param()

$ErrorActionPreference = "Stop"
$Drives = @("G:\", "H:\")
$TargetRootName = "scan-allfiles-test"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UserResourceDir = Join-Path $ScriptDir "user_resources"

function Ensure-Dir {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

if (-not (Test-Path -LiteralPath $UserResourceDir -PathType Container)) {
    Ensure-Dir $UserResourceDir
    Write-Warning "Resource folder not found. Put your files here first: $UserResourceDir"
    exit 1
}

$resourceFiles = Get-ChildItem -LiteralPath $UserResourceDir -File -Recurse -Force
if (-not $resourceFiles) {
    Write-Warning "No files found in: $UserResourceDir"
    exit 1
}

$runTag = "{0}_{1}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), (Get-Random -Minimum 1000 -Maximum 9999)

foreach ($drive in $Drives) {
    if (-not (Test-Path -LiteralPath $drive)) {
        Write-Warning "Drive not found, skipped: $drive"
        continue
    }

    $root = Join-Path $drive $TargetRootName
    $sourceDir = Join-Path $root "source"
    $deepDir = Join-Path $root "deep\nested\level1\level2"
    Ensure-Dir $sourceDir
    Ensure-Dir $deepDir

    Write-Host "Staging resources to $root"

    $index = 1
    foreach ($file in $resourceFiles) {
        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $ext = [System.IO.Path]::GetExtension($file.Name)
        $driveLabel = $drive.TrimEnd("\").TrimEnd(":")
        $uniqueName = "{0}_{1}_{2}_{3:000}{4}" -f $nameWithoutExt, $driveLabel, $runTag, $index, $ext
        $target = Join-Path $sourceDir $uniqueName
        Copy-Item -LiteralPath $file.FullName -Destination $target -Force
        $index++
    }

    # Keep one same-name file in deep path for path display checks.
    if ($resourceFiles.Count -gt 0) {
        $sample = $resourceFiles[0]
        $sampleNameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($sample.Name)
        $sampleExt = [System.IO.Path]::GetExtension($sample.Name)
        $driveLabel = $drive.TrimEnd("\").TrimEnd(":")
        $deepUniqueName = "deep_copy_{0}_{1}_{2}{3}" -f $sampleNameWithoutExt, $driveLabel, $runTag, $sampleExt
        $deepTarget = Join-Path $deepDir $deepUniqueName
        Copy-Item -LiteralPath $sample.FullName -Destination $deepTarget -Force
    }

    # Existing control file: should remain on disk and not be shown in lost-file results.
    $liveFile = Join-Path $root "live_existing_keep.txt"
    "existing control file: should NOT appear in lost-file scan results" |
        Set-Content -LiteralPath $liveFile -Encoding UTF8

    # Hidden file sample
    $hiddenDir = Join-Path $root "hidden"
    Ensure-Dir $hiddenDir
    $hiddenFile = Join-Path $hiddenDir "hidden_keep_for_scan.txt"
    "hidden sample for scan range check" | Set-Content -LiteralPath $hiddenFile -Encoding UTF8
    (Get-Item -LiteralPath $hiddenDir -Force).Attributes = "Directory,Hidden"
    (Get-Item -LiteralPath $hiddenFile -Force).Attributes = "Hidden"

    # Manifest
    $manifest = Join-Path $root "manifest.csv"
    Get-ChildItem -LiteralPath $root -Recurse -Force -File |
        Select-Object FullName, Length, LastWriteTime, Attributes |
        Export-Csv -LiteralPath $manifest -NoTypeInformation -Encoding UTF8

    Write-Host "Done staging on $drive ; total source files copied: $($resourceFiles.Count)"
    Write-Host ""
}

Write-Host "All done."
