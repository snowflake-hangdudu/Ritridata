param(
    [string]$TargetDriveRoot = "G:\",
    [string[]]$Sources = @(),
    [int]$FilesPerFormat = 3,
    [switch]$ClearImageFolder
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PyScript = Join-Path $ScriptDir "stage_g_image_formats.py"

if (-not (Test-Path -LiteralPath $PyScript -PathType Leaf)) {
    Write-Error "Missing script: $PyScript"
    exit 1
}

$pyArgs = @(
    $PyScript,
    "--drive", $TargetDriveRoot,
    "--per-format", "$FilesPerFormat"
)
if ($ClearImageFolder) {
    $pyArgs += "--clear-image-folder"
}
if ($Sources -and $Sources.Count -gt 0) {
    $pyArgs += "--sources"
    $pyArgs += $Sources
}

Write-Host "Running: python $($pyArgs -join ' ')"
& python @pyArgs
exit $LASTEXITCODE
