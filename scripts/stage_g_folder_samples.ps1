param(
    [switch]$ClearExistingNest
)

$ErrorActionPreference = "Stop"
$TargetDriveRoot = "G:\"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UserResourceDir = Join-Path $ScriptDir "user_resources"

$FolderCases = @(
    @{ Depth = 1;  Id = "L01"; RelPath = "nest_L1" },
    @{ Depth = 3;  Id = "L03"; RelPath = "nest_L1\nest_L2\nest_L3" },
    @{ Depth = 5;  Id = "L05"; RelPath = "nest_L1\nest_L2\nest_L3\nest_L4\nest_L5" },
    @{ Depth = 10; Id = "L10"; RelPath = "nest_L1\nest_L2\nest_L3\nest_L4\nest_L5\nest_L6\nest_L7\nest_L8\nest_L9\nest_L10" }
)

function Get-TemplateFile {
    $candidates = @()
    if (Test-Path -LiteralPath $UserResourceDir -PathType Container) {
        $candidates = Get-ChildItem -LiteralPath $UserResourceDir -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in @(".pptx", ".xlsx", ".jpg", ".pdf", ".txt") }
    }
    if ($candidates.Count -eq 0) {
        return $null
    }
    $pptx = $candidates | Where-Object { $_.Extension -eq ".pptx" } | Select-Object -First 1
    if ($pptx) { return $pptx }
    return ($candidates | Select-Object -First 1)
}

if (-not (Test-Path -LiteralPath $TargetDriveRoot)) {
    Write-Error "Target drive not found: $TargetDriveRoot"
    exit 1
}

$template = Get-TemplateFile
if (-not $template) {
    Write-Error "No template found under $UserResourceDir (need at least one pptx/xlsx/jpg/pdf/txt)."
    exit 1
}

$ext = $template.Extension
Write-Host "Using template: $($template.FullName)"

if ($ClearExistingNest -and (Test-Path -LiteralPath (Join-Path $TargetDriveRoot "nest_L1") -PathType Container)) {
    Remove-Item -LiteralPath (Join-Path $TargetDriveRoot "nest_L1") -Recurse -Force
    Write-Host "Removed existing G:\nest_L1\ tree"
}

foreach ($case in $FolderCases) {
    $dir = Join-Path $TargetDriveRoot $case.RelPath
    New-Item -Path $dir -ItemType Directory -Force | Out-Null
    $fileName = "folder_test_$($case.Id)_sample$ext"
    $target = Join-Path $dir $fileName
    Copy-Item -LiteralPath $template.FullName -Destination $target -Force
    Write-Host "Created: $target"
}

Write-Host "Done. Folder depth samples ready on G: (1/3/5/10 layers)."
