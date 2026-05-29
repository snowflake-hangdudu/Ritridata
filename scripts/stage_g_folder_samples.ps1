param(
    [string]$TargetDriveRoot = "G:\",
    [switch]$ClearExistingNest
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UserResourceDir = Join-Path $ScriptDir "user_resources"
$SkipTemplateNames = @("user_resources.zip")

# Case 6: nested path depth 1 / 3 / 5 / 10 (one file per depth)
$FolderCases = @(
    @{ Depth = 1;  Id = "L01"; RelPath = "nest_L1" },
    @{ Depth = 3;  Id = "L03"; RelPath = "nest_L1\nest_L2\nest_L3" },
    @{ Depth = 5;  Id = "L05"; RelPath = "nest_L1\nest_L2\nest_L3\nest_L4\nest_L5" },
    @{ Depth = 10; Id = "L10"; RelPath = "nest_L1\nest_L2\nest_L3\nest_L4\nest_L5\nest_L6\nest_L7\nest_L8\nest_L9\nest_L10" }
)

function Get-TemplateFile {
    if (-not (Test-Path -LiteralPath $UserResourceDir -PathType Container)) {
        return $null
    }
    $preferExt = @(".pptx", ".docx", ".xlsx", ".jpg", ".pdf", ".txt")
  $candidates = Get-ChildItem -LiteralPath $UserResourceDir -File -Recurse -Force -ErrorAction SilentlyContinue |
        Where-Object { $SkipTemplateNames -notcontains $_.Name }

    foreach ($ext in $preferExt) {
        $hit = $candidates | Where-Object { $_.Extension -eq $ext } | Select-Object -First 1
        if ($hit) { return $hit }
    }
    return ($candidates | Select-Object -First 1)
}

if (-not (Test-Path -LiteralPath $TargetDriveRoot)) {
    Write-Error "Target drive not found: $TargetDriveRoot"
    exit 1
}

$template = Get-TemplateFile
if (-not $template) {
    Write-Error "No template under $UserResourceDir (need pptx/docx/xlsx/jpg/pdf/txt)."
    exit 1
}

$runTime = Get-Date -Format "yyMMdd_HHmm"
Write-Host "Target: $TargetDriveRoot"
Write-Host "Template: $($template.FullName)"

$nestRoot = Join-Path $TargetDriveRoot "nest_L1"
if ($ClearExistingNest -and (Test-Path -LiteralPath $nestRoot -PathType Container)) {
    Remove-Item -LiteralPath $nestRoot -Recurse -Force
    Write-Host "Removed: $nestRoot"
}

foreach ($case in $FolderCases) {
    $dir = Join-Path $TargetDriveRoot $case.RelPath
    [void][System.IO.Directory]::CreateDirectory($dir)
    $fileName = "G_nest_$($case.Id)_$runTime$($template.Extension)"
    $target = Join-Path $dir $fileName
    Copy-Item -LiteralPath $template.FullName -Destination $target -Force
    Write-Host "Created ($($case.Depth) layers): $target"
}

Write-Host "Done. Nested samples: 1/3/5/10 layer paths under G:\nest_L1\..."
Write-Host "Naming: G_nest_L<depth>_<yyMMdd_HHmm>.<ext>"
