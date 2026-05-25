param()

$ErrorActionPreference = "Stop"
$TargetDriveRoot = "G:\"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UserResourceDir = Join-Path $ScriptDir "user_resources"
$SkipNames = @("System Volume Information")
$TargetExtensions = @("pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "txt", "rtf", "odt", "pages", "numbers", "key")
$FilesPerExtension = 3
$TemplateRequiredExtensions = @("doc", "docx", "xls", "xlsx", "ppt", "pptx")

function Normalize-PathToken {
    param([string]$PathText)
    $token = $PathText.Replace(":\", "").Replace(":", "").Replace("\", "_").Replace("/", "_")
    if ([string]::IsNullOrWhiteSpace($token)) {
        return "ROOT"
    }
    return $token
}

function New-PlaceholderFile {
    param(
        [string]$Path,
        [string]$Extension,
        [string]$TitleText
    )

    $ext = $Extension.ToLowerInvariant()
    switch ($ext) {
        "txt" {
            [System.IO.File]::WriteAllText($Path, "Generated TXT`r`n$TitleText`r`n", [System.Text.Encoding]::UTF8)
            return
        }
        "rtf" {
            $rtf = "{\rtf1\ansi\deff0 {\fonttbl {\f0 Calibri;}}\f0\fs24 Generated RTF\par " + $TitleText.Replace("\", "\\") + "\par }"
            [System.IO.File]::WriteAllText($Path, $rtf, [System.Text.Encoding]::ASCII)
            return
        }
        "pdf" {
            $pdf = @"
%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << >> >>
endobj
4 0 obj
<< /Length 65 >>
stream
BT /F1 12 Tf 72 720 Td (Generated PDF - $TitleText) Tj ET
endstream
endobj
xref
0 5
0000000000 65535 f
0000000010 00000 n
0000000060 00000 n
0000000120 00000 n
0000000216 00000 n
trailer
<< /Root 1 0 R /Size 5 >>
startxref
340
%%EOF
"@
            [System.IO.File]::WriteAllText($Path, $pdf, [System.Text.Encoding]::ASCII)
            return
        }
        default {
            [System.IO.File]::WriteAllText($Path, "Generated .$ext placeholder`r`n$TitleText`r`n", [System.Text.Encoding]::UTF8)
            return
        }
    }
}

function Clear-TargetDriveRoot {
    param([string]$DriveRoot)

    Write-Host "Pre-step: clearing existing items under $DriveRoot"
    $items = Get-ChildItem -LiteralPath $DriveRoot -Force -ErrorAction SilentlyContinue
    if (-not $items) {
        Write-Host "Drive root is already empty: $DriveRoot"
        return
    }

    foreach ($item in $items) {
        if ($SkipNames -contains $item.Name) {
            Write-Host "Skipped system folder: $($item.FullName)"
            continue
        }
        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
            Write-Host "Removed: $($item.FullName)"
        }
        catch {
            Write-Warning "Failed to remove: $($item.FullName) - $($_.Exception.Message)"
        }
    }
}

if (-not (Test-Path -LiteralPath $UserResourceDir -PathType Container)) {
    New-Item -Path $UserResourceDir -ItemType Directory -Force | Out-Null
    Write-Host "Template folder created (optional): $UserResourceDir"
}

$runTime = Get-Date -Format "yyyyMMdd_HHmmss"
$pathToken = Normalize-PathToken -PathText "G:"

if (-not (Test-Path -LiteralPath $TargetDriveRoot)) {
    Write-Error "Target drive not found: $TargetDriveRoot"
    exit 1
}

Clear-TargetDriveRoot -DriveRoot $TargetDriveRoot

Write-Host "Generating files to $TargetDriveRoot (single level only)"
Write-Host "Target formats: $($TargetExtensions -join ', ')"

$templateFiles = @()
if (Test-Path -LiteralPath $UserResourceDir -PathType Container) {
    $templateFiles = Get-ChildItem -LiteralPath $UserResourceDir -File -Recurse -Force -ErrorAction SilentlyContinue
}

$templatesByExt = @{}
foreach ($ext in $TargetExtensions) {
    $templatesByExt[$ext] = @()
}

foreach ($f in $templateFiles) {
    $ext = $f.Extension.TrimStart(".").ToLowerInvariant()
    if ($templatesByExt.ContainsKey($ext)) {
        $templatesByExt[$ext] += $f
    }
}

$globalIndex = 1
foreach ($ext in $TargetExtensions) {
    $extTemplates = $templatesByExt[$ext]
    $needsTemplate = $TemplateRequiredExtensions -contains $ext

    if ($needsTemplate -and $extTemplates.Count -eq 0) {
        Write-Warning "Skipped .$ext generation: no .$ext template found in $UserResourceDir"
        continue
    }

    for ($i = 1; $i -le $FilesPerExtension; $i++) {
        $baseName = "AUTO_{0}_{1:00}" -f $ext.ToUpperInvariant(), $i
        $uniqueName = "{0}_TIME_{1}_PATH_{2}_{3:000}.{4}" -f $baseName, $runTime, $pathToken, $globalIndex, $ext
        $targetPath = Join-Path $TargetDriveRoot $uniqueName

        if ($extTemplates.Count -gt 0) {
            $template = $extTemplates[($i - 1) % $extTemplates.Count]
            Copy-Item -LiteralPath $template.FullName -Destination $targetPath -Force
        }
        else {
            $titleText = "TIME=$runTime PATH=G: TYPE=$ext INDEX=$globalIndex"
            New-PlaceholderFile -Path $targetPath -Extension $ext -TitleText $titleText
        }

        $globalIndex++
    }
}

Write-Host "Done generation on G: ; total files created: $(($TargetExtensions.Count) * $FilesPerExtension)"
Write-Host "Naming pattern: <name>_TIME_<yyyyMMdd_HHmmss>_PATH_G_<index>.<ext>"
