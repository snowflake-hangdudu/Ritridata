param(
    [switch]$SkipOfficeCom,
  # OpenXmlOnly: skip legacy doc/xls/ppt (product preview prefers docx/xlsx/pptx)
    [switch]$OpenXmlOnly
)

$ErrorActionPreference = "Stop"
$TargetDriveRoot = "G:\"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$UserResourceDir = Join-Path $ScriptDir "user_resources"
$SkipNames = @("System Volume Information")
$SkipTemplateNames = @("user_resources.zip")
$FilesPerExtension = 3

function Get-CategoryFolderName {
    param([string]$Extension)
    $ext = $Extension.ToLowerInvariant()
    switch ($ext) {
        { $_ -in @("jpg", "jpeg", "png", "gif", "bmp", "webp") } {
            return [string]::Concat([char]0x56FE, [char]0x7247)
        }
        { $_ -in @("doc", "docx", "pdf", "odt", "pages") } {
            return [string]::Concat([char]0x6587, [char]0x6863)
        }
        { $_ -in @("xls", "xlsx", "numbers") } {
            return [string]::Concat([char]0x8868, [char]0x683C)
        }
        { $_ -in @("ppt", "pptx", "key") } {
            return [string]::Concat([char]0x6F14, [char]0x793A)
        }
        { $_ -in @("txt", "rtf") } {
            return [string]::Concat([char]0x6587, [char]0x672C)
        }
        default {
            return [string]::Concat([char]0x5176, [char]0x4ED6)
        }
    }
}

# Product document filters: PDF/DOC/DOCX/XLS/XLSX/PPT/PPTX + TXT/RTF/ODT/PAGES/NUMBERS/KEY
$ProductDocumentExtensions = @(
    "pdf", "doc", "docx",
    "xls", "xlsx",
    "ppt", "pptx",
    "txt", "rtf",
    "odt", "pages", "numbers", "key"
)

$ImageExtensions = @("jpg", "png", "gif", "bmp", "webp")
$NativeTextExtensions = @("txt", "rtf")
$TemplateOnlyExtensions = @("odt", "pages", "numbers", "key")

$TargetExtensions = $ImageExtensions + $ProductDocumentExtensions
if ($OpenXmlOnly) {
    $TargetExtensions = $ImageExtensions + @(
        "pdf", "docx", "xlsx", "pptx",
        "txt", "rtf", "odt", "pages", "numbers", "key"
    )
}

$WordExtensions = if ($OpenXmlOnly) { @("docx", "pdf") } else { @("docx", "pdf", "doc") }
$ExcelExtensions = if ($OpenXmlOnly) { @("xlsx") } else { @("xlsx", "xls") }
$PowerPointExtensions = if ($OpenXmlOnly) { @("pptx") } else { @("pptx", "ppt") }

function Get-CategoryOutputDir {
    param(
        [string]$DriveRoot,
        [string]$Extension
    )
    $folderName = Get-CategoryFolderName -Extension $Extension
    $dir = Join-Path $DriveRoot $folderName
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        [void][System.IO.Directory]::CreateDirectory($dir)
        Write-Host "Created folder: $dir"
    }
    return $dir
}

function New-UniqueFileName {
    param(
        [string]$Extension,
        [int]$Index,
        [string]$RunTime
    )
    return ("G_{0}_{1:000}_{2}.{0}" -f $Extension, $Index, $RunTime)
}

function New-RealTextFile {
    param(
        [string]$Path,
        [string]$Extension,
        [string]$Label
    )

    switch ($Extension.ToLowerInvariant()) {
        "txt" {
            $utf8 = [System.Text.Encoding]::UTF8
            $body = $utf8.GetString([byte[]](
                0x52, 0x69, 0x74, 0x72, 0x69, 0x64, 0x61, 0x74, 0x61, 0x20, 0x47, 0xE7, 0x9B, 0x98, 0xE6, 0xB5, 0x8B, 0xE8, 0xAF, 0x95, 0xE6, 0xA0, 0xB7, 0xE6, 0x9C, 0xAC, 0x0D, 0x0A,
                0xE6, 0x96, 0x87, 0xE4, 0xBB, 0xB6, 0x3A, 0x20
            )) + $Label + $utf8.GetString([byte[]](
                0x0D, 0x0A, 0xE8, 0xAF, 0xB4, 0xE6, 0x98, 0x8E, 0x3A, 0x20, 0xE6, 0x89, 0xAB, 0xE6, 0x8F, 0x8F, 0x2F, 0xE9, 0xA2, 0x84, 0xE8, 0xA7, 0x88, 0x2F, 0xE6, 0x81, 0xA2, 0xE5, 0xA4, 0x8D, 0xE6, 0xB5, 0x8B, 0xE8, 0xAF, 0x95, 0xE3, 0x80, 0x82, 0x0D, 0x0A,
                0xE6, 0x97, 0xB6, 0xE9, 0x97, 0xB4, 0x3A, 0x20
            )) + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            $utf8Bom = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllText($Path, $body, $utf8Bom)
        }
        "rtf" {
            $safe = $Label.Replace('\', '\\')
            $rtf = @"
{\rtf1\ansi\deff0{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
\f0\fs22 Ritridata G-drive preview test sample.\par
File: $safe\par
Open in WordPad or Word for preview validation.\par
}
"@
            [System.IO.File]::WriteAllText($Path, $rtf.Trim(), [System.Text.Encoding]::ASCII)
        }
    }
}

function Stop-OfficeProcess {
    param([string]$Name)
    Get-Process -Name $Name -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300
}

function Invoke-SafeQuit {
    param($App)
    if (-not $App) { return }
    try { $App.Quit() } catch { }
    try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($App) | Out-Null } catch { }
}

function Copy-FromTemplate {
    param(
        [string]$DriveRoot,
        [string]$Extension,
        [int]$Count,
        [ref]$Index,
        [string]$RunTime,
        [hashtable]$TemplatesByExt
    )
    if (-not $TemplatesByExt.ContainsKey($Extension) -or $TemplatesByExt[$Extension].Count -eq 0) {
        return $false
    }
    $outDir = Get-CategoryOutputDir -DriveRoot $DriveRoot -Extension $Extension
    $pool = @($TemplatesByExt[$Extension])
    for ($i = 1; $i -le $Count; $i++) {
        $uniqueName = New-UniqueFileName -Extension $Extension -Index $Index.Value -RunTime $RunTime
        $targetPath = Join-Path $outDir $uniqueName
        $tpl = $pool[($i - 1) % $pool.Count]
        Copy-Item -LiteralPath $tpl.FullName -Destination $targetPath -Force
        Write-Host "Created (template): $targetPath"
        $Index.Value++
    }
    return $true
}

function New-WordSamples {
    param(
        [string]$DriveRoot,
        [string[]]$Extensions,
        [int]$CountPerExt,
        [ref]$Index,
        [string]$RunTime
    )
    $wdFormatDocument = 0
    $wdFormatDocumentDefault = 16
    $wdFormatPDF = 17

    foreach ($ext in $Extensions) {
        $outDir = Get-CategoryOutputDir -DriveRoot $DriveRoot -Extension $ext
        for ($i = 1; $i -le $CountPerExt; $i++) {
            $uniqueName = New-UniqueFileName -Extension $ext -Index $Index.Value -RunTime $RunTime
            $targetPath = Join-Path $outDir $uniqueName
            $word = $null
            $doc = $null
            try {
                Stop-OfficeProcess "WINWORD"
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0
                $doc = $word.Documents.Add()
                $doc.Range().Text = "Ritridata G-drive test. File: $uniqueName"
                if ($ext -eq "pdf") {
                    $null = $doc.SaveAs2($targetPath, $wdFormatPDF)
                }
                elseif ($ext -eq "doc") {
                    $null = $doc.SaveAs2($targetPath, $wdFormatDocument)
                }
                else {
                    $null = $doc.SaveAs2($targetPath, $wdFormatDocumentDefault)
                }
                Write-Host "Created (Word): $targetPath"
                $Index.Value++
            }
            catch {
                Write-Warning "Word failed .$ext #$i : $($_.Exception.Message)"
            }
            finally {
                if ($doc) { try { $doc.Close($false) } catch { } }
                Invoke-SafeQuit $word
                Stop-OfficeProcess "WINWORD"
            }
        }
    }
}

function New-ExcelSamples {
    param(
        [string]$DriveRoot,
        [string[]]$Extensions,
        [int]$CountPerExt,
        [ref]$Index,
        [string]$RunTime
    )
    $excel = $null
    $xlExcel8 = 56
    $xlOpenXMLWorkbook = 51

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        foreach ($ext in $Extensions) {
            $outDir = Get-CategoryOutputDir -DriveRoot $DriveRoot -Extension $ext
            if ($ext -eq "xls") { $fmt = $xlExcel8 } else { $fmt = $xlOpenXMLWorkbook }
            for ($i = 1; $i -le $CountPerExt; $i++) {
                $uniqueName = New-UniqueFileName -Extension $ext -Index $Index.Value -RunTime $RunTime
                $targetPath = Join-Path $outDir $uniqueName
                $wb = $null
                try {
                    $wb = $excel.Workbooks.Add()
                    $wb.Worksheets.Item(1).Cells.Item(1, 1).Value2 = "Ritridata test $uniqueName"
                    $wb.SaveAs($targetPath, $fmt)
                    Write-Host "Created (Excel): $targetPath"
                    $Index.Value++
                }
                catch {
                    Write-Warning "Excel failed .$ext #$i : $($_.Exception.Message)"
                }
                finally {
                    if ($wb) {
                        try { $wb.Close($false) } catch { }
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Excel COM unavailable: $($_.Exception.Message)"
    }
    finally {
        Invoke-SafeQuit $excel
        Stop-OfficeProcess "EXCEL"
    }
}

function New-PowerPointSamples {
    param(
        [string]$DriveRoot,
        [string[]]$Extensions,
        [int]$CountPerExt,
        [ref]$Index,
        [string]$RunTime,
        [hashtable]$TemplatesByExt
    )
    $ppSaveAsPresentation = 1
    $ppSaveAsOpenXMLPresentation = 24
    $hasPptxTpl = $TemplatesByExt.ContainsKey("pptx") -and $TemplatesByExt["pptx"].Count -gt 0

    foreach ($ext in $Extensions) {
        if ($ext -eq "pptx" -and $hasPptxTpl) { continue }

        $pp = $null
        try {
            Stop-OfficeProcess "POWERPNT"
            $pp = New-Object -ComObject PowerPoint.Application
            $pp.Visible = 0
            $outDir = Get-CategoryOutputDir -DriveRoot $DriveRoot -Extension $ext
            $saveFmt = if ($ext -eq "ppt") { $ppSaveAsPresentation } else { $ppSaveAsOpenXMLPresentation }

            for ($i = 1; $i -le $CountPerExt; $i++) {
                $uniqueName = New-UniqueFileName -Extension $ext -Index $Index.Value -RunTime $RunTime
                $targetPath = Join-Path $outDir $uniqueName
                $pres = $null
                try {
                    if ($ext -eq "ppt" -and $hasPptxTpl) {
                        $tpl = $TemplatesByExt["pptx"][($i - 1) % $TemplatesByExt["pptx"].Count]
                        $pres = $pp.Presentations.Open($tpl.FullName, $true, $false, $false)
                        $pres.SaveAs($targetPath, $saveFmt)
                    }
                    else {
                        $pres = $pp.Presentations.Add()
                        $slide = $pres.Slides.Add(1, 12)
                        $slide.Shapes.Title.TextFrame.TextRange.Text = "Ritridata test"
                        $pres.SaveAs($targetPath, $saveFmt)
                    }
                    Write-Host "Created (PowerPoint): $targetPath"
                    $Index.Value++
                }
                catch {
                    Write-Warning "PowerPoint failed .$ext #$i : $($_.Exception.Message)"
                }
                finally {
                    if ($pres) {
                        try { $pres.Close() } catch { }
                    }
                }
            }
        }
        catch {
            Write-Warning "PowerPoint COM unavailable for .$ext : $($_.Exception.Message)"
        }
        finally {
            Invoke-SafeQuit $pp
            Stop-OfficeProcess "POWERPNT"
        }
    }
}

function New-ImageSamples {
    param(
        [string]$DriveRoot,
        [string]$Extension,
        [int]$Count,
        [ref]$Index,
        [string]$RunTime,
        [hashtable]$TemplatesByExt
    )
    if ($Extension -ne "jpg") {
        if (-not $TemplatesByExt.ContainsKey($Extension) -or $TemplatesByExt[$Extension].Count -eq 0) {
            Write-Warning "Skip .$Extension : add templates under user_resources"
            return
        }
        Copy-FromTemplate -DriveRoot $DriveRoot -Extension $Extension -Count $Count -Index $Index -RunTime $RunTime -TemplatesByExt $TemplatesByExt | Out-Null
        return
    }
    if (-not $TemplatesByExt.ContainsKey("jpg") -or $TemplatesByExt["jpg"].Count -eq 0) {
        Write-Warning "Skip jpg: no jpg template in user_resources"
        return
    }
    Copy-FromTemplate -DriveRoot $DriveRoot -Extension "jpg" -Count $Count -Index $Index -RunTime $RunTime -TemplatesByExt $TemplatesByExt | Out-Null
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
    Write-Host "Created template folder: $UserResourceDir"
    Write-Host "Please add real pptx/xlsx/jpg files under user_resources, then run again."
    exit 0
}

if (-not (Test-Path -LiteralPath $TargetDriveRoot)) {
    Write-Error "Target drive not found: $TargetDriveRoot"
    exit 1
}

$runTime = Get-Date -Format "yyMMdd_HHmm"
Write-Host "Document types: $($ProductDocumentExtensions -join ', ')"
if ($OpenXmlOnly) {
    Write-Host "Mode: OpenXmlOnly (no doc/xls/ppt)"
}
Clear-TargetDriveRoot -DriveRoot $TargetDriveRoot

$templateFiles = Get-ChildItem -LiteralPath $UserResourceDir -File -Recurse -Force -ErrorAction SilentlyContinue |
    Where-Object { $SkipTemplateNames -notcontains $_.Name }

$templatesByExt = @{}
foreach ($f in $templateFiles) {
    $ext = $f.Extension.TrimStart(".").ToLowerInvariant()
    if ($ext -eq "jpeg") { $ext = "jpg" }
    if ($TargetExtensions -notcontains $ext) { continue }
    if (-not $templatesByExt.ContainsKey($ext)) {
        $templatesByExt[$ext] = @()
    }
    $templatesByExt[$ext] += $f
}

$globalIndex = 1
$idxRef = [ref]$globalIndex
$generated = @{}

foreach ($ext in $NativeTextExtensions) {
    $outDir = Get-CategoryOutputDir -DriveRoot $TargetDriveRoot -Extension $ext
    $category = Get-CategoryFolderName -Extension $ext
    for ($i = 1; $i -le $FilesPerExtension; $i++) {
        $uniqueName = New-UniqueFileName -Extension $ext -Index $idxRef.Value -RunTime $runTime
        $targetPath = Join-Path $outDir $uniqueName
        New-RealTextFile -Path $targetPath -Extension $ext -Label (Join-Path $category $uniqueName)
        Write-Host "Created: $targetPath"
        $idxRef.Value++
    }
    $generated[$ext] = $true
}

foreach ($ext in $ImageExtensions) {
    if ($ext -eq "jpg" -or ($templatesByExt.ContainsKey($ext) -and $templatesByExt[$ext].Count -gt 0)) {
        $before = $idxRef.Value
        New-ImageSamples -DriveRoot $TargetDriveRoot -Extension $ext -Count $FilesPerExtension -Index $idxRef -RunTime $runTime -TemplatesByExt $templatesByExt
        if ($idxRef.Value -gt $before) { $generated[$ext] = $true }
    }
}

$templateCopyOrder = @(
    "pdf", "docx", "doc", "xlsx", "xls", "pptx", "ppt",
    "odt", "pages", "numbers", "key"
)
foreach ($ext in $templateCopyOrder) {
    if ($generated.ContainsKey($ext)) { continue }
    if ($OpenXmlOnly -and $ext -in @("doc", "xls", "ppt")) { continue }
    if (Copy-FromTemplate -DriveRoot $TargetDriveRoot -Extension $ext -Count $FilesPerExtension -Index $idxRef -RunTime $runTime -TemplatesByExt $templatesByExt) {
        $generated[$ext] = $true
    }
}

if (-not $SkipOfficeCom) {
    foreach ($ext in $WordExtensions) {
        if ($generated.ContainsKey($ext)) { continue }
        $before = $idxRef.Value
        New-WordSamples -DriveRoot $TargetDriveRoot -Extensions @($ext) -CountPerExt $FilesPerExtension -Index $idxRef -RunTime $runTime
        if ($idxRef.Value -gt $before) { $generated[$ext] = $true }
    }

    foreach ($ext in $ExcelExtensions) {
        if ($generated.ContainsKey($ext)) { continue }
        $before = $idxRef.Value
        New-ExcelSamples -DriveRoot $TargetDriveRoot -Extensions @($ext) -CountPerExt $FilesPerExtension -Index $idxRef -RunTime $runTime
        if ($idxRef.Value -gt $before) { $generated[$ext] = $true }
    }

    foreach ($ext in $PowerPointExtensions) {
        if ($generated.ContainsKey($ext)) { continue }
        $before = $idxRef.Value
        New-PowerPointSamples -DriveRoot $TargetDriveRoot -Extensions @($ext) -CountPerExt $FilesPerExtension -Index $idxRef -RunTime $runTime -TemplatesByExt $templatesByExt
        if ($idxRef.Value -gt $before) { $generated[$ext] = $true }
    }
}
else {
    Write-Warning "SkipOfficeCom: Office types need templates or remove -SkipOfficeCom"
}

foreach ($ext in $TemplateOnlyExtensions) {
    if ($generated.ContainsKey($ext)) { continue }
    if (Copy-FromTemplate -DriveRoot $TargetDriveRoot -Extension $ext -Count $FilesPerExtension -Index $idxRef -RunTime $runTime -TemplatesByExt $templatesByExt) {
        $generated[$ext] = $true
    }
}

$docSkipped = $ProductDocumentExtensions | Where-Object { -not $generated.ContainsKey($_) }
if ($docSkipped.Count -gt 0) {
    Write-Warning "Document types not generated: $($docSkipped -join ', ')"
    $needTpl = $docSkipped | Where-Object { $_ -in $TemplateOnlyExtensions }
    if ($needTpl.Count -gt 0) {
        Write-Warning "Add real templates to scripts\user_resources\ for: $($needTpl -join ', ')"
    }
}

$imgSkipped = $ImageExtensions | Where-Object { $_ -ne "jpg" -and -not $generated.ContainsKey($_) }
if ($imgSkipped.Count -gt 0) {
    Write-Warning "Image types not generated: $($imgSkipped -join ', ')"
}

$createdCount = $idxRef.Value - 1
Write-Host "Done. Total files: $createdCount"
Write-Host "Naming: G_<ext>_<seq>_<yyMMdd_HHmm>.<ext>"
