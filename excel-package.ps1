param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateRange(96, 300)][int]$Dpi = 180
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$extractor = Join-Path $root 'extract-excel-package.py'
$renderer = Join-Path $root 'render-pdf-pages.py'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'Input must be a file.' }
$extension = [System.IO.Path]::GetExtension($inputPath).ToLowerInvariant()
if ($extension -notin @('.xlsx', '.xlsm', '.xltx', '.xltm')) { throw 'Supported formats: .xlsx, .xlsm, .xltx, .xltm' }

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$build = New-ContextPackBuild -InputPath $inputPath -PreferredName ($baseName + '_excel_package')

function Invoke-ExcelRetry {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        try { return & $Action }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 8) { throw }
            Start-Sleep -Milliseconds (300 * $attempt)
        }
    }
}

try {
    $packageDir = $build.BuildPath
    $renderedDir = Join-Path $packageDir 'rendered-sheets'
    $pagesFolder = Join-Path $renderedDir 'pages'
    New-Item -ItemType Directory -Path $pagesFolder -Force | Out-Null
    $env:PYTHONUTF8 = '1'
    & $python $extractor $inputPath $packageDir
    if ($LASTEXITCODE -ne 0) { throw "Excel extraction failed with exit code: $LASTEXITCODE" }

    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 800
        Invoke-ExcelRetry { $excel.Visible = $false }
        Invoke-ExcelRetry { $excel.DisplayAlerts = $false }
        Invoke-ExcelRetry { $excel.ScreenUpdating = $false }
        Invoke-ExcelRetry { $excel.EnableEvents = $false }
        Invoke-ExcelRetry { $excel.AutomationSecurity = 3 }
        $workbook = Invoke-ExcelRetry { $excel.Workbooks.Open($inputPath, 0, $true) }
        $workbookPdf = Join-Path $renderedDir 'workbook.pdf'
        Invoke-ExcelRetry { $workbook.ExportAsFixedFormat(0, $workbookPdf, 0, $true, $false) }
        & $python $renderer $workbookPdf $pagesFolder --dpi $Dpi
        if ($LASTEXITCODE -ne 0) { throw 'Rendering workbook PDF failed' }
    } finally {
        if ($workbook -ne $null) {
            try { Invoke-ExcelRetry { $workbook.Close($false) } } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch { }
        }
        if ($excel -ne $null) {
            try { Invoke-ExcelRetry { $excel.Quit() } } catch { }
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    $metricsPath = Join-Path $packageDir 'excel-metrics.json'
    $metrics = Get-Content -LiteralPath $metricsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $handoffEnglish = @(
        "# AI Handoff — Excel: $baseName", '', 'Read `workbook-info.md` first. Open only the relevant per-sheet files under `sheets-data`. Use `quality-report.md` to find cached formula errors, external links, sparse sheets, or other warnings. Use the original workbook and rendered pages only for exact verification.', '', 'Goal: [describe the goal]', 'Scope: [sheets/ranges/period]', 'Desired output: [format]'
    ) -join [Environment]::NewLine
    $handoffGeorgian = @(
        "# AI Handoff — Excel: $baseName", '', 'ჯერ წაიკითხე `workbook-info.md`. შემდეგ `sheets-data`-დან გახსენი მხოლოდ საჭირო ფურცლის ფაილები. `quality-report.md` გამოიყენე formula error-ების, external link-ების, sparse ფურცლებისა და სხვა გაფრთხილებების სანახავად. ორიგინალი workbook და დარენდერებული გვერდები გამოიყენე მხოლოდ ზუსტი გადამოწმებისთვის.', '', 'მიზანი: [აღწერე მიზანი]', 'ფარგლები: [ფურცლები/დიაპაზონები/პერიოდი]', 'შედეგი: [ფორმატი]'
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.md') -Value $handoffEnglish -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.ka.md') -Value $handoffGeorgian -Encoding UTF8

    $warnings = @($metrics.warnings)
    $outputs = [ordered]@{
        source_workbook = [System.IO.Path]::GetFileName($inputPath)
        workbook_info = 'workbook-info.md'
        values_index = 'values.md'
        formulas_index = 'formulas.md'
        sheet_data = 'sheets-data/'
        rendered_pdf = 'rendered-sheets/workbook.pdf'
        rendered_pages = 'rendered-sheets/pages/'
        quality_report = 'quality-report.md'
    }
    Write-ContextPackManifest -Build $build -InputPath $inputPath -PackageType 'excel' -Outputs $outputs -Settings ([ordered]@{ dpi = $Dpi; macros_disabled = $true; events_disabled = $true }) -Warnings $warnings
    $finalPath = Complete-ContextPackBuild $build
    Write-Host "Excel package ready: $finalPath" -ForegroundColor Green
} catch {
    Remove-ContextPackBuild $build
    throw
}
