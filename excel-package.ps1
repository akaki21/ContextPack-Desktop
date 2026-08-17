param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [ValidateSet('Workbook', 'AutoFit', 'Both')][string]$RenderMode = 'Both',
    [ValidateRange(10, 200)][int]$MaxAutoFitColumns = 60,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
. (Join-Path $root 'ContextPack.ExcelCom.ps1')
. (Join-Path $root 'ContextPack.ExcelDiagnostics.ps1')
. (Join-Path $root 'ContextPack.ExcelPagination.ps1')
. (Join-Path $root 'ContextPack.ExcelAutoFit.ps1')
. (Join-Path $root 'ContextPack.ExcelWorkbookLayout.ps1')
. (Join-Path $root 'ContextPack.ExcelLayoutReport.ps1')
$python = Get-ContextPackPython
$extractor = Join-Path $root 'extract-excel-package.py'
$renderer = Join-Path $root 'render-pdf-pages.py'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'Input must be a file.' }
$extension = [System.IO.Path]::GetExtension($inputPath).ToLowerInvariant()
if ($extension -notin @('.xlsx', '.xlsm', '.xltx', '.xltm')) { throw 'Supported formats: .xlsx, .xlsm, .xltx, .xltm' }

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$build = New-ContextPackBuild -InputPath $inputPath -PreferredName ($baseName + '_excel_package') -OutputDirectory $OutputDirectory

function Export-ExcelLayout {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('Workbook', 'AutoFit')][string]$Layout,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)]$SheetMetrics
    )
    $excel = $null
    $workbook = $null
    $diagnostics = @()
    try {
        $excel = New-ContextPackExcelApplication
        $workbook = Open-ContextPackExcelWorkbook -Application $excel -Path $inputPath

        $worksheetCount = [int](Invoke-ExcelRetry { $workbook.Worksheets.Count })
        for ($worksheetIndex = 1; $worksheetIndex -le $worksheetCount; $worksheetIndex++) {
            $worksheet = $null
            try {
                $worksheet = Invoke-ExcelRetry { $workbook.Worksheets.Item($worksheetIndex) }
                $title = [string]$worksheet.Name
                $metric = $SheetMetrics[$title]
                $visible = ([int]$worksheet.Visible -eq -1)
                $printAreaBefore = ''
                $titleRows = ''
                $titleColumns = ''
                try { $printAreaBefore = [string]$worksheet.PageSetup.PrintArea } catch { }
                try { $titleRows = [string]$worksheet.PageSetup.PrintTitleRows } catch { }
                try { $titleColumns = [string]$worksheet.PageSetup.PrintTitleColumns } catch { }
                $horizontalBreaks = Get-ExcelManualPageBreakCount $worksheet 'HPageBreaks'
                $verticalBreaks = Get-ExcelManualPageBreakCount $worksheet 'VPageBreaks'
                $shapeCount = Get-ExcelShapeCount $worksheet
                $status = if ($Layout -eq 'Workbook') { 'preserved' } else { 'skipped' }
                $reasons = @()
                $printAreaAfter = $printAreaBefore
                $fitToPagesWide = $null

                if ($Layout -eq 'AutoFit') {
                    $autoFitDecision = Get-ContextPackExcelAutoFitDecision -Visible $visible -Metric $metric -ShapeCount $shapeCount -HorizontalPageBreaks $horizontalBreaks -VerticalPageBreaks $verticalBreaks -MaxAutoFitColumns $MaxAutoFitColumns
                    $reasons += @($autoFitDecision.Reasons)
                    if ($autoFitDecision.CanApply) {
                        $pagination = Set-ContextPackExcelAutoFitPagination -Worksheet $worksheet -Metric $metric
                        $printAreaAfter = $pagination.PrintArea
                        $fitToPagesWide = $pagination.FitToPagesWide
                        $reasons += @($pagination.Notes)
                        $status = 'applied'
                    }
                }

                $diagnostics += [pscustomobject]@{
                    sheet = $title
                    visible = $visible
                    layout = $Layout
                    status = $status
                    reasons = @($reasons)
                    print_area_before = $printAreaBefore
                    print_area_after = $printAreaAfter
                    fit_to_pages_wide = $fitToPagesWide
                    print_title_rows = $titleRows
                    print_title_columns = $titleColumns
                    manual_horizontal_page_breaks = $horizontalBreaks
                    manual_vertical_page_breaks = $verticalBreaks
                    drawing_objects = $shapeCount
                }
            } finally {
                Release-ExcelComObject $worksheet
            }
        }
        Invoke-ExcelRetry { $workbook.ExportAsFixedFormat(0, $PdfPath, 0, $true, $false) } | Out-Null
        return @($diagnostics)
    } finally {
        Complete-ContextPackExcelComCleanup -Workbook $workbook -Application $excel
    }
}

try {
    $packageDir = $build.BuildPath
    $renderedDir = Join-Path $packageDir 'rendered-sheets'
    New-Item -ItemType Directory -Path $renderedDir -Force | Out-Null
    $env:PYTHONUTF8 = '1'
    & $python $extractor $inputPath $packageDir
    if ($LASTEXITCODE -ne 0) { throw "Excel extraction failed with exit code: $LASTEXITCODE" }

    $metricsPath = Join-Path $packageDir 'excel-metrics.json'
    $metrics = Get-Content -LiteralPath $metricsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $sheetMetricMap = @{}
    foreach ($sheetMetric in $metrics.sheets) { $sheetMetricMap[[string]$sheetMetric.title] = $sheetMetric }
    $layoutDiagnostics = @()
    $layoutWarnings = @()
    $maxRenderedPages = 1000

    if ($RenderMode -in @('Workbook', 'Both')) {
        $workbookLayoutResult = Invoke-ContextPackExcelWorkbookLayout -RenderedDirectory $renderedDir -Python $python -Renderer $renderer -Dpi $Dpi -MaxRenderedPages $maxRenderedPages -ExportPdf {
            param($PdfPath)
            Export-ExcelLayout -Layout Workbook -PdfPath $PdfPath -SheetMetrics $sheetMetricMap
        }
        $layoutDiagnostics += @($workbookLayoutResult.Diagnostics)
        $layoutWarnings += @($workbookLayoutResult.Warnings)
    }

    if ($RenderMode -in @('AutoFit', 'Both')) {
        $autoLayoutDir = Join-Path $renderedDir 'auto-layout'
        $autoPages = Join-Path $autoLayoutDir 'pages'
        New-Item -ItemType Directory -Path $autoPages -Force | Out-Null
        $autoPdf = Join-Path $autoLayoutDir 'workbook.pdf'
        $autoRenderMetricsPath = Join-Path $autoLayoutDir 'page-render-metrics.json'
        $layoutDiagnostics += Export-ExcelLayout -Layout AutoFit -PdfPath $autoPdf -SheetMetrics $sheetMetricMap
        & $python $renderer $autoPdf $autoPages --dpi $Dpi --metrics $autoRenderMetricsPath --max-pages $maxRenderedPages
        if ($LASTEXITCODE -ne 0) { throw 'Rendering auto-layout PDF failed' }
        $autoRenderMetrics = Get-Content -LiteralPath $autoRenderMetricsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($autoRenderMetrics.render_skipped) {
            $layoutWarnings += "Auto-layout PNG rendering skipped: $($autoRenderMetrics.reason) The complete PDF is preserved."
        }
    }

    $layoutReportResult = Write-ContextPackExcelLayoutReport -PackageDirectory $packageDir -Diagnostics $layoutDiagnostics
    $autoApplied = @($layoutReportResult.AutoFitApplied)
    $autoSkipped = @($layoutReportResult.AutoFitSkipped)
    $layoutWarnings += @($layoutReportResult.Warnings)

    $qualityPath = Join-Path $packageDir 'quality-report.md'
    Add-Content -LiteralPath $qualityPath -Encoding UTF8 -Value @(
        ''
        '## Print layout rendering'
        ''
        "- Requested mode: $RenderMode"
        "- AutoFit applied sheets: $($autoApplied.Count)"
        "- AutoFit skipped sheets: $($autoSkipped.Count)"
        '- Workbook layout always preserves the workbook print settings.'
        '- AutoFit keeps hidden rows/columns/sheets hidden, preserves orientation and print titles, and adaptively splits wide tables across horizontal pages to protect readability.'
        $(if ($layoutWarnings.Count) { $layoutWarnings | ForEach-Object { '- ' + $_ } } else { '- No AutoFit safety warning was detected.' })
        ''
        'Auto-layout is a convenience view. Treat workbook-layout and the original workbook as authoritative.'
    )

    $handoffEnglish = @(
        "# AI Handoff — Excel: $baseName", '', 'Read `workbook-info.md` first and open only relevant files under `sheets-data`. Use `quality-report.md` and `print-layout-report.json` for warnings. `workbook-layout` preserves author print settings; `auto-layout` is only a convenience view and must be verified against the original workbook.', '', 'Goal: [describe the goal]', 'Scope: [sheets/ranges/period]', 'Desired output: [format]'
    ) -join [Environment]::NewLine
    $handoffGeorgian = @(
        "# AI Handoff — Excel: $baseName", '', 'ჯერ წაიკითხე `workbook-info.md` და `sheets-data`-დან გახსენი მხოლოდ საჭირო ფურცლები. გაფრთხილებებისთვის გამოიყენე `quality-report.md` და `print-layout-report.json`. `workbook-layout` ინარჩუნებს ავტორის print settings-ს; `auto-layout` მხოლოდ დამხმარე ხედია და ორიგინალ workbook-თან უნდა გადამოწმდეს.', '', 'მიზანი: [აღწერე მიზანი]', 'ფარგლები: [ფურცლები/დიაპაზონები/პერიოდი]', 'შედეგი: [ფორმატი]'
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.md') -Value $handoffEnglish -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.ka.md') -Value $handoffGeorgian -Encoding UTF8

    $warnings = @($metrics.warnings) + @($layoutWarnings)
    $outputs = [ordered]@{
        source_workbook = [System.IO.Path]::GetFileName($inputPath)
        workbook_info = 'workbook-info.md'
        values_index = 'values.md'
        formulas_index = 'formulas.md'
        sheet_data = 'sheets-data/'
        workbook_layout = $(if ($RenderMode -in @('Workbook','Both')) { 'rendered-sheets/workbook-layout/' } else { $null })
        auto_layout = $(if ($RenderMode -in @('AutoFit','Both')) { 'rendered-sheets/auto-layout/' } else { $null })
        print_layout_report = 'print-layout-report.json'
        quality_report = 'quality-report.md'
    }
    Write-ContextPackManifest -Build $build -InputPath $inputPath -PackageType 'excel' -Outputs $outputs -Settings ([ordered]@{ dpi = $Dpi; render_mode = $RenderMode; max_autofit_columns = $MaxAutoFitColumns; max_rendered_pages_per_layout = $maxRenderedPages; macros_disabled = $true; events_disabled = $true; source_saved = $false }) -Warnings $warnings
    $finalPath = Complete-ContextPackBuild $build
    Write-Host "Excel package ready: $finalPath" -ForegroundColor Green
} catch {
    Remove-ContextPackBuild $build
    throw
}
