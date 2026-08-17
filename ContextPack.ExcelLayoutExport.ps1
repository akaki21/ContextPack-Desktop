# Owns Excel workbook layout export and per-sheet diagnostics.
# Excel COM lifecycle, diagnostics, AutoFit decisions, and pagination come from the shared helpers.

function Export-ContextPackExcelLayout {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][ValidateSet('Workbook', 'AutoFit')][string]$Layout,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)]$SheetMetrics,
        [Parameter(Mandatory = $true)][ValidateRange(10, 200)][int]$MaxAutoFitColumns
    )

    $excel = $null
    $workbook = $null
    $diagnostics = @()
    try {
        $excel = New-ContextPackExcelApplication
        $workbook = Open-ContextPackExcelWorkbook -Application $excel -Path $InputPath

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
