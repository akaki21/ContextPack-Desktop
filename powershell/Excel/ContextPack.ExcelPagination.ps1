# Excel: plan and apply AutoFit horizontal pagination without saving the workbook.
# Eight populated columns per horizontal page is the reviewed readability rule.

function Get-ContextPackExcelWideSheetSkipReason {
    param(
        [Parameter(Mandatory = $true)][int]$PopulatedColumnSpan,
        [Parameter(Mandatory = $true)][ValidateRange(10, 200)][int]$MaxAutoFitColumns
    )

    if ($PopulatedColumnSpan -gt $MaxAutoFitColumns) {
        return "populated range exceeds the $MaxAutoFitColumns-column AutoFit safety limit"
    }
    return $null
}

function Get-ContextPackExcelHorizontalPageCount {
    param([Parameter(Mandatory = $true)][ValidateRange(1, 16384)][int]$PopulatedColumnSpan)

    return [Math]::Max(1, [Math]::Ceiling($PopulatedColumnSpan / 8.0))
}

function Set-ContextPackExcelAutoFitPagination {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)]$Metric
    )

    $startCell = $null
    $endCell = $null
    $usedDataRange = $null
    $pageSetup = $null
    try {
        $startCell = $Worksheet.Cells([int]$Metric.min_row, [int]$Metric.min_column)
        $endCell = $Worksheet.Cells([int]$Metric.max_row, [int]$Metric.max_column)
        $usedDataRange = $Worksheet.Range($startCell, $endCell)
        $pageSetup = Invoke-ExcelRetry { $Worksheet.PageSetup }
        $printArea = [string]$usedDataRange.Address()
        $fitToPagesWide = Get-ContextPackExcelHorizontalPageCount -PopulatedColumnSpan ([int]$Metric.populated_column_span)

        Invoke-ExcelRetry { $pageSetup.PrintArea = $printArea } | Out-Null
        Invoke-ExcelRetry { $pageSetup.Zoom = $false } | Out-Null
        Invoke-ExcelRetry { $pageSetup.FitToPagesWide = $fitToPagesWide } | Out-Null
        Invoke-ExcelRetry { $pageSetup.FitToPagesTall = $false } | Out-Null

        $notes = [System.Collections.Generic.List[string]]::new()
        if ($fitToPagesWide -gt 1) { $notes.Add("wide sheet split across $fitToPagesWide pages to preserve readability") }
        if ([int]$Metric.merged_ranges -gt 0) { $notes.Add('merged cells are present; verify page boundaries visually') }
        return [pscustomobject]@{
            PrintArea = $printArea
            FitToPagesWide = $fitToPagesWide
            Notes = $notes
        }
    } finally {
        Release-ExcelComObject $pageSetup
        Release-ExcelComObject $usedDataRange
        Release-ExcelComObject $startCell
        Release-ExcelComObject $endCell
    }
}
