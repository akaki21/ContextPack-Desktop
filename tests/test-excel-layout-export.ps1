$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'ContextPack.ExcelLayoutExport.ps1')

function Assert-ExcelLayoutExportTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel layout-export test failed: $Message" }
}

$pageSetup = [pscustomobject]@{
    PrintArea = '$A$1:$Q$20'
    PrintTitleRows = '$1:$1'
    PrintTitleColumns = '$A:$A'
}
$worksheet = [pscustomobject]@{
    Name = 'Estimate'
    Visible = -1
    PageSetup = $pageSetup
}
$worksheets = [pscustomobject]@{
    Count = 1
    Items = @($worksheet)
}
$worksheets | Add-Member -MemberType ScriptMethod -Name Item -Value {
    param($index)
    return $this.Items[$index - 1]
}
$workbook = [pscustomobject]@{
    Worksheets = $worksheets
    ExportArguments = $null
    ThrowOnExport = $false
}
$workbook | Add-Member -MemberType ScriptMethod -Name ExportAsFixedFormat -Value {
    param($type, $path, $quality, $includeDocumentProperties, $ignorePrintAreas)
    if ($this.ThrowOnExport) { throw 'Simulated PDF export failure.' }
    $this.ExportArguments = @($type, $path, $quality, $includeDocumentProperties, $ignorePrintAreas)
}

$script:application = [pscustomobject]@{ Marker = 'excel-application' }
$script:openedPaths = [System.Collections.Generic.List[string]]::new()
$script:releasedWorksheets = [System.Collections.Generic.List[object]]::new()
$script:cleanupCalls = [System.Collections.Generic.List[object]]::new()
$script:autoFitCanApply = $true
$script:autoFitDecisionCalls = 0
$script:paginationCalls = 0
$script:lastMaxAutoFitColumns = $null

function Invoke-ExcelRetry {
    param([scriptblock]$Action)
    return & $Action
}
function New-ContextPackExcelApplication { return $script:application }
function Open-ContextPackExcelWorkbook {
    param($Application, [string]$Path)
    if ($Application -ne $script:application) { throw 'Unexpected application fixture.' }
    $script:openedPaths.Add($Path)
    return $script:workbookFixture
}
function Get-ExcelManualPageBreakCount {
    param($Worksheet, [string]$PropertyName)
    if ($PropertyName -eq 'HPageBreaks') { return 2 }
    return 1
}
function Get-ExcelShapeCount { param($Worksheet) return 3 }
function Get-ContextPackExcelAutoFitDecision {
    param($Visible, $Metric, $ShapeCount, $HorizontalPageBreaks, $VerticalPageBreaks, $MaxAutoFitColumns)
    $script:autoFitDecisionCalls++
    $script:lastMaxAutoFitColumns = $MaxAutoFitColumns
    return [pscustomobject]@{
        CanApply = $script:autoFitCanApply
        Reasons = $(if ($script:autoFitCanApply) { @() } else { @('fixture safety reason') })
    }
}
function Set-ContextPackExcelAutoFitPagination {
    param($Worksheet, $Metric)
    $script:paginationCalls++
    return [pscustomobject]@{
        PrintArea = '$B$2:$R$10'
        FitToPagesWide = 3
        Notes = @('wide sheet split across 3 pages to preserve readability')
    }
}
function Release-ExcelComObject {
    param([AllowNull()]$ComObject)
    if ($null -ne $ComObject) { $script:releasedWorksheets.Add($ComObject) }
}
function Complete-ContextPackExcelComCleanup {
    param([AllowNull()]$Workbook, [AllowNull()]$Application)
    $script:cleanupCalls.Add([pscustomobject]@{ Workbook = $Workbook; Application = $Application })
}

$script:workbookFixture = $workbook
$metrics = @{
    Estimate = [pscustomobject]@{
        min_row = 2
        min_column = 2
        max_row = 10
        max_column = 18
        populated_column_span = 17
        charts = 0
        images = 0
        merged_ranges = 0
    }
}

$workbookResult = @(Export-ContextPackExcelLayout -InputPath 'fixture.xlsx' -Layout Workbook -PdfPath 'workbook.pdf' -SheetMetrics $metrics -MaxAutoFitColumns 60)
Assert-ExcelLayoutExportTest ($workbookResult.Count -eq 1) 'Workbook diagnostics were not returned.'
Assert-ExcelLayoutExportTest ($workbookResult[0].status -eq 'preserved') 'Workbook layout is no longer marked preserved.'
Assert-ExcelLayoutExportTest ($workbookResult[0].print_area_before -eq '$A$1:$Q$20' -and $workbookResult[0].print_area_after -eq '$A$1:$Q$20') 'Workbook print area was not preserved.'
Assert-ExcelLayoutExportTest ($workbookResult[0].manual_horizontal_page_breaks -eq 2 -and $workbookResult[0].manual_vertical_page_breaks -eq 1) 'Page-break diagnostics changed.'
Assert-ExcelLayoutExportTest ($workbookResult[0].drawing_objects -eq 3) 'Drawing-object diagnostics changed.'
Assert-ExcelLayoutExportTest ($script:autoFitDecisionCalls -eq 0 -and $script:paginationCalls -eq 0) 'Workbook mode invoked AutoFit logic.'
Assert-ExcelLayoutExportTest ($workbook.ExportArguments[0] -eq 0 -and $workbook.ExportArguments[1] -eq 'workbook.pdf' -and $workbook.ExportArguments[2] -eq 0) 'PDF export arguments changed.'
Assert-ExcelLayoutExportTest ($workbook.ExportArguments[3] -eq $true -and $workbook.ExportArguments[4] -eq $false) 'PDF export no longer preserves document properties or print areas.'

$script:autoFitCanApply = $true
$autoFitResult = @(Export-ContextPackExcelLayout -InputPath 'fixture.xlsx' -Layout AutoFit -PdfPath 'autofit.pdf' -SheetMetrics $metrics -MaxAutoFitColumns 77)
Assert-ExcelLayoutExportTest ($autoFitResult[0].status -eq 'applied') 'Eligible AutoFit layout was not applied.'
Assert-ExcelLayoutExportTest ($autoFitResult[0].print_area_after -eq '$B$2:$R$10' -and $autoFitResult[0].fit_to_pages_wide -eq 3) 'AutoFit pagination results were lost.'
Assert-ExcelLayoutExportTest (($autoFitResult[0].reasons -join ' ') -match 'split across 3 pages') 'AutoFit pagination guidance was lost.'
Assert-ExcelLayoutExportTest ($script:lastMaxAutoFitColumns -eq 77) 'The custom AutoFit width limit was not forwarded.'
Assert-ExcelLayoutExportTest ($script:paginationCalls -eq 1) 'Eligible AutoFit pagination was not called exactly once.'

$script:autoFitCanApply = $false
$skippedResult = @(Export-ContextPackExcelLayout -InputPath 'fixture.xlsx' -Layout AutoFit -PdfPath 'skipped.pdf' -SheetMetrics $metrics -MaxAutoFitColumns 60)
Assert-ExcelLayoutExportTest ($skippedResult[0].status -eq 'skipped') 'An unsafe AutoFit sheet was not skipped.'
Assert-ExcelLayoutExportTest ($skippedResult[0].reasons[0] -eq 'fixture safety reason') 'The AutoFit skip reason was lost.'
Assert-ExcelLayoutExportTest ($skippedResult[0].print_area_after -eq '$A$1:$Q$20' -and $null -eq $skippedResult[0].fit_to_pages_wide) 'Skipped AutoFit changed pagination settings.'
Assert-ExcelLayoutExportTest ($script:paginationCalls -eq 1) 'Skipped AutoFit invoked pagination.'

$workbook.ThrowOnExport = $true
$exportFailurePropagated = $false
try {
    Export-ContextPackExcelLayout -InputPath 'fixture.xlsx' -Layout Workbook -PdfPath 'failure.pdf' -SheetMetrics $metrics -MaxAutoFitColumns 60
} catch {
    $exportFailurePropagated = $true
}
Assert-ExcelLayoutExportTest $exportFailurePropagated 'A PDF export failure was swallowed.'
Assert-ExcelLayoutExportTest ($script:cleanupCalls.Count -eq 4) 'Success or failure did not run COM cleanup exactly once.'
Assert-ExcelLayoutExportTest ($script:releasedWorksheets.Count -eq 4) 'A worksheet was not released after success or failure.'
Assert-ExcelLayoutExportTest ($script:openedPaths.Count -eq 4 -and ($script:openedPaths | Where-Object { $_ -ne 'fixture.xlsx' }).Count -eq 0) 'The source workbook path changed before opening.'

Write-Host 'Excel layout-export orchestration tests passed.' -ForegroundColor Green
