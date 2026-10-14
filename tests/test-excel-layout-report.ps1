$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelModules = Join-Path $root 'powershell\Excel'
. (Join-Path $excelModules 'ContextPack.ExcelLayoutReport.ps1')

function Assert-ExcelLayoutReportTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel layout-report test failed: $Message" }
}

function New-ExcelLayoutDiagnosticFixture {
    param(
        [string]$Sheet,
        [string]$Layout,
        [string]$Status,
        [object[]]$Reasons = @(),
        $FitToPagesWide = $null
    )

    return [pscustomobject][ordered]@{
        sheet = $Sheet
        visible = $true
        layout = $Layout
        status = $Status
        reasons = @($Reasons)
        print_area_before = '$A$1:$Q$20'
        print_area_after = '$A$1:$Q$20'
        fit_to_pages_wide = $FitToPagesWide
        print_title_rows = '$1:$1'
        print_title_columns = '$A:$A'
        manual_horizontal_page_breaks = 0
        manual_vertical_page_breaks = 0
        drawing_objects = 0
    }
}

$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-LayoutReport-Test-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $testDirectory -Force | Out-Null
    $diagnostics = @(
        (New-ExcelLayoutDiagnosticFixture -Sheet 'Authoritative' -Layout 'Workbook' -Status 'preserved')
        (New-ExcelLayoutDiagnosticFixture -Sheet 'ფართო ცხრილი' -Layout 'AutoFit' -Status 'applied' -Reasons @('Wide table split across 3 pages') -FitToPagesWide 3)
        (New-ExcelLayoutDiagnosticFixture -Sheet 'Dashboard' -Layout 'AutoFit' -Status 'skipped' -Reasons @('Drawing objects detected', 'Manual page breaks detected'))
    )

    $result = Write-ContextPackExcelLayoutReport -PackageDirectory $testDirectory -Diagnostics $diagnostics
    Assert-ExcelLayoutReportTest ($result.ReportPath -eq (Join-Path $testDirectory 'print-layout-report.json')) 'The report path changed.'
    Assert-ExcelLayoutReportTest (Test-Path -LiteralPath $result.ReportPath -PathType Leaf) 'The report was not written.'
    Assert-ExcelLayoutReportTest ($result.AutoFitApplied.Count -eq 1) 'Applied AutoFit diagnostics were counted incorrectly.'
    Assert-ExcelLayoutReportTest ($result.AutoFitSkipped.Count -eq 1) 'Skipped AutoFit diagnostics were counted incorrectly.'
    Assert-ExcelLayoutReportTest ($result.Warnings.Count -eq 2) 'AutoFit warnings were counted incorrectly.'
    Assert-ExcelLayoutReportTest ($result.Warnings[0] -eq "AutoFit skipped for sheet 'Dashboard': Drawing objects detected; Manual page breaks detected.") 'The skipped warning text or order changed.'
    Assert-ExcelLayoutReportTest ($result.Warnings[1] -eq "AutoFit note for sheet 'ფართო ცხრილი': Wide table split across 3 pages.") 'The applied note text or order changed.'

    $reportText = Get-Content -LiteralPath $result.ReportPath -Raw -Encoding UTF8
    Assert-ExcelLayoutReportTest ($reportText.TrimStart().StartsWith('[')) 'A multi-record report is not a JSON array.'
    Assert-ExcelLayoutReportTest ($reportText -match 'ფართო ცხრილი') 'UTF-8 worksheet text was not preserved.'
    $report = ConvertFrom-Json -InputObject $reportText
    Assert-ExcelLayoutReportTest ($report.Count -eq 3) 'The report lost diagnostics.'
    $expectedFields = @('sheet', 'visible', 'layout', 'status', 'reasons', 'print_area_before', 'print_area_after', 'fit_to_pages_wide', 'print_title_rows', 'print_title_columns', 'manual_horizontal_page_breaks', 'manual_vertical_page_breaks', 'drawing_objects')
    $actualFields = @($report[0].PSObject.Properties.Name)
    Assert-ExcelLayoutReportTest (($actualFields -join '|') -eq ($expectedFields -join '|')) 'The stable report schema changed.'
    Assert-ExcelLayoutReportTest ($report[1].fit_to_pages_wide -eq 3) 'The pagination value was not preserved.'
    Assert-ExcelLayoutReportTest ($null -eq $report[0].fit_to_pages_wide) 'A null pagination value was not preserved.'
    Assert-ExcelLayoutReportTest ($report[0].reasons.Count -eq 0) 'An empty reasons array was not preserved.'

    $singleDirectory = Join-Path $testDirectory 'single'
    New-Item -ItemType Directory -Path $singleDirectory -Force | Out-Null
    $singleResult = Write-ContextPackExcelLayoutReport -PackageDirectory $singleDirectory -Diagnostics @($diagnostics[0])
    $singleText = Get-Content -LiteralPath $singleResult.ReportPath -Raw -Encoding UTF8
    Assert-ExcelLayoutReportTest ($singleText.TrimStart().StartsWith('[')) 'A one-record report collapsed into a JSON object.'
    $singleReport = ConvertFrom-Json -InputObject $singleText
    Assert-ExcelLayoutReportTest ($singleReport.Count -eq 1) 'The one-record report does not contain exactly one diagnostic.'

    $emptyDirectory = Join-Path $testDirectory 'empty'
    New-Item -ItemType Directory -Path $emptyDirectory -Force | Out-Null
    $emptyResult = Write-ContextPackExcelLayoutReport -PackageDirectory $emptyDirectory -Diagnostics @()
    $emptyText = Get-Content -LiteralPath $emptyResult.ReportPath -Raw -Encoding UTF8
    Assert-ExcelLayoutReportTest ($emptyText -match '^\s*\[\]\s*$') 'An empty report is not an empty JSON array.'
    Assert-ExcelLayoutReportTest ($emptyResult.AutoFitApplied.Count -eq 0 -and $emptyResult.AutoFitSkipped.Count -eq 0 -and $emptyResult.Warnings.Count -eq 0) 'An empty report produced a non-empty summary.'

    $invalidPackagePath = Join-Path $testDirectory 'not-a-directory.txt'
    Set-Content -LiteralPath $invalidPackagePath -Value 'fixture' -Encoding Ascii
    $writeFailurePropagated = $false
    try {
        Write-ContextPackExcelLayoutReport -PackageDirectory $invalidPackagePath -Diagnostics @($diagnostics[0]) | Out-Null
    } catch {
        $writeFailurePropagated = $true
    }
    Assert-ExcelLayoutReportTest $writeFailurePropagated 'A report write failure was swallowed.'

    Write-Host 'Excel layout-report tests passed.' -ForegroundColor Green
} finally {
    if (Test-Path -LiteralPath $testDirectory) {
        $resolvedTestDirectory = (Resolve-Path -LiteralPath $testDirectory).Path
        $expectedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if (-not $resolvedTestDirectory.StartsWith($expectedTemp, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolvedTestDirectory) -notlike 'ContextPack-LayoutReport-Test-*') {
            throw "Refusing to clean an unexpected layout-report test directory: $resolvedTestDirectory"
        }
        Remove-Item -LiteralPath $resolvedTestDirectory -Recurse -Force
    }
}
