$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelModules = Join-Path $root 'powershell\Excel'
. (Join-Path $excelModules 'ContextPack.ExcelCom.ps1')
. (Join-Path $excelModules 'ContextPack.ExcelPagination.ps1')

function Assert-ExcelPaginationTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel pagination test failed: $Message" }
}

foreach ($case in @(
    @{ Columns = 1; Pages = 1 },
    @{ Columns = 8; Pages = 1 },
    @{ Columns = 9; Pages = 2 },
    @{ Columns = 16; Pages = 2 },
    @{ Columns = 17; Pages = 3 },
    @{ Columns = 24; Pages = 3 },
    @{ Columns = 60; Pages = 8 }
)) {
    $actual = Get-ContextPackExcelHorizontalPageCount -PopulatedColumnSpan $case.Columns
    Assert-ExcelPaginationTest ($actual -eq $case.Pages) "The $($case.Columns)-column boundary produced $actual page(s), expected $($case.Pages)."
}

Assert-ExcelPaginationTest ($null -eq (Get-ContextPackExcelWideSheetSkipReason -PopulatedColumnSpan 60 -MaxAutoFitColumns 60)) 'The default width limit rejected 60 columns.'
Assert-ExcelPaginationTest ((Get-ContextPackExcelWideSheetSkipReason -PopulatedColumnSpan 61 -MaxAutoFitColumns 60) -eq 'populated range exceeds the 60-column AutoFit safety limit') 'The default width safeguard did not reject 61 columns.'
Assert-ExcelPaginationTest ($null -eq (Get-ContextPackExcelWideSheetSkipReason -PopulatedColumnSpan 61 -MaxAutoFitColumns 80)) 'A custom width limit was ignored.'

$pageSetup = [pscustomobject]@{
    PrintArea = ''
    Zoom = $true
    FitToPagesWide = 1
    FitToPagesTall = 1
}
$range = [pscustomobject]@{ ThrowOnAddress = $false }
$range | Add-Member -MemberType ScriptMethod -Name Address -Value {
    if ($this.ThrowOnAddress) { throw 'Simulated range-address failure.' }
    return '$B$2:$R$10'
}
$worksheet = [pscustomobject]@{
    PageSetup = $pageSetup
    RangeObject = $range
    CellsReturned = [System.Collections.Generic.List[object]]::new()
    RangeStart = $null
    RangeEnd = $null
}
$worksheet | Add-Member -MemberType ScriptMethod -Name Cells -Value {
    param($row, $column)
    $cell = [pscustomobject]@{ Row = $row; Column = $column }
    $this.CellsReturned.Add($cell)
    return $cell
}
$worksheet | Add-Member -MemberType ScriptMethod -Name Range -Value {
    param($startCell, $endCell)
    $this.RangeStart = $startCell
    $this.RangeEnd = $endCell
    return $this.RangeObject
}

$script:releasedPaginationObjects = [System.Collections.Generic.List[object]]::new()
function Release-ExcelComObject {
    param([AllowNull()]$ComObject)
    if ($null -ne $ComObject) { $script:releasedPaginationObjects.Add($ComObject) }
}

$metric = [pscustomobject]@{
    min_row = 2
    min_column = 2
    max_row = 10
    max_column = 18
    populated_column_span = 17
    merged_ranges = 1
}
$pagination = Set-ContextPackExcelAutoFitPagination -Worksheet $worksheet -Metric $metric
Assert-ExcelPaginationTest ($pagination.PrintArea -eq '$B$2:$R$10') 'The inferred print area changed.'
Assert-ExcelPaginationTest ($pagination.FitToPagesWide -eq 3) 'The applied horizontal page count changed.'
Assert-ExcelPaginationTest ($pageSetup.PrintArea -eq '$B$2:$R$10') 'PrintArea was not applied.'
Assert-ExcelPaginationTest ($pageSetup.Zoom -eq $false) 'PageSetup Zoom was not disabled.'
Assert-ExcelPaginationTest ($pageSetup.FitToPagesWide -eq 3) 'FitToPagesWide was not applied.'
Assert-ExcelPaginationTest ($pageSetup.FitToPagesTall -eq $false) 'FitToPagesTall no longer allows unlimited vertical pages.'
Assert-ExcelPaginationTest ($pagination.Notes.Count -eq 2) 'Wide-sheet or merged-cell guidance was lost.'
Assert-ExcelPaginationTest ($script:releasedPaginationObjects.Count -eq 4) 'A PageSetup, range, or cell COM object was not released.'

$script:releasedPaginationObjects.Clear()
$range.ThrowOnAddress = $true
$failurePropagated = $false
try {
    Set-ContextPackExcelAutoFitPagination -Worksheet $worksheet -Metric $metric
} catch {
    $failurePropagated = $true
}
Assert-ExcelPaginationTest $failurePropagated 'A pagination failure was swallowed.'
Assert-ExcelPaginationTest ($script:releasedPaginationObjects.Count -eq 4) 'Pagination failure did not release all acquired objects.'

Write-Host 'Excel pagination and width-safeguard tests passed.' -ForegroundColor Green
