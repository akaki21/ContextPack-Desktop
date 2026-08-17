$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'ContextPack.ExcelAutoFit.ps1')

function Assert-ExcelAutoFitTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel AutoFit decision test failed: $Message" }
}

function New-ExcelAutoFitMetric {
    param(
        [int]$MaxRow = 10,
        [int]$MaxColumn = 8,
        [int]$PopulatedColumnSpan = 8,
        [int]$Charts = 0,
        [int]$Images = 0
    )

    return [pscustomobject]@{
        max_row = $MaxRow
        max_column = $MaxColumn
        populated_column_span = $PopulatedColumnSpan
        charts = $Charts
        images = $Images
    }
}

function Get-TestAutoFitDecision {
    param(
        [bool]$Visible = $true,
        [AllowNull()]$Metric = $(New-ExcelAutoFitMetric),
        [int]$ShapeCount = 0,
        [int]$HorizontalPageBreaks = 0,
        [int]$VerticalPageBreaks = 0,
        [int]$MaxAutoFitColumns = 60
    )

    return Get-ContextPackExcelAutoFitDecision -Visible $Visible -Metric $Metric -ShapeCount $ShapeCount -HorizontalPageBreaks $HorizontalPageBreaks -VerticalPageBreaks $VerticalPageBreaks -MaxAutoFitColumns $MaxAutoFitColumns
}

$hidden = Get-TestAutoFitDecision -Visible $false -Metric (New-ExcelAutoFitMetric -MaxRow 0 -PopulatedColumnSpan 100 -Charts 1) -ShapeCount 1 -HorizontalPageBreaks 1
Assert-ExcelAutoFitTest (-not $hidden.CanApply) 'A hidden sheet was accepted.'
Assert-ExcelAutoFitTest ($hidden.Reasons[0] -eq 'sheet is hidden') 'Hidden-sheet precedence changed.'

$empty = Get-TestAutoFitDecision -Metric (New-ExcelAutoFitMetric -MaxRow 0 -PopulatedColumnSpan 100 -Charts 1) -ShapeCount 1 -HorizontalPageBreaks 1
Assert-ExcelAutoFitTest (-not $empty.CanApply) 'An empty sheet was accepted.'
Assert-ExcelAutoFitTest ($empty.Reasons[0] -eq 'sheet has no populated cells') 'Empty-sheet precedence changed.'

$tooWide = Get-TestAutoFitDecision -Metric (New-ExcelAutoFitMetric -PopulatedColumnSpan 61 -Charts 1) -ShapeCount 1 -HorizontalPageBreaks 1
Assert-ExcelAutoFitTest (-not $tooWide.CanApply) 'A sheet above the width limit was accepted.'
Assert-ExcelAutoFitTest ($tooWide.Reasons[0] -eq 'populated range exceeds the 60-column AutoFit safety limit') 'Wide-sheet precedence or warning changed.'

foreach ($objectMetric in @(
    (New-ExcelAutoFitMetric -Charts 1),
    (New-ExcelAutoFitMetric -Images 1),
    (New-ExcelAutoFitMetric)
)) {
    $shapeCount = if ($objectMetric.charts -eq 0 -and $objectMetric.images -eq 0) { 1 } else { 0 }
    $drawingObject = Get-TestAutoFitDecision -Metric $objectMetric -ShapeCount $shapeCount -HorizontalPageBreaks 1
    Assert-ExcelAutoFitTest (-not $drawingObject.CanApply) 'A sheet with a drawing object was accepted.'
    Assert-ExcelAutoFitTest ($drawingObject.Reasons[0] -match 'charts, images, or drawing objects') 'Drawing-object precedence or warning changed.'
}

$manualBreak = Get-TestAutoFitDecision -HorizontalPageBreaks 1 -VerticalPageBreaks 1
Assert-ExcelAutoFitTest (-not $manualBreak.CanApply) 'A sheet with manual page breaks was accepted.'
Assert-ExcelAutoFitTest ($manualBreak.Reasons[0] -eq 'sheet contains manual page breaks') 'Manual page-break warning changed.'

$eligible = Get-TestAutoFitDecision
Assert-ExcelAutoFitTest $eligible.CanApply 'A safe populated sheet was rejected.'
Assert-ExcelAutoFitTest ($eligible.Reasons.Count -eq 0) 'An eligible sheet returned a skip reason.'

Write-Host 'Excel AutoFit decision tests passed.' -ForegroundColor Green
