$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelModules = Join-Path $root 'powershell\Excel'
. (Join-Path $excelModules 'ContextPack.ExcelCom.ps1')
. (Join-Path $excelModules 'ContextPack.ExcelDiagnostics.ps1')

function Assert-ExcelDiagnosticsTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel diagnostics test failed: $Message" }
}

function New-FakePageBreakCollection {
    param([int[]]$Types)

    $items = @($Types | ForEach-Object { [pscustomobject]@{ Type = $_ } })
    $collection = [pscustomobject]@{
        Count = $items.Count
        Items = $items
    }
    $collection | Add-Member -MemberType ScriptMethod -Name Item -Value {
        param($index)
        return $this.Items[$index - 1]
    }
    return $collection
}

$worksheet = [pscustomobject]@{
    HPageBreaks = (New-FakePageBreakCollection @(-4135, -4105, -4135))
    VPageBreaks = (New-FakePageBreakCollection @(-4105, -4135))
    Shapes = [pscustomobject]@{ Count = 4 }
}

$script:releasedPageBreaks = @()
function Release-ExcelComObject {
    param([AllowNull()]$ComObject)
    if ($null -ne $ComObject) { $script:releasedPageBreaks += $ComObject }
}

$horizontalCount = Get-ExcelManualPageBreakCount $worksheet 'HPageBreaks'
Assert-ExcelDiagnosticsTest ($horizontalCount -eq 2) 'Horizontal manual page breaks were counted incorrectly.'
Assert-ExcelDiagnosticsTest ($script:releasedPageBreaks.Count -eq 3) 'A horizontal page-break object was not released.'

$script:releasedPageBreaks = @()
$verticalCount = Get-ExcelManualPageBreakCount $worksheet 'VPageBreaks'
Assert-ExcelDiagnosticsTest ($verticalCount -eq 1) 'Vertical manual page breaks were counted incorrectly.'
Assert-ExcelDiagnosticsTest ($script:releasedPageBreaks.Count -eq 2) 'A vertical page-break object was not released.'

$shapeCount = Get-ExcelShapeCount $worksheet
Assert-ExcelDiagnosticsTest ($shapeCount -eq 4) 'Worksheet shapes were counted incorrectly.'

$brokenWorksheet = [pscustomobject]@{}
$brokenWorksheet | Add-Member -MemberType ScriptProperty -Name HPageBreaks -Value { throw 'Page breaks unavailable.' }
$brokenWorksheet | Add-Member -MemberType ScriptProperty -Name Shapes -Value { throw 'Shapes unavailable.' }
Assert-ExcelDiagnosticsTest ((Get-ExcelManualPageBreakCount $brokenWorksheet 'HPageBreaks') -eq 0) 'Page-break failure did not use the safe zero fallback.'
Assert-ExcelDiagnosticsTest ((Get-ExcelShapeCount $brokenWorksheet) -eq 0) 'Shape failure did not use the safe zero fallback.'

Write-Host 'Excel diagnostics tests passed.' -ForegroundColor Green
