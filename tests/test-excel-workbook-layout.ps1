$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelModules = Join-Path $root 'powershell\Excel'
. (Join-Path $excelModules 'ContextPack.ExcelWorkbookLayout.ps1')

function Assert-ExcelWorkbookLayoutTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel workbook-layout test failed: $Message" }
}

$script:rendererSkipped = $false
$script:rendererShouldFail = $false
$script:rendererCall = $null
function Invoke-ContextPackExcelPageRenderer {
    param(
        [string]$Python,
        [string]$Renderer,
        [string]$PdfPath,
        [string]$PagesPath,
        [string]$MetricsPath,
        [int]$Dpi,
        [int]$MaxPages
    )

    if ($script:rendererShouldFail) { throw 'Simulated renderer failure.' }
    $script:rendererCall = [pscustomobject]@{
        Python = $Python
        Renderer = $Renderer
        PdfPath = $PdfPath
        PagesPath = $PagesPath
        MetricsPath = $MetricsPath
        Dpi = $Dpi
        MaxPages = $MaxPages
    }
    Set-Content -LiteralPath (Join-Path $PagesPath 'page-001.png') -Value 'fixture' -Encoding Ascii
    [ordered]@{
        render_skipped = $script:rendererSkipped
        reason = $(if ($script:rendererSkipped) { 'page safety limit reached' } else { $null })
    } | ConvertTo-Json | Set-Content -LiteralPath $MetricsPath -Encoding UTF8
}

$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-WorkbookLayout-Test-' + [guid]::NewGuid().ToString('N'))
$renderedDirectory = Join-Path $testDirectory 'rendered-sheets'
try {
    New-Item -ItemType Directory -Path $renderedDirectory -Force | Out-Null
    $result = Invoke-ContextPackExcelWorkbookLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
        param($PdfPath)
        Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
        return [pscustomobject]@{ sheet = 'Estimate'; layout = 'Workbook'; status = 'preserved' }
    }

    $layoutDirectory = Join-Path $renderedDirectory 'workbook-layout'
    Assert-ExcelWorkbookLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'workbook.pdf') -PathType Leaf) 'The authoritative workbook PDF path changed.'
    Assert-ExcelWorkbookLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'pages\page-001.png') -PathType Leaf) 'The workbook page output path changed.'
    Assert-ExcelWorkbookLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'page-render-metrics.json') -PathType Leaf) 'The workbook render-metrics path changed.'
    Assert-ExcelWorkbookLayoutTest ($result.Diagnostics.Count -eq 1) 'Workbook diagnostics were not returned.'
    Assert-ExcelWorkbookLayoutTest ($result.Diagnostics[0].status -eq 'preserved') 'Workbook diagnostic status changed.'
    Assert-ExcelWorkbookLayoutTest ($result.Warnings.Count -eq 0) 'A successful render produced a warning.'
    Assert-ExcelWorkbookLayoutTest ($script:rendererCall.Dpi -eq 144) 'DPI was not forwarded to the renderer.'
    Assert-ExcelWorkbookLayoutTest ($script:rendererCall.MaxPages -eq 77) 'The page safety limit was not forwarded to the renderer.'

    $script:rendererSkipped = $true
    $skippedResult = Invoke-ContextPackExcelWorkbookLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
        param($PdfPath)
        Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
    }
    Assert-ExcelWorkbookLayoutTest ($skippedResult.Warnings.Count -eq 1) 'A skipped PNG render did not produce a warning.'
    Assert-ExcelWorkbookLayoutTest ($skippedResult.Warnings[0] -match 'complete PDF is preserved') 'The skipped-render warning lost the PDF fallback guidance.'

    $script:rendererShouldFail = $true
    $rendererFailurePropagated = $false
    try {
        Invoke-ContextPackExcelWorkbookLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
            param($PdfPath)
            Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
        }
    } catch {
        $rendererFailurePropagated = $true
    }
    Assert-ExcelWorkbookLayoutTest $rendererFailurePropagated 'A renderer failure was swallowed.'

    Write-Host 'Excel workbook-layout orchestration tests passed.' -ForegroundColor Green
} finally {
    if (Test-Path -LiteralPath $testDirectory) {
        $resolvedTestDirectory = (Resolve-Path -LiteralPath $testDirectory).Path
        $expectedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if (-not $resolvedTestDirectory.StartsWith($expectedTemp, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolvedTestDirectory) -notlike 'ContextPack-WorkbookLayout-Test-*') {
            throw "Refusing to clean an unexpected workbook-layout test directory: $resolvedTestDirectory"
        }
        Remove-Item -LiteralPath $resolvedTestDirectory -Recurse -Force
    }
}
