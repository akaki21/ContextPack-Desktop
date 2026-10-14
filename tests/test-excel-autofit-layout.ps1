$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$excelModules = Join-Path $root 'powershell\Excel'
. (Join-Path $excelModules 'ContextPack.ExcelAutoFitLayout.ps1')

function Assert-ExcelAutoFitLayoutTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel AutoFit-layout test failed: $Message" }
}

$script:rendererSkipped = $false
$script:rendererShouldFail = $false
$script:rendererCall = $null
function Invoke-ContextPackExcelAutoFitPageRenderer {
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

$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-AutoFitLayout-Test-' + [guid]::NewGuid().ToString('N'))
$renderedDirectory = Join-Path $testDirectory 'rendered-sheets'
try {
    New-Item -ItemType Directory -Path $renderedDirectory -Force | Out-Null
    $result = Invoke-ContextPackExcelAutoFitLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
        param($PdfPath)
        Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
        return [pscustomobject]@{ sheet = 'Estimate'; layout = 'AutoFit'; status = 'applied' }
    }

    $layoutDirectory = Join-Path $renderedDirectory 'auto-layout'
    Assert-ExcelAutoFitLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'workbook.pdf') -PathType Leaf) 'The AutoFit PDF path changed.'
    Assert-ExcelAutoFitLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'pages\page-001.png') -PathType Leaf) 'The AutoFit page output path changed.'
    Assert-ExcelAutoFitLayoutTest (Test-Path -LiteralPath (Join-Path $layoutDirectory 'page-render-metrics.json') -PathType Leaf) 'The AutoFit render-metrics path changed.'
    Assert-ExcelAutoFitLayoutTest ($result.Diagnostics.Count -eq 1 -and $result.Diagnostics[0].status -eq 'applied') 'AutoFit diagnostics were not returned.'
    Assert-ExcelAutoFitLayoutTest ($result.Warnings.Count -eq 0) 'A successful AutoFit render produced a warning.'
    Assert-ExcelAutoFitLayoutTest ($script:rendererCall.Dpi -eq 144 -and $script:rendererCall.MaxPages -eq 77) 'DPI or the page safety limit was not forwarded.'
    Assert-ExcelAutoFitLayoutTest ($script:rendererCall.PdfPath -eq (Join-Path $layoutDirectory 'workbook.pdf')) 'The renderer received the wrong PDF path.'

    $script:rendererSkipped = $true
    $skippedResult = Invoke-ContextPackExcelAutoFitLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
        param($PdfPath)
        Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
    }
    Assert-ExcelAutoFitLayoutTest ($skippedResult.Warnings.Count -eq 1) 'A skipped AutoFit PNG render did not produce a warning.'
    Assert-ExcelAutoFitLayoutTest ($skippedResult.Warnings[0] -eq 'Auto-layout PNG rendering skipped: page safety limit reached The complete PDF is preserved.') 'The skipped-render warning changed.'

    $script:rendererShouldFail = $true
    $rendererFailurePropagated = $false
    try {
        Invoke-ContextPackExcelAutoFitLayout -RenderedDirectory $renderedDirectory -Python 'python-fixture' -Renderer 'renderer-fixture.py' -Dpi 144 -MaxRenderedPages 77 -ExportPdf {
            param($PdfPath)
            Set-Content -LiteralPath $PdfPath -Value 'pdf fixture' -Encoding Ascii
        }
    } catch {
        $rendererFailurePropagated = $true
    }
    Assert-ExcelAutoFitLayoutTest $rendererFailurePropagated 'A renderer failure was swallowed.'

    Write-Host 'Excel AutoFit-layout orchestration tests passed.' -ForegroundColor Green
} finally {
    if (Test-Path -LiteralPath $testDirectory) {
        $resolvedTestDirectory = (Resolve-Path -LiteralPath $testDirectory).Path
        $expectedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if (-not $resolvedTestDirectory.StartsWith($expectedTemp, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolvedTestDirectory) -notlike 'ContextPack-AutoFitLayout-Test-*') {
            throw "Refusing to clean an unexpected AutoFit-layout test directory: $resolvedTestDirectory"
        }
        Remove-Item -LiteralPath $resolvedTestDirectory -Recurse -Force
    }
}
