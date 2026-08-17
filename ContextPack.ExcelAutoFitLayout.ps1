# Orchestrates the optional AutoFit PDF/PNG output and renderer warnings.
# The supplied export action owns Excel COM work and must never save the source workbook.

function Invoke-ContextPackExcelAutoFitPageRenderer {
    param(
        [Parameter(Mandatory = $true)][string]$Python,
        [Parameter(Mandatory = $true)][string]$Renderer,
        [Parameter(Mandatory = $true)][string]$PdfPath,
        [Parameter(Mandatory = $true)][string]$PagesPath,
        [Parameter(Mandatory = $true)][string]$MetricsPath,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [Parameter(Mandatory = $true)][int]$MaxPages
    )

    & $Python $Renderer $PdfPath $PagesPath --dpi $Dpi --metrics $MetricsPath --max-pages $MaxPages | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Rendering auto-layout PDF failed' }
}

function Invoke-ContextPackExcelAutoFitLayout {
    param(
        [Parameter(Mandatory = $true)][string]$RenderedDirectory,
        [Parameter(Mandatory = $true)][string]$Python,
        [Parameter(Mandatory = $true)][string]$Renderer,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [Parameter(Mandatory = $true)][int]$MaxRenderedPages,
        [Parameter(Mandatory = $true)][scriptblock]$ExportPdf
    )

    $layoutDirectory = Join-Path $RenderedDirectory 'auto-layout'
    $pagesDirectory = Join-Path $layoutDirectory 'pages'
    New-Item -ItemType Directory -Path $pagesDirectory -Force | Out-Null
    $pdfPath = Join-Path $layoutDirectory 'workbook.pdf'
    $renderMetricsPath = Join-Path $layoutDirectory 'page-render-metrics.json'

    $diagnostics = @(& $ExportPdf $pdfPath)
    Invoke-ContextPackExcelAutoFitPageRenderer -Python $Python -Renderer $Renderer -PdfPath $pdfPath -PagesPath $pagesDirectory -MetricsPath $renderMetricsPath -Dpi $Dpi -MaxPages $MaxRenderedPages

    $warnings = @()
    $renderMetrics = Get-Content -LiteralPath $renderMetricsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($renderMetrics.render_skipped) {
        $warnings += "Auto-layout PNG rendering skipped: $($renderMetrics.reason) The complete PDF is preserved."
    }

    return [pscustomobject]@{
        Diagnostics = @($diagnostics)
        Warnings = @($warnings)
    }
}
