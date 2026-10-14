# Processing routes used by the stable contextpack.ps1 entry point.

function Resolve-ContextPackInput {
    param([Parameter(Mandatory = $true)][string]$InputFile)

    $resolved = (Resolve-Path -LiteralPath $InputFile).Path
    if (-not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw 'Input must be a file.'
    }
    return $resolved
}

function Get-ContextPackInputType {
    param([Parameter(Mandatory = $true)][string]$InputPath)

    $extension = [System.IO.Path]::GetExtension($InputPath).ToLowerInvariant()
    if ($extension -eq '.pdf') { return 'Pdf' }
    if ($extension -in @('.html', '.htm')) { return 'Html' }
    if ($extension -in @('.xlsx', '.xlsm', '.xltx', '.xltm')) { return 'Excel' }
    if ($extension -in @('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')) { return 'Image' }
    return 'Document'
}

function Invoke-ContextPackPdf {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [string]$OutputDirectory,
        [string]$OcrLanguages = 'kat+eng',
        [string]$PdfRenderMode = 'All',
        [int]$MaxRenderPages = 100,
        [string]$SourcePolicy = 'Copy',
        [string]$OcrPdfPolicy = 'Keep',
        [string]$PdfPages,
        [switch]$Force,
        [string]$CancelFile
    )

    if ($Mode -eq 'Fast') {
        if ($PdfPages) { throw 'PdfPages requires a PDF package mode (Auto, Full or Ocr).' }
        & (Join-Path $Root 'convert-to-markdown.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
        return
    }

    $useOcr = $Mode -eq 'Ocr'
    if ($Mode -eq 'Auto') {
        $python = Get-ContextPackPython
        $inspectionJson = & $python (Join-Path $Root 'inspect-pdf.py') $InputPath
        if ($LASTEXITCODE -ne 0) { throw 'Could not inspect the PDF for OCR auto-detection.' }
        $inspection = $inspectionJson | ConvertFrom-Json
        $useOcr = [bool]$inspection.needs_ocr
        Write-Host "Auto-detection: pages=$($inspection.page_count), OCR=$useOcr" -ForegroundColor Cyan
    }

    $packageScript = Join-Path $Root 'pdf-package.ps1'
    $packageOptions = @{ InputFile = $InputPath; Dpi = $Dpi; OutputDirectory = $OutputDirectory; OcrLanguages = $OcrLanguages; PdfRenderMode = $PdfRenderMode; MaxRenderPages = $MaxRenderPages; SourcePolicy = $SourcePolicy; OcrPdfPolicy = $OcrPdfPolicy; PdfPages = $PdfPages; Force = $Force; CancelFile = $CancelFile }
    if ($useOcr) {
        & $packageScript @packageOptions -Ocr
    } else {
        & $packageScript @packageOptions
    }
}

function Invoke-ContextPackExcel {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [Parameter(Mandatory = $true)][string]$ExcelRenderMode,
        [Parameter(Mandatory = $true)][int]$MaxAutoFitColumns,
        [string]$OutputDirectory
    )

    if ($Mode -eq 'Fast') {
        & (Join-Path $Root 'convert-to-markdown.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
        return
    }
    & (Join-Path $Root 'excel-package.ps1') -InputFile $InputPath -Dpi $Dpi -RenderMode $ExcelRenderMode -MaxAutoFitColumns $MaxAutoFitColumns -OutputDirectory $OutputDirectory
}

function Invoke-ContextPackProcessor {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][ValidateSet('Pdf', 'Html', 'Excel', 'Image', 'Document')][string]$InputType,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [Parameter(Mandatory = $true)][string]$ExcelRenderMode,
        [Parameter(Mandatory = $true)][int]$MaxAutoFitColumns,
        [string]$OutputDirectory,
        [string]$OcrLanguages = 'kat+eng',
        [string]$PdfRenderMode = 'All',
        [int]$MaxRenderPages = 100,
        [string]$SourcePolicy = 'Copy',
        [string]$OcrPdfPolicy = 'Keep',
        [string]$PdfPages,
        [switch]$Force,
        [string]$CancelFile
    )

    switch ($InputType) {
        'Pdf' {
            Invoke-ContextPackPdf -Root $Root -InputPath $InputPath -Mode $Mode -Dpi $Dpi -OutputDirectory $OutputDirectory -OcrLanguages $OcrLanguages -PdfRenderMode $PdfRenderMode -MaxRenderPages $MaxRenderPages -SourcePolicy $SourcePolicy -OcrPdfPolicy $OcrPdfPolicy -PdfPages $PdfPages -Force:$Force -CancelFile $CancelFile
        }
        'Html' {
            if ($Mode -eq 'Ocr') { throw 'OCR mode supports PDF and image inputs only.' }
            if ($Mode -eq 'Fast') {
                & (Join-Path $Root 'convert-to-markdown.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
            } else {
                & (Join-Path $Root 'html-package.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory -SourcePolicy $SourcePolicy -Force:$Force
            }
        }
        'Excel' {
            Invoke-ContextPackExcel -Root $Root -InputPath $InputPath -Mode $Mode -Dpi $Dpi -ExcelRenderMode $ExcelRenderMode -MaxAutoFitColumns $MaxAutoFitColumns -OutputDirectory $OutputDirectory
        }
        'Image' {
            & (Join-Path $Root 'ocr-image.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory -OcrLanguages $OcrLanguages
        }
        'Document' {
            if ($Mode -eq 'Ocr') { throw 'OCR mode supports PDF and image inputs only.' }
            & (Join-Path $Root 'convert-to-markdown.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
        }
    }
}

