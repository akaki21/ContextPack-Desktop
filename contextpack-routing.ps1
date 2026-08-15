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
        [string]$OutputDirectory
    )

    if ($Mode -eq 'Fast') {
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
    if ($useOcr) {
        & $packageScript -InputFile $InputPath -Dpi $Dpi -Ocr -OutputDirectory $OutputDirectory
    } else {
        & $packageScript -InputFile $InputPath -Dpi $Dpi -OutputDirectory $OutputDirectory
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
        [Parameter(Mandatory = $true)][ValidateSet('Pdf', 'Excel', 'Image', 'Document')][string]$InputType,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$Dpi,
        [Parameter(Mandatory = $true)][string]$ExcelRenderMode,
        [Parameter(Mandatory = $true)][int]$MaxAutoFitColumns,
        [string]$OutputDirectory
    )

    switch ($InputType) {
        'Pdf' {
            Invoke-ContextPackPdf -Root $Root -InputPath $InputPath -Mode $Mode -Dpi $Dpi -OutputDirectory $OutputDirectory
        }
        'Excel' {
            Invoke-ContextPackExcel -Root $Root -InputPath $InputPath -Mode $Mode -Dpi $Dpi -ExcelRenderMode $ExcelRenderMode -MaxAutoFitColumns $MaxAutoFitColumns -OutputDirectory $OutputDirectory
        }
        'Image' {
            & (Join-Path $Root 'ocr-image.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
        }
        'Document' {
            if ($Mode -eq 'Ocr') { throw 'OCR mode supports PDF and image inputs only.' }
            & (Join-Path $Root 'convert-to-markdown.ps1') -InputFile $InputPath -OutputDirectory $OutputDirectory
        }
    }
}

