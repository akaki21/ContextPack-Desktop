param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateSet('Auto', 'Fast', 'Full', 'Ocr')][string]$Mode = 'Auto',
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [ValidateSet('Workbook', 'AutoFit', 'Both')][string]$ExcelRenderMode = 'Both'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'Input must be a file.' }
$extension = [System.IO.Path]::GetExtension($inputPath).ToLowerInvariant()

if ($extension -eq '.pdf') {
    if ($Mode -eq 'Fast') { & (Join-Path $root 'convert-to-markdown.ps1') $inputPath; return }
    $useOcr = $Mode -eq 'Ocr'
    if ($Mode -eq 'Auto') {
        $python = Get-ContextPackPython
        $inspectionJson = & $python (Join-Path $root 'inspect-pdf.py') $inputPath
        if ($LASTEXITCODE -ne 0) { throw 'Could not inspect the PDF for OCR auto-detection.' }
        $inspection = $inspectionJson | ConvertFrom-Json
        $useOcr = [bool]$inspection.needs_ocr
        Write-Host "Auto-detection: pages=$($inspection.page_count), OCR=$useOcr" -ForegroundColor Cyan
    }
    if ($useOcr) { & (Join-Path $root 'pdf-package.ps1') $inputPath -Dpi $Dpi -Ocr }
    else { & (Join-Path $root 'pdf-package.ps1') $inputPath -Dpi $Dpi }
    return
}

if ($extension -in @('.xlsx', '.xlsm', '.xltx', '.xltm')) {
    if ($Mode -eq 'Fast') { & (Join-Path $root 'convert-to-markdown.ps1') $inputPath }
    else { & (Join-Path $root 'excel-package.ps1') $inputPath -Dpi $Dpi -RenderMode $ExcelRenderMode }
    return
}

if ($extension -in @('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.webp')) {
    & (Join-Path $root 'ocr-image.ps1') $inputPath
    return
}

if ($Mode -eq 'Ocr') { throw 'OCR mode supports PDF and image inputs only.' }
& (Join-Path $root 'convert-to-markdown.ps1') $inputPath
return
