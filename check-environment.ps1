$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')

Write-Host 'Python:' -ForegroundColor Cyan
try {
    $python = Get-ContextPackPython
    & $python --version
    Write-Host 'MarkItDown:' -ForegroundColor Cyan
    & $python -m markitdown --version
    Write-Host 'OCRmyPDF:' -ForegroundColor Cyan
    & $python -m ocrmypdf --version
} catch { Write-Warning $_.Exception.Message }

Write-Host 'Tesseract and project OCR languages:' -ForegroundColor Cyan
try {
    $tesseract = Enable-ContextPackOcr
    & $tesseract --tessdata-dir (Join-Path $root 'tessdata') --list-langs
} catch { Write-Warning $_.Exception.Message }

Write-Host 'Microsoft Excel (optional, required for Excel PDF/PNG rendering):' -ForegroundColor Cyan
$excelType = [Type]::GetTypeFromProgID('Excel.Application')
if ($excelType) { Write-Host 'Available' -ForegroundColor Green } else { Write-Host 'Not detected' -ForegroundColor Yellow }
