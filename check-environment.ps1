$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$failures = 0

function Test-NativeCommand {
    param([string]$Label, [scriptblock]$Action)
    Write-Host ($Label + ':') -ForegroundColor Cyan
    try {
        & $Action
        if ($LASTEXITCODE -ne 0) { throw "$Label exited with code $LASTEXITCODE" }
    } catch { $script:failures++; Write-Warning $_.Exception.Message }
}

try { $python = Get-ContextPackPython } catch { $failures++; Write-Warning $_.Exception.Message; $python = $null }
if ($python) {
    Test-NativeCommand 'Python' { & $python --version }
    Test-NativeCommand 'MarkItDown' { & $python -m markitdown --version }
    Test-NativeCommand 'OCRmyPDF' { & $python -m ocrmypdf --version }
}

Write-Host 'Tesseract and project OCR languages:' -ForegroundColor Cyan
try {
    $tesseract = Enable-ContextPackOcr
    & $tesseract --tessdata-dir (Join-Path $root 'tessdata') --list-langs
    if ($LASTEXITCODE -ne 0) { throw "Tesseract exited with code $LASTEXITCODE" }
} catch { $failures++; Write-Warning $_.Exception.Message }

Write-Host 'Microsoft Excel (optional):' -ForegroundColor Cyan
$excelType = [Type]::GetTypeFromProgID('Excel.Application')
if ($excelType) { Write-Host 'Available' -ForegroundColor Green } else { Write-Host 'Not detected (Excel visual rendering will be unavailable).' -ForegroundColor Yellow }

if ($failures -gt 0) { Write-Host "Environment check failed: $failures required component(s)." -ForegroundColor Red; exit 1 }
Write-Host 'Environment check passed.' -ForegroundColor Green
exit 0
