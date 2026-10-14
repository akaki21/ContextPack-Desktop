param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [string]$OutputDirectory,
    [ValidatePattern('^(eng|kat|rus)(\+(eng|kat|rus))*$')][string]$OcrLanguages = 'kat+eng'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$null = Enable-ContextPackOcr -Languages $OcrLanguages
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$outputDir = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { Join-Path $root 'output' } else { [System.IO.Path]::GetFullPath($OutputDirectory) }
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$ocrPdf = Join-Path $outputDir ($baseName + '_ocr.pdf')
$markdown = Join-Path $outputDir ($baseName + '.md')

if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne '.pdf') {
    throw 'This script accepts PDF files only. Use ocr-image.ps1 for images.'
}
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Invoke-ContextPackNative -Executable $python -Arguments @('-m', 'ocrmypdf', '--skip-text', '--rotate-pages', '--deskew', '--output-type', 'pdf', '--optimize', '0', '-l', $OcrLanguages, $inputPath, $ocrPdf)

Invoke-ContextPackNative -Executable $python -Arguments @('-m', 'markitdown', $ocrPdf, '-o', $markdown)

Write-Host "OCR PDF: $ocrPdf" -ForegroundColor Green
Write-Host "Markdown: $markdown" -ForegroundColor Green
