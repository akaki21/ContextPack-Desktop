param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [string]$OutputDirectory,
    [ValidatePattern('^(eng|kat|rus)(\+(eng|kat|rus))*$')][string]$OcrLanguages = 'kat+eng'
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$tesseract = Enable-ContextPackOcr -Languages $OcrLanguages
$tessdata = Get-ContextPackTessdata
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$outputDir = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { Join-Path $root 'output' } else { [System.IO.Path]::GetFullPath($OutputDirectory) }
$outputBase = Join-Path $outputDir ([System.IO.Path]::GetFileNameWithoutExtension($inputPath))

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

Invoke-ContextPackNative -Executable $tesseract -Arguments @($inputPath, $outputBase, '--tessdata-dir', $tessdata, '-l', $OcrLanguages)
Write-Host "Ready: $outputBase.txt" -ForegroundColor Green
