param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$tesseract = Enable-ContextPackOcr
$tessdata = Join-Path $root 'tessdata'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$outputDir = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { Join-Path $root 'output' } else { [System.IO.Path]::GetFullPath($OutputDirectory) }
$outputBase = Join-Path $outputDir ([System.IO.Path]::GetFileNameWithoutExtension($inputPath))

New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

& $tesseract $inputPath $outputBase --tessdata-dir $tessdata -l kat+eng
if ($LASTEXITCODE -ne 0) { throw "Tesseract failed with exit code: $LASTEXITCODE" }
Write-Host "Ready: $outputBase.txt" -ForegroundColor Green
