$ErrorActionPreference = 'Stop'

function Get-ContextPackPython {
    $python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw "ContextPack environment is missing. Run .\setup.ps1 first."
    }
    return $python
}

function Get-ContextPackTesseract {
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($env:CONTEXTPACK_TESSERACT)) { $candidates += $env:CONTEXTPACK_TESSERACT }
    $command = Get-Command tesseract.exe -ErrorAction SilentlyContinue
    if ($command) { $candidates += $command.Source }
    if ($env:ProgramFiles) { $candidates += (Join-Path $env:ProgramFiles 'Tesseract-OCR\tesseract.exe') }
    if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'Tesseract-OCR\tesseract.exe') }
    if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA 'Programs\Tesseract-OCR\tesseract.exe') }
    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    throw 'Tesseract was not found. Run .\setup.ps1 or set CONTEXTPACK_TESSERACT to tesseract.exe.'
}

function Enable-ContextPackOcr {
    $tesseract = Get-ContextPackTesseract
    $tessdata = Join-Path $PSScriptRoot 'tessdata'
    foreach ($language in @('eng', 'kat', 'osd')) {
        $model = Join-Path $tessdata ($language + '.traineddata')
        if (-not (Test-Path -LiteralPath $model -PathType Leaf)) { throw "OCR model is missing: $model. Run .\setup.ps1 again." }
    }
    $tesseractDir = Split-Path -Parent $tesseract
    if (($env:PATH -split ';') -notcontains $tesseractDir) { $env:PATH = "$tesseractDir;$env:PATH" }
    $env:TESSDATA_PREFIX = $tessdata
    return $tesseract
}
