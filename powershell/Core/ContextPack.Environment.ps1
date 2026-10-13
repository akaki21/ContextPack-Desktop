# Core: locate and prepare external tools required by processing pipelines.
$script:ContextPackEnvironmentRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Invoke-ContextPackNative {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$AllowFailure,
        [ref]$ExitCode
    )
    # Windows PowerShell 5 treats redirected native stderr as ErrorRecord.
    # Tesseract/OCRmyPDF also write progress there: the exit code is authoritative.
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & $Executable @Arguments
        $nativeCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previousPreference }
    if ($ExitCode) { $ExitCode.Value = $nativeCode }
    if ($nativeCode -ne 0 -and -not $AllowFailure) { throw "Native tool failed with exit code $nativeCode : $([System.IO.Path]::GetFileName($Executable))" }
}

function Get-ContextPackPython {
    if (-not [string]::IsNullOrWhiteSpace($env:CONTEXTPACK_PYTHON)) {
        if (-not (Test-Path -LiteralPath $env:CONTEXTPACK_PYTHON -PathType Leaf)) { throw 'CONTEXTPACK_PYTHON must name an existing Python executable.' }
        return (Resolve-Path -LiteralPath $env:CONTEXTPACK_PYTHON).Path
    }
    $python = Join-Path $script:ContextPackEnvironmentRoot '.venv\Scripts\python.exe'
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

function Get-ContextPackTessdata {
    if (-not [string]::IsNullOrWhiteSpace($env:CONTEXTPACK_TESSDATA)) { return [System.IO.Path]::GetFullPath($env:CONTEXTPACK_TESSDATA) }
    return (Join-Path $script:ContextPackEnvironmentRoot 'tessdata')
}

function Enable-ContextPackOcr {
    param([ValidatePattern('^(eng|kat|rus)(\+(eng|kat|rus))*$')][string]$Languages = 'kat+eng')
    $tesseract = Get-ContextPackTesseract
    $tessdata = Get-ContextPackTessdata
    foreach ($language in @(($Languages -split '\+') + @('osd') | Select-Object -Unique)) {
        $model = Join-Path $tessdata ($language + '.traineddata')
        if (-not (Test-Path -LiteralPath $model -PathType Leaf)) { throw "OCR model is missing: $model. Run .\setup.ps1 to install English, Georgian and Russian models." }
    }
    $tesseractDir = Split-Path -Parent $tesseract
    if (($env:PATH -split ';') -notcontains $tesseractDir) { $env:PATH = "$tesseractDir;$env:PATH" }
    $env:TESSDATA_PREFIX = $tessdata
    return $tesseract
}

