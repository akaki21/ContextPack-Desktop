param([switch]$SkipTesseractInstall, [switch]$ForceRecreate)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$venv = Join-Path $root '.venv'

function Find-PythonLauncher {
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) { return @{ Command = $py.Source; Arguments = @('-3') } }
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return @{ Command = $python.Source; Arguments = @() } }
    throw 'Python 3 was not found. Install Python 3.10+ from https://www.python.org/downloads/windows/ and run setup.ps1 again.'
}

$launcher = Find-PythonLauncher
$pythonCommand = $launcher.Command
$pythonArguments = $launcher.Arguments
& $pythonCommand @pythonArguments -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
if ($LASTEXITCODE -ne 0) { throw 'ContextPack requires Python 3.10 or newer.' }

if ($ForceRecreate -and (Test-Path -LiteralPath $venv)) {
    $resolvedRoot = (Resolve-Path -LiteralPath $root).Path
    $resolvedVenv = (Resolve-Path -LiteralPath $venv).Path
    if ($resolvedVenv -ne (Join-Path $resolvedRoot '.venv')) { throw "Unexpected environment path: $resolvedVenv" }
    Remove-Item -LiteralPath $resolvedVenv -Recurse -Force
}

if (-not (Test-Path -LiteralPath $venv)) {
    Write-Host 'Creating Python environment...' -ForegroundColor Cyan
    & $pythonCommand @pythonArguments -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create .venv.' }
}

$python = Join-Path $venv 'Scripts\python.exe'
Write-Host 'Installing Python dependencies...' -ForegroundColor Cyan
& $python -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw 'Failed to upgrade pip.' }
& $python -m pip install -r (Join-Path $root 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Failed to install Python dependencies.' }

. (Join-Path $root 'common.ps1')
$tesseract = $null
try { $tesseract = Get-ContextPackTesseract } catch { }
if (-not $tesseract -and -not $SkipTesseractInstall) {
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) { throw 'Tesseract is missing and winget is unavailable. Install Tesseract 5, then rerun setup.ps1.' }
    Write-Host 'Installing Tesseract OCR...' -ForegroundColor Cyan
    & $winget.Source install --id UB-Mannheim.TesseractOCR --exact --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -ne 0) { throw 'Tesseract installation failed.' }
    $tesseract = Get-ContextPackTesseract
}

New-Item -ItemType Directory -Path (Join-Path $root 'input'), (Join-Path $root 'output'), (Join-Path $root 'tessdata') -Force | Out-Null
if ($tesseract) {
    $installedTessdata = Join-Path (Split-Path -Parent $tesseract) 'tessdata'
    if (Test-Path -LiteralPath $installedTessdata) { Copy-Item -Path (Join-Path $installedTessdata '*') -Destination (Join-Path $root 'tessdata') -Recurse -Force }
    foreach ($language in @('eng', 'kat', 'osd')) {
        $destination = Join-Path (Join-Path $root 'tessdata') ($language + '.traineddata')
        if (-not (Test-Path -LiteralPath $destination)) {
            $url = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/$language.traineddata"
            Write-Host "Downloading OCR model: $language" -ForegroundColor Cyan
            Invoke-WebRequest -Uri $url -OutFile $destination
        }
    }
}

Write-Host ''
Write-Host 'ContextPack is ready.' -ForegroundColor Green
Write-Host 'Run: .\check-environment.ps1'
