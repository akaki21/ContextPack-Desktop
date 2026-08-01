param([switch]$SkipTesseractInstall, [switch]$ForceRecreate)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$venv = Join-Path $root '.venv'
if ($env:OS -ne 'Windows_NT') { throw 'ContextPack Desktop currently supports Windows only.' }

function Find-PythonLauncher {
    if (-not [string]::IsNullOrWhiteSpace($env:CONTEXTPACK_PYTHON) -and (Test-Path -LiteralPath $env:CONTEXTPACK_PYTHON -PathType Leaf)) {
        return @{ Command = (Resolve-Path -LiteralPath $env:CONTEXTPACK_PYTHON).Path; Arguments = @() }
    }
    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) { return @{ Command = $py.Source; Arguments = @('-3') } }
    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) { return @{ Command = $python.Source; Arguments = @() } }
    throw 'Python 3 was not found. Install Python 3.10+ from https://www.python.org/downloads/windows/ and run setup.ps1 again.'
}

function Install-VerifiedModel {
    param([string]$Language, [string]$ExpectedHash)
    $destination = Join-Path (Join-Path $root 'tessdata') ($Language + '.traineddata')
    if (Test-Path -LiteralPath $destination) {
        $actualHash = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash
        if ($actualHash -eq $ExpectedHash) { return }
    }
    $commit = '87416418657359cb625c412a48b6e1d6d41c29bd'
    $url = "https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/$commit/$Language.traineddata"
    $temporary = $destination + '.download-' + [guid]::NewGuid().ToString('N')
    try {
        Write-Host "Downloading verified OCR model: $Language" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $url -OutFile $temporary
        $downloadHash = (Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash
        if ($downloadHash -ne $ExpectedHash) { throw "Checksum verification failed for $Language.traineddata." }
        if (Test-Path -LiteralPath $destination) {
            $backup = $destination + '.backup-' + [guid]::NewGuid().ToString('N')
            try { [System.IO.File]::Replace($temporary, $destination, $backup, $true) }
            finally { if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force } }
        } else { Move-Item -LiteralPath $temporary -Destination $destination }
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

$launcher = Find-PythonLauncher
$pythonCommand = $launcher.Command
$pythonArguments = $launcher.Arguments
& $pythonCommand @pythonArguments -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)"
if ($LASTEXITCODE -ne 0) { throw 'ContextPack requires Python 3.10 or newer.' }

$venvPython = Join-Path $venv 'Scripts\python.exe'
if ((Test-Path -LiteralPath $venv) -and -not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
    $venvItem = Get-Item -LiteralPath $venv -Force
    if (($venvItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Refusing to recreate an incomplete .venv reparse point.' }
    $resolvedRoot = (Resolve-Path -LiteralPath $root).Path
    $resolvedVenv = (Resolve-Path -LiteralPath $venv).Path
    if ($resolvedVenv -ne (Join-Path $resolvedRoot '.venv')) { throw "Unexpected incomplete environment path: $resolvedVenv" }
    Write-Warning 'The existing .venv is incomplete and will be recreated.'
    Remove-Item -LiteralPath $resolvedVenv -Recurse -Force
}
if ($ForceRecreate -and (Test-Path -LiteralPath $venv)) {
    $venvItem = Get-Item -LiteralPath $venv -Force
    if (($venvItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Refusing to remove a .venv reparse point with ForceRecreate.' }
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

$python = $venvPython
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
    Install-VerifiedModel 'eng' '7D4322BD2A7749724879683FC3912CB542F19906C83BCC1A52132556427170B2'
    Install-VerifiedModel 'kat' '557ABB6F1C68BC1B286F1BDD00BB6B82F85A427A91899807DAB6C2F6C7986731'
    Install-VerifiedModel 'osd' '9CF5D576FCC47564F11265841E5CA839001E7E6F38FF7F7AACF46D15A96B00FF'
} else {
    Write-Warning 'Tesseract setup was skipped. Markdown and Excel extraction can work, but OCR commands will fail until Tesseract and OCR models are configured.'
}

Write-Host ''
Write-Host 'ContextPack is ready.' -ForegroundColor Green
Write-Host 'Run: .\check-environment.ps1'
