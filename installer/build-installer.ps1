param(
    [string]$Version = '2.2.0',
    [switch]$InstallInnoSetup
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $PSScriptRoot 'ContextPack.iss'
$iconBuilder = Join-Path $PSScriptRoot 'make-icon.py'
$iconPath = Join-Path $root 'assets\contextpack.ico'
$dist = Join-Path $root 'dist'

function Find-Iscc {
    $command = Get-Command iscc.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    foreach ($candidate in @(
        (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
        (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe')
    )) {
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) { return $candidate }
    }
    return $null
}

$iscc = Find-Iscc
if (-not $iscc -and $InstallInnoSetup) {
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) { throw 'Inno Setup is missing and winget is unavailable.' }
    & $winget.Source install --id JRSoftware.InnoSetup --exact --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
    if ($LASTEXITCODE -ne 0) { throw "Inno Setup installation failed with exit code $LASTEXITCODE." }
    $iscc = Find-Iscc
}
if (-not $iscc) { throw 'Inno Setup 6 was not found. Rerun with -InstallInnoSetup.' }

$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'The project .venv is missing. Run setup.ps1 before building the installer.' }
& $python $iconBuilder
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $iconPath -PathType Leaf)) { throw 'ContextPack icon generation failed.' }

New-Item -ItemType Directory -Path $dist -Force | Out-Null
& $iscc "/DMyAppVersion=$Version" $scriptPath
if ($LASTEXITCODE -ne 0) { throw "Inno Setup compilation failed with exit code $LASTEXITCODE." }

$installer = Join-Path $dist 'ContextPack-Setup.exe'
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) { throw "Expected installer was not created: $installer" }
$hash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash.ToLowerInvariant()
Write-Host "Installer ready: $installer" -ForegroundColor Green
Write-Host "SHA-256: $hash" -ForegroundColor Green
