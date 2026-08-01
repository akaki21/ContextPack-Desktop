param(
    [string]$InstallRoot = (Split-Path -Parent $PSScriptRoot),
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
$InstallRoot = [System.IO.Path]::GetFullPath($InstallRoot)
$logDirectory = Join-Path $InstallRoot 'logs'
$logPath = Join-Path $logDirectory 'install.log'
$errorPath = Join-Path $logDirectory 'last-install-error.txt'
$exitCode = 0
$errorMessage = ''

function Show-ContextPackMessage {
    param([string]$Message, [string]$Title = 'ContextPack Desktop')
    try {
        Add-Type -AssemblyName PresentationFramework
        [System.Windows.MessageBox]::Show($Message, $Title) | Out-Null
    } catch { Write-Host $Message }
}

function Test-PythonCandidate {
    param([string]$Command, [string[]]$Arguments = @())
    try {
        $resolved = & $Command @Arguments -c "import pathlib,sys; print(pathlib.Path(sys.executable).resolve()); raise SystemExit(0 if sys.version_info >= (3,10) else 1)" 2>$null
        if ($LASTEXITCODE -eq 0 -and $resolved) {
            $path = [string]($resolved | Select-Object -Last 1)
            if (Test-Path -LiteralPath $path -PathType Leaf) { return (Resolve-Path -LiteralPath $path).Path }
        }
    } catch { }
    return $null
}

function Find-CompatiblePython {
    $configured = $env:CONTEXTPACK_PYTHON
    if ($configured) {
        $result = Test-PythonCandidate -Command $configured
        if ($result) { return $result }
    }

    $py = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($py) {
        $result = Test-PythonCandidate -Command $py.Source -Arguments @('-3')
        if ($result) { return $result }
    }

    $python = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($python) {
        $result = Test-PythonCandidate -Command $python.Source
        if ($result) { return $result }
    }

    $roots = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Python'),
        (Join-Path $env:ProgramFiles 'Python312'),
        (Join-Path $env:ProgramFiles 'Python311'),
        (Join-Path $env:ProgramFiles 'Python310')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }
    foreach ($root in $roots) {
        $candidates = @(Get-ChildItem -LiteralPath $root -Filter python.exe -File -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)
        foreach ($candidate in $candidates) {
            $result = Test-PythonCandidate -Command $candidate.FullName
            if ($result) { return $result }
        }
    }
    return $null
}

function Install-PythonWithWinget {
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw 'Python 3.10+ is missing and Windows Package Manager (winget) is unavailable. Install Python 3.12 from python.org, then run ContextPack Setup / Repair.'
    }
    Write-Host 'Python 3.10+ was not found. Installing Python 3.12...' -ForegroundColor Cyan
    $arguments = @('install', '--id', 'Python.Python.3.12', '--exact', '--silent', '--scope', 'user', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
    & $winget.Source @arguments
    if ($LASTEXITCODE -ne 0) {
        Write-Warning 'The per-user Python install failed; retrying with the package default scope.'
        $arguments = @('install', '--id', 'Python.Python.3.12', '--exact', '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
        & $winget.Source @arguments
    }
    if ($LASTEXITCODE -ne 0) { throw "Python installation failed with winget exit code $LASTEXITCODE." }
}

New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
if (Test-Path -LiteralPath $errorPath) { Remove-Item -LiteralPath $errorPath -Force }

try {
    Start-Transcript -LiteralPath $logPath -Append | Out-Null
    Write-Host 'Preparing ContextPack Desktop...' -ForegroundColor Cyan
    if ($env:OS -ne 'Windows_NT') { throw 'ContextPack Desktop currently supports Windows only.' }
    if (-not (Test-Path -LiteralPath (Join-Path $InstallRoot 'setup.ps1') -PathType Leaf)) { throw "ContextPack setup.ps1 is missing from $InstallRoot." }

    $python = Find-CompatiblePython
    if (-not $python) {
        Install-PythonWithWinget
        $python = Find-CompatiblePython
    }
    if (-not $python) { throw 'Python was installed but could not be discovered. Restart Windows, then run ContextPack Setup / Repair from the Start menu.' }
    $env:CONTEXTPACK_PYTHON = $python
    Write-Host "Using Python: $python" -ForegroundColor Green

    $powerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    & $powerShell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $InstallRoot 'setup.ps1')
    if ($LASTEXITCODE -ne 0) { throw "ContextPack dependency setup failed with exit code $LASTEXITCODE." }
    & $powerShell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $InstallRoot 'check-environment.ps1')
    if ($LASTEXITCODE -ne 0) { throw "ContextPack environment verification failed with exit code $LASTEXITCODE." }
    Write-Host 'ContextPack Desktop is ready.' -ForegroundColor Green
} catch {
    $exitCode = 1
    $errorMessage = $_.Exception.Message
    $details = @(
        'ContextPack setup could not finish.',
        '',
        "Error: $errorMessage",
        "Log: $logPath",
        '',
        'ContextPack-ის დაყენება ვერ დასრულდა.',
        'გახსენი log ფაილი ან Start მენიუდან ხელახლა გაუშვი ContextPack Setup / Repair.'
    ) -join [Environment]::NewLine
    $details | Set-Content -LiteralPath $errorPath -Encoding UTF8
    Write-Error $details -ErrorAction Continue
} finally {
    try { Stop-Transcript | Out-Null } catch { }
}

if ($Interactive) {
    if ($exitCode -eq 0) {
        Show-ContextPackMessage "ContextPack Desktop is ready.`n`nContextPack Desktop მზადაა."
    } else {
        Show-ContextPackMessage "ContextPack setup could not finish.`n$errorMessage`n`nSee:`n$errorPath`n`nContextPack-ის დაყენება ვერ დასრულდა. დეტალები იხილე log ფაილში."
    }
}
exit $exitCode
