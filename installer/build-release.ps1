param([string]$Version = '2.2.0')

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $root 'dist'
$stageRoot = Join-Path $dist ('.release-stage-' + [guid]::NewGuid().ToString('N'))
$packageName = "ContextPack-Desktop-$Version"
$packageRoot = Join-Path $stageRoot $packageName

function Copy-ReleaseDirectory {
    param([string]$Name)
    $source = Join-Path $root $Name
    if (Test-Path -LiteralPath $source -PathType Container) {
        Copy-Item -LiteralPath $source -Destination (Join-Path $packageRoot $Name) -Recurse -Force
    }
}

try {
    & (Join-Path $PSScriptRoot 'build-installer.ps1') -Version $Version
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed.' }

    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    $rootExtensions = @('.ps1', '.py', '.cmd', '.txt', '.md')
    foreach ($file in Get-ChildItem -LiteralPath $root -File) {
        if ($file.Extension -in $rootExtensions -or $file.Name -eq 'LICENSE') {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $packageRoot $file.Name) -Force
        }
    }
    foreach ($directory in @('assets', 'docs', 'examples')) { Copy-ReleaseDirectory $directory }
    foreach ($directory in @('input', 'output', 'tessdata')) {
        $destination = Join-Path $packageRoot $directory
        New-Item -ItemType Directory -Path $destination -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $destination '.gitkeep') -Value '' -Encoding Ascii
    }

    $baseInstaller = Join-Path $dist 'ContextPack-Setup.exe'
    $versionedInstaller = Join-Path $dist "ContextPack-Setup-$Version.exe"
    Copy-Item -LiteralPath $baseInstaller -Destination $versionedInstaller -Force

    $zip = Join-Path $dist "$packageName-portable.zip"
    if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }
    Compress-Archive -LiteralPath $packageRoot -DestinationPath $zip -CompressionLevel Optimal

    $checksumPath = Join-Path $dist "SHA256SUMS-$Version.txt"
    $checksumLines = foreach ($artifact in @($versionedInstaller, $zip)) {
        $hash = (Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToLowerInvariant()
        "$hash  $([System.IO.Path]::GetFileName($artifact))"
    }
    Set-Content -LiteralPath $checksumPath -Value $checksumLines -Encoding Ascii

    Write-Host "Release artifacts ready in: $dist" -ForegroundColor Green
    $checksumLines | ForEach-Object { Write-Host $_ -ForegroundColor Green }
} finally {
    if (Test-Path -LiteralPath $stageRoot) {
        $resolvedStage = (Resolve-Path -LiteralPath $stageRoot).Path
        $resolvedDist = (Resolve-Path -LiteralPath $dist).Path
        if ((Split-Path -Parent $resolvedStage) -ne $resolvedDist -or (Split-Path -Leaf $resolvedStage) -notlike '.release-stage-*') {
            throw "Refusing to remove an unexpected release staging directory: $resolvedStage"
        }
        Remove-Item -LiteralPath $resolvedStage -Recurse -Force
    }
}
