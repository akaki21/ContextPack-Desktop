param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [Parameter(Position = 1)]
    [string]$OutputFile,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'Input must be a file.' }

if (-not [string]::IsNullOrWhiteSpace($OutputFile) -and -not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    throw 'Use OutputFile or OutputDirectory, not both.'
}

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $outputDir = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { Join-Path $root 'output' } else { [System.IO.Path]::GetFullPath($OutputDirectory) }
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    $OutputFile = Join-Path $outputDir (([System.IO.Path]::GetFileNameWithoutExtension($inputPath)) + '.md')
}
else {
    $parent = Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputFile))
    if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
}

& $python -m markitdown $inputPath -o $OutputFile
if ($LASTEXITCODE -ne 0) { throw "MarkItDown failed with exit code: $LASTEXITCODE" }
Write-Host "Ready: $OutputFile" -ForegroundColor Green
