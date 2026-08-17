$script:ContextPackVersion = '2.2.0'

# Stable compatibility loader used by the existing PDF, Excel, OCR, and
# document-processing entry points.
$coreModules = Join-Path $PSScriptRoot 'powershell\Core'
. (Join-Path $coreModules 'ContextPack.Environment.ps1')
. (Join-Path $coreModules 'ContextPack.Build.ps1')
. (Join-Path $coreModules 'ContextPack.Manifest.ps1')
