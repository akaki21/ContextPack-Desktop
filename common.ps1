$script:ContextPackVersion = '2.2.0'

# Stable compatibility loader used by the existing PDF, Excel, OCR, and
# document-processing entry points.
. (Join-Path $PSScriptRoot 'ContextPack.Environment.ps1')
. (Join-Path $PSScriptRoot 'ContextPack.Build.ps1')
. (Join-Path $PSScriptRoot 'ContextPack.Manifest.ps1')
