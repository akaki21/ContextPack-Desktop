param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateSet('Auto', 'Fast', 'Full', 'Ocr')][string]$Mode = 'Auto',
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [ValidateSet('Workbook', 'AutoFit', 'Both')][string]$ExcelRenderMode = 'Both',
    [ValidateRange(10, 200)][int]$MaxAutoFitColumns = 60,
    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
. (Join-Path $root 'contextpack-routing.ps1')

$inputPath = Resolve-ContextPackInput -InputFile $InputFile
$inputType = Get-ContextPackInputType -InputPath $inputPath
Invoke-ContextPackProcessor -Root $root -InputPath $inputPath -InputType $inputType -Mode $Mode -Dpi $Dpi -ExcelRenderMode $ExcelRenderMode -MaxAutoFitColumns $MaxAutoFitColumns -OutputDirectory $OutputDirectory
