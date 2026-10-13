param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateSet('Auto', 'Fast', 'Full', 'Ocr')][string]$Mode = 'Auto',
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [ValidateSet('Workbook', 'AutoFit', 'Both')][string]$ExcelRenderMode = 'Both',
    [ValidateRange(10, 200)][int]$MaxAutoFitColumns = 60,
    [string]$OutputDirectory,
    [ValidatePattern('^(eng|kat|rus)(\+(eng|kat|rus))*$')][string]$OcrLanguages = 'kat+eng',
    [ValidateSet('All', 'Sparse', 'None')][string]$PdfRenderMode = 'All',
    [ValidateRange(0, 100000)][int]$MaxRenderPages = 100,
    [ValidateSet('Copy', 'Reference')][string]$SourcePolicy = 'Copy',
    [switch]$Force,
    [string]$CancelFile
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
. (Join-Path $root 'contextpack-routing.ps1')

$inputPath = Resolve-ContextPackInput -InputFile $InputFile
$inputType = Get-ContextPackInputType -InputPath $inputPath
Invoke-ContextPackProcessor -Root $root -InputPath $inputPath -InputType $inputType -Mode $Mode -Dpi $Dpi -ExcelRenderMode $ExcelRenderMode -MaxAutoFitColumns $MaxAutoFitColumns -OutputDirectory $OutputDirectory -OcrLanguages $OcrLanguages -PdfRenderMode $PdfRenderMode -MaxRenderPages $MaxRenderPages -SourcePolicy $SourcePolicy -Force:$Force -CancelFile $CancelFile
