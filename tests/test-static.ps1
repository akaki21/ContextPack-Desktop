$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

foreach ($file in Get-ChildItem -LiteralPath $root -Filter '*.ps1' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]\.venv[\\/]' }) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count) { $failures += "$($file.Name): $($errors.Message -join ' | ')" }
}

$sourceText = Get-ChildItem -LiteralPath $root -File -Recurse | Where-Object { $_.Extension -in @('.ps1', '.py') -and $_.FullName -notmatch '[\\/]\.venv[\\/]' } | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 }
if (($sourceText -join "`n") -match 'C:\\Users\\') { $failures += 'A user-specific Windows path was found.' }
$excelScript = Get-Content -LiteralPath (Join-Path $root 'excel-package.ps1') -Raw -Encoding UTF8
$excelComScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.ExcelCom.ps1') -Raw -Encoding UTF8
$excelDiagnosticsScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.ExcelDiagnostics.ps1') -Raw -Encoding UTF8
$excelPaginationScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.ExcelPagination.ps1') -Raw -Encoding UTF8
$excelAutoFitScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.ExcelAutoFit.ps1') -Raw -Encoding UTF8
$excelWorkbookLayoutScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.ExcelWorkbookLayout.ps1') -Raw -Encoding UTF8
if ($excelScript -notmatch 'ContextPack\.ExcelCom\.ps1') { $failures += 'Excel packaging does not load the COM lifecycle helper.' }
if ($excelScript -notmatch 'ContextPack\.ExcelDiagnostics\.ps1') { $failures += 'Excel packaging does not load the layout diagnostics helper.' }
if ($excelScript -notmatch 'ContextPack\.ExcelPagination\.ps1') { $failures += 'Excel packaging does not load the pagination helper.' }
if ($excelScript -notmatch 'ContextPack\.ExcelAutoFit\.ps1') { $failures += 'Excel packaging does not load the AutoFit decision helper.' }
if ($excelScript -notmatch 'ContextPack\.ExcelWorkbookLayout\.ps1') { $failures += 'Excel packaging does not load the workbook-layout orchestration helper.' }
if ($excelComScript -notmatch 'AutomationSecurity\s*=\s*3') { $failures += 'Excel macros are not force-disabled.' }
if ($excelComScript -notmatch 'EnableEvents\s*=\s*\$false') { $failures += 'Excel events are not disabled.' }
if ($excelComScript -notmatch 'AskToUpdateLinks\s*=\s*\$false') { $failures += 'Automatic external-link updates are not disabled.' }
if ($excelComScript -notmatch 'function\s+Release-ExcelComObject') { $failures += 'The shared Excel COM release helper is missing.' }
if ($excelScript -match 'Marshal\]::ReleaseComObject') { $failures += 'Excel packaging bypasses the shared COM release helper.' }
if ($excelDiagnosticsScript -notmatch 'Get-ExcelManualPageBreakCount') { $failures += 'Excel manual page-break diagnostics are missing.' }
if ($excelDiagnosticsScript -notmatch 'Get-ExcelShapeCount') { $failures += 'Excel shape diagnostics are missing.' }
if ($excelDiagnosticsScript -notmatch '-4135') { $failures += 'Excel manual page-break type detection is missing.' }
if ($excelAutoFitScript -notmatch 'Get-ContextPackExcelAutoFitDecision') { $failures += 'Excel AutoFit decision helper is missing.' }
if ($excelPaginationScript -notmatch 'Get-ContextPackExcelHorizontalPageCount') { $failures += 'Excel horizontal pagination planning is missing.' }
if ($excelWorkbookLayoutScript -notmatch 'Join-Path\s+\$RenderedDirectory\s+''workbook-layout''') { $failures += 'The authoritative workbook-layout output path changed.' }
if ($excelWorkbookLayoutScript -notmatch 'complete PDF is preserved') { $failures += 'Workbook rendering no longer documents the complete-PDF fallback.' }
if ($excelScript -notmatch "ValidateSet\('Workbook',\s*'AutoFit',\s*'Both'\)") { $failures += 'Excel render-mode validation is missing.' }
if ($excelScript -notmatch "RenderMode\s*=\s*'Both'") { $failures += 'Safe dual Excel rendering is not the default.' }
if ($excelScript -notmatch 'if\s*\(\$RenderMode\s+-in\s+@\(''Workbook'',\s*''Both''\)\)') { $failures += 'Both mode no longer includes the authoritative workbook layout.' }
if ($excelPaginationScript -notmatch 'Ceiling\(\$PopulatedColumnSpan\s*/\s*8\.0\)') { $failures += 'AutoFit does not adapt horizontal pagination for readability.' }
if ($excelScript -notmatch '\$maxRenderedPages\s*=\s*1000') { $failures += 'Excel page rendering does not use the reviewed safety limit.' }
if ($excelComScript -notmatch '\.Open\(\$Path,\s*0,\s*\$true\)') { $failures += 'Excel workbook is not opened read-only.' }
$contextScript = Get-Content -LiteralPath (Join-Path $root 'contextpack.ps1') -Raw -Encoding UTF8
if ($contextScript -notmatch 'MaxAutoFitColumns') { $failures += 'The main router does not expose the AutoFit width limit.' }
if ($contextScript -notmatch 'OutputDirectory') { $failures += 'The main router does not expose a custom output directory.' }
$routingScriptPath = Join-Path $root 'contextpack-routing.ps1'
. $routingScriptPath
if ((Get-ContextPackInputType 'sample.PDF') -ne 'Pdf') { $failures += 'PDF routing is not case-insensitive.' }
if ((Get-ContextPackInputType 'sample.XLSM') -ne 'Excel') { $failures += 'Excel routing does not recognize macro-enabled workbooks.' }
if ((Get-ContextPackInputType 'sample.WEBP') -ne 'Image') { $failures += 'Image routing does not recognize WebP.' }
if ((Get-ContextPackInputType 'sample.docx') -ne 'Document') { $failures += 'Generic document routing is unavailable.' }
$guiRunner = Get-Content -LiteralPath (Join-Path $root 'contextpack-gui-runner.ps1') -Raw -Encoding UTF8
if ($guiRunner -notmatch 'contextpack_event') { $failures += 'The GUI runner does not emit structured events.' }
if ($guiRunner -notmatch 'OperationCanceledException') { $failures += 'The GUI runner does not support cooperative cancellation.' }
$commonScript = Get-Content -LiteralPath (Join-Path $root 'common.ps1') -Raw -Encoding UTF8
$buildScript = Get-Content -LiteralPath (Join-Path $root 'ContextPack.Build.ps1') -Raw -Encoding UTF8
if ($buildScript -notmatch 'OutputRoot') { $failures += 'Atomic package builds do not retain their selected output root.' }
foreach ($module in @('ContextPack.Environment.ps1', 'ContextPack.Build.ps1', 'ContextPack.Manifest.ps1')) {
    if ($commonScript -notmatch [regex]::Escape($module)) { $failures += "The common compatibility loader does not include $module." }
}
$installerScript = Get-Content -LiteralPath (Join-Path $root 'installer\ContextPack.iss') -Raw -Encoding UTF8
if ($installerScript -match 'createallsubdirs') { $failures += 'Installer creates excluded directories and can break first-run setup.' }
if ($installerScript -notmatch 'PrivilegesRequired=lowest') { $failures += 'Installer is not configured for a per-user install.' }

if ($failures.Count) { $failures | ForEach-Object { Write-Error $_ }; exit 1 }
& (Join-Path $PSScriptRoot 'test-excel-com.ps1')
& (Join-Path $PSScriptRoot 'test-excel-diagnostics.ps1')
& (Join-Path $PSScriptRoot 'test-excel-pagination.ps1')
& (Join-Path $PSScriptRoot 'test-excel-autofit.ps1')
& (Join-Path $PSScriptRoot 'test-excel-workbook-layout.ps1')
Write-Host 'Static PowerShell safety checks passed.' -ForegroundColor Green
