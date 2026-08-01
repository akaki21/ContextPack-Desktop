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
if ($excelScript -notmatch 'AutomationSecurity\s*=\s*3') { $failures += 'Excel macros are not force-disabled.' }
if ($excelScript -notmatch 'EnableEvents\s*=\s*\$false') { $failures += 'Excel events are not disabled.' }
if ($excelScript -notmatch 'AskToUpdateLinks\s*=\s*\$false') { $failures += 'Automatic external-link updates are not disabled.' }
if ($excelScript -notmatch "ValidateSet\('Workbook',\s*'AutoFit',\s*'Both'\)") { $failures += 'Excel render-mode validation is missing.' }
if ($excelScript -notmatch "RenderMode\s*=\s*'Both'") { $failures += 'Safe dual Excel rendering is not the default.' }
if ($excelScript -notmatch 'Ceiling\(\[int\]\$metric\.populated_column_span\s*/\s*8\.0\)') { $failures += 'AutoFit does not adapt horizontal pagination for readability.' }
if ($excelScript -notmatch '\$maxRenderedPages\s*=\s*1000') { $failures += 'Excel page rendering does not use the reviewed safety limit.' }
if ($excelScript -notmatch 'Workbooks\.Open\(\$inputPath,\s*0,\s*\$true\)') { $failures += 'Excel workbook is not opened read-only.' }
$contextScript = Get-Content -LiteralPath (Join-Path $root 'contextpack.ps1') -Raw -Encoding UTF8
if ($contextScript -notmatch 'MaxAutoFitColumns') { $failures += 'The main router does not expose the AutoFit width limit.' }
if ($contextScript -notmatch 'OutputDirectory') { $failures += 'The main router does not expose a custom output directory.' }
$guiRunner = Get-Content -LiteralPath (Join-Path $root 'contextpack-gui-runner.ps1') -Raw -Encoding UTF8
if ($guiRunner -notmatch 'contextpack_event') { $failures += 'The GUI runner does not emit structured events.' }
if ($guiRunner -notmatch 'OperationCanceledException') { $failures += 'The GUI runner does not support cooperative cancellation.' }
$commonScript = Get-Content -LiteralPath (Join-Path $root 'common.ps1') -Raw -Encoding UTF8
if ($commonScript -notmatch 'OutputRoot') { $failures += 'Atomic package builds do not retain their selected output root.' }
$installerScript = Get-Content -LiteralPath (Join-Path $root 'installer\ContextPack.iss') -Raw -Encoding UTF8
if ($installerScript -match 'createallsubdirs') { $failures += 'Installer creates excluded directories and can break first-run setup.' }
if ($installerScript -notmatch 'PrivilegesRequired=lowest') { $failures += 'Installer is not configured for a per-user install.' }

if ($failures.Count) { $failures | ForEach-Object { Write-Error $_ }; exit 1 }
Write-Host 'Static PowerShell safety checks passed.' -ForegroundColor Green
