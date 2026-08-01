$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = @()

foreach ($file in Get-ChildItem -LiteralPath $root -Filter '*.ps1') {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count) { $failures += "$($file.Name): $($errors.Message -join ' | ')" }
}

$sourceText = Get-ChildItem -LiteralPath $root -File | Where-Object { $_.Extension -in @('.ps1', '.py') } | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw -Encoding UTF8 }
if (($sourceText -join "`n") -match 'C:\\Users\\') { $failures += 'A user-specific Windows path was found.' }
$excelScript = Get-Content -LiteralPath (Join-Path $root 'excel-package.ps1') -Raw -Encoding UTF8
if ($excelScript -notmatch 'AutomationSecurity\s*=\s*3') { $failures += 'Excel macros are not force-disabled.' }
if ($excelScript -notmatch 'EnableEvents\s*=\s*\$false') { $failures += 'Excel events are not disabled.' }

if ($failures.Count) { $failures | ForEach-Object { Write-Error $_ }; exit 1 }
Write-Host 'Static PowerShell safety checks passed.' -ForegroundColor Green
