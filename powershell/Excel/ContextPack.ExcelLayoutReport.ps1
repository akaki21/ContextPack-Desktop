# Excel: write the stable print-layout report and summarize AutoFit outcomes.

function Write-ContextPackExcelLayoutReport {
    param(
        [Parameter(Mandatory = $true)][string]$PackageDirectory,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Diagnostics
    )

    $diagnosticList = @($Diagnostics | Where-Object { $null -ne $_ })
    $reportPath = Join-Path $PackageDirectory 'print-layout-report.json'
    if ($diagnosticList.Count -eq 0) {
        Set-Content -LiteralPath $reportPath -Value '[]' -Encoding UTF8
    } else {
        ConvertTo-Json -InputObject @($diagnosticList) -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    }

    $autoDiagnostics = @($diagnosticList | Where-Object { $_.layout -eq 'AutoFit' })
    $autoApplied = @($autoDiagnostics | Where-Object { $_.status -eq 'applied' })
    $autoSkipped = @($autoDiagnostics | Where-Object { $_.status -eq 'skipped' })
    $warnings = @()

    foreach ($diagnostic in $autoSkipped) {
        $warnings += ("AutoFit skipped for sheet '{0}': {1}." -f $diagnostic.sheet, ($diagnostic.reasons -join '; '))
    }
    foreach ($diagnostic in $autoApplied | Where-Object { $_.reasons.Count -gt 0 }) {
        $warnings += ("AutoFit note for sheet '{0}': {1}." -f $diagnostic.sheet, ($diagnostic.reasons -join '; '))
    }

    return [pscustomobject]@{
        ReportPath = $reportPath
        AutoFitApplied = @($autoApplied)
        AutoFitSkipped = @($autoSkipped)
        Warnings = @($warnings)
    }
}
