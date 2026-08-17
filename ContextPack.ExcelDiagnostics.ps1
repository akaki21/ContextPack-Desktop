# Reads Excel layout-risk signals without changing workbook or worksheet settings.
# Retry and COM release behavior comes from ContextPack.ExcelCom.ps1.

function Get-ExcelManualPageBreakCount {
    param(
        [Parameter(Mandatory = $true)]$Worksheet,
        [Parameter(Mandatory = $true)][ValidateSet('HPageBreaks', 'VPageBreaks')][string]$PropertyName
    )

    $manualCount = 0
    try {
        if ($PropertyName -eq 'HPageBreaks') {
            $breakCount = [int](Invoke-ExcelRetry { $Worksheet.HPageBreaks.Count })
        } else {
            $breakCount = [int](Invoke-ExcelRetry { $Worksheet.VPageBreaks.Count })
        }

        for ($breakIndex = 1; $breakIndex -le $breakCount; $breakIndex++) {
            $pageBreak = $null
            try {
                if ($PropertyName -eq 'HPageBreaks') {
                    $pageBreak = Invoke-ExcelRetry { $Worksheet.HPageBreaks.Item($breakIndex) }
                } else {
                    $pageBreak = Invoke-ExcelRetry { $Worksheet.VPageBreaks.Item($breakIndex) }
                }
                if ([int](Invoke-ExcelRetry { $pageBreak.Type }) -eq -4135) { $manualCount++ }
            } finally {
                Release-ExcelComObject $pageBreak
            }
        }
        return $manualCount
    } catch {
        return 0
    }
}

function Get-ExcelShapeCount {
    param([Parameter(Mandatory = $true)]$Worksheet)

    try {
        return [int](Invoke-ExcelRetry { $Worksheet.Shapes.Count })
    } catch {
        return 0
    }
}
