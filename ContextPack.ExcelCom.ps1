# Owns Excel COM retry, safety configuration, read-only opening, and cleanup.
# Keep workbook layout and rendering decisions in the calling package script.
$script:ExcelCallRejectedHResult = -2147418111

function Invoke-ExcelRetry {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    for ($attempt = 1; $attempt -le 8; $attempt++) {
        try {
            return & $Action
        } catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne $script:ExcelCallRejectedHResult -or $attempt -eq 8) { throw }
            Start-Sleep -Milliseconds (300 * $attempt)
        }
    }
}

function Release-ExcelComObject {
    param([AllowNull()]$ComObject)

    if ($null -eq $ComObject -or -not [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) { return }
    try {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject) | Out-Null
    } catch {
        # Cleanup must continue even when Excel has already invalidated an RCW.
    }
}

function Set-ContextPackExcelApplicationSafety {
    param([Parameter(Mandatory = $true)]$Application)

    Invoke-ExcelRetry { $Application.Visible = $false } | Out-Null
    Invoke-ExcelRetry { $Application.DisplayAlerts = $false } | Out-Null
    Invoke-ExcelRetry { $Application.ScreenUpdating = $false } | Out-Null
    Invoke-ExcelRetry { $Application.EnableEvents = $false } | Out-Null
    Invoke-ExcelRetry { $Application.AskToUpdateLinks = $false } | Out-Null
    Invoke-ExcelRetry { $Application.AutomationSecurity = 3 } | Out-Null
}

function New-ContextPackExcelApplication {
    $application = $null
    try {
        $application = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 800
        Set-ContextPackExcelApplicationSafety -Application $application
        return $application
    } catch {
        Close-ContextPackExcelApplication -Application $application
        throw
    }
}

function Open-ContextPackExcelWorkbook {
    param(
        [Parameter(Mandatory = $true)]$Application,
        [Parameter(Mandatory = $true)][string]$Path
    )

    return Invoke-ExcelRetry {
        $workbooks = $Application.Workbooks
        try {
            return $workbooks.Open($Path, 0, $true)
        } finally {
            Release-ExcelComObject $workbooks
        }
    }
}

function Close-ContextPackExcelWorkbook {
    param([AllowNull()]$Workbook)

    if ($null -eq $Workbook) { return }
    try {
        Invoke-ExcelRetry { $Workbook.Close($false) } | Out-Null
    } catch {
        # Release the COM wrapper even if Excel refuses the close request.
    } finally {
        Release-ExcelComObject $Workbook
    }
}

function Close-ContextPackExcelApplication {
    param([AllowNull()]$Application)

    if ($null -eq $Application) { return }
    try {
        Invoke-ExcelRetry { $Application.Quit() } | Out-Null
    } catch {
        # Release the COM wrapper even if Excel is already shutting down.
    } finally {
        Release-ExcelComObject $Application
    }
}

function Complete-ContextPackExcelComCleanup {
    param(
        [AllowNull()]$Workbook,
        [AllowNull()]$Application
    )

    Close-ContextPackExcelWorkbook -Workbook $Workbook
    Close-ContextPackExcelApplication -Application $Application
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
