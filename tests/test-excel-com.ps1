$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'ContextPack.ExcelCom.ps1')

function Assert-ExcelComTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Excel COM test failed: $Message" }
}

$retryState = [pscustomobject]@{ Attempts = 0 }
$retryResult = Invoke-ExcelRetry {
    $retryState.Attempts++
    if ($retryState.Attempts -eq 1) {
        throw [System.Runtime.InteropServices.COMException]::new('Excel is busy.', -2147418111)
    }
    return 'retried'
}
Assert-ExcelComTest ($retryResult -eq 'retried') 'A rejected Excel call was not retried.'
Assert-ExcelComTest ($retryState.Attempts -eq 2) 'Retry used an unexpected number of attempts.'

$nonRetryState = [pscustomobject]@{ Attempts = 0 }
$nonRetryThrown = $false
try {
    Invoke-ExcelRetry {
        $nonRetryState.Attempts++
        throw [System.Runtime.InteropServices.COMException]::new('Access denied.', -2147024891)
    }
} catch [System.Runtime.InteropServices.COMException] {
    $nonRetryThrown = $true
}
Assert-ExcelComTest $nonRetryThrown 'A non-retryable COM error was swallowed.'
Assert-ExcelComTest ($nonRetryState.Attempts -eq 1) 'A non-retryable COM error was retried.'

$application = [pscustomobject]@{
    Visible = $true
    DisplayAlerts = $true
    ScreenUpdating = $true
    EnableEvents = $true
    AskToUpdateLinks = $true
    AutomationSecurity = 1
}
Set-ContextPackExcelApplicationSafety -Application $application
Assert-ExcelComTest (-not $application.Visible) 'Excel visibility was not disabled.'
Assert-ExcelComTest (-not $application.DisplayAlerts) 'Excel alerts were not disabled.'
Assert-ExcelComTest (-not $application.ScreenUpdating) 'Excel screen updating was not disabled.'
Assert-ExcelComTest (-not $application.EnableEvents) 'Excel events were not disabled.'
Assert-ExcelComTest (-not $application.AskToUpdateLinks) 'External-link prompts were not disabled.'
Assert-ExcelComTest ($application.AutomationSecurity -eq 3) 'Excel macros were not force-disabled.'

$workbooks = [pscustomobject]@{
    OpenPath = $null
    UpdateLinks = $null
    ReadOnly = $null
}
$workbooks | Add-Member -MemberType ScriptMethod -Name Open -Value {
    param($path, $updateLinks, $readOnly)
    $this.OpenPath = $path
    $this.UpdateLinks = $updateLinks
    $this.ReadOnly = $readOnly
    return [pscustomobject]@{ Marker = 'opened-workbook' }
}
$openApplication = [pscustomobject]@{ Workbooks = $workbooks }
$openedWorkbook = Open-ContextPackExcelWorkbook -Application $openApplication -Path 'fixture.xlsx'
Assert-ExcelComTest ($openedWorkbook.Marker -eq 'opened-workbook') 'Workbook open result was not returned.'
Assert-ExcelComTest ($workbooks.OpenPath -eq 'fixture.xlsx') 'Workbook path changed before opening.'
Assert-ExcelComTest ($workbooks.UpdateLinks -eq 0) 'Workbook open can update external links.'
Assert-ExcelComTest ($workbooks.ReadOnly -eq $true) 'Workbook was not opened read-only.'

$closeWorkbook = [pscustomobject]@{ ClosedWithSave = $null }
$closeWorkbook | Add-Member -MemberType ScriptMethod -Name Close -Value {
    param($saveChanges)
    $this.ClosedWithSave = $saveChanges
}
Close-ContextPackExcelWorkbook -Workbook $closeWorkbook
Assert-ExcelComTest ($closeWorkbook.ClosedWithSave -eq $false) 'Workbook cleanup requested a save.'

$closeApplication = [pscustomobject]@{ QuitCalled = $false }
$closeApplication | Add-Member -MemberType ScriptMethod -Name Quit -Value { $this.QuitCalled = $true }
Close-ContextPackExcelApplication -Application $closeApplication
Assert-ExcelComTest $closeApplication.QuitCalled 'Excel application cleanup did not call Quit.'

Release-ExcelComObject $null
Write-Host 'Excel COM lifecycle tests passed.' -ForegroundColor Green
