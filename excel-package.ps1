param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [ValidateRange(96, 300)]
    [int]$Dpi = 180
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$extractor = Join-Path $root 'extract-excel-package.py'
$renderer = Join-Path $root 'render-pdf-pages.py'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$extension = [System.IO.Path]::GetExtension($inputPath).ToLowerInvariant()

if ($extension -notin @('.xlsx', '.xlsm', '.xltx', '.xltm')) {
    throw 'Supported formats: .xlsx, .xlsm, .xltx, .xltm'
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$packageDir = Join-Path (Join-Path $root 'output') ($baseName + '_excel_package')
$sheetsDir = Join-Path $packageDir 'sheets'
New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
if (Test-Path -LiteralPath $sheetsDir) {
    Remove-Item -LiteralPath $sheetsDir -Recurse -Force
}
New-Item -ItemType Directory -Path $sheetsDir -Force | Out-Null

$env:PYTHONUTF8 = '1'
& $python $extractor $inputPath $packageDir
if ($LASTEXITCODE -ne 0) { throw "Excel extraction failed with exit code: $LASTEXITCODE" }

function Invoke-ExcelRetry {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)
    for ($attempt = 1; $attempt -le 8; $attempt++) {
        try {
            return & $Action
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 8) { throw }
            Start-Sleep -Milliseconds (300 * $attempt)
        }
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    Start-Sleep -Milliseconds 800
    Invoke-ExcelRetry { $excel.Visible = $false }
    Invoke-ExcelRetry { $excel.DisplayAlerts = $false }
    Invoke-ExcelRetry { $excel.ScreenUpdating = $false }
    $workbook = Invoke-ExcelRetry { $excel.Workbooks.Open($inputPath, 0, $true) }
    Start-Sleep -Milliseconds 500
    Write-Host ("Workbook opened. Worksheets: " + $workbook.Worksheets.Count)

    $workbookPdf = Join-Path $sheetsDir 'workbook.pdf'
    $pagesFolder = Join-Path $sheetsDir 'pages'
    New-Item -ItemType Directory -Path $pagesFolder -Force | Out-Null
    Write-Host 'Rendering visible workbook sheets'
    Invoke-ExcelRetry { $workbook.ExportAsFixedFormat(0, $workbookPdf, 0, $true, $false) }

    & $python $renderer $workbookPdf $pagesFolder --dpi $Dpi
    if ($LASTEXITCODE -ne 0) { throw 'Rendering workbook PDF failed' }
}
catch {
    Write-Error ("Excel automation failed: " + $_.Exception.Message + " | script line " + $_.InvocationInfo.ScriptLineNumber + ": " + $_.InvocationInfo.Line)
    throw
}
finally {
    if ($workbook -ne $null) {
        try { Invoke-ExcelRetry { $workbook.Close($false) } } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch { }
    }
    if ($excel -ne $null) {
        try { Invoke-ExcelRetry { $excel.Quit() } } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$manifest = @(
    "Source workbook: $([System.IO.Path]::GetFileName($inputPath))"
    "Package type: ContextPack Excel"
    "Displayed values: values.md"
    "Formulas: formulas.md"
    "Workbook structure: workbook-info.md"
    "Rendered visible sheets: sheets\workbook.pdf and sheets\pages"
    "DPI: $Dpi"
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'PACKAGE.txt') -Value $manifest -Encoding UTF8

$handoff = @(
    ('# AI Handoff — Excel: {0}' -f $baseName)
    ''
    '## რა მივაწოდო AI-ს'
    ''
    '1. `workbook-info.md` — workbook-ის რუკა.'
    '2. `formulas.md` — გამოთვლის ლოგიკა და cached შედეგები.'
    '3. `values.md` — შენახული მონაცემები.'
    ('4. `{0}` — ზუსტი შემოწმების ან ცვლილებისთვის.' -f ([System.IO.Path]::GetFileName($inputPath)))
    '5. `sheets\pages`-დან მხოლოდ საჭირო გვერდები — ვიზუალური განლაგებისთვის.'
    ''
    '## მზა მოთხოვნა'
    ''
    'ჯერ workbook-info.md-ით განსაზღვრე მნიშვნელოვანი ფურცლები და დიაპაზონები. შემდეგ formulas.md-ით შეამოწმე ლოგიკა და values.md-ით შედეგები. ორიგინალი Excel გამოიყენე საჭირო უჯრედების ზუსტი გადამოწმებისთვის. ჩემი მიზანია: [ჩაწერე მიზანი].'
    ''
    'ფარგლები: [ფურცლები/დიაპაზონები/პერიოდი].'
    'შედეგის ფორმატი: [ცხრილი/შეჯამება/შეცდომების სია/განახლებული ფაილი].'
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.ka.md') -Value $handoff -Encoding UTF8

$handoffEnglish = @(
    ('# AI Handoff — Excel: {0}' -f $baseName)
    ''
    '## What to give the AI'
    ''
    '1. `workbook-info.md` for the workbook map.'
    '2. `formulas.md` for calculation logic and cached results.'
    '3. `values.md` for stored values.'
    ('4. `{0}` for exact verification or editing.' -f ([System.IO.Path]::GetFileName($inputPath)))
    '5. Only relevant files from `sheets\pages` when visual layout matters.'
    ''
    '## Ready-to-use prompt'
    ''
    'Use workbook-info.md to identify relevant sheets and ranges. Check calculation logic in formulas.md and results in values.md. Use the original workbook only for exact cell-level verification. My goal: [describe the goal].'
    ''
    'Scope: [sheets, ranges, or period].'
    'Desired output: [table, summary, issue list, or updated file].'
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.md') -Value $handoffEnglish -Encoding UTF8

Write-Host "Excel package ready: $packageDir" -ForegroundColor Green

