param(
    [string]$WorkDirectory,
    [switch]$RequireExcel,
    [switch]$SkipOcr,
    [switch]$SkipPdfOcr,
    [switch]$KeepArtifacts
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$python = Join-Path $root '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $python -PathType Leaf)) { throw 'Run setup.ps1 before the end-to-end tests.' }

$ownsWorkDirectory = [string]::IsNullOrWhiteSpace($WorkDirectory)
if ($ownsWorkDirectory) {
    $WorkDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-E2E-' + [guid]::NewGuid().ToString('N'))
}
$work = [System.IO.Path]::GetFullPath($WorkDirectory)
$fixtures = Join-Path $work 'fixtures'
$output = Join-Path $work 'output'
New-Item -ItemType Directory -Path $fixtures, $output -Force | Out-Null

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "E2E assertion failed: $Message" }
}

function Read-JsonFile {
    param([string]$Path)
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

try {
    & $python (Join-Path $PSScriptRoot 'create-e2e-fixtures.py') $fixtures
    if ($LASTEXITCODE -ne 0) { throw 'Could not create E2E fixtures.' }

    # Markdown router: copy a Markdown file through MarkItDown.
    & (Join-Path $root 'contextpack.ps1') (Join-Path $fixtures 'sample.md') -Mode Fast -OutputDirectory $output
    $markdownOutput = Join-Path $output 'sample.md'
    Assert-True (Test-Path -LiteralPath $markdownOutput -PathType Leaf) 'Markdown output is missing.'
    Assert-True ((Get-Content -LiteralPath $markdownOutput -Raw -Encoding UTF8) -match 'ContextPack Markdown fixture') 'Markdown text changed or disappeared.'

    # PDF package without OCR: source, Markdown, page text, images and manifest.
    & (Join-Path $root 'contextpack.ps1') (Join-Path $fixtures 'sample-scan.pdf') -Mode Full -Dpi 96 -OutputDirectory $output
    $pdfPackage = Join-Path $output 'sample-scan_pdf_package'
    $pdfManifest = Read-JsonFile (Join-Path $pdfPackage 'manifest.json')
    Assert-True ($pdfManifest.package_type -eq 'pdf') 'PDF manifest has the wrong package type.'
    Assert-True (-not [bool]$pdfManifest.settings.ocr) 'Full PDF test unexpectedly enabled OCR.'
    Assert-True (Test-Path -LiteralPath (Join-Path $pdfPackage 'pages\page-001.png') -PathType Leaf) 'Rendered PDF page is missing.'
    Assert-True (Test-Path -LiteralPath (Join-Path $pdfPackage 'quality-report.md') -PathType Leaf) 'PDF quality report is missing.'

    # Excel extraction is portable; visual rendering is added when Microsoft Excel is available.
    $excelAvailable = $null -ne [Type]::GetTypeFromProgID('Excel.Application')
    if ($RequireExcel -and -not $excelAvailable) { throw 'Microsoft Excel is required for this E2E run but was not detected.' }
    $excelFixture = Join-Path $fixtures 'sample.xlsx'
    if ($excelAvailable) {
        & (Join-Path $root 'contextpack.ps1') $excelFixture -Mode Full -Dpi 96 -ExcelRenderMode Workbook -OutputDirectory $output
        $excelPackage = Join-Path $output 'sample_excel_package'
        $excelManifest = Read-JsonFile (Join-Path $excelPackage 'manifest.json')
        Assert-True ($excelManifest.package_type -eq 'excel') 'Excel manifest has the wrong package type.'
        Assert-True ($excelManifest.settings.render_mode -eq 'Workbook') 'Excel render mode was not recorded.'
        Assert-True (Test-Path -LiteralPath (Join-Path $excelPackage 'rendered-sheets\workbook-layout\pages\page-001.png') -PathType Leaf) 'Excel rendered page is missing.'
    } else {
        $excelPackage = Join-Path $output 'sample_excel_extracted'
        & $python (Join-Path $root 'extract-excel-package.py') $excelFixture $excelPackage
        if ($LASTEXITCODE -ne 0) { throw 'Portable Excel extraction failed.' }
    }
    $formulaText = Get-Content -LiteralPath (Join-Path $excelPackage 'sheets-data\01-Estimate\formulas.md') -Raw -Encoding UTF8
    Assert-True ($formulaText -match '=B2\*C2') 'Excel formula B2*C2 was not preserved.'
    Assert-True ((Get-FileHash -LiteralPath $excelFixture -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath (Join-Path $excelPackage 'sample.xlsx') -Algorithm SHA256).Hash) 'Packaged workbook differs from the source.'

    if (-not $SkipOcr) {
        # Image OCR through the main router.
        & (Join-Path $root 'contextpack.ps1') (Join-Path $fixtures 'sample-scan.png') -Mode Ocr -OutputDirectory $output
        $imageText = Get-Content -LiteralPath (Join-Path $output 'sample-scan.txt') -Raw -Encoding UTF8
        Assert-True ($imageText -match 'CONTEXT\s+PACK\s+OCR') 'Image OCR did not recover the expected text.'

        if (-not $SkipPdfOcr) {
            # Scanned PDF OCR through OCRmyPDF and the complete PDF packaging path.
            $ocrOutput = Join-Path $work 'ocr-output'
            & (Join-Path $root 'contextpack.ps1') (Join-Path $fixtures 'sample-scan.pdf') -Mode Ocr -Dpi 96 -OutputDirectory $ocrOutput
            $ocrPackage = Join-Path $ocrOutput 'sample-scan_pdf_package'
            $ocrManifest = Read-JsonFile (Join-Path $ocrPackage 'manifest.json')
            Assert-True ([bool]$ocrManifest.settings.ocr) 'OCR PDF manifest did not record OCR.'
            Assert-True (Test-Path -LiteralPath (Join-Path $ocrPackage 'sample-scan_ocr.pdf') -PathType Leaf) 'Searchable OCR PDF is missing.'
            $ocrMarkdown = Get-Content -LiteralPath (Join-Path $ocrPackage 'sample-scan.md') -Raw -Encoding UTF8
            Assert-True ($ocrMarkdown -match 'CONTEXT\s+PACK\s+OCR') 'OCR PDF Markdown did not recover the expected text.'
        }
    }

    Write-Host "ContextPack E2E tests passed: $work" -ForegroundColor Green
} finally {
    if ($ownsWorkDirectory -and -not $KeepArtifacts -and (Test-Path -LiteralPath $work)) {
        $expectedPrefix = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if (-not $work.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $work) -notlike 'ContextPack-E2E-*') {
            throw "Refusing to clean an unexpected test directory: $work"
        }
        Remove-Item -LiteralPath $work -Recurse -Force
    }
}
