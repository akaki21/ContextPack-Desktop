param(
    [Parameter(Mandatory = $true, Position = 0)][string]$InputFile,
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [switch]$Ocr
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$renderer = Join-Path $root 'render-pdf-pages.py'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
if (-not (Test-Path -LiteralPath $inputPath -PathType Leaf)) { throw 'Input must be a file.' }
if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne '.pdf') { throw 'This script accepts PDF files only.' }

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$build = New-ContextPackBuild -InputPath $inputPath -PreferredName ($baseName + '_pdf_package')
try {
    $packageDir = $build.BuildPath
    $pagesDir = Join-Path $packageDir 'pages'
    $markdown = Join-Path $packageDir ($baseName + '.md')
    $pageText = Join-Path $packageDir 'page-text.md'
    $metricsPath = Join-Path $packageDir 'pdf-metrics.json'
    $ocrPdf = Join-Path $packageDir ($baseName + '_ocr.pdf')
    New-Item -ItemType Directory -Path $pagesDir -Force | Out-Null
    $sourceName = [System.IO.Path]::GetFileName($inputPath)
    Copy-Item -LiteralPath $inputPath -Destination (Join-Path $packageDir $sourceName) -Force

    $textSource = $inputPath
    if ($Ocr) {
        $null = Enable-ContextPackOcr
        & $python -m ocrmypdf --skip-text --rotate-pages --deskew --output-type pdf --optimize 0 -l kat+eng $inputPath $ocrPdf
        if ($LASTEXITCODE -ne 0) { throw "OCRmyPDF failed with exit code: $LASTEXITCODE" }
        $textSource = $ocrPdf
    }

    & $python -m markitdown $textSource -o $markdown
    if ($LASTEXITCODE -ne 0) { throw "MarkItDown failed with exit code: $LASTEXITCODE" }
    $env:PYTHONUTF8 = '1'
    & $python $renderer $textSource $pagesDir --dpi $Dpi --page-text $pageText --metrics $metricsPath
    if ($LASTEXITCODE -ne 0) { throw "PDF rendering failed with exit code: $LASTEXITCODE" }

    $metrics = Get-Content -LiteralPath $metricsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $warnings = @()
    if ($metrics.sparse_pages.Count -gt 0) { $warnings += ('Sparse or missing text on page(s): ' + ($metrics.sparse_pages -join ', ')) }
    $quality = @(
        "# Quality report — $sourceName"
        ''
        "- Pages: $($metrics.page_count)"
        "- Extracted characters: $($metrics.total_characters)"
        "- OCR used: $Ocr"
        "- OCR languages: $(if ($Ocr) { 'Georgian + English (kat+eng)' } else { 'not applicable' })"
        "- Sparse-text pages: $(if ($metrics.sparse_pages.Count) { $metrics.sparse_pages -join ', ' } else { 'none' })"
        ''
        '## Review guidance'
        ''
        $(if ($warnings.Count) { '- Verify sparse-text pages against the source PDF or PNG pages.' } else { '- No page-level text coverage warning was detected.' })
        '- Verify critical names, figures, signatures, stamps, tables, and drawings against the source.'
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath (Join-Path $packageDir 'quality-report.md') -Value $quality -Encoding UTF8

    $handoffEnglish = @(
        "# AI Handoff — PDF: $baseName", '', 'Read the main Markdown first. Use `page-text.md` to map extracted text to source page numbers, `quality-report.md` to find pages requiring review, and the source PDF/selected PNG pages for exact visual verification.', '', 'Goal: [describe the goal]', 'Scope: [pages/topics]', 'Desired output: [format]'
    ) -join [Environment]::NewLine
    $handoffGeorgian = @(
        "# AI Handoff — PDF: $baseName", '', 'ჯერ წაიკითხე მთავარი Markdown. `page-text.md` გამოიყენე ტექსტის საწყის გვერდებთან დასაკავშირებლად, `quality-report.md` — გადასამოწმებელი გვერდების საპოვნელად, ხოლო საწყისი PDF და შერჩეული PNG გვერდები — ზუსტი ვიზუალური შემოწმებისთვის.', '', 'მიზანი: [აღწერე მიზანი]', 'ფარგლები: [გვერდები/თემები]', 'შედეგი: [ფორმატი]'
    ) -join [Environment]::NewLine
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.md') -Value $handoffEnglish -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.ka.md') -Value $handoffGeorgian -Encoding UTF8

    $outputs = [ordered]@{
        source_pdf = $sourceName
        markdown = [System.IO.Path]::GetFileName($markdown)
        page_text = 'page-text.md'
        page_images = 'pages/'
        ocr_pdf = $(if ($Ocr) { [System.IO.Path]::GetFileName($ocrPdf) } else { $null })
        quality_report = 'quality-report.md'
    }
    Write-ContextPackManifest -Build $build -InputPath $inputPath -PackageType 'pdf' -Outputs $outputs -Settings ([ordered]@{ dpi = $Dpi; ocr = [bool]$Ocr; languages = $(if ($Ocr) { @('kat','eng') } else { @() }) }) -Warnings $warnings
    $finalPath = Complete-ContextPackBuild $build
    Write-Host "PDF package ready: $finalPath" -ForegroundColor Green
} catch {
    Remove-ContextPackBuild $build
    throw
}
