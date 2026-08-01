param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,
    [ValidateRange(96, 300)]
    [int]$Dpi = 180,
    [switch]$Ocr
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$renderer = Join-Path $root 'render-pdf-pages.py'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path

if ([System.IO.Path]::GetExtension($inputPath).ToLowerInvariant() -ne '.pdf') {
    throw 'This script accepts PDF files only.'
}

$baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputPath)
$packageDir = Join-Path (Join-Path $root 'output') $baseName
$pagesDir = Join-Path $packageDir 'pages'
$markdown = Join-Path $packageDir ($baseName + '.md')
$ocrPdf = Join-Path $packageDir ($baseName + '_ocr.pdf')

New-Item -ItemType Directory -Path $packageDir -Force | Out-Null
New-Item -ItemType Directory -Path $pagesDir -Force | Out-Null
$sourceCopy = Join-Path $packageDir ([System.IO.Path]::GetFileName($inputPath))
if ($inputPath -ne $sourceCopy) { Copy-Item -LiteralPath $inputPath -Destination $sourceCopy -Force }

if ($Ocr) {
    $null = Enable-ContextPackOcr
    & $python -m ocrmypdf --skip-text --rotate-pages --deskew --output-type pdf --optimize 0 -l kat+eng $inputPath $ocrPdf
    if ($LASTEXITCODE -ne 0) { throw "OCRmyPDF failed with exit code: $LASTEXITCODE" }
    & $python -m markitdown $ocrPdf -o $markdown
} else {
    & $python -m markitdown $inputPath -o $markdown
}
if ($LASTEXITCODE -ne 0) { throw "MarkItDown failed with exit code: $LASTEXITCODE" }

$env:PYTHONUTF8 = '1'
& $python $renderer $inputPath $pagesDir --dpi $Dpi
if ($LASTEXITCODE -ne 0) { throw "PDF rendering failed with exit code: $LASTEXITCODE" }

$manifest = @(
    "Source PDF: $([System.IO.Path]::GetFileName($sourceCopy))"
    "Markdown: $([System.IO.Path]::GetFileName($markdown))"
    "Page images: pages\"
    "DPI: $Dpi"
    "OCR used: $Ocr"
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'PACKAGE.txt') -Value $manifest -Encoding UTF8

$pdfReference = if ($Ocr) {
    '2. `{0}_ocr.pdf` — OCR-ისა და ფაქტების გადასამოწმებლად.' -f $baseName
} else {
    '2. ორიგინალი PDF — ფაქტებისა და განლაგების გადასამოწმებლად.'
}
$handoff = @(
    ('# AI Handoff — PDF: {0}' -f $baseName)
    ''
    '## რა მივაწოდო AI-ს'
    ''
    ('1. `{0}.md` — პირველადი სწრაფი წაკითხვისთვის.' -f $baseName)
    $pdfReference
    '3. `pages`-დან მხოლოდ საჭირო PNG გვერდები — ნახაზის, ბეჭდის, ხელმოწერის ან რთული განლაგებისთვის.'
    ''
    '## მზა მოთხოვნა'
    ''
    'ჯერ წაიკითხე Markdown და შეადგინე დოკუმენტის მოკლე რუკა. შემდეგ PDF/გვერდების სურათებში გადაამოწმე მხოლოდ მნიშვნელოვანი ან საეჭვო ადგილები. მიუთითე OCR-ის ან კონვერტაციის შესაძლო შეცდომები. ჩემი მიზანია: [ჩაწერე მიზანი]. პასუხი მინდა: [ჩაწერე ფორმატი].'
    ''
    'მნიშვნელოვანი გვერდები ან თემები: [ჩაწერე გვერდები/თემები].'
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.ka.md') -Value $handoff -Encoding UTF8

$handoffEnglish = @(
    ('# AI Handoff — PDF: {0}' -f $baseName)
    ''
    '## What to give the AI'
    ''
    ('1. `{0}.md` for the first, token-efficient reading pass.' -f $baseName)
    ('2. `{0}` for fact and layout verification.' -f ([System.IO.Path]::GetFileName($sourceCopy)))
    '3. Only relevant files from `pages` when drawings, signatures, stamps, tables, or layout matter.'
    ''
    '## Ready-to-use prompt'
    ''
    'Read the Markdown first and build a short content map. Then use the PDF or page images only to verify important or uncertain details. Flag possible OCR or conversion errors. My goal: [describe the goal]. Desired output: [describe the format].'
    ''
    'Important pages or topics: [list them].'
) -join [Environment]::NewLine
Set-Content -LiteralPath (Join-Path $packageDir 'AI-HANDOFF.md') -Value $handoffEnglish -Encoding UTF8

Write-Host "Package ready: $packageDir" -ForegroundColor Green
Write-Host "Markdown: $markdown" -ForegroundColor Green
Write-Host "Page images: $pagesDir" -ForegroundColor Green

