param([Parameter(Mandatory = $true)][string]$WorkDirectory)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$work = [System.IO.Path]::GetFullPath($WorkDirectory)
$fixtures = Join-Path $work 'fixtures'
& $python (Join-Path $PSScriptRoot 'create-e2e-fixtures.py') $fixtures
if ($LASTEXITCODE -ne 0) { throw 'Fixtures failed.' }
function Assert-Storage([bool]$Ok, [string]$Message) { if (-not $Ok) { throw $Message } }
$source = Join-Path $fixtures 'sample-scan.pdf'
$before = (Get-FileHash -LiteralPath $source).Hash
$out = Join-Path $work 'output'
$options = @{ InputFile=$source; OutputDirectory=$out; Mode='Ocr'; OcrLanguages='eng'; PdfRenderMode='None'; SourcePolicy='Reference'; OcrPdfPolicy='Temporary' }
& (Join-Path $root 'contextpack.ps1') @options
$package = Join-Path $out 'sample-scan_pdf_package'
$manifestPath = Join-Path $package 'manifest.json'
$m = Get-Content $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
Assert-Storage ($m.settings.ocr_pdf_policy -eq 'Temporary') 'OCR policy is missing.'
Assert-Storage (-not $m.outputs.ocr_pdf -and -not $m.outputs.page_images) 'Temporary outputs were advertised as retained.'
Assert-Storage (@(Get-ChildItem $package -Recurse -File | Where-Object Extension -in @('.pdf','.png','.jpg')).Count -eq 0) 'Economical package retained raster/PDF files.'
Assert-Storage ((Get-Content (Join-Path $package 'page-text.md') -Raw) -match '987.65') 'Actual OCR text was not recovered.'
$review = Get-Content (Join-Path $package 'review.json') -Raw | ConvertFrom-Json
$journal = Join-Path $review.journal_directory 'events.jsonl'
$entry = @{ record_id='amount'; source_sha256=$m.source.sha256; source_revision='synthetic fixture'; locator=@{page=1;region=@(100,360,900,420);coordinate_system='original fixture pixels'}; original='987.65'; corrected='987.65'; status='source_verified'; reviewer='automated synthetic test'; method='known fixture generator'; raw_snapshot=$review.raw_snapshot }
$entryPath = Join-Path $work 'entry.json'
$entry | ConvertTo-Json -Depth 8 | Set-Content $entryPath -Encoding UTF8
& $python (Join-Path $root 'review-journal.py') append $review.journal_directory $entryPath | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Review append failed.' }
$journalHash = (Get-FileHash $journal).Hash
$time = (Get-Item $manifestPath).LastWriteTimeUtc.Ticks
& (Join-Path $root 'contextpack.ps1') @options
Assert-Storage ((Get-Item $manifestPath).LastWriteTimeUtc.Ticks -eq $time) 'Review edit invalidated raw cache reuse.'
& (Join-Path $root 'contextpack.ps1') @options -Force
Assert-Storage ((Get-FileHash $journal).Hash -eq $journalHash) 'Forced conversion lost a correction.'
Assert-Storage (@(Get-ChildItem (Join-Path $review.journal_directory 'raw') -Directory).Count -eq 1) 'Identical extraction duplicated snapshots.'
$options.OcrPdfPolicy = 'Keep'
$options.SourcePolicy = 'Copy'
& (Join-Path $root 'contextpack.ps1') @options
$m = Get-Content $manifestPath -Raw | ConvertFrom-Json
Assert-Storage (Test-Path (Join-Path $package $m.outputs.ocr_pdf)) 'Keep mode did not retain OCR PDF.'
Assert-Storage (Test-Path (Join-Path $package $m.outputs.source_pdf)) 'Portable source copy missing.'
Assert-Storage ((Get-FileHash $journal).Hash -eq $journalHash) 'Policy switch lost a correction.'
$oldManifest = (Get-FileHash $manifestPath).Hash
$token = Join-Path $work 'cancel'
New-Item -ItemType File -Path $token -Force | Out-Null
$cancelled=$false
try { & (Join-Path $root 'pdf-package.ps1') -InputFile $source -OutputDirectory $out -PdfRenderMode None -OcrPdfPolicy Temporary -CancelFile $token -Force } catch [System.OperationCanceledException] { $cancelled=$true }
Assert-Storage $cancelled 'Cancellation did not propagate.'
Assert-Storage ((Get-FileHash $manifestPath).Hash -eq $oldManifest) 'Cancelled build replaced completed package.'
Assert-Storage (@(Get-ChildItem $out -Directory -Force | Where-Object Name -like '.contextpack-building-*').Count -eq 0) 'Cancellation left build scratch.'
Assert-Storage ((Get-FileHash $source).Hash -eq $before) 'Source changed.'
Write-Host 'Temporary/Keep OCR, no binary outputs, reuse, durable review, forced conversion and cancellation passed.'
