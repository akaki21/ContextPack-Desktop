$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'common.ps1')
$python = Get-ContextPackPython
$work = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-Library-Test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $work | Out-Null

function Assert-LibraryTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Library assertion failed: $Message" }
}

try {
    $source = Join-Path $work 'norm.html'
    $out = Join-Path $work 'output'
    '<meta charset="utf-8"><p id="note">See <a href="#table">table</a>.</p><table id="table"><tr><th rowspan="2">Unit</th><th>A</th></tr><tr><td>14,5</td></tr></table>' | Set-Content -LiteralPath $source -Encoding UTF8
    $before = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    & (Join-Path $root 'contextpack.ps1') -InputFile $source -OutputDirectory $out -SourcePolicy Reference
    $package = Join-Path $out 'norm_html_package'
    $manifestPath = Join-Path $package 'manifest.json'
    $firstTime = (Get-Item -LiteralPath $manifestPath).LastWriteTimeUtc.Ticks
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-LibraryTest ($manifest.package_type -eq 'html') 'HTML was routed to the generic text converter.'
    Assert-LibraryTest (Test-Path -LiteralPath (Join-Path $package 'source-reference.json')) 'Original reference is missing.'
    Assert-LibraryTest (-not (Test-Path -LiteralPath (Join-Path $package 'original'))) 'Reference mode duplicated the original.'
    & (Join-Path $root 'contextpack.ps1') -InputFile $source -OutputDirectory $out -SourcePolicy Reference
    Assert-LibraryTest ((Get-Item -LiteralPath $manifestPath).LastWriteTimeUtc.Ticks -eq $firstTime) 'Unchanged package was regenerated.'

    'Human note must survive' | Set-Content -LiteralPath (Join-Path $package 'my-note.md') -Encoding UTF8
    & (Join-Path $root 'contextpack.ps1') -InputFile $source -OutputDirectory $out -SourcePolicy Reference
    $backups = @(Get-ChildItem -LiteralPath $out -Directory | Where-Object { $_.Name -like 'norm_html_package.previous-*' })
    Assert-LibraryTest ($backups.Count -eq 1) 'Edited previous package was not retained.'
    Assert-LibraryTest (Test-Path -LiteralPath (Join-Path $backups[0].FullName 'my-note.md')) 'Human note was lost.'
    Assert-LibraryTest ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -eq $before) 'HTML source was modified.'

    $markdown = Join-Path $work 'keep.md'
    '# Original' | Set-Content -LiteralPath $markdown -Encoding UTF8
    $rejected = $false
    try { & (Join-Path $root 'convert-to-markdown.ps1') -InputFile $markdown -OutputDirectory $work } catch { $rejected = $true }
    Assert-LibraryTest $rejected 'The text converter allowed overwriting its input.'

    $badPdf = Join-Path $work 'broken.pdf'
    'This is not a PDF' | Set-Content -LiteralPath $badPdf -Encoding Ascii
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $events = & $powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'contextpack-gui-runner.ps1') -InputFile $badPdf -OutputDirectory $out
    $code = $LASTEXITCODE
    Assert-LibraryTest ($code -ne 0) 'GUI runner reported success after PDF failure.'
    Assert-LibraryTest (-not (($events -join "`n") -match '"type":"result"')) 'GUI runner emitted a success result after failure.'
    # The child failed as expected. GitHub's PowerShell wrapper propagates
    # LASTEXITCODE after this script, so clear only this asserted failure.
    $global:LASTEXITCODE = 0
    Write-Host 'HTML routing, cache reuse, edit preservation, input protection and GUI error tests passed.'
} finally {
    if (Test-Path -LiteralPath $work) {
        $resolved = (Resolve-Path -LiteralPath $work).Path
        $expected = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        if (-not $resolved.StartsWith($expected, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolved) -notlike 'ContextPack-Library-Test-*') { throw 'Unexpected test cleanup path.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
