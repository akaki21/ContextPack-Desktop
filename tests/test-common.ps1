$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
. (Join-Path $root 'common.ps1')

$work = Join-Path ([System.IO.Path]::GetTempPath()) ('ContextPack-Common-Test-' + [guid]::NewGuid().ToString('N'))
$output = Join-Path $work 'output'
$failures = @()

function Assert-CommonTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { $script:failures += $Message }
}

try {
    New-Item -ItemType Directory -Path $work, $output -Force | Out-Null
    $sourceA = Join-Path $work 'source-a.txt'
    $sourceB = Join-Path $work 'source-b.txt'
    Set-Content -LiteralPath $sourceA -Value 'first source' -Encoding UTF8
    Set-Content -LiteralPath $sourceB -Value 'second source' -Encoding UTF8

    $first = New-ContextPackBuild -InputPath $sourceA -PreferredName 'shared-package' -OutputDirectory $output
    Assert-CommonTest ($first.BuildPath.StartsWith($output, [System.StringComparison]::OrdinalIgnoreCase)) 'Build directory was created outside the selected output directory.'
    Set-Content -LiteralPath (Join-Path $first.BuildPath 'first.txt') -Value 'first package' -Encoding UTF8
    Write-ContextPackManifest -Build $first -InputPath $sourceA -PackageType 'test' -Outputs ([ordered]@{ artifact = 'first.txt' })
    $firstFinal = Complete-ContextPackBuild $first
    $firstManifest = Get-Content -LiteralPath (Join-Path $firstFinal 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectedHash = (Get-FileHash -LiteralPath $sourceA -Algorithm SHA256).Hash.ToLowerInvariant()
    Assert-CommonTest ($firstManifest.source.sha256 -eq $expectedHash) 'Manifest source hash is incorrect.'

    $differentSource = New-ContextPackBuild -InputPath $sourceB -PreferredName 'shared-package' -OutputDirectory $output
    Assert-CommonTest ($differentSource.FinalPath -ne $firstFinal) 'Different sources with the same name would overwrite each other.'
    Assert-CommonTest ((Split-Path -Leaf $differentSource.FinalPath) -match '^shared-package_[0-9a-f]{8}$') 'Different-source package does not include the expected short hash.'
    $unrelated = Join-Path $output 'do-not-delete.txt'
    Set-Content -LiteralPath $unrelated -Value 'keep' -Encoding UTF8
    Remove-ContextPackBuild $differentSource
    Assert-CommonTest (-not (Test-Path -LiteralPath $differentSource.BuildPath)) 'Failed temporary build was not removed.'
    Assert-CommonTest (Test-Path -LiteralPath $unrelated -PathType Leaf) 'Build cleanup removed an unrelated output file.'

    $replacement = New-ContextPackBuild -InputPath $sourceA -PreferredName 'shared-package' -OutputDirectory $output
    Assert-CommonTest ($replacement.FinalPath -eq $firstFinal) 'The same source did not select its existing package path.'
    Set-Content -LiteralPath (Join-Path $replacement.BuildPath 'second.txt') -Value 'replacement package' -Encoding UTF8
    Write-ContextPackManifest -Build $replacement -InputPath $sourceA -PackageType 'test' -Outputs ([ordered]@{ artifact = 'second.txt' })
    $replacementFinal = Complete-ContextPackBuild $replacement
    Assert-CommonTest (Test-Path -LiteralPath (Join-Path $replacementFinal 'second.txt') -PathType Leaf) 'Replacement package was not finalized.'
    Assert-CommonTest (-not (Test-Path -LiteralPath (Join-Path $replacementFinal 'first.txt'))) 'Previous package content remained after replacement.'

    if ($failures.Count) {
        $failures | ForEach-Object { Write-Error $_ }
        exit 1
    }
    Write-Host 'Common build and manifest tests passed.' -ForegroundColor Green
} finally {
    if (Test-Path -LiteralPath $work) {
        $resolved = (Resolve-Path -LiteralPath $work).Path
        $tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())
        if (-not $resolved.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolved) -notlike 'ContextPack-Common-Test-*') {
            throw "Refusing to clean an unexpected test directory: $resolved"
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
