# Core: create, finalize, replace, and clean package build directories safely.
$script:ContextPackBuildRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-ContextPackSafeName {
    param([Parameter(Mandatory = $true)][string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = $Name
    foreach ($character in $invalid) { $safe = $safe.Replace([string]$character, '_') }
    $safe = $safe.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'document' }
    return $safe
}

function New-ContextPackBuild {
    param(
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$PreferredName,
        [string]$OutputDirectory
    )
    $outputRoot = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        Join-Path $script:ContextPackBuildRoot 'output'
    } else {
        [System.IO.Path]::GetFullPath($OutputDirectory)
    }
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
    $outputRoot = (Resolve-Path -LiteralPath $outputRoot).Path
    $sourceHash = (Get-FileHash -LiteralPath $InputPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $safeName = Get-ContextPackSafeName $PreferredName
    $finalPath = Join-Path $outputRoot $safeName
    if (Test-Path -LiteralPath $finalPath) {
        $existingManifest = Join-Path $finalPath 'manifest.json'
        $sameSource = $false
        if (Test-Path -LiteralPath $existingManifest) {
            try { $sameSource = ((Get-Content -LiteralPath $existingManifest -Raw -Encoding UTF8 | ConvertFrom-Json).source.sha256 -eq $sourceHash) } catch { }
        }
        if (-not $sameSource) { $finalPath = Join-Path $outputRoot ($safeName + '_' + $sourceHash.Substring(0, 8)) }
    }
    $buildPath = Join-Path $outputRoot ('.contextpack-building-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $buildPath | Out-Null
    return [pscustomobject]@{ BuildPath = $buildPath; FinalPath = $finalPath; OutputRoot = $outputRoot; SourceHash = $sourceHash }
}

function Complete-ContextPackBuild {
    param([Parameter(Mandatory = $true)]$Build)
    $buildPath = [System.IO.Path]::GetFullPath($Build.BuildPath)
    $finalPath = [System.IO.Path]::GetFullPath($Build.FinalPath)
    $outputRoot = if ($Build.OutputRoot) { [System.IO.Path]::GetFullPath($Build.OutputRoot) } else { [System.IO.Path]::GetFullPath((Join-Path $script:ContextPackBuildRoot 'output')) }
    $outputPrefix = $outputRoot.TrimEnd([char[]]@('\', '/')) + [System.IO.Path]::DirectorySeparatorChar
    if (-not $buildPath.StartsWith($outputPrefix, [System.StringComparison]::OrdinalIgnoreCase) -or -not $finalPath.StartsWith($outputPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'Refusing to finalize a package outside the output directory.'
    }
    $backupPath = $null
    try {
        if (Test-Path -LiteralPath $finalPath) {
            $backupPath = $finalPath + '.previous-' + [guid]::NewGuid().ToString('N')
            Move-Item -LiteralPath $finalPath -Destination $backupPath
        }
        Move-Item -LiteralPath $buildPath -Destination $finalPath
        if ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
            try { Remove-Item -LiteralPath $backupPath -Recurse -Force } catch { Write-Warning "The new package is ready, but the previous backup could not be removed: $backupPath" }
        }
    } catch {
        if ((-not (Test-Path -LiteralPath $finalPath)) -and $backupPath -and (Test-Path -LiteralPath $backupPath)) { Move-Item -LiteralPath $backupPath -Destination $finalPath }
        throw
    }
    return $finalPath
}

function Remove-ContextPackBuild {
    param([Parameter(Mandatory = $true)]$Build)
    if (-not $Build -or -not $Build.BuildPath) { return }
    $buildPath = [System.IO.Path]::GetFullPath($Build.BuildPath)
    $outputRoot = if ($Build.OutputRoot) { [System.IO.Path]::GetFullPath($Build.OutputRoot) } else { [System.IO.Path]::GetFullPath((Join-Path $script:ContextPackBuildRoot 'output')) }
    $outputPrefix = $outputRoot.TrimEnd([char[]]@('\', '/')) + [System.IO.Path]::DirectorySeparatorChar
    if ($buildPath.StartsWith($outputPrefix, [System.StringComparison]::OrdinalIgnoreCase) -and (Split-Path -Leaf $buildPath).StartsWith('.contextpack-building-') -and (Test-Path -LiteralPath $buildPath)) {
        Remove-Item -LiteralPath $buildPath -Recurse -Force
    }
}

