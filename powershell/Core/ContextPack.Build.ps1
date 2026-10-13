# Core: create, finalize, replace, and clean package build directories safely.
$script:ContextPackBuildRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-ContextPackCacheSignature {
    param([Parameter(Mandatory = $true)]$Settings)
    $python = Get-ContextPackPython
    $OutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $signature = ($Settings | ConvertTo-Json -Depth 10 -Compress) | & $python (Join-Path $script:ContextPackBuildRoot 'package-cache.py') signature $script:ContextPackBuildRoot
    if ($LASTEXITCODE -ne 0) { throw 'Could not calculate the processing signature.' }
    return ([string]$signature).Trim()
}

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
        [string]$OutputDirectory,
        [string]$CacheSignature,
        [switch]$Force
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
    $preservePrevious = $false
    if (Test-Path -LiteralPath $finalPath) {
        $existingManifest = Join-Path $finalPath 'manifest.json'
        $sameSource = $false
        if (Test-Path -LiteralPath $existingManifest) {
            try { $sameSource = ((Get-Content -LiteralPath $existingManifest -Raw -Encoding UTF8 | ConvertFrom-Json).source.sha256 -eq $sourceHash) } catch { }
        }
        if (-not $sameSource) {
            $finalPath = Join-Path $outputRoot ($safeName + '_' + $sourceHash.Substring(0, 8))
            if (Test-Path -LiteralPath $finalPath) {
                $candidateManifest = Join-Path $finalPath 'manifest.json'
                $candidateHash = $null
                try { $candidateHash = (Get-Content -LiteralPath $candidateManifest -Raw -Encoding UTF8 | ConvertFrom-Json).source.sha256 } catch { }
                if ($candidateHash -ne $sourceHash) { $finalPath = Join-Path $outputRoot ($safeName + '_' + $sourceHash + '_' + [guid]::NewGuid().ToString('N').Substring(0, 8)) }
            }
        }
    }
    $finalPrefix = [System.IO.Path]::GetFullPath($finalPath).TrimEnd([char[]]@('\', '/')) + [System.IO.Path]::DirectorySeparatorChar
    if ([System.IO.Path]::GetFullPath($InputPath).StartsWith($finalPrefix, [System.StringComparison]::OrdinalIgnoreCase)) { throw 'The source is inside the destination package. Choose another output directory.' }
    if (Test-Path -LiteralPath $finalPath) {
        if (((Get-Item -LiteralPath $finalPath -Force).Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Refusing to replace a package reparse point.' }
        if ($CacheSignature) {
            $python = Get-ContextPackPython
            $cacheTool = Join-Path $script:ContextPackBuildRoot 'package-cache.py'
            $cacheJson = & $python $cacheTool probe $finalPath --signature $CacheSignature --source-hash $sourceHash
            if ($LASTEXITCODE -ne 0) { throw 'Package reuse validation failed.' }
            $cache = $cacheJson | ConvertFrom-Json
            if ($cache.valid -and -not $Force) {
                return [pscustomobject]@{ BuildPath = $null; FinalPath = $finalPath; OutputRoot = $outputRoot; SourceHash = $sourceHash; CacheHit = $true; CacheSignature = $CacheSignature; PreservePrevious = $false }
            }
            # Keep the previous directory if its generated files have been edited.
            try {
                $old = Get-Content -LiteralPath (Join-Path $finalPath 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($old.cache.signature) {
                    $oldJson = & $python $cacheTool probe $finalPath --signature $old.cache.signature --source-hash $sourceHash
                    if ($LASTEXITCODE -ne 0) { throw 'Previous package validation failed.' }
                    $preservePrevious = -not ($oldJson | ConvertFrom-Json).valid
                } else { $preservePrevious = $true }
            } catch { $preservePrevious = $true }
        }
    }
    $buildPath = Join-Path $outputRoot ('.contextpack-building-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $buildPath | Out-Null
    return [pscustomobject]@{ BuildPath = $buildPath; FinalPath = $finalPath; OutputRoot = $outputRoot; SourceHash = $sourceHash; CacheHit = $false; CacheSignature = $CacheSignature; PreservePrevious = $preservePrevious }
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
        if ($backupPath -and $Build.PreservePrevious) {
            Write-Warning "An edited or unverified previous package was preserved: $backupPath"
        } elseif ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
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

