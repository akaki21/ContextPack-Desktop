$script:ContextPackVersion = '2.0.0'

function Get-ContextPackPython {
    $python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw "ContextPack environment is missing. Run .\setup.ps1 first."
    }
    return $python
}

function Get-ContextPackTesseract {
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($env:CONTEXTPACK_TESSERACT)) { $candidates += $env:CONTEXTPACK_TESSERACT }
    $command = Get-Command tesseract.exe -ErrorAction SilentlyContinue
    if ($command) { $candidates += $command.Source }
    if ($env:ProgramFiles) { $candidates += (Join-Path $env:ProgramFiles 'Tesseract-OCR\tesseract.exe') }
    if (${env:ProgramFiles(x86)}) { $candidates += (Join-Path ${env:ProgramFiles(x86)} 'Tesseract-OCR\tesseract.exe') }
    if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA 'Programs\Tesseract-OCR\tesseract.exe') }
    foreach ($candidate in $candidates | Select-Object -Unique) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }
    throw 'Tesseract was not found. Run .\setup.ps1 or set CONTEXTPACK_TESSERACT to tesseract.exe.'
}

function Enable-ContextPackOcr {
    $tesseract = Get-ContextPackTesseract
    $tessdata = Join-Path $PSScriptRoot 'tessdata'
    foreach ($language in @('eng', 'kat', 'osd')) {
        $model = Join-Path $tessdata ($language + '.traineddata')
        if (-not (Test-Path -LiteralPath $model -PathType Leaf)) { throw "OCR model is missing: $model. Run .\setup.ps1 again." }
    }
    $tesseractDir = Split-Path -Parent $tesseract
    if (($env:PATH -split ';') -notcontains $tesseractDir) { $env:PATH = "$tesseractDir;$env:PATH" }
    $env:TESSDATA_PREFIX = $tessdata
    return $tesseract
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
        [Parameter(Mandatory = $true)][string]$PreferredName
    )
    $outputRoot = Join-Path $PSScriptRoot 'output'
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
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
    return [pscustomobject]@{ BuildPath = $buildPath; FinalPath = $finalPath; SourceHash = $sourceHash }
}

function Complete-ContextPackBuild {
    param([Parameter(Mandatory = $true)]$Build)
    $buildPath = [System.IO.Path]::GetFullPath($Build.BuildPath)
    $finalPath = [System.IO.Path]::GetFullPath($Build.FinalPath)
    $outputRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'output'))
    if (-not $buildPath.StartsWith($outputRoot + [System.IO.Path]::DirectorySeparatorChar) -or -not $finalPath.StartsWith($outputRoot + [System.IO.Path]::DirectorySeparatorChar)) {
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
    $outputRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'output'))
    if ($buildPath.StartsWith($outputRoot + [System.IO.Path]::DirectorySeparatorChar) -and (Split-Path -Leaf $buildPath).StartsWith('.contextpack-building-') -and (Test-Path -LiteralPath $buildPath)) {
        Remove-Item -LiteralPath $buildPath -Recurse -Force
    }
}

function Write-ContextPackManifest {
    param(
        [Parameter(Mandatory = $true)]$Build,
        [Parameter(Mandatory = $true)][string]$InputPath,
        [Parameter(Mandatory = $true)][string]$PackageType,
        [Parameter(Mandatory = $true)]$Outputs,
        [Parameter(Mandatory = $false)]$Settings = @{},
        [Parameter(Mandatory = $false)]$Warnings = @()
    )
    $source = Get-Item -LiteralPath $InputPath
    $manifest = [ordered]@{
        schema_version = 1
        contextpack_version = $script:ContextPackVersion
        created_utc = [DateTime]::UtcNow.ToString('o')
        package_type = $PackageType
        source = [ordered]@{ file_name = $source.Name; size_bytes = $source.Length; sha256 = $Build.SourceHash }
        outputs = $Outputs
        settings = $Settings
        warnings = @($Warnings)
    }
    $manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $Build.BuildPath 'manifest.json') -Encoding UTF8
}
