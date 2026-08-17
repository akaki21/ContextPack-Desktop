# Core: write the stable machine-readable description of a completed package.

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

