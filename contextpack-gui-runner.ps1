param(
    [ValidateSet('Convert', 'Check', 'Setup')][string]$Action = 'Convert',
    [string]$InputFile,
    [ValidateSet('Auto', 'Fast', 'Full', 'Ocr')][string]$Mode = 'Auto',
    [ValidateRange(96, 300)][int]$Dpi = 180,
    [ValidateSet('Workbook', 'AutoFit', 'Both')][string]$ExcelRenderMode = 'Both',
    [ValidateRange(10, 200)][int]$MaxAutoFitColumns = 60,
    [string]$OutputDirectory,
    [string]$CancelFile
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$utf8 = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = $utf8
$OutputEncoding = $utf8
$script:resultPath = $null

function Write-GuiEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Type,
        [string]$Stage = '',
        [string]$Message = '',
        [Nullable[int]]$Progress = $null,
        [string]$OutputPath = ''
    )
    $event = [ordered]@{
        contextpack_event = 1
        type = $Type
        stage = $Stage
        message = $Message
        progress = $Progress
        output_path = $OutputPath
    }
    Write-Output ($event | ConvertTo-Json -Compress -Depth 4)
}

function Assert-NotCancelled {
    if ($CancelFile -and (Test-Path -LiteralPath $CancelFile)) {
        throw [System.OperationCanceledException]::new('The job was cancelled by the user.')
    }
}

function Write-ToolLine {
    param([Parameter(Mandatory = $true)]$Value)
    Assert-NotCancelled
    $line = [string]$Value
    if ([string]::IsNullOrWhiteSpace($line)) { return }

    if ($line -match '^Auto-detection:') {
        Write-GuiEvent -Type 'status' -Stage 'inspect' -Message 'PDF inspection finished.' -Progress 25
    } elseif ($line -match '^Extracted\s+\d+\s+sheets') {
        Write-GuiEvent -Type 'status' -Stage 'extract' -Message 'Workbook data extracted.' -Progress 40
    } elseif ($line -match '^Rendered\s+\d+\s+pages') {
        Write-GuiEvent -Type 'status' -Stage 'render' -Message 'Visual pages rendered.' -Progress 78
    } elseif ($line -match '^(?:PDF|Excel) package ready:\s*(.+)$') {
        $script:resultPath = $Matches[1].Trim()
        Write-GuiEvent -Type 'status' -Stage 'finalize' -Message 'Package finalized.' -Progress 95 -OutputPath $script:resultPath
    } elseif ($line -match '^Ready:\s*(.+)$') {
        $script:resultPath = $Matches[1].Trim()
        Write-GuiEvent -Type 'status' -Stage 'finalize' -Message 'Output finalized.' -Progress 95 -OutputPath $script:resultPath
    }
    Write-GuiEvent -Type 'log' -Stage 'tool' -Message $line
}

function Invoke-ContextPackScript {
    param(
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [hashtable]$Arguments = @{}
    )
    Assert-NotCancelled
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & $ScriptPath @Arguments *>&1 | ForEach-Object { Write-ToolLine $_ }
    } finally {
        $ErrorActionPreference = $previousPreference
    }
    Assert-NotCancelled
}

function Invoke-ChildPowerShellScript {
    param([Parameter(Mandatory = $true)][string]$ScriptPath)
    Assert-NotCancelled
    $powerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $previousPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        & $powerShell -NoProfile -ExecutionPolicy Bypass -File $ScriptPath *>&1 | ForEach-Object { Write-ToolLine $_ }
        $childExitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $previousPreference
    }
    if ($childExitCode -ne 0) { throw "The environment command failed with exit code $childExitCode." }
    Assert-NotCancelled
}

try {
    if ($Action -eq 'Check') {
        Write-GuiEvent -Type 'status' -Stage 'environment' -Message 'Checking the environment...' -Progress 10
        Invoke-ChildPowerShellScript -ScriptPath (Join-Path $root 'check-environment.ps1')
        Write-GuiEvent -Type 'result' -Stage 'complete' -Message 'Environment check passed.' -Progress 100
        exit 0
    }

    if ($Action -eq 'Setup') {
        Write-GuiEvent -Type 'status' -Stage 'setup' -Message 'Installing or repairing dependencies...' -Progress 10
        Invoke-ContextPackScript -ScriptPath (Join-Path $root 'setup.ps1')
        Write-GuiEvent -Type 'status' -Stage 'environment' -Message 'Verifying the environment...' -Progress 85
        Invoke-ChildPowerShellScript -ScriptPath (Join-Path $root 'check-environment.ps1')
        Write-GuiEvent -Type 'result' -Stage 'complete' -Message 'Setup and environment check completed.' -Progress 100
        exit 0
    }

    if ([string]::IsNullOrWhiteSpace($InputFile)) { throw 'Select an input file.' }
    $resolvedInput = (Resolve-Path -LiteralPath $InputFile).Path
    if (-not (Test-Path -LiteralPath $resolvedInput -PathType Leaf)) { throw 'Input must be an existing file.' }

    Write-GuiEvent -Type 'status' -Stage 'prepare' -Message 'Validating options and preparing the job...' -Progress 5
    Assert-NotCancelled
    $arguments = @{
        InputFile = $resolvedInput
        Mode = $Mode
        Dpi = $Dpi
        ExcelRenderMode = $ExcelRenderMode
        MaxAutoFitColumns = $MaxAutoFitColumns
        OutputDirectory = $OutputDirectory
    }
    Write-GuiEvent -Type 'status' -Stage 'process' -Message 'ContextPack is processing the file...' -Progress 15
    Invoke-ContextPackScript -ScriptPath (Join-Path $root 'contextpack.ps1') -Arguments $arguments

    if (-not $script:resultPath) {
        $script:resultPath = if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
            Join-Path $root 'output'
        } else {
            [System.IO.Path]::GetFullPath($OutputDirectory)
        }
    }
    Write-GuiEvent -Type 'result' -Stage 'complete' -Message 'The job completed successfully.' -Progress 100 -OutputPath $script:resultPath
    exit 0
} catch [System.OperationCanceledException] {
    Write-GuiEvent -Type 'cancelled' -Stage 'cancelled' -Message 'The job was cancelled safely.' -Progress 0
    exit 2
} catch {
    Write-GuiEvent -Type 'error' -Stage 'error' -Message $_.Exception.Message -Progress 0
    exit 1
} finally {
    if ($CancelFile -and (Test-Path -LiteralPath $CancelFile)) {
        try { Remove-Item -LiteralPath $CancelFile -Force } catch { }
    }
}
