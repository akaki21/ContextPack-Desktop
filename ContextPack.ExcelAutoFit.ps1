# Makes pure AutoFit eligibility decisions without touching Excel COM objects.
# Keep this condition order stable because the first matching reason is reported.

function Get-ContextPackExcelAutoFitDecision {
    param(
        [Parameter(Mandatory = $true)][bool]$Visible,
        [AllowNull()]$Metric,
        [Parameter(Mandatory = $true)][int]$ShapeCount,
        [Parameter(Mandatory = $true)][int]$HorizontalPageBreaks,
        [Parameter(Mandatory = $true)][int]$VerticalPageBreaks,
        [Parameter(Mandatory = $true)][ValidateRange(10, 200)][int]$MaxAutoFitColumns
    )

    $reason = $null
    $wideSheetReason = $null
    if (-not $Visible) {
        $reason = 'sheet is hidden'
    } elseif (-not $Metric -or [int]$Metric.max_row -eq 0 -or [int]$Metric.max_column -eq 0) {
        $reason = 'sheet has no populated cells'
    } elseif ($null -ne ($wideSheetReason = Get-ContextPackExcelWideSheetSkipReason -PopulatedColumnSpan ([int]$Metric.populated_column_span) -MaxAutoFitColumns $MaxAutoFitColumns)) {
        $reason = $wideSheetReason
    } elseif ([int]$Metric.charts -gt 0 -or [int]$Metric.images -gt 0 -or $ShapeCount -gt 0) {
        $reason = 'sheet contains charts, images, or drawing objects that could fall outside an inferred print area'
    } elseif (($HorizontalPageBreaks + $VerticalPageBreaks) -gt 0) {
        $reason = 'sheet contains manual page breaks'
    }

    $reasons = [System.Collections.Generic.List[string]]::new()
    if ($null -ne $reason) { $reasons.Add($reason) }
    return [pscustomobject]@{
        CanApply = $null -eq $reason
        Reasons = $reasons
    }
}
