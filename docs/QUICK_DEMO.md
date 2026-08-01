# ContextPack Desktop — 60-second demo

This demo uses the repository's tiny, non-sensitive `examples\sample.csv` file. Complete `setup.ps1` once before starting.

## GUI

1. Double-click `Start-ContextPack-GUI.cmd`.
2. Choose `examples\sample.csv`.
3. Keep **Automatic — recommended**.
4. Click **Start processing** and choose a destination such as Desktop.
5. Open the result when the status reaches 100%.

The destination should contain `sample.md`. Its table should match [the expected result](../examples/demo-output/sample.md).

## Command line

```powershell
$demoOutput = Join-Path $env:TEMP 'ContextPack-Demo'
.\contextpack.ps1 ".\examples\sample.csv" -OutputDirectory $demoOutput
```

Expected final message:

```text
Ready: ...\ContextPack-Demo\sample.md
```

This quick demo proves the launcher, environment, router, MarkItDown conversion, custom output routing, and final result handling. Use a PDF or Excel workbook afterward to exercise the complete visual package workflow.

[ქართული ვერსია](QUICK_DEMO.ka.md)
