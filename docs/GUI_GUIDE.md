# ContextPack Desktop GUI — first use

This is the local Windows interface for ContextPack. Files remain on the computer, and the GUI uses the existing PowerShell/Python processing engine.

## Launch

1. Open the ContextPack project folder.
2. Double-click `Start-ContextPack-GUI.cmd`.
3. If the local environment is missing, run `setup.ps1` first. No additional GUI framework is required in an already configured project.

The language selector in the upper-right corner changes the whole interface between Georgian and English immediately. Processing-mode choices keep their underlying value when the language changes.

You can also drag one file onto `Start-ContextPack-GUI.cmd` to open the window with that file selected.

## Basic workflow

1. Choose a PDF, Excel workbook, image, or another supported Office/data file.
2. Keep the recommended **Auto** mode for the first run.
3. For Excel, keep **Both** and 180 DPI unless you have a specific requirement.
4. Start processing and choose a destination folder. The dialog opens near Desktop by default.
5. Open the result or copy the prepared AI task after completion.

The selected destination is used directly. Complete packages are built atomically inside that folder; ContextPack does not first create a duplicate under the project `output` folder.

## Safe cancellation

Cancel is cooperative: it requests cancellation and waits for the current safe operation boundary instead of force-killing Excel. A large Excel/PDF operation may therefore take a short time to stop. This protects the source and any previously valid package.

## Environment actions

**Environment Check** is read-only. **Setup / Repair** installs or updates missing project dependencies and should be used when the check reports a problem.

The quick indicators in the window header show engine, OCR, and Excel availability. Use the detailed check whenever an indicator shows `!`.

## AI handoff

The **Copy AI task** action places the result path and recommended reading order on the clipboard. Add your concrete goal before sending it to Codex. When Codex runs on the same computer, attaching every generated PDF/PNG page is unnecessary.
