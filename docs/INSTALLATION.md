# ContextPack Desktop installation

## Recommended: Windows installer

1. Open the latest GitHub Release and download `ContextPack-Setup-2.2.0.exe`.
2. Optionally verify the file against `SHA256SUMS-2.2.0.txt`.
3. Double-click the installer. Keep the optional Desktop shortcut selected if you want one.
4. Keep the internet connection available during the first setup. ContextPack checks Python, Tesseract OCR, its virtual environment, Python packages, and English/Georgian OCR models. Missing required components are installed automatically.
5. Open **ContextPack Desktop** from the Start menu or Desktop.

Microsoft Excel is optional. It is required only for faithful Excel PDF/PNG visual rendering; values and formulas can still be extracted without it.

## If setup reports an error

Open **ContextPack Setup Repair** from the Start menu. It retries dependency preparation and writes details to:

```text
%LOCALAPPDATA%\Programs\ContextPack Desktop\logs\install.log
%LOCALAPPDATA%\Programs\ContextPack Desktop\logs\last-install-error.txt
```

The GUI also provides **Check environment** and **Repair environment** actions.

## Portable ZIP

1. Extract `ContextPack-Desktop-2.2.0-portable.zip` to a normal writable folder.
2. Open PowerShell in the extracted folder.
3. Run:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\Start-ContextPack-GUI.cmd
```

Do not run directly from inside the ZIP. `.venv` activation is not required.

## Uninstall

Open **Settings → Apps → Installed apps**, select **ContextPack Desktop**, and choose **Uninstall**. You can also use **Uninstall ContextPack Desktop** in the Start menu. The application, local Python environment, OCR models, logs, and shortcuts are removed. Non-empty `input` and `output` folders are preserved to avoid deleting user documents.

## Verify a download with SHA-256

```powershell
Get-FileHash .\ContextPack-Setup-2.2.0.exe -Algorithm SHA256
```

Compare the result with the matching line in `SHA256SUMS-2.2.0.txt`.
