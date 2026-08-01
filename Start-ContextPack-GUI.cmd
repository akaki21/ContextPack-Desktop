@echo off
setlocal
set "CONTEXTPACK_ROOT=%~dp0"
set "CONTEXTPACK_PYTHONW=%CONTEXTPACK_ROOT%.venv\Scripts\pythonw.exe"

if not exist "%CONTEXTPACK_PYTHONW%" (
    if exist "%CONTEXTPACK_ROOT%installer\bootstrap.ps1" (
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CONTEXTPACK_ROOT%installer\bootstrap.ps1" -InstallRoot "%CONTEXTPACK_ROOT%" -Interactive
    ) else (
        powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('ContextPack environment is missing. Run setup.ps1 first.','ContextPack Desktop')"
    )
)

if not exist "%CONTEXTPACK_PYTHONW%" (
    exit /b 1
)

if "%~1"=="" (
    start "ContextPack Desktop" "%CONTEXTPACK_PYTHONW%" "%CONTEXTPACK_ROOT%contextpack_gui.py"
) else (
    start "ContextPack Desktop" "%CONTEXTPACK_PYTHONW%" "%CONTEXTPACK_ROOT%contextpack_gui.py" "%~f1"
)
