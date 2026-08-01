# Building ContextPack releases

This directory contains the Windows installer bootstrap, Inno Setup definition, icon builder, and release packager.

## Prerequisites

- Windows 10/11
- the project environment created by `setup.ps1`
- Inno Setup 6 (the project build helper can install it with winget)

## Build the installer

```powershell
.\installer\build-installer.ps1 -Version 2.2.0 -InstallInnoSetup
```

The unversioned development artifact is `dist\ContextPack-Setup.exe`.

## Build all release assets

```powershell
.\installer\build-release.ps1 -Version 2.2.0
```

This creates the versioned installer, portable ZIP, and SHA-256 checksum file. The portable ZIP still performs its first-run Python/Tesseract/dependency setup through `setup.ps1`; it is not an offline bundle.
