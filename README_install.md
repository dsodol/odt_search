# Windows Installer Build and Installation Guide

This project includes a PowerShell script to produce a Windows setup file (MSI and/or EXE) for the Java desktop application "Document Search Tool". The installer is self-contained and bundles a private Java runtime, so end users do not need to install Java separately.

## Prerequisites
- Windows 10 or Windows 11
- JDK 17 or newer installed and on PATH (must include `javac`, `jar`, and `jpackage`)
- If building an MSI: WiX Toolset v3.x installed and on PATH (`candle.exe` and `light.exe`)
  - If WiX is not installed, the script will automatically fall back to building an EXE and warn you.
- PowerShell 5+ or PowerShell Core (`pwsh`)

## Build the installer
From the project root (where this README resides), run one of the following:

- Build an MSI (requires WiX):
  ```powershell
  pwsh -File scripts\package_windows.ps1 -Version 1.0.0 -Type msi -Vendor "Your Company"
  ```

- Build an EXE installer:
  ```powershell
  pwsh -File scripts\package_windows.ps1 -Version 1.0.0 -Type exe -Vendor "Your Company"
  ```

- Build both MSI and EXE (MSI requires WiX; otherwise it will fall back to EXE only):
  ```powershell
  pwsh -File scripts\package_windows.ps1 -Version 1.0.0 -Type msi+exe -Vendor "Your Company"
  ```

- Build a portable ZIP (no installer, just a runnable JAR in a zip):
  ```powershell
  pwsh -File scripts\package_windows.ps1 -Type portable
  ```

Artifacts will be placed in the `dist\` folder.

## Customization
- Product Name: `-Name "Document Search Tool"`
- Vendor/Publisher: `-Vendor "Your Company"`
- Version: `-Version 1.0.0`
- Description: `-Description "Search ODT, DOCX, and XLSX files for text"`
- App Icon: If omitted, the build script auto-generates an `.ico` from the in-app renderer; override with `-IconPath path\to\app.ico` (must be a `.ico` file)
- Install Scope: per-user (no admin) is default; omit `-PerUser` switch to request per-machine (admin) install
- Target architecture (JDK 21+ jpackage): `-Arch x64` (also supports `x86`, `aarch64` if supported by your jpackage)
- Stable upgrade UUID (MSI upgrades in place): `-UpgradeUUID <your-stable-GUID>`

Example with customizations:
```powershell
pwsh -File scripts\package_windows.ps1 -Version 1.2.3 -Type msi+exe -Vendor "Acme Corp" -Name "Acme Document Search" -Description "Search Office/ODF documents" -IconPath resources\app.ico -UpgradeUUID 12345678-90ab-cdef-1234-567890abcdef
```

## Installing and running
- Double-click the generated `.msi` or `.exe` under `dist\` and follow the installer prompts.
- A Start Menu shortcut will be created under the vendor menu group.
- To uninstall, use Windows Apps & Features (or Programs and Features) and remove the application as usual.

## Code signing (optional)
If you need a signed installer, you can sign the resulting `.msi`/`.exe` with `signtool.exe` using your code signing certificate, or integrate signing into the packaging pipeline. Provide a `.pfx` certificate, password, and a timestamp server URL.

## Troubleshooting
- `jpackage not found`: Ensure you are using JDK 17+ and `jpackage` is on PATH.
- MSI build fails: Install WiX Toolset v3.x and ensure `candle.exe` and `light.exe` are on PATH; or build with `-Type exe`.
- App fails to launch after install: Delete the `dist\` folder, rebuild, and reinstall. Ensure no old versions are conflicting. If using MSI, keep the same `-UpgradeUUID` across versions to allow upgrades.
