# Document Search Tool (ODT/DOCX/XLSX)

A lightweight desktop utility to search through OpenDocument Text (ODT), Microsoft Word (DOCX), and Excel (XLSX) files for text. No external libraries required; everything runs on a standard Java runtime.

## Features
- Fast text search across ODT, DOCX, and XLSX files in a chosen folder (recursive)
- Simple, native-looking Swing UI
- Search history with quick re-run
- Result table with file, match count, and snippet
- Log window with detailed progress and errors
- Remembers last used root folder
- Self-contained packaging for Windows (MSI/EXE) and a portable ZIP/JAR option

## Screenshots
Coming soon.

## Requirements
- Windows 10/11 (for the provided packaging scripts)
- Java JDK 17 or newer (to build or run from source). JRE is sufficient for running the packaged app.
- Optional for MSI packaging: WiX Toolset v3.x (puts `candle.exe` and `light.exe` on PATH)
- No Microsoft Office or LibreOffice required for searching — the app reads ODT/DOCX/XLSX files directly using the JDK. To open a result by double‑clicking, any associated app works (e.g., LibreOffice or Microsoft Office), but it’s optional.

## Quick Start (End Users)
Prebuilt artifacts are placed under `dist/` when you package the app.

- Installer (EXE): double-click `dist/Document Search Tool-<version>.exe`
- Portable ZIP: unzip `dist/document-search-tool-<version>-portable.zip` and run `document-search-tool.jar` with Java 17+

If you received an MSI, double-click the MSI. If SmartScreen warns you, choose "More info" and "Run anyway" (you may need to allow it depending on your Windows policies).

## Using the App
1. Choose Root… to select the folder to scan
2. Enter a search term in the combo box
3. Click Search (Stop is available to cancel)
4. Results show matching files, their match count, and a short snippet
5. Open Log Window for detailed progress (errors, skipped files, timings)
6. The app remembers your last root and previous search terms

Tip: The tool reads content locally and does not upload files anywhere.

## Build from Source
From the project root:

```powershell
# Compile and run directly (requires Java 17+ on PATH)
javac -d build src\OdtSearchApp.java
java -cp build OdtSearchApp
```

Or use the packaging script (recommended on Windows):

```powershell
# Build installers into dist/ (default: MSI if WiX available; else EXE)
pwsh -File scripts\package_windows.ps1
```

## Packaging on Windows
The PowerShell script `scripts\package_windows.ps1` compiles the app, builds a runnable JAR, and invokes `jpackage` to produce MSI and/or EXE installers, or a portable ZIP.

Common options:

```powershell
# Basic
pwsh -File scripts\package_windows.ps1 -Version 1.0.0 -Type msi

# Build EXE instead of MSI
pwsh -File scripts\package_windows.ps1 -Type exe

# Build both MSI and EXE
pwsh -File scripts\package_windows.ps1 -Type msi+exe

# Build portable ZIP (no installer)
pwsh -File scripts\package_windows.ps1 -Type portable

# Target architecture (if supported by your jpackage)
pwsh -File scripts\package_windows.ps1 -Arch x64   # or x86, aarch64

# Install per-user (default) vs. per-machine
pwsh -File scripts\package_windows.ps1 -PerUser    # adds --win-per-user-install
```

Icon handling:

```powershell
# Provide your own icon (.ico)
pwsh -File scripts\package_windows.ps1 -IconPath "C:\path\to\icon.ico"

# Build without any icon (bypasses known Windows icon injection issues)
pwsh -File scripts\package_windows.ps1 -NoIcon
```

Notes:
- Without `-IconPath`, the script tries to auto-generate an ICO from the app. Some Windows setups reject certain ICO formats during installer creation. If you encounter an error like:
  "Not enough memory resources are available to process this command.\nFailed to update icon ... .ico"
  use `-NoIcon` or supply a known-good `.ico` (sizes 16, 32, 48, 256 at 32-bit BGRA; keep file size modest).
- MSI requires WiX Toolset v3.x on PATH (`candle.exe`, `light.exe`). If WiX isn’t found, the script falls back to EXE automatically.
- The script attempts to add `--target-arch` only if supported by your `jpackage` version.

## Troubleshooting
- jpackage not found
  - Install a full JDK (e.g., Adoptium Temurin 17/21), ensure `%JAVA_HOME%` points to the JDK and `%JAVA_HOME%\bin` is on `PATH`.
- MSI build fails due to WiX
  - Install WiX v3.x and add to PATH, or build `-Type exe`.
- Icon errors during packaging
  - Build with `-NoIcon`, or supply a standard `.ico` via `-IconPath`.
- App doesn’t see files / permissions
  - Run as a user who can access the directories; network shares may need permissions.

## Development Notes
- Main class: `OdtSearchApp`
- Source: `src\OdtSearchApp.java` (all logic, UI, and simple file parsers)
- Packaging script: `scripts\package_windows.ps1`
- No external dependencies; parsing relies on core JDK APIs (ZIP + XML DOM)

### Supported Formats
- ODT: reads `content.xml` from the ODT ZIP
- DOCX: extracts visible text from `document.xml`
- XLSX: extracts values from worksheet XML and shared strings

### Limitations
- Does not render complex formatting; extracts plain text
- Password-protected/encrypted office files are skipped
- Very large directories can take time; use Stop to cancel

## Related Files
- `README_install.md` — supplementary installation notes. This README consolidates quick start and packaging instructions; see the other file for any additional environment-specific tips.

## License
Add your license here (e.g., MIT). Replace this section with the actual license text or a link to `LICENSE`.

## Changelog
- 2025-10-29: Added `-NoIcon` packaging switch and created this README
