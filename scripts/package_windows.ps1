param(
    [string]$Version = "1.0.0",
    [ValidateSet("msi","exe","msi+exe","portable")]
    [string]$Type = "msi",
    [ValidateSet("x64","x86","aarch64")]
    [string]$Arch = "x64",
    [switch]$PerUser = $true,
    [string]$IconPath = "",
    [string]$Name = "Document Search Tool",
    [string]$Vendor = "Your Company",
    [string]$Description = "Search ODT, DOCX, and XLSX files for text",
    [string]$UpgradeUUID = "8f230aa1-9d0e-4e4e-9e3a-1a2b3c4d5e6f" # replace with your stable GUID
)

# Packaging script for Document Search Tool (Windows)
# Prerequisites:
# - Windows 10/11
# - JDK 17+ installed and available on PATH (javac, jar, jpackage)
# - For MSI output: WiX Toolset v3.x installed (candle.exe, light.exe on PATH)
#
# Usage examples (PowerShell):
#   pwsh -File scripts/package_windows.ps1
#   pwsh -File scripts/package_windows.ps1 -Version 1.2.3 -Type msi+exe -Arch x64 -Vendor "Acme Corp"
#   pwsh -File scripts/package_windows.ps1 -Type portable

$ErrorActionPreference = 'Stop'

# Resolve project root as the parent of the scripts directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot

$MainClass = "OdtSearchApp"

$SrcDir = Join-Path $ProjectRoot "src"
$BuildDir = Join-Path $ProjectRoot "build"
$DistDir = Join-Path $ProjectRoot "dist"

# Prepare directories
if (Test-Path $BuildDir) { Remove-Item -Recurse -Force $BuildDir }
New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null

Write-Host "[1/4] Compiling sources..."
& javac -d $BuildDir (Join-Path $SrcDir "OdtSearchApp.java")

Write-Host "[2/4] Creating runnable JAR..."
$JarPath = Join-Path $BuildDir "app.jar"
# Create a simple manifest with Main-Class
$manifest = "Manifest-Version: 1.0`nMain-Class: $MainClass`n"
$manifestFile = Join-Path $BuildDir "MANIFEST.MF"
$manifest | Out-File -Encoding ascii $manifestFile -NoNewline

Push-Location $BuildDir
# Include all compiled classes into the JAR
& jar cfm app.jar MANIFEST.MF *
Pop-Location

# Optional portable ZIP
if ($Type -eq 'portable') {
    Write-Host "[3/3] Creating portable ZIP..."
    $PortableDir = Join-Path $BuildDir "portable"
    New-Item -ItemType Directory -Force -Path $PortableDir | Out-Null
    Copy-Item $JarPath (Join-Path $PortableDir "document-search-tool.jar") -Force
    $ZipOut = Join-Path $DistDir ("document-search-tool-" + $Version + "-portable.zip")
    if (Test-Path $ZipOut) { Remove-Item $ZipOut -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($PortableDir, $ZipOut)
    Write-Host "Portable ZIP created: $ZipOut"
    exit 0
}

# Verify jpackage availability
$jpackage = (Get-Command jpackage -ErrorAction SilentlyContinue)
if (-not $jpackage) {
    throw "jpackage not found. Please use JDK 17+ and ensure jpackage is on PATH."
}

# Determine installer types to build
$buildMsi = $false
$buildExe = $false
switch ($Type) {
    'msi' { $buildMsi = $true }
    'exe' { $buildExe = $true }
    'msi+exe' { $buildMsi = $true; $buildExe = $true }
    default { $buildMsi = $true }
}

# If building MSI, ensure WiX is present; otherwise fallback to EXE with a warning
if ($buildMsi) {
    $candle = Get-Command candle.exe -ErrorAction SilentlyContinue
    $light  = Get-Command light.exe  -ErrorAction SilentlyContinue
    if (-not $candle -or -not $light) {
        Write-Warning "WiX Toolset not found (required for MSI). Falling back to EXE. Install WiX v3.x and add to PATH to build MSI."
        $buildMsi = $false
        if (-not $buildExe) { $buildExe = $true }
    }
}

# Common jpackage args
$CommonArgs = @(
    '--name', $Name,
    '--app-version', $Version,
    '--vendor', $Vendor,
    '--dest', $DistDir,
    '--input', $BuildDir,
    '--main-jar', 'app.jar',
    '--description', $Description,
    '--copyright', "© " + (Get-Date).Year + " $Vendor",
    '--win-menu',
    '--win-menu-group', $Vendor,
    '--win-shortcut',
    '--win-dir-chooser'
)

if ($PerUser) { $CommonArgs += '--win-per-user-install' }

# Add icon: use provided path if valid; otherwise auto-generate from app renderer
$IconToUse = $null
if ($IconPath -and (Test-Path $IconPath)) {
    $IconToUse = (Resolve-Path $IconPath)
} else {
    Write-Host "[2.5/4] Generating app icon (.ico) from renderer..."
    $GenIco = Join-Path $BuildDir 'app.ico'
    try {
        & java -cp $BuildDir OdtSearchApp --export-icon $GenIco
        if (Test-Path $GenIco) { $IconToUse = (Resolve-Path $GenIco) }
    } catch {
        Write-Warning "Icon generation failed: $($_.Exception.Message). Proceeding without custom icon."
    }
}
if ($IconToUse) {
    $CommonArgs += @('--icon', $IconToUse)
}

# Add upgrade UUID for Windows installers so newer versions upgrade in place
if ($UpgradeUUID -and $UpgradeUUID -ne '') {
    $CommonArgs += @('--win-upgrade-uuid', $UpgradeUUID)
}

# Target architecture (only applied if supported by jpackage version)
try {
    $help = & jpackage --help 2>&1 | Out-String
    if ($help -match '--target-arch') {
        $CommonArgs += @('--target-arch', $Arch)
    }
} catch {}

if ($buildMsi) {
    Write-Host "[3/4] Building MSI installer..."
    & jpackage @CommonArgs --type msi
}

if ($buildExe) {
    Write-Host "[4/4] Building EXE installer..."
    & jpackage @CommonArgs --type exe
}

Write-Host "Done. Check output in: $DistDir"