param(
    [string]$Version = "1.0.0",
    [ValidateSet("msi","exe","msi+exe","portable")]
    [string]$Type = "msi",
    [ValidateSet("x64","x86","aarch64")]
    [string]$Arch = "x64",
    [switch]$PerUser = $true,
    [switch]$NoIcon = $false,
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
# Always show verbose messages by default and route them to standard output
$VerbosePreference = 'Continue'
function Write-Verbose {
    param([Parameter(Mandatory=$true, Position=0)][object]$Message)
    Write-Output $Message
}

function Join-Args {
    param([Parameter(Mandatory=$true)][object[]]$Args)
    return ($Args | ForEach-Object {
        if ($_ -is [string] -and $_ -match '\s') { '"{0}"' -f $_ } else { $_ }
    } | ForEach-Object { $_.ToString() }) -join ' '
}

# Resolve project root as the parent of the scripts directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot

$MainClass = "OdtSearchApp"

$SrcDir = Join-Path $ProjectRoot "src"
$BuildDir = Join-Path $ProjectRoot "build"
$DistDir = Join-Path $ProjectRoot "dist"

Write-Verbose ("ScriptDir    : {0}" -f $ScriptDir)
Write-Verbose ("ProjectRoot  : {0}" -f $ProjectRoot)
Write-Verbose ("MainClass    : {0}" -f $MainClass)
Write-Verbose ("SrcDir       : {0}" -f $SrcDir)
Write-Verbose ("BuildDir     : {0}" -f $BuildDir)
Write-Verbose ("DistDir      : {0}" -f $DistDir)

# Prepare directories
if (Test-Path $BuildDir) { Remove-Item -Recurse -Force $BuildDir }
New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null

Write-Host "[1/4] Compiling sources..."
$srcFile = (Join-Path $SrcDir "OdtSearchApp.java")
Write-Verbose ("javac -d `"{0}`" `"{1}`"" -f $BuildDir, $srcFile)
& javac -d $BuildDir $srcFile

Write-Host "[2/4] Creating runnable JAR..."
$JarPath = Join-Path $BuildDir "app.jar"
# Create a simple manifest with Main-Class
$manifest = "Manifest-Version: 1.0`nMain-Class: $MainClass`n"
$manifestFile = Join-Path $BuildDir "MANIFEST.MF"
Write-Verbose ("Writing manifest to {0}" -f $manifestFile)
$manifest | Out-File -Encoding ascii $manifestFile -NoNewline

Write-Verbose ("Creating jar at {0}" -f $JarPath)
Push-Location $BuildDir
Write-Verbose ("cwd: {0}" -f (Get-Location))
# Include all compiled classes into the JAR
Write-Verbose "jar cfm app.jar MANIFEST.MF *"
& jar cfm app.jar MANIFEST.MF *
Pop-Location

# Optional portable ZIP
if ($Type -eq 'portable') {
    Write-Host "[3/3] Creating portable ZIP..."
    $PortableDir = Join-Path $BuildDir "portable"
    Write-Verbose ("PortableDir: {0}" -f $PortableDir)
    New-Item -ItemType Directory -Force -Path $PortableDir | Out-Null
    $PortableJar = (Join-Path $PortableDir "document-search-tool.jar")
    Write-Verbose ("Copying jar -> {0}" -f $PortableJar)
    Copy-Item $JarPath $PortableJar -Force
    $ZipOut = Join-Path $DistDir ("document-search-tool-" + $Version + "-portable.zip")
    Write-Verbose ("Zip output: {0}" -f $ZipOut)
    if (Test-Path $ZipOut) { Remove-Item $ZipOut -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($PortableDir, $ZipOut)
    Write-Host "Portable ZIP created: $ZipOut"
    exit 0
}

# Resolve jpackage (prefer PATH, then JAVA_HOME)
$JPACKAGE = $null
$cmd = Get-Command jpackage -ErrorAction SilentlyContinue
if ($cmd) {
    $JPACKAGE = $cmd.Source
    Write-Verbose ("jpackage found on PATH: {0}" -f $JPACKAGE)
} elseif ($env:JAVA_HOME -and (Test-Path (Join-Path $env:JAVA_HOME 'bin\jpackage.exe'))) {
    $JPACKAGE = (Join-Path $env:JAVA_HOME 'bin\jpackage.exe')
    Write-Verbose ("jpackage resolved via JAVA_HOME: {0}" -f $JPACKAGE)
} else {
    $javaVer = ''
    try { $javaVer = (& java -version) 2>&1 | Out-String } catch {}
    $javacVer = ''
    try { $javacVer = (& javac -version) 2>&1 | Out-String } catch {}
    $jh = $env:JAVA_HOME
    $pathJava = (& where.exe java) 2>$null
    $msg = @"
jpackage not found.

What to do:
- Install a JDK 17 or newer that includes jpackage (e.g., Adoptium Temurin or Oracle JDK).
- Ensure %JAVA_HOME% points to the JDK and %JAVA_HOME%\bin is on PATH.
Detected:
- JAVA_HOME: $jh
- where java: $pathJava
- java -version:
$javaVer
- javac -version:
$javacVer
"@
    throw $msg
}

try {
    $jpVer = (& $JPACKAGE --version) 2>&1 | Out-String
    Write-Verbose ("jpackage --version: {0}" -f ($jpVer.Trim()))
} catch { Write-Verbose ("Failed to get jpackage version: {0}" -f $_.Exception.Message) }

# Determine installer types to build
$buildMsi = $false
$buildExe = $false
switch ($Type) {
    'msi' { $buildMsi = $true }
    'exe' { $buildExe = $true }
    'msi+exe' { $buildMsi = $true; $buildExe = $true }
    default { $buildMsi = $true }
}
Write-Verbose ("Build type requested: {0} (MSI={1}, EXE={2})" -f $Type, $buildMsi, $buildExe)

# If building MSI, ensure WiX is present; otherwise fallback to EXE with a warning
if ($buildMsi) {
    Write-Verbose "Checking WiX Toolset (candle.exe, light.exe) on PATH..."
    $candle = Get-Command candle.exe -ErrorAction SilentlyContinue
    $light  = Get-Command light.exe  -ErrorAction SilentlyContinue
    Write-Verbose ("candle.exe: {0}" -f ($candle?.Source))
    Write-Verbose ("light.exe : {0}" -f ($light?.Source))
    if (-not $candle -or -not $light) {
        Write-Warning "WiX Toolset not found (required for MSI). Falling back to EXE. Install WiX v3.x and add to PATH to build MSI."
        $buildMsi = $false
        if (-not $buildExe) { $buildExe = $true }
        Write-Verbose ("Adjusted build targets -> MSI={0}, EXE={1}" -f $buildMsi, $buildExe)
    }
}

# Common jpackage args
# Precompute copyright as a single scalar string to avoid argument splitting
$Copyright = ('Copyright (c) {0} {1}' -f (Get-Date).Year, $Vendor)
$CommonArgs = @(
    '--name', $Name,
    '--app-version', $Version,
    '--vendor', $Vendor,
    '--dest', $DistDir,
    '--input', $BuildDir,
    '--main-jar', 'app.jar',
    '--description', $Description,
    '--copyright', $Copyright,
    '--win-menu',
    '--win-menu-group', $Vendor,
    '--win-shortcut',
    '--win-dir-chooser'
)

if ($PerUser) {
    Write-Verbose "Using per-user installation (--win-per-user-install)."
    $CommonArgs += '--win-per-user-install'
}

# Icon handling
if ($NoIcon) {
    Write-Verbose "No icon requested via -NoIcon. Skipping --icon."
} else {
    # Add icon: use provided path if valid; otherwise auto-generate from app renderer
    $IconToUse = $null
    if ($IconPath -and (Test-Path $IconPath)) {
        $IconToUse = (Resolve-Path $IconPath)
        Write-Verbose ("Using provided icon: {0}" -f $IconToUse)
    } else {
        Write-Host "[2.5/4] Generating app icon (.ico) from renderer..."
        $GenIco = Join-Path $BuildDir 'app.ico'
        Write-Verbose ("Attempting icon generation: java -cp `"{0}`" OdtSearchApp --export-icon `"{1}`"" -f $BuildDir, $GenIco)
        try {
            & java -cp $BuildDir OdtSearchApp --export-icon $GenIco
            if (Test-Path $GenIco) {
                $IconToUse = (Resolve-Path $GenIco)
                Write-Verbose ("Generated icon at: {0}" -f $IconToUse)
            } else {
                Write-Verbose "Icon generation did not produce a file."
            }
        } catch {
            Write-Warning "Icon generation failed: $($_.Exception.Message). Proceeding without custom icon."
        }
    }
    if ($IconToUse) {
        Write-Verbose ("Adding --icon argument: {0}" -f $IconToUse)
        $CommonArgs += @('--icon', $IconToUse)
    } else {
        Write-Verbose "No icon will be used."
    }
}

# Add upgrade UUID for Windows installers so newer versions upgrade in place
if ($UpgradeUUID -and $UpgradeUUID -ne '') {
    Write-Verbose ("Adding --win-upgrade-uuid: {0}" -f $UpgradeUUID)
    $CommonArgs += @('--win-upgrade-uuid', $UpgradeUUID)
}

# Target architecture (only applied if supported by jpackage version)
try {
    Write-Verbose "Checking jpackage --help for --target-arch support..."
    $help = & $JPACKAGE --help 2>&1 | Out-String
    if ($help -match '--target-arch') {
        Write-Verbose ("Adding --target-arch {0}" -f $Arch)
        $CommonArgs += @('--target-arch', $Arch)
    } else {
        Write-Verbose "--target-arch not supported by this jpackage version."
    }
} catch {
    Write-Verbose ("Failed to inspect jpackage help: {0}" -f $_.Exception.Message)
}

Write-Verbose ("CommonArgs: {0}" -f (Join-Args -Args $CommonArgs))

if ($buildMsi) {
    Write-Host "[3/4] Building MSI installer..."
    $msiCmd = @($JPACKAGE) + $CommonArgs + @('--type','msi')
    Write-Verbose ("Run: {0}" -f (Join-Args -Args $msiCmd))
    & $JPACKAGE @CommonArgs --type msi
}

if ($buildExe) {
    Write-Host "[4/4] Building EXE installer..."
    $exeCmd = @($JPACKAGE) + $CommonArgs + @('--type','exe')
    Write-Verbose ("Run: {0}" -f (Join-Args -Args $exeCmd))
    & $JPACKAGE @CommonArgs --type exe
}

Write-Host "Done. Check output in: $DistDir"