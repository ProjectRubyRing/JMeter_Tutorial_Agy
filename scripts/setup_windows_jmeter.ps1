# ==============================================================================
# Windows 11 JMeter Setup & Verification Script (PowerShell)
# Verifies Java, generates setenv.bat, and downloads Plugins Manager if requested.
# ==============================================================================
param(
    [string]$JMeterHome = "C:\Tools\apache-jmeter-5.6.3",
    [switch]$DownloadPluginsManager = $true
)

$ErrorActionPreference = "Stop"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " JMeter Windows 11 Setup & Environment Verification" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Check Java
Write-Host "`n[1/4] Checking Java Environment..." -ForegroundColor Yellow
try {
    $javaVer = & java -version 2>&1
    Write-Host "Java is detected:" -ForegroundColor Green
    $javaVer | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
} catch {
    Write-Host "[ERROR] Java is not detected in PATH. Please install JDK 17 or 21." -ForegroundColor Red
    exit 1
}

# 2. Check JMeter Directory
Write-Host "`n[2/4] Checking JMeter Installation Directory..." -ForegroundColor Yellow
if (-not (Test-Path $JMeterHome)) {
    Write-Host "[WARNING] JMeter directory not found at $JMeterHome." -ForegroundColor Magenta
    Write-Host "Please download JMeter from https://jmeter.apache.org/download_jmeter.cgi and extract to $JMeterHome."
} else {
    Write-Host "Found JMeter at: $JMeterHome" -ForegroundColor Green
}

# 3. Create / Verify setenv.bat
Write-Host "`n[3/4] Configuring JVM Heap Settings (setenv.bat)..." -ForegroundColor Yellow
$binDir = Join-Path $JMeterHome "bin"
if (Test-Path $binDir) {
    $setenvPath = Join-Path $binDir "setenv.bat"
    $setenvContent = "@echo off`r`nrem JVM Heap settings for JMeter GUI & Local Run`r`nset HEAP=-Xms1g -Xmx2g -XX:MaxMetaspaceSize=512m`r`n"
    Set-Content -Path $setenvPath -Value $setenvContent -Encoding ASCII
    Write-Host "Created/Updated setenv.bat at: $setenvPath" -ForegroundColor Green
    Get-Content $setenvPath | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
} else {
    Write-Host "Skipping setenv.bat creation ($binDir does not exist yet)." -ForegroundColor Gray
}

# 4. Download Plugins Manager
if ($DownloadPluginsManager) {
    Write-Host "`n[4/4] Checking JMeter Plugins Manager..." -ForegroundColor Yellow
    $libExtDir = Join-Path $JMeterHome "lib\ext"
    if (Test-Path $libExtDir) {
        $pmJar = Join-Path $libExtDir "jmeter-plugins-manager-1.10.jar"
        if (-not (Test-Path $pmJar)) {
            Write-Host "Downloading jmeter-plugins-manager-1.10.jar to $libExtDir..." -ForegroundColor Cyan
            $url = "https://repo1.maven.org/maven2/kg/apc/jmeter-plugins-manager/1.10/jmeter-plugins-manager-1.10.jar"
            try {
                Invoke-WebRequest -Uri $url -OutFile $pmJar
                Write-Host "Plugins Manager downloaded successfully!" -ForegroundColor Green
            } catch {
                Write-Host "[WARNING] Could not auto-download Plugins Manager: $_" -ForegroundColor Yellow
                Write-Host "Please manually download from https://jmeter-plugins.org/wiki/PluginsManager/ and place into $libExtDir"
            }
        } else {
            Write-Host "JMeter Plugins Manager jar already exists in lib\ext." -ForegroundColor Green
        }
    }
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host " Setup Check Finished. Launch JMeter by running:" -ForegroundColor Green
Write-Host "   $JMeterHome\bin\jmeter.bat" -ForegroundColor White
Write-Host "==========================================================" -ForegroundColor Cyan
