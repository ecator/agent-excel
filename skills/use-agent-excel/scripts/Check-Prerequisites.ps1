# Check-Prerequisites.ps1
# This script verifies that Microsoft Excel is installed via registry and that
# AgentExcel.exe is present in the parent directory's 'bin' folder.

$ErrorActionPreference = "Stop"

# Verify operating system is Windows
$runningOnWindows = $true
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $runningOnWindows = $IsWindows
}
if (-not $runningOnWindows) {
    Write-Host "Error: AgentExcel is only supported on Windows operating systems." -ForegroundColor Red
    exit 1
}

# 1. Check if Excel is installed via registry
$excelInstalled = $false

# Define registry paths to check for Excel App Path (User-specific and System-wide)
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
)

foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        $excelInstalled = $true
        break
    }
}

# Fallback check: Check Excel COM ProgID registration
if (-not $excelInstalled) {
    if (Test-Path "Registry::HKEY_CLASSES_ROOT\Excel.Application") {
        $excelInstalled = $true
    }
}

# If Excel is not installed, output error message and exit
if (-not $excelInstalled) {
    Write-Host "Error: Microsoft Excel is not installed on this system." -ForegroundColor Red
    Write-Host "Please install Microsoft Office 2016 or later." -ForegroundColor Yellow
    exit 1
}

# 2. Check if the parent directory contains a 'bin' folder with 'AgentExcel.exe'
$scriptDir = $PSScriptRoot
if (-not $scriptDir) {
    $scriptDir = Get-Location
}

$parentDir = Split-Path -Parent $scriptDir
$binDir = Join-Path $parentDir "bin"
$exePath = Join-Path $binDir "AgentExcel.exe"

if (-not (Test-Path -Path $binDir -PathType Container)) {
    Write-Host "Error: The 'bin' directory does not exist in the parent directory: $parentDir" -ForegroundColor Red
    Write-Host "Please go to https://github.com/ecator/agent-excel to download the binary." -ForegroundColor Yellow
    exit 1
}

if (-not (Test-Path -Path $exePath -PathType Leaf)) {
    Write-Host "Error: 'AgentExcel.exe' was not found under the 'bin' directory: $binDir" -ForegroundColor Red
    Write-Host "Please go to https://github.com/ecator/agent-excel to download the binary." -ForegroundColor Yellow
    exit 1
}

Write-Host "All prerequisites are satisfied. Excel is installed and AgentExcel.exe is present." -ForegroundColor Green
exit 0
