$scriptDir = $PSScriptRoot
$parentDir = Split-Path -Parent $scriptDir
$binDir = Join-Path $parentDir "bin"
$exePath = Join-Path $binDir "AgentExcel.exe"
if (-not (Test-Path -Path $exePath -PathType Leaf)) {
    Write-Host "Error: 'AgentExcel.exe' was not found under the 'bin' directory: $binDir" -ForegroundColor Red
    Write-Host "Please go to https://github.com/ecator/agent-excel to download the binary." -ForegroundColor Yellow
    exit 1
}
$exeFullPath = Resolve-Path $exePath
Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList "`"$exeFullPath`" start" | Out-Null

Start-Sleep 3

Write-Host "Please run 'AgentExcel.exe status' to confirm whether the service has been started."