
# Check if Excel process is running, and start it if not
$ExcelProcess = Get-Process -Name "excel" -ErrorAction SilentlyContinue
if (-not $ExcelProcess) {
    Write-Host "Excel is not running. Starting Excel..."
    Start-Process excel
    Start-Sleep -Seconds 3
}

$TestPath = Resolve-Path $(Join-Path $PSScriptRoot "..\tests\AgentExcel.Tests")

dotnet test $TestPath @args

