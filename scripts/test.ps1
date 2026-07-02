
# Check if Excel process is running, and start it if not
$ExcelProcess = Get-Process -Name "excel" -ErrorAction SilentlyContinue
if (-not $ExcelProcess) {
    Write-Host "Excel is not running. Starting Excel..."
    $CmdLine = "powershell -Command `"Start-Process excel`""
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList $CmdLine | Out-Null
    }
    else {
        Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine = $CmdLine } | Out-Null
    }
    Start-Sleep -Seconds 5
}

$TestPath = Resolve-Path $(Join-Path $PSScriptRoot "..\tests\AgentExcel.Tests")

dotnet test $TestPath @args

