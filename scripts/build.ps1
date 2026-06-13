$SolutionPath = Resolve-Path $(Join-Path $PSScriptRoot "..\AgentExcel.slnx")
$ExePath = Resolve-Path $(Join-Path $PSScriptRoot "..\src\AgentExcel\bin\Debug\net10.0-windows7.0\win-x64\AgentExcel.exe")
if (Test-Path $ExePath) {
    & $ExePath stop
}
dotnet build $SolutionPath