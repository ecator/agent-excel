$projectPath = Resolve-Path $(Join-Path $PSScriptRoot "..\src\AgentExcel\AgentExcel.csproj")
$skillBinPath = Resolve-Path $(Join-Path $PSScriptRoot "..\skills\use-agent-excel\bin")
if (Test-Path $skillBinPath) {
    Remove-Item -Recurse -Force $skillBinPath
}
dotnet publish $projectPath  -c Release -r win-x64 --self-contained true -o $skillBinPath