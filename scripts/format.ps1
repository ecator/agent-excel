$SolutionPath = Resolve-Path $(Join-Path $PSScriptRoot "..\AgentExcel.slnx")
dotnet format $SolutionPath --no-restore -v n