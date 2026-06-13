$SolutionRootPath = Resolve-Path $(Join-Path $PSScriptRoot "..\")
$ExePath = Resolve-Path $(Join-Path $PSScriptRoot "..\src\AgentExcel\bin\Debug\net10.0-windows7.0\win-x64\AgentExcel.exe")
$DllPath = Resolve-Path $(Join-Path $PSScriptRoot "..\src\AgentExcel\bin\Debug\net10.0-windows7.0\win-x64\AgentExcel.dll")
$BuildScript = Resolve-Path $(Join-Path $PSScriptRoot "build.ps1")
$LastSourceUpdateTime = $(Get-ChildItem -File -Recurse -Filter "*.cs" -Path $SolutionRootPath | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
$LastBuildTime = $(Get-Item $DllPath).LastWriteTime
$IsStop = $false
if ($args.Count -gt 0 -and $args[0] -eq "stop") {
    $IsStop = $true
}
if (!$IsStop -and $LastSourceUpdateTime -gt $LastBuildTime) {
    & $BuildScript
}

& $ExePath @args