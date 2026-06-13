$CliPath = Resolve-Path $(Join-Path $PSScriptRoot "cli.ps1")

& $CliPath --run-server 4869