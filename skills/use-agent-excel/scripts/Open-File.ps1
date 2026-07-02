param (
    [Parameter(Mandatory=$true)]
    [string]$File
)
if (-not (Test-Path -Path $File -PathType Leaf)) {
    Write-Host "Error: '$File' was not found" -ForegroundColor Red
    exit 1
}
$cmd = "Start-Process `"$File`""
$base64 = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($cmd))
$CmdLine = "powershell -EncodedCommand $base64"
if ($PSVersionTable.PSVersion.Major -lt 6) {
    Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList $CmdLine | Out-Null
}
else {
    Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{CommandLine = $CmdLine } | Out-Null
}

Start-Sleep 3

Write-Host "Please list workbooks to confirm whether the file have been opened."