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
Invoke-WmiMethod -Class Win32_Process -Name Create -ArgumentList "powershell -EncodedCommand $base64" | Out-Null

Start-Sleep 3

Write-Host "Please list workbooks to confirm whether the file have been opened."