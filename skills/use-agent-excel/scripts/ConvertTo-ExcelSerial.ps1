param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string[]]$Date,
    [Parameter(Mandatory = $false)]
    [string]$Format = "yyyy/MM/dd HH:mm:ss"
)

process {
    $baseDate = [DateTime]::ParseExact("1899-12-30", "yyyy-MM-dd", $null)
    foreach ($d in $Date) {
        $parsedDate = [DateTime]::MinValue
        if (-not [System.DateTime]::TryParseExact($d, $Format, $null, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            $parsedDate = [DateTime]$d
        }
        $rawSerial = ($parsedDate - $baseDate).TotalDays
        $serial = if ($rawSerial -lt 61) { $rawSerial - 1 } else { $rawSerial }
        Write-Host $("$d -> $serial")
    }
}
