param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [double[]]$Serial,
    [Parameter(Mandatory = $false)]
    [string]$Format = "yyyy/MM/dd HH:mm:ss"
)

process {
    $baseDate = [DateTime]::ParseExact("1899-12-30", "yyyy-MM-dd", $null)
    foreach ($s in $Serial) {
        $days = if ($s -lt 60) { $s + 1 } else { $s }
        Write-Host $("$s -> " + $baseDate.AddDays($days).ToString($Format))
    }
}