<#
.SYNOPSIS
    Updates the version number across the repository files.
.DESCRIPTION
    This script updates the version fields in the following files:
    - src/AgentExcel/AgentExcel.csproj
    - skills/use-agent-excel/SKILL.md
.PARAMETER Version
    The target version in 'x.y.z' format.
.EXAMPLE
    .\scripts\version.ps1 0.2.1
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Version
)


# Validate version format (x.y.z)
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    Write-Error "Invalid version format '$Version'. Version must follow 'x.y.z' format (e.g. 1.0.0)."
    exit 1
}

$csprojPath = Resolve-Path $(Join-Path $PSScriptRoot "..\src\AgentExcel\AgentExcel.csproj") -ErrorAction SilentlyContinue
if (-not $csprojPath) {
    $csprojPath = Join-Path $PSScriptRoot "..\src\AgentExcel\AgentExcel.csproj"
}

$skillMdPath = Resolve-Path $(Join-Path $PSScriptRoot "..\skills\use-agent-excel\SKILL.md") -ErrorAction SilentlyContinue
if (-not $skillMdPath) {
    $skillMdPath = Join-Path $PSScriptRoot "..\skills\use-agent-excel\SKILL.md"
}

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 1. Update src/AgentExcel/AgentExcel.csproj
if (Test-Path $csprojPath) {
    $csprojContent = [System.IO.File]::ReadAllText($csprojPath, $utf8NoBom)
    $csprojVersionPattern = '<Version>(\d+\.\d+\.\d+)</Version>'
    if ($csprojContent -match $csprojVersionPattern) {
        $oldVersion = $Matches[1]
        $csprojContent = $csprojContent -replace $csprojVersionPattern, "<Version>$Version</Version>"
        [System.IO.File]::WriteAllText($csprojPath, $csprojContent, $utf8NoBom)
        Write-Host "Updated AgentExcel.csproj version from '$oldVersion' to '$Version'"
    }
    else {
        Write-Warning "Could not find <Version> tag in AgentExcel.csproj"
    }
}
else {
    Write-Error "Could not find AgentExcel.csproj at $csprojPath"
}

# 2. Update skills/use-agent-excel/SKILL.md
if (Test-Path $skillMdPath) {
    $skillMdContent = [System.IO.File]::ReadAllText($skillMdPath, $utf8NoBom)
    $skillMdVersionPattern = '(?m)^( +version:\s*)"([^"]+)"'
    if ($skillMdContent -match $skillMdVersionPattern) {
        $versionTag = $Matches[1]
        $oldVersion = $Matches[2]
        $skillMdContent = $skillMdContent -replace $skillMdVersionPattern, "${versionTag}`"$Version`""
        [System.IO.File]::WriteAllText($skillMdPath, $skillMdContent, $utf8NoBom)
        Write-Host "Updated SKILL.md version from '$oldVersion' to '$Version'"
    }
    else {
        Write-Warning "Could not find version tag in YAML frontmatter of SKILL.md"
    }
}
else {
    Write-Error "Could not find SKILL.md at $skillMdPath"
}
