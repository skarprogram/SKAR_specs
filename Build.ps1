# SPDX-License-Identifier: GPL-3.0-only
# Copyright (c) 2025-2026 SKAR_specs contributors

#Requires -Version 5.1

<#
.SYNOPSIS
Builds SKAR_specs from the repository root.

.DESCRIPTION
Builds the SKAR_specs solution in Release mode by default, verifies the
expected application files, and prints the executable SHA-256 hash.

.PARAMETER Configuration
Build configuration. Release is the default.

.PARAMETER Platform
Solution platform. This solution currently defines Any CPU.

.PARAMETER Clean
Runs dotnet clean before building.

.PARAMETER NoRestore
Passes --no-restore to dotnet build. Use only after dependencies have already been restored.

.PARAMETER Verbosity
MSBuild output verbosity passed to dotnet build.

.EXAMPLE
.\Build.ps1

.EXAMPLE
.\Build.ps1 -Configuration Debug

.EXAMPLE
.\Build.ps1 -Clean
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release',

    [Parameter()]
    [ValidateSet('Any CPU')]
    [string] $Platform = 'Any CPU',

    [Parameter()]
    [switch] $Clean,

    [Parameter()]
    [switch] $NoRestore,

    [Parameter()]
    [ValidateSet('quiet', 'minimal', 'normal', 'detailed', 'diagnostic')]
    [string] $Verbosity = 'minimal'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-CheckedCommand {
    param(
        [Parameter(Mandatory = $true)][string] $Description,
        [Parameter(Mandatory = $true)][string] $FilePath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]] $Arguments
    )

    Write-Host $Description -ForegroundColor Cyan
    & $FilePath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $LASTEXITCODE."
    }
}

$repoRoot = [IO.Path]::GetFullPath($PSScriptRoot)
$solutionPath = Join-Path $repoRoot 'SKAR_specs.sln'
$applicationOutputDirectory = Join-Path $repoRoot "bin\$Configuration"
$applicationExecutable = Join-Path $applicationOutputDirectory 'SKAR_specs.exe'

if (-not (Test-Path -LiteralPath $solutionPath -PathType Leaf)) {
    throw "Solution file not found: $solutionPath"
}

$dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
if ($null -eq $dotnetCommand) {
    throw 'The dotnet SDK command was not found. Install a .NET SDK capable of building .NET Framework 4.6 projects.'
}

Write-Host 'SKAR_specs build' -ForegroundColor White
Write-Host "Repository:    $repoRoot"
Write-Host "Solution:      $solutionPath"
Write-Host "Configuration: $Configuration"
Write-Host "Platform:      $Platform"
Write-Host ''

Push-Location $repoRoot
try {
    if ($Clean) {
        $cleanArguments = @(
            'clean',
            $solutionPath,
            '--configuration', $Configuration,
            '--nologo',
            "-p:Platform=$Platform"
        )
        Invoke-CheckedCommand `
            -Description 'Cleaning solution' `
            -FilePath $dotnetCommand.Source `
            -Arguments $cleanArguments
    }

    $buildArguments = @(
        'build',
        $solutionPath,
        '--configuration', $Configuration,
        '--nologo',
        '--verbosity', $Verbosity,
        "-p:Platform=$Platform"
    )
    if ($NoRestore) {
        $buildArguments += '--no-restore'
    }

    Invoke-CheckedCommand `
        -Description 'Building solution' `
        -FilePath $dotnetCommand.Source `
        -Arguments $buildArguments

    $requiredOutputFiles = @(
        $applicationExecutable,
        (Join-Path $applicationOutputDirectory 'SKAR_specs.exe.config')
    )
    $missingOutputFiles = @($requiredOutputFiles | Where-Object {
        -not (Test-Path -LiteralPath $_ -PathType Leaf)
    })
    if ($missingOutputFiles.Count -gt 0) {
        throw "The build succeeded but expected output files are missing:`n - " +
            [string]::Join("`n - ", $missingOutputFiles)
    }

    $executableHash = Get-FileHash -LiteralPath $applicationExecutable -Algorithm SHA256
    Write-Host ''
    Write-Host 'Build completed successfully.' -ForegroundColor Green
    Write-Host "Application output: $applicationOutputDirectory"
    Write-Host "Executable SHA-256: $($executableHash.Hash)"
}
finally {
    Pop-Location
}
