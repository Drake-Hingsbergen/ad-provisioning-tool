<#
.SYNOPSIS
    Compiles ADProvisioningTool.ps1 into ADProvisioningTool.exe with ps2exe.
.DESCRIPTION
    Runs Build\Invoke-CIChecks.ps1 first and stops if any check fails, then compiles with
    -noConsole -STA (WinForms needs a single-threaded apartment) and the icon in assets\ when present.
    Installs the ps2exe module for the current user if it is missing.
.PARAMETER OutputPath
    Where to write the exe. Defaults to ADProvisioningTool.exe in the repo root.
.PARAMETER SkipChecks
    Builds without running the checks. CI uses this because it runs the checks in an earlier step.
.EXAMPLE
    .\Build\Build-Exe.ps1
.NOTES
    Copy the exe to C:\ADProvisioning on the management server. ADProvisioningTool.json stays on the server and is not in the repo.
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$SkipChecks
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$source   = Join-Path $repoRoot 'ADProvisioningTool.ps1'
$icon     = Join-Path $repoRoot 'assets\ADProvisioningTool.ico'
if (-not $OutputPath) { $OutputPath = Join-Path $repoRoot 'ADProvisioningTool.exe' }
$OutputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)

if (-not $SkipChecks) {
    & (Join-Path $PSScriptRoot 'Invoke-CIChecks.ps1') -RepoRoot $repoRoot
    if ($LASTEXITCODE -ne 0) { throw 'Checks failed; build stopped.' }
}

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'Installing ps2exe for the current user...' -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

$outDir = Split-Path -Parent $OutputPath
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$buildArgs = @{
    InputFile   = $source
    OutputFile  = $OutputPath
    NoConsole   = $true
    STA         = $true
    Title       = 'AD Provisioning Tool'
    Description = 'Active Directory provisioning tool'
    Company     = 'Contoso'
    Version     = '2.0.0.0'
}
if (Test-Path $icon) { $buildArgs.IconFile = $icon }

Invoke-ps2exe @buildArgs
if (-not (Test-Path $OutputPath)) { throw "ps2exe finished but $OutputPath was not created." }
Write-Host "Built $OutputPath" -ForegroundColor Green
