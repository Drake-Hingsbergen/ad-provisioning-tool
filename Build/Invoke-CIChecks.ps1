<#
.SYNOPSIS
    Runs every pre-build check for the AD Provisioning Tool.
.DESCRIPTION
    1. Parse check of every script in the repo.
    2. ps2exe compatibility check of ADProvisioningTool.ps1 (UTF-8 BOM, ASCII only, no constructor syntax ps2exe cannot compile).
    3. PSScriptAnalyzer on ADProvisioningTool.ps1, compared with Build\PSScriptAnalyzerBaseline.json. Any finding not in the baseline fails.
    4. Pester tests in Tests\.
    Exits with code 1 if any check fails, so a CI pipeline stops before building.
.PARAMETER RepoRoot
    Repo folder. Defaults to the parent of this script's folder.
.PARAMETER TestResultPath
    Optional path for a Pester NUnit XML report (used by CI).
.PARAMETER UpdateBaseline
    Rewrites the analyzer baseline from the current findings instead of comparing. Use only after reviewing every new finding.
.EXAMPLE
    .\Build\Invoke-CIChecks.ps1
.EXAMPLE
    .\Build\Invoke-CIChecks.ps1 -UpdateBaseline
.NOTES
    Requires Pester 5 or later and PSScriptAnalyzer. Runs on Windows PowerShell 5.1 and PowerShell 7.
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$TestResultPath,
    [switch]$UpdateBaseline
)

$ErrorActionPreference = 'Stop'
$toolPath     = Join-Path $RepoRoot 'ADProvisioningTool.ps1'
$baselinePath = Join-Path $RepoRoot 'Build\PSScriptAnalyzerBaseline.json'
$results      = New-Object System.Collections.Generic.List[object]

function Add-CheckResult {
    param([string]$Check, [bool]$Passed, [string]$Detail)
    $results.Add([PSCustomObject]@{ Check = $Check; Result = $(if ($Passed) { 'PASS' } else { 'FAIL' }); Detail = $Detail })
}

function Get-FindingKey {
    # Line numbers move with every edit, so findings are matched on rule plus the offending code text.
    param($Finding)
    $text = "$($Finding.Extent.Text)".Split("`n")[0].Trim()
    return "$($Finding.RuleName)|$text"
}

#region Parse check
$scripts = @(Get-ChildItem -Path $RepoRoot -Recurse -Include '*.ps1', '*.psm1', '*.psd1' -File)
$parseFailures = @()
foreach ($script in $scripts) {
    $tokens = $null; $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($script.FullName, [ref]$tokens, [ref]$errors) | Out-Null
    foreach ($e in @($errors)) { $parseFailures += "$($script.Name):$($e.Extent.StartLineNumber) $($e.Message)" }
}
Add-CheckResult 'Parse' (-not $parseFailures) $(if ($parseFailures) { $parseFailures -join '; ' } else { "$($scripts.Count) files" })
#endregion

#region ps2exe compatibility
$bytes = [System.IO.File]::ReadAllBytes($toolPath)
$hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
$nonAscii = @()
for ($i = 3; $i -lt $bytes.Length; $i++) { if ($bytes[$i] -gt 127) { $nonAscii += $i; if ($nonAscii.Count -ge 5) { break } } }
$source = [System.IO.File]::ReadAllText($toolPath)
$ctorLines = @(($source -split "`n") | Select-String -Pattern 'New-Object\s+[\w\.]+\(|\]::new\(' | ForEach-Object { $_.LineNumber })
$compatProblems = @()
if (-not $hasBom)      { $compatProblems += 'missing UTF-8 BOM' }
if ($nonAscii)         { $compatProblems += "non-ASCII bytes at offsets $($nonAscii -join ', ')" }
if ($ctorLines)        { $compatProblems += "constructor syntax on lines $($ctorLines -join ', ')" }
Add-CheckResult 'ps2exe compatibility' (-not $compatProblems) $(if ($compatProblems) { $compatProblems -join '; ' } else { 'BOM, ASCII, no constructor syntax' })
#endregion

#region PSScriptAnalyzer
try {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
    $findings = @(Invoke-ScriptAnalyzer -Path $toolPath -Severity Error, Warning)
    $current = @{}
    foreach ($f in $findings) { $k = Get-FindingKey $f; $current[$k] = 1 + [int]$current[$k] }

    if ($UpdateBaseline) {
        $current.GetEnumerator() | Sort-Object Key | ForEach-Object { [PSCustomObject]@{ Key = $_.Key; Count = $_.Value } } |
            ConvertTo-Json -Depth 3 | Set-Content -Path $baselinePath -Encoding UTF8
        Add-CheckResult 'PSScriptAnalyzer' $true "Baseline rewritten: $($findings.Count) accepted findings"
    } else {
        $baseline = @{}
        if (Test-Path $baselinePath) {
            foreach ($b in @(Get-Content $baselinePath -Raw | ConvertFrom-Json)) { $baseline[$b.Key] = [int]$b.Count }
        }
        $new = @()
        foreach ($k in $current.Keys) {
            $accepted = [int]$baseline[$k]
            if ($current[$k] -le $accepted) { continue }
            $new += $k
            $lines = @($findings | Where-Object { (Get-FindingKey $_) -eq $k } | ForEach-Object Line) -join ', '
            Write-Host ("  New finding: {0} now {1}, accepted {2}. Lines: {3}" -f $k, $current[$k], $accepted, $lines) -ForegroundColor Yellow
        }
        Add-CheckResult 'PSScriptAnalyzer' (-not $new) $(if ($new) { "$($new.Count) new finding(s) not in baseline" } else { "$($findings.Count) findings, all in baseline" })
    }
} catch {
    Add-CheckResult 'PSScriptAnalyzer' $false "Could not run: $($_.Exception.Message)"
}
#endregion

#region Pester
try {
    $pester = Get-Module Pester -ListAvailable | Where-Object { $_.Version -ge [version]'5.0' } | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $pester) { throw 'Pester 5 or later is not installed (Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser -SkipPublisherCheck)' }
    Import-Module $pester.Path -Force -ErrorAction Stop
    $config = New-PesterConfiguration
    $config.Run.Path = (Join-Path $RepoRoot 'Tests')
    $config.Run.PassThru = $true
    $config.Output.Verbosity = 'Normal'
    if ($TestResultPath) {
        $config.TestResult.Enabled = $true
        $config.TestResult.OutputFormat = 'NUnitXml'
        $config.TestResult.OutputPath = $TestResultPath
    }
    $run = Invoke-Pester -Configuration $config
    Add-CheckResult 'Pester' ($run.FailedCount -eq 0 -and $run.PassedCount -gt 0) "$($run.PassedCount) passed, $($run.FailedCount) failed (Pester $($pester.Version))"
} catch {
    Add-CheckResult 'Pester' $false "Could not run: $($_.Exception.Message)"
}
#endregion

$results | Format-Table -AutoSize -Wrap | Out-String -Width 300 | Write-Host
if ($results | Where-Object Result -eq 'FAIL') {
    Write-Host 'Checks FAILED. Do not build.' -ForegroundColor Red
    exit 1
}
Write-Host 'All checks passed.' -ForegroundColor Green
exit 0
