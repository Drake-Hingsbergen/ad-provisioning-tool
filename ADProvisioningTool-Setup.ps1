# ============================================================
#  ADProvisioningTool-Setup.ps1
#  Run ONCE on the management server to encrypt the Graph
#  client secret and write ADProvisioningTool.json.
#  Replaces the legacy ADProvisioningTool.config (flat key=value).
# ============================================================
$ErrorActionPreference = "Continue"
trap {
    Write-Host "  [FATAL] $_" -ForegroundColor Red
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$JsonPath   = Join-Path $ScriptDir "ADProvisioningTool.json"
$LegacyPath = Join-Path $ScriptDir "ADProvisioningTool.config"

Write-Host ""
Write-Host "  AD Provisioning Tool -- Setup" -ForegroundColor Green
Write-Host "  ======================" -ForegroundColor Green
Write-Host "  Config: $JsonPath" -ForegroundColor Gray
Write-Host ""

# Read existing values (from JSON or legacy config)
$existingTenant = ""; $existingClient = ""
if (Test-Path $JsonPath) {
    $ex = Get-Content $JsonPath -Raw | ConvertFrom-Json
    $existingTenant = $ex.TenantId; $existingClient = $ex.ClientId
    Write-Host "  Existing JSON config found." -ForegroundColor Gray
} elseif (Test-Path $LegacyPath) {
    $ex = Get-Content $LegacyPath | ConvertFrom-StringData
    $existingTenant = $ex.TenantId; $existingClient = $ex.ClientId
    Write-Host "  Legacy .config found -- will migrate to JSON." -ForegroundColor Yellow
}

# Prompt for values (pre-fill from existing)
Write-Host ""
$tenantId = Read-Host "  TenantId  [$existingTenant]"
if (-not $tenantId) { $tenantId = $existingTenant }

$clientId = Read-Host "  ClientId  [$existingClient]"
if (-not $clientId) { $clientId = $existingClient }

Write-Host "  ClientSecret (hidden):"
$secureSecret = Read-Host -AsSecureString
$plainSecret  = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret))

if ([string]::IsNullOrWhiteSpace($plainSecret)) {
    Write-Host "  [X] No secret entered." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); exit 1
}

# Test DPAPI
Write-Host "  Testing DPAPI..." -ForegroundColor Gray
try {
    Add-Type -AssemblyName System.Security -ErrorAction Stop
    $t = [System.Security.Cryptography.ProtectedData]::Protect(
        [Text.Encoding]::UTF8.GetBytes("test"), $null,
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
    [System.Security.Cryptography.ProtectedData]::Unprotect(
        $t, $null,
        [System.Security.Cryptography.DataProtectionScope]::LocalMachine) | Out-Null
    Write-Host "  DPAPI OK." -ForegroundColor Green
} catch {
    Write-Host "  [X] DPAPI failed: $_" -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); exit 1
}

# Encrypt secret
$secretBytes     = [Text.Encoding]::UTF8.GetBytes($plainSecret)
$encryptedBytes  = [System.Security.Cryptography.ProtectedData]::Protect(
    $secretBytes, $null,
    [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
$encryptedBase64 = [Convert]::ToBase64String($encryptedBytes)

# Write JSON config
$config = [ordered]@{
    TenantId    = $tenantId
    ClientId    = $clientId
    ClientSecret= $encryptedBase64
    IsEncrypted = $true
}
# Keep an existing AuthMode (for example ManagedIdentity) so rotating the secret does not switch modes.
if (Test-Path $JsonPath) {
    try {
        $prior = Get-Content $JsonPath -Raw | ConvertFrom-Json
        if ($prior.AuthMode) {
            $config.AuthMode = $prior.AuthMode
            Write-Host "  Keeping AuthMode: $($prior.AuthMode)" -ForegroundColor Gray
        }
    } catch { Write-Host "  [!] Could not read the existing config to keep AuthMode: $_" -ForegroundColor Yellow }
}
$config | ConvertTo-Json | Set-Content -Path $JsonPath -Encoding UTF8

Write-Host ""
Write-Host "  [OK] ADProvisioningTool.json written and secret encrypted." -ForegroundColor Green
Write-Host "       Encrypted value is only decryptable on THIS machine." -ForegroundColor Green

# Verify
$v = Get-Content $JsonPath -Raw | ConvertFrom-Json
$vb = [Convert]::FromBase64String($v.ClientSecret)
$vd = [System.Security.Cryptography.ProtectedData]::Unprotect(
    $vb, $null,
    [System.Security.Cryptography.DataProtectionScope]::LocalMachine)
if ([Text.Encoding]::UTF8.GetString($vd) -eq $plainSecret) {
    Write-Host "  [OK] Verification passed." -ForegroundColor Green
} else {
    Write-Host "  [!] Verification mismatch." -ForegroundColor Yellow
}

# Offer to delete legacy config
if (Test-Path $LegacyPath) {
    $del = Read-Host "  Delete legacy ADProvisioningTool.config? (Y/N)"
    if ($del -match "^[Yy]$") { Remove-Item $LegacyPath; Write-Host "  Legacy config deleted." -ForegroundColor Gray }
}

$plainSecret = $null; [GC]::Collect()
Write-Host ""
Write-Host "  Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
