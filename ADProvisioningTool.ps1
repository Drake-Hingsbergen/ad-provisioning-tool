# ============================================================
#  ADProvisioningTool.ps1  v2.0
#  Full account creation flow:
#  Pre-flight -> New-RemoteMailbox -> Entra sync wait -> License group -> AUs
#  Compiled to ADProvisioningTool.exe with ps2exe; run from C:\ADProvisioning as a normal operator.
#  Every name in this file (company, domains, divisions, groups, servers) is a placeholder.
# ============================================================

Import-Module ActiveDirectory -ErrorAction Stop
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Security   # required for ProtectedData (DPAPI)

# Suppress ALL warnings and progress bars at the script level.
# ps2exe converts PowerShell warnings into blocking modal dialog boxes.
# This prevents ANY cmdlet (AD, Exchange, Graph) from triggering popups.
$WarningPreference  = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

[System.Windows.Forms.Application]::EnableVisualStyles()

# Force TLS 1.2 for all web requests
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ---- COLORS ------------------------------------
$ColorDark  = [System.Drawing.ColorTranslator]::FromHtml("#1F2A44")
$ColorAccent = [System.Drawing.ColorTranslator]::FromHtml("#0B5CAD")
$ColorLight = [System.Drawing.ColorTranslator]::FromHtml("#A9C7E8")
$ColorPale  = [System.Drawing.ColorTranslator]::FromHtml("#DCE6F2")

# Pre-defined custom RGB colors (single-int ARGB for ps2exe compatibility)
$ColorRGB_100_40_40 = [System.Drawing.Color]::FromArgb(-10213336)
$ColorRGB_120_80_0 = [System.Drawing.Color]::FromArgb(-8892416)
$ColorRGB_15_15_15 = [System.Drawing.Color]::FromArgb(-15790321)
$ColorRGB_150_150_150 = [System.Drawing.Color]::FromArgb(-6908266)
$ColorRGB_180_255_180 = [System.Drawing.Color]::FromArgb(-4915276)
$ColorRGB_180_60_60 = [System.Drawing.Color]::FromArgb(-4965316)
$ColorRGB_20_20_20 = [System.Drawing.Color]::FromArgb(-15461356)
$ColorRGB_20_20_50 = [System.Drawing.Color]::FromArgb(-15461326)
$ColorRGB_30_30_30 = [System.Drawing.Color]::FromArgb(-14803426)
$ColorRGB_50_50_50 = [System.Drawing.Color]::FromArgb(-13487566)
$ColorRGB_55_45_5 = [System.Drawing.Color]::FromArgb(-13161211)
$ColorRGB_60_50_10 = [System.Drawing.Color]::FromArgb(-12832246)
$ColorRGB_60_60_60 = [System.Drawing.Color]::FromArgb(-12829636)
$ColorRGB_70_15_15 = [System.Drawing.Color]::FromArgb(-12185841)
$ColorRGB_70_45_5 = [System.Drawing.Color]::FromArgb(-12178171)
$ColorRGB_80_20_20 = [System.Drawing.Color]::FromArgb(-11529196)
$ColorGray       = [System.Drawing.ColorTranslator]::FromHtml("#F5F4F0")
$ColorWhite      = [System.Drawing.Color]::White
$ColorRed        = [System.Drawing.ColorTranslator]::FromHtml("#CC0000")
$ColorGold       = [System.Drawing.ColorTranslator]::FromHtml("#B8860B")

# ============================================================
#  >>> SITE SETTINGS - YOU MUST SET THESE <<<
# ============================================================
# Exchange on-prem PowerShell endpoint
$ExchangeServer = "http://exch01.res.contoso.local/PowerShell/"

# Audit log. Lives in the 'logs' subfolder under the tool folder.
# The folder is created automatically on first run if it's missing.
# Whoever runs the tool needs Modify (write/append) on that folder.
$AuditLogPath   = "C:\ADProvisioning\logs\ADProvisioningTool-Audit.csv"
# ============================================================

# ---- M365 LICENSE GROUPS -----------------------------------
$LicenseGroups = @{
    "Frontline Worker"  = "M365 License - Frontline Worker"
    "Knowledge Worker"  = "M365 License - Knowledge Worker"
}

# ---- CUSTOM AD ATTRIBUTES (orgCategoryID / orgAccountType) --
# Written after New-RemoteMailbox through ADSI as the operator (Set-ADUser raises a warning that ps2exe turns into a pop-up).
# Sample values. Replace both lists with your directory's own.
$CategoryList = @(
    "1 - General", "2 - Supervisor", "3 - Manager", "4 - Director", "5 - Executive", "9 - Non-Employee"
)
$AccountTypeList = @(
    "Primary", "Admin", "Service", "Contractor"
)

# Sample division map. Replace with your own divisions.
#   Code        = division code used in the display name: "First Last - CODE"
#   Domain      = AD domain the account is created in
#   EmailDomain = vanity UPN/email domain (omit for contoso.com)
#   AU          = Entra Administrative Unit(s), comma-separated for more than one
$DivisionMap = @(
    @{ Code="FIELD SVC"; Name="Field Service"; Domain="east.contoso.local"; AU="FIELD,OPS" },
    @{ Code="FIN"; Name="Finance"; Domain="east.contoso.local"; AU="FIN" },
    @{ Code="HR"; Name="Human Resources"; Domain="east.contoso.local"; AU="HR" },
    @{ Code="ITS"; Name="IT Services"; Domain="central.contoso.local"; AU="ITS" },
    @{ Code="LEGAL"; Name="Legal"; Domain="central.contoso.local"; AU="LEGAL" },
    @{ Code="MKT"; Name="Marketing"; Domain="west.contoso.local"; EmailDomain="fabrikam.com"; AU="MKT" },
    @{ Code="OPS"; Name="Operations"; Domain="west.contoso.local"; AU="OPS" },
    @{ Code="R-D"; Name="Research and Development"; Domain="central.contoso.local"; AU="RD" },
    @{ Code="SALES"; Name="Sales"; Domain="east.contoso.local"; EmailDomain="northwindtraders.com"; AU="SALES" }
)

# Sort divisions alphabetically by code
$DivisionMap = $DivisionMap | Sort-Object { $_.Code }

# ---- ALL DOMAINS FOR SEARCH --------------------------------
$AllDomains = @(
    "contoso.local","east.contoso.local","west.contoso.local",
    "central.contoso.local","res.contoso.local"
)

# ---- AU (Administrative Unit) MAPPING ----------------------
function Get-AUCodes {
    # Returns the Entra AU display name(s) for a division code, from the AU
    # field in $DivisionMap. Falls back to the division code itself.
    param([string]$DivCode)
    $d = $DivisionMap | Where-Object { $_.Code -eq $DivCode }
    if ($d -and $d.AU) { return @($d.AU -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
    return @($DivCode)
}

function Get-DivCodeFromItem {
    # Pulls the division code out of a dropdown item "CODE - Name".
    # Splits on " - " so codes with spaces or hyphens (FIELD SVC, R-D) work.
    param($Item)
    return (("$Item") -split ' - ', 2)[0].Trim()
}

function Wait-UiDelay {
    # Start-Sleep alone freezes the window, Windows marks it Not Responding, and a tech who
    # ends the task leaves a half-built account with no audit row.
    param([int]$Seconds)
    $end = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $end) {
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.Application]::DoEvents()
    }
}

function Get-GraphErrorDetail {
    # Reads the full Graph error body from a failed Invoke-RestMethod call.
    param($ErrRecord)
    $detail = "$ErrRecord"
    try {
        if ($ErrRecord.Exception.Response) {
            $errStream = $ErrRecord.Exception.Response.GetResponseStream()
            $errReader = New-Object System.IO.StreamReader -ArgumentList $errStream
            $errReader.BaseStream.Position = 0
            $detail = $errReader.ReadToEnd()
            $errReader.Close()
        }
    } catch {}
    return $detail
}

function Get-AdministrativeUnitList {
    # Every AU in the tenant (id, displayName), following @odata.nextLink. Throws on a Graph error.
    param($Headers)
    $items = @()
    $uri = "https://graph.microsoft.com/v1.0/directory/administrativeUnits?`$select=id,displayName&`$top=999"
    while ($uri) {
        $page = Invoke-RestMethod -Uri $uri -Headers $Headers -ErrorAction Stop
        $items += @($page.value)
        $uri = $page.'@odata.nextLink'
    }
    return $items
}

function Get-EntraTargetProblem {
    # Checks at Preview time that the license group and the division's AUs exist in Entra,
    # so a wrong name in $DivisionMap shows up before the 45 minute sync wait, not after.
    # Returns one message per problem; an empty result means everything was found.
    param($AUNames, [string]$GroupName, $Headers)
    $problems = @()
    if (-not $Headers) { return @('Graph token unavailable -- AU and license group names not checked') }
    try {
        $auItems = Get-AdministrativeUnitList -Headers $Headers
        foreach ($n in @($AUNames)) {
            if (-not @($auItems | Where-Object { $_.displayName -eq $n }).Count) { $problems += "AU '$n' not found in Entra -- STEP 5 will fail for it" }
        }
    } catch { $problems += "AU list could not be read: $(Get-GraphErrorDetail $_)" }
    if ($GroupName) {
        try {
            $gf = [Uri]::EscapeDataString("displayName eq '$($GroupName -replace "'", "''")'")
            $g = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$gf&`$select=id" -Headers $Headers -ErrorAction Stop
            if (-not @($g.value).Count) { $problems += "License group '$GroupName' not found in Entra -- STEP 4 will fail" }
        } catch { $problems += "License group could not be checked: $(Get-GraphErrorDetail $_)" }
    }
    return $problems
}

function Add-UserToAUs {
    # Adds a user to every AU in AUNames. Returns one result per AU: AU, OK, Detail.
    # A user who is already a member counts as OK (safe to re-run).
    # A user that synced seconds ago can still be unknown to the AU service
    # (Request_ResourceNotFound), so that error is retried before it counts as a failure.
    param($UserId, $AUNames, $Headers, [int]$RetryCount = 3, [int]$RetryDelaySeconds = 15)
    $results = @()
    try {
        $auItems = Get-AdministrativeUnitList -Headers $Headers
    } catch {
        $d = Get-GraphErrorDetail $_
        foreach ($n in $AUNames) { $results += [PSCustomObject]@{ AU = $n; OK = $false; Detail = "AU list lookup failed: $d" } }
        return $results
    }
    foreach ($n in $AUNames) {
        $match = @($auItems | Where-Object { $_.displayName -eq $n })
        if ($match.Count -eq 0) { $results += [PSCustomObject]@{ AU = $n; OK = $false; Detail = "AU not found in Entra" }; continue }
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$UserId" } | ConvertTo-Json
        $result = $null
        for ($try = 1; $try -le $RetryCount -and -not $result; $try++) {
            try {
                Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/directory/administrativeUnits/$($match[0].id)/members/`$ref" `
                    -Method POST -Headers $Headers -Body $body -ContentType "application/json" -ErrorAction Stop | Out-Null
                $result = [PSCustomObject]@{ AU = $n; OK = $true; Detail = "ID=$($match[0].id)" }
            } catch {
                $d = Get-GraphErrorDetail $_
                if ($d -match 'already exist') { $result = [PSCustomObject]@{ AU = $n; OK = $true; Detail = "already a member" } }
                elseif ($d -match 'Request_ResourceNotFound' -and $try -lt $RetryCount) { Wait-UiDelay -Seconds $RetryDelaySeconds }
                else { $result = [PSCustomObject]@{ AU = $n; OK = $false; Detail = $d } }
            }
        }
        $results += $result
    }
    return $results
}

function Get-OUPath {
    param($Domain)
    $parts = $Domain.Split(".")
    $dc = ($parts | ForEach-Object { "DC=$_" }) -join ","
    return "OU=New Accounts,$dc"
}

function Get-EmailDomain {
    # Returns the division's vanity email/UPN domain, or contoso.com if none is set.
    # To give a division a vanity domain, add  EmailDomain="fabrikam.com"  to its
    # entry in $DivisionMap above. Divisions with no EmailDomain default here.
    param([string]$DivCode)
    $d = $DivisionMap | Where-Object { $_.Code -eq $DivCode }
    if ($d -and $d.EmailDomain) { return $d.EmailDomain }
    return "contoso.com"
}

function Format-ProperCase {
    # Mixed-case input (McDonald, DeLuca, van der Berg) is kept as typed; only all-lower or
    # all-upper input is re-cased. Each word and each part after a hyphen or apostrophe gets a
    # capital, Mc names keep the inner capital, and II/III/IV stay upper case.
    param([string]$Value)
    $v = ("$Value".Trim() -replace '\s+', ' ')
    if (-not $v) { return "" }
    if ($v -cmatch '[a-z]' -and $v -cmatch '[A-Z]') { return $v }
    $v = $v.ToLower()
    $v = [regex]::Replace($v, "(^|[ \-'])([a-z])", { param($m) $m.Groups[1].Value + $m.Groups[2].Value.ToUpper() })
    $v = [regex]::Replace($v, "(^|[ \-])Mc([a-z])", { param($m) $m.Groups[1].Value + 'Mc' + $m.Groups[2].Value.ToUpper() })
    $v = [regex]::Replace($v, "(?<=^| )(Ii|Iii|Iv)(?= |$)", { param($m) $m.Value.ToUpper() })
    return $v
}

function Get-ArcChallengePath {
    # The Arc agent answers the first token request with a 401 whose WWW-Authenticate
    # header points at a key file. Only local admins (elevated) and members of
    # "Hybrid agent extension applications" can read that file, which is how the
    # agent decides who may use the server's managed identity.
    param([string]$Uri)
    try {
        Invoke-WebRequest -Uri $Uri -Headers @{ Metadata = 'true' } -UseBasicParsing -ErrorAction Stop | Out-Null
    } catch [System.Net.WebException] {
        $response = $_.Exception.Response
        if (-not $response) { throw "Arc agent not reachable at $Uri : $($_.Exception.Message)" }
        $wwwAuth = $response.Headers['WWW-Authenticate']
        if ($wwwAuth -notmatch '^Basic realm=(.+)$') { throw "Arc agent returned no challenge (HTTP $([int]$response.StatusCode))." }
        return $Matches[1].Trim()
    }
    throw "Arc agent issued a token without a challenge. Unexpected agent behavior."
}

function Get-ArcManagedIdentityToken {
    param([string]$Resource = 'https://graph.microsoft.com')
    $endpoint = $env:IDENTITY_ENDPOINT
    if (-not $endpoint) { $endpoint = 'http://localhost:40342/metadata/identity/oauth2/token' }
    $uri = '{0}?api-version=2020-06-01&resource={1}' -f $endpoint, [Uri]::EscapeDataString($Resource)

    $keyPath = Get-ArcChallengePath -Uri $uri
    try {
        $key = (Get-Content -Path $keyPath -Raw -ErrorAction Stop).Trim()
    } catch {
        throw "Cannot read Arc challenge file. Run as a member of 'Hybrid agent extension applications' (or elevated as local admin). $($_.Exception.Message)"
    }
    $resp = Invoke-RestMethod -Uri $uri -Headers @{ Metadata = 'true'; Authorization = "Basic $key" } -ErrorAction Stop
    if (-not $resp.access_token) { throw "Arc agent returned no access token." }
    return $resp.access_token
}

function Get-GraphToken {
    # Returns Graph request headers, or $null with the reason written to the debug log.
    # AuthMode "ManagedIdentity" in ADProvisioningTool.json uses the server's Azure Arc identity;
    # anything else uses the app registration client secret (DPAPI machine scope).
    # The token is cached until 5 minutes before it expires, so callers can ask for it
    # before every Graph step without a new request each time. -Force skips the cache
    # (use it after a 401).
    param([string]$ScriptDir, [switch]$Force)
    if ($Force) { $Script:GraphTokenCache = $null }
    $cached = $Script:GraphTokenCache
    if (-not $Force -and $cached -and $cached.ScriptDir -eq $ScriptDir -and (Get-Date) -lt $cached.RefreshAt) {
        return $cached.Headers
    }
    try {
        $jsonPath = Join-Path $ScriptDir "ADProvisioningTool.json"
        $cfgPath  = Join-Path $ScriptDir "ADProvisioningTool.config"
        $cfg = $null
        if (Test-Path $jsonPath) {
            $cfg = Get-Content $jsonPath -Raw | ConvertFrom-Json
        } elseif (Test-Path $cfgPath) {
            $flat = Get-Content $cfgPath | ConvertFrom-StringData
            $cfg  = [PSCustomObject]@{
                TenantId     = $flat.TenantId
                ClientId     = $flat.ClientId
                ClientSecret = $flat.ClientSecret
                IsEncrypted  = ($flat.IsEncrypted -eq "true")
            }
        } else {
            throw "No config file found in $ScriptDir"
        }

        if ("$($cfg.AuthMode)" -eq 'ManagedIdentity') {
            $accessToken = Get-ArcManagedIdentityToken -Resource 'https://graph.microsoft.com'
            # The Arc agent can return a token it cached earlier, so its remaining life is unknown.
            # Asking again every 5 minutes is cheap because the agent answers from that cache.
            $lifetime = 600
        } else {
            $clientSecret = $cfg.ClientSecret
            if ($cfg.IsEncrypted -eq $true -or $cfg.IsEncrypted -eq "true") {
                try {
                    $encBytes     = [Convert]::FromBase64String($cfg.ClientSecret)
                    $decBytes     = [System.Security.Cryptography.ProtectedData]::Unprotect(
                        $encBytes, $null,
                        [System.Security.Cryptography.DataProtectionScope]::LocalMachine
                    )
                    $clientSecret = [System.Text.Encoding]::UTF8.GetString($decBytes)
                } catch {
                    throw "DPAPI decryption failed -- wrong machine or corrupt config: $_"
                }
            }
            $body = @{
                grant_type    = "client_credentials"
                scope         = "https://graph.microsoft.com/.default"
                client_id     = $cfg.ClientId
                client_secret = $clientSecret
            }
            $resp = Invoke-RestMethod `
                -Uri "https://login.microsoftonline.com/$($cfg.TenantId)/oauth2/v2.0/token" `
                -Method POST -Body $body -ErrorAction Stop
            $accessToken = $resp.access_token
            $lifetime = 3599
            if ($resp.expires_in) { $lifetime = [int]$resp.expires_in }
        }
        $headers = @{ Authorization = "Bearer $accessToken"; "Content-Type" = "application/json" }
        $Script:GraphTokenCache = @{ ScriptDir = $ScriptDir; Headers = $headers; RefreshAt = (Get-Date).AddSeconds([Math]::Max(60, $lifetime - 300)) }
        return $headers
    } catch {
        $Script:GraphTokenCache = $null
        DBG "GRAPH_TOKEN" "FAIL" "$($_.Exception.Message)"
        return $null
    }
}

function Get-NameInputProblem {
    # Rejects names that would produce a broken or misleading address. Get-NameCandidate keeps
    # only a-z and 0-9, so an accented letter is dropped (Jose with an accent became jos) and a
    # name with no ASCII letters at all gives ".smith". Accent conversion waits on a decision to
    # match the HR onboarding workflow, so for now the tech types the name without accents.
    param([string]$First, [string]$Middle, [string]$Last)
    $problems = @()
    foreach ($field in @(@{ N = 'First name'; V = $First }, @{ N = 'Last name'; V = $Last }, @{ N = 'Middle initial'; V = $Middle })) {
        $v = "$($field.V)".Trim()
        if (-not $v) { continue }
        if ($v -match '[0-9]') { $problems += "$($field.N) contains a digit" }
        if ($v -match "[^A-Za-z0-9 \-'\.]") {
            if ($v -match '[^\x00-\x7F]') { $problems += "$($field.N) has accented or non-English letters -- type it without accents" }
            else { $problems += "$($field.N) has characters other than letters, space, hyphen, apostrophe or period" }
        }
        if ($field.N -ne 'Middle initial' -and $v -notmatch '[A-Za-z]') { $problems += "$($field.N) has no letters" }
    }
    return $problems
}

function Get-NameLengthProblem {
    # AD rejects a CN over 64 characters and Exchange rejects an address local part over 64,
    # but only when New-RemoteMailbox runs. Checking at Preview gives the tech the reason.
    param([string]$CnName, [string]$LocalPart)
    $problems = @()
    if ($CnName.Length -gt 64) { $problems += "AD name '$CnName' is $($CnName.Length) characters (limit 64)" }
    if ($LocalPart.Length -gt 64) { $problems += "Email name '$LocalPart' is $($LocalPart.Length) characters (limit 64)" }
    return $problems
}

function Get-NameCandidate {
    # Returns the UPN/email local part for one attempt, full length.
    # Order matches the HR onboarding workflow so both paths produce the
    # same address for the same person: first.last, first.m.last (only when the
    # middle initial is exactly one letter or digit), then first.last1, first.last2, ...
    param($First, $Middle, $Last, [int]$Attempt)
    $f = ("$First".ToLower()  -replace "[^a-z0-9]", "")
    $m = ("$Middle".ToLower() -replace "[^a-z0-9]", "")
    $l = ("$Last".ToLower()   -replace "[^a-z0-9]", "")
    if ($m.Length -ne 1) { $m = "" }
    $candidates = @("$f.$l")
    if ($m) { $candidates += "$f.$m.$l" }
    for ($n = 1; $candidates.Count -lt 5; $n++) { $candidates += "$f.$l$n" }
    return $candidates[$Attempt - 1]
}

function Truncate-SAM {
    # sAMAccountName is limited to 20 characters and cannot end with a period.
    param([string]$Sam)
    if ($Sam.Length -gt 20) { $Sam = $Sam.Substring(0, 20) }
    return $Sam.TrimEnd('.')
}

function Get-SamCandidate {
    # Index 0 is the local part cut to 20 characters. Indexes 1-9 cut to 19 and add
    # the digit, so a collision on a long name still yields distinct SAMs
    # (christopher.lancaste -> christopher.lancast1). Same rule as the HR onboarding workflow.
    param([string]$LocalPart, [int]$Index)
    $base = Truncate-SAM $LocalPart
    if ($Index -eq 0) { return $base }
    return Truncate-SAM ($base.Substring(0, [Math]::Min($base.Length, 19)) + "$Index")
}

function Get-ForestGlobalCatalog {
    # One Global Catalog query covers every domain in the forest, so a UPN or
    # address held in any domain is found, synced or not.
    if (-not $Script:ForestGC) {
        $dc = Get-ADDomainController -Discover -Service GlobalCatalog -ErrorAction Stop
        $Script:ForestGC = "$(@($dc.HostName)[0]):3268"
    }
    return $Script:ForestGC
}

function Test-GraphAddressTaken {
    # Graph OData filters on proxyAddresses are case-sensitive, so both SMTP: and smtp: are checked.
    param([string]$Address, $GraphHeaders, [switch]$IncludeMail)
    $encL = [Uri]::EscapeDataString("smtp:$Address")
    $encU = [Uri]::EscapeDataString("SMTP:$Address")
    $filter = "proxyAddresses/any(x:x eq '$encL') or proxyAddresses/any(x:x eq '$encU')"
    if ($IncludeMail) { $filter += " or mail eq '$([Uri]::EscapeDataString($Address))'" }
    $resp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=id,userPrincipalName&`$top=1" `
        -Headers $GraphHeaders -ErrorAction Stop
    if ($resp.value -and @($resp.value).Count -gt 0) { return @($resp.value)[0].userPrincipalName }
    return $null
}

function Get-BatchClash {
    # Returns the row index an earlier bulk row already holds for this name, or $null.
    # Local parts clash batch-wide (routing address and contoso.com alias are shared by
    # every email domain); SAMs only clash inside the same AD domain.
    param([hashtable]$ClaimedLocal, [hashtable]$ClaimedSam, $LocalPart, $Domain, $Sam)
    if ($ClaimedLocal.ContainsKey($LocalPart)) { return $ClaimedLocal[$LocalPart] }
    if ($ClaimedSam.ContainsKey("$Domain|$Sam")) { return $ClaimedSam["$Domain|$Sam"] }
    return $null
}

function Get-BatchExclusion {
    param([hashtable]$ClaimedLocal, [hashtable]$ClaimedSam, $Domain)
    return @{
        LocalParts = @($ClaimedLocal.Keys)
        SAMs       = @($ClaimedSam.Keys | Where-Object { $_ -like "$Domain|*" } | ForEach-Object { $_.Split('|')[1] })
    }
}

function ConvertTo-LdapFilterValue {
    # RFC 4515 escaping, so names like "Pat O'Neil (Jr)" are safe inside an LDAP filter.
    param([string]$Value)
    return ($Value -replace '\\', '\5c' -replace '\*', '\2a' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00')
}

function Test-ProxyAddressPresent {
    # Exchange returns proxy addresses as "smtp:x" / "SMTP:x"; address comparison is case-insensitive.
    param($Addresses, [string]$Address)
    foreach ($a in @($Addresses)) {
        if ("$a" -match '^(smtp|SMTP):(.+)$' -and $Matches[2] -ieq $Address) { return $true }
    }
    return $false
}

function Get-NewRemoteMailboxParameter {
    # PrimarySmtpAddress is always set so the mailbox's email matches the UPN. Without it the
    # email address policy builds First.Last2@contoso.com on its own when first.last is taken.
    param($Resolved, $CnName, $DispName, $First, $Last, $Ou, $Pdc, [System.Security.SecureString]$Password)
    return @{
        Name                         = $CnName
        DisplayName                  = $DispName
        FirstName                    = $First
        LastName                     = $Last
        UserPrincipalName            = $Resolved.UPN
        SamAccountName               = $Resolved.SAM
        Password                     = $Password
        ResetPasswordOnNextLogon     = $true
        OnPremisesOrganizationalUnit = $Ou
        RemoteRoutingAddress         = $Resolved.Routing
        PrimarySmtpAddress           = $Resolved.PrimarySmtp
        DomainController             = $Pdc
    }
}

function Get-CreatedObjectState {
    # After New-RemoteMailbox throws, the account may still have been written (for example
    # the WinRM call dropped after Exchange finished). Retrying or re-running Create would
    # then make a second, numbered account, so the caller checks first.
    # Returns Exists (with DN), Absent, or Unknown (with the lookup error).
    param([string]$Sam, [string]$Pdc)
    try {
        $hit = Get-ADObject -Server $Pdc -LDAPFilter "(sAMAccountName=$(ConvertTo-LdapFilterValue $Sam))" -ErrorAction Stop | Select-Object -First 1
    } catch {
        return [PSCustomObject]@{ State = 'Unknown'; Detail = "$($_.Exception.Message)" }
    }
    if ($hit) { return [PSCustomObject]@{ State = 'Exists'; Detail = "$($hit.DistinguishedName)" } }
    return [PSCustomObject]@{ State = 'Absent'; Detail = '' }
}

function Get-SyncedEntraUser {
    # One sync poll. State is Synced, NotYet, WrongObject or UpnChanged. Graph errors other
    # than 404 are rethrown so the caller can refresh the token on 401.
    # UpnChanged: the object we created is in Entra but under another UPN, which is what
    # Entra Connect does when the UPN is already taken in the tenant.
    param([string]$Upn, [string]$ImmutableId, $Headers)
    $sel = '$select=id,displayName,userPrincipalName,mail,onPremisesImmutableId,onPremisesSyncEnabled'
    $user = $null
    try {
        $user = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$([Uri]::EscapeDataString($Upn))?$sel" -Headers $Headers -ErrorAction Stop
    } catch {
        $code = 0
        try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 }
        if ($code -ne 404) { throw }
    }
    if ($user -and $user.id) {
        if (-not $ImmutableId) { return [PSCustomObject]@{ State = 'Synced'; User = $user; Detail = 'UPN match (no ImmutableId to verify)' } }
        if ($user.onPremisesImmutableId -eq $ImmutableId) { return [PSCustomObject]@{ State = 'Synced'; User = $user; Detail = 'ImmutableId match' } }
    }
    # Entra Connect renames our object exactly when another object already holds the UPN,
    # so the ImmutableId lookup runs both on a 404 and when the UPN belongs to someone else.
    if ($ImmutableId) {
        $f = [Uri]::EscapeDataString("onPremisesImmutableId eq '$ImmutableId'")
        $r = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$filter=$f&$sel" -Headers $Headers -ErrorAction Stop
        $byId = @($r.value) | Select-Object -First 1
        if ($byId -and $byId.id) {
            return [PSCustomObject]@{ State = 'UpnChanged'; User = $byId; Detail = "Synced as $($byId.userPrincipalName) instead of $Upn" }
        }
    }
    if ($user -and $user.id) {
        $kind = if ($user.onPremisesImmutableId) { 'a different synced account' } else { 'a cloud-only account' }
        return [PSCustomObject]@{ State = 'WrongObject'; User = $user; Detail = "UPN $Upn belongs to $kind" }
    }
    return [PSCustomObject]@{ State = 'NotYet'; User = $null; Detail = 'Not in Entra yet' }
}

function Resolve-AccountCandidate {
    # Pre-flight resolution before any write. Returns a hashtable whose Status is:
    #   Resolved   - every check completed and the values are free
    #   Exhausted  - every candidate is taken
    #   Unverified - a check could not be completed; never treated as free
    # UPN, primary address, contoso.com alias and routing address share the full
    # local part. The SAM is resolved separately (20-character limit).
    # ExcludeLocalParts / ExcludeSAMs: values already claimed by other rows in a bulk batch.
    param(
        $First, $Middle, $Last, $Domain, $EmailDomain,
        $GraphHeaders = $null, [switch]$ExchSessionActive,
        $ExcludeLocalParts = @(), $ExcludeSAMs = @()
    )
    $unverified = {
        param($Reason)
        DBG "PREFLIGHT_UNVERIFIED" "FAIL" $Reason
        return @{ Status = 'Unverified'; Reason = $Reason }
    }

    if (-not $GraphHeaders) { return (& $unverified "Graph token unavailable -- Entra cannot be checked for cloud-only conflicts") }
    try { $gc = Get-ForestGlobalCatalog } catch { return (& $unverified "No Global Catalog reachable for the forest-wide check: $($_.Exception.Message)") }
    try { $pdc = (Get-ADDomain -Identity $Domain -Server $Domain -ErrorAction Stop).PDCEmulator }
    catch { return (& $unverified "Cannot reach the PDC for ${Domain}: $($_.Exception.Message)") }
    DBG "PREFLIGHT_SERVERS" "INFO" "GC=$gc PDC=$pdc"

    $vanity = ($EmailDomain -ne "contoso.com")
    $localPart = $null
    $holders = New-Object System.Collections.Generic.List[string]
    for ($i = 1; $i -le 5 -and -not $localPart; $i++) {
        $lp      = Get-NameCandidate -First $First -Middle $Middle -Last $Last -Attempt $i
        $upn     = "$lp@$EmailDomain"
        $alias   = "$lp@contoso.com"
        $routing = "$lp@contoso.mail.onmicrosoft.com"
        DBG "PREFLIGHT_ATTEMPT" "INFO" "Attempt $i -- UPN=$upn Alias=$(if($vanity){$alias}else{'n/a'}) Routing=$routing"

        if ($ExcludeLocalParts -contains $lp) {
            DBG "PREFLIGHT_BATCH" "WARN" "Attempt $i -- $lp claimed by another row in batch"; $holders.Add("$lp (another row in this batch)"); continue
        }

        try {
            $ldap = "(|(userPrincipalName=$upn)(mail=$upn)(proxyAddresses=smtp:$upn)(proxyAddresses=smtp:$routing)"
            if ($vanity) { $ldap += "(mail=$alias)(proxyAddresses=smtp:$alias)" }
            $ldap += ")"
            $hit = Get-ADObject -Server $gc -SearchBase "" -LDAPFilter $ldap -Properties userPrincipalName -ErrorAction Stop | Select-Object -First 1
        } catch { return (& $unverified "Forest-wide AD check failed for ${upn}: $($_.Exception.Message)") }
        if ($hit) { DBG "PREFLIGHT_AD" "WARN" "Attempt $i -- in use in AD by $($hit.DistinguishedName)"; $holders.Add("$lp (AD: $($hit.DistinguishedName))"); continue }

        try {
            $encUpn = [Uri]::EscapeDataString($upn)
            $gUpn = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$encUpn'&`$select=id,userPrincipalName&`$top=1" `
                -Headers $GraphHeaders -ErrorAction Stop
            if ($gUpn.value -and @($gUpn.value).Count -gt 0) { DBG "PREFLIGHT_ENTRA_UPN" "WARN" "Attempt $i -- $upn taken in Entra"; $holders.Add("$lp (Entra UPN)"); continue }
            $by = Test-GraphAddressTaken -Address $upn -GraphHeaders $GraphHeaders -IncludeMail
            if ($by) { DBG "PREFLIGHT_ENTRA_SMTP" "WARN" "Attempt $i -- $upn held by $by in Entra"; $holders.Add("$lp (Entra: $by)"); continue }
            if ($vanity) {
                $by = Test-GraphAddressTaken -Address $alias -GraphHeaders $GraphHeaders
                if ($by) { DBG "PREFLIGHT_ENTRA_ALIAS" "WARN" "Attempt $i -- $alias held by $by in Entra"; $holders.Add("$lp (alias, Entra: $by)"); continue }
            }
            $by = Test-GraphAddressTaken -Address $routing -GraphHeaders $GraphHeaders
            if ($by) { DBG "PREFLIGHT_ENTRA_ROUTING" "WARN" "Attempt $i -- $routing held by $by in Entra"; $holders.Add("$lp (routing, Entra: $by)"); continue }
        } catch { return (& $unverified "Entra check failed for ${upn}: $($_.Exception.Message)") }

        if ($vanity -and $ExchSessionActive) {
            try { $eHit = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$alias'" -ErrorAction Stop }
            catch { return (& $unverified "Exchange check failed for ${alias}: $($_.Exception.Message)") }
            if ($eHit) { DBG "PREFLIGHT_EXCH_ALIAS" "WARN" "Attempt $i -- $alias taken in Exchange by $($eHit.Name)"; $holders.Add("$lp (alias, Exchange: $($eHit.Name))"); continue }
        }

        $localPart = $lp
        $attempt   = $i
    }
    if (-not $localPart) {
        DBG "PREFLIGHT_EXHAUSTED" "FAIL" "All 5 name candidates taken"
        return @{ Status = 'Exhausted'; Reason = "All 5 UPN/email candidates are taken in AD or Entra: $($holders -join '; ')" }
    }

    # sAMAccountName is unique per domain across users, groups and computers,
    # so the check uses Get-ADObject against the PDC the account will be written to.
    $sam = $null
    $samTried = @{}
    for ($k = 0; $k -le 9 -and -not $sam; $k++) {
        $candidate = Get-SamCandidate -LocalPart $localPart -Index $k
        if ($samTried.ContainsKey($candidate)) { continue }
        $samTried[$candidate] = $true
        if ($ExcludeSAMs -contains $candidate) { DBG "PREFLIGHT_BATCH" "WARN" "SAM $candidate claimed by another row in batch"; continue }
        try { $samHit = Get-ADObject -Server $pdc -LDAPFilter "(sAMAccountName=$candidate)" -ErrorAction Stop | Select-Object -First 1 }
        catch { return (& $unverified "SAM check failed for ${candidate} on ${pdc}: $($_.Exception.Message)") }
        if ($samHit) { DBG "PREFLIGHT_SAM" "WARN" "SAM $candidate taken in $Domain by $($samHit.DistinguishedName)"; continue }
        $sam = $candidate
    }
    if (-not $sam) {
        DBG "PREFLIGHT_EXHAUSTED" "FAIL" "No free SAM for $localPart in $Domain"
        return @{ Status = 'Exhausted'; Reason = "No free SAM in $Domain for $localPart (base plus 9 numbered variants taken)" }
    }

    DBG "PREFLIGHT_RESOLVED" "OK" "Attempt $attempt -- SAM=$sam UPN=$localPart@$EmailDomain"
    return @{
        Status      = 'Resolved'
        Reason      = ''
        LocalPart   = $localPart
        SAM         = $sam
        UPN         = "$localPart@$EmailDomain"
        PrimarySmtp = "$localPart@$EmailDomain"
        TenantAlias  = if ($vanity) { "$localPart@contoso.com" } else { $null }
        Routing     = "$localPart@contoso.mail.onmicrosoft.com"
        Truncated   = ($sam -ne $localPart)
        Attempt     = $attempt
        PDC         = $pdc
    }
}

function Write-Log {
    param($RichBox, $Message, $Color)
    if (-not $Color) { $Color = [System.Drawing.Color]::Black }
    $RichBox.SelectionStart  = $RichBox.TextLength
    $RichBox.SelectionLength = 0
    $RichBox.SelectionColor  = $Color
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    $RichBox.AppendText("[$timestamp] $Message`n")
    $RichBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function New-TempPassword {
    # 16 random characters with at least one upper, lower, digit and symbol, from a cryptographic
    # generator. Look-alike characters (0/O, 1/l/I) are left out because the tech reads it to the user.
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnpqrstuvwxyz', '23456789', '!#%*+-=?@')
    $all  = -join $sets
    $rng  = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $next  = { param([string]$From) $b = New-Object byte[] 4; $rng.GetBytes($b); $From[[int]([BitConverter]::ToUInt32($b, 0) % $From.Length)] }
        $chars = @(foreach ($set in $sets) { & $next $set })
        while ($chars.Count -lt 16) { $chars += & $next $all }
        # Shuffle so the four guaranteed characters are not always first.
        $keys = New-Object byte[] 16; $rng.GetBytes($keys)
        $order = 0..15 | Sort-Object { $keys[$_] }
        return -join ($order | ForEach-Object { $chars[$_] })
    } finally { $rng.Dispose() }
}

function Write-AuditLog {
    # Appends one row per action to the central CSV. NEVER logs passwords.
    # A named mutex serializes writes so two techs running the tool at the
    # same second on the same server don't corrupt the file.
    param($Action, $TargetSAM, $TargetUPN, $Division, $EmpType, $Result, $Details)
    # The ticket goes on every row of a run, so a partial account can be traced to its ticket.
    if ($Script:AuditTicket -and "$Details" -notmatch 'Ticket=') { $Details = "$Details | Ticket=$($Script:AuditTicket)" }
    $row = [PSCustomObject]@{
        Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Operator  = "$env:USERDOMAIN\$env:USERNAME"
        Machine   = $env:COMPUTERNAME
        Action    = $Action
        TargetSAM = $TargetSAM
        TargetUPN = $TargetUPN
        Division  = $Division
        EmpType   = $EmpType
        Result    = $Result
        Details   = ($Details -replace "[`r`n]+", " ")
    }
    $mtx = New-Object System.Threading.Mutex -ArgumentList $false, "Global\ADProvisioningToolAudit"
    try {
        [void]$mtx.WaitOne(5000)
        $dir = Split-Path -Parent $AuditLogPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        # Build the CSV text ourselves so we control the file handle.
        $needHeader = -not (Test-Path $AuditLogPath)
        $csvLines   = $row | ConvertTo-Csv -NoTypeInformation
        # ConvertTo-Csv yields header + data row; drop the header if the file already exists.
        $payload    = if ($needHeader) { $csvLines } else { $csvLines | Select-Object -Skip 1 }
        $text       = ($payload -join "`r`n") + "`r`n"

        # Retry on transient locks (AV scan, lingering handle from a prior write).
        $written = $false
        for ($attempt = 0; $attempt -lt 5 -and -not $written; $attempt++) {
            try {
                $fs = [System.IO.File]::Open($AuditLogPath, [System.IO.FileMode]::Append,
                        [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
                $fs.Write($bytes, 0, $bytes.Length)
                $fs.Close(); $fs.Dispose()
                $written = $true
            } catch {
                Start-Sleep -Milliseconds 150
            }
        }
        if (-not $written -and $RichLog) {
            Write-Log $RichLog "[!] Audit log busy -- row skipped after 5 retries." $ColorGold
        }
    } catch {
        # Logging must never block account creation. Surface it only if a log box exists.
        if ($RichLog) { Write-Log $RichLog "[!] Audit log write failed: $_" $ColorGold }
    } finally {
        try { $mtx.ReleaseMutex() } catch {}
        $mtx.Dispose()
    }
}

# ---- DEBUG LOG ---------------------------------------------
# Always writes silently regardless of the UI toggle.
# Columns: Timestamp, SAM, DisplayName, Operator, Step, Status, Detail
$DebugLogPath = $AuditLogPath -replace 'Audit','Debug'
$Script:DebugSessionSAM  = ""
$Script:DebugSessionName = ""

function Write-DebugLog {
    param($Step, $Status, $Detail)
    $row = [PSCustomObject]@{
        Timestamp   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        SAM         = $Script:DebugSessionSAM
        DisplayName = $Script:DebugSessionName
        Operator    = "$env:USERDOMAIN\$env:USERNAME"
        Step        = $Step
        Status      = $Status
        Detail      = ($Detail -replace "[`r`n]+"," ")
    }
    $mtx = New-Object System.Threading.Mutex -ArgumentList $false, "Global\ADProvisioningToolDebug"
    try {
        [void]$mtx.WaitOne(3000)
        $dir = Split-Path -Parent $DebugLogPath
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $needHeader = -not (Test-Path $DebugLogPath)
        $csvLines   = $row | ConvertTo-Csv -NoTypeInformation
        $payload    = if ($needHeader) { $csvLines } else { $csvLines | Select-Object -Skip 1 }
        $text       = ($payload -join "`r`n") + "`r`n"
        $done = $false
        for ($a = 0; $a -lt 3 -and -not $done; $a++) {
            try {
                $fs = [System.IO.File]::Open($DebugLogPath,[System.IO.FileMode]::Append,
                        [System.IO.FileAccess]::Write,[System.IO.FileShare]::Read)
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
                $fs.Write($bytes,0,$bytes.Length); $fs.Close(); $fs.Dispose()
                $done = $true
            } catch { Start-Sleep -Milliseconds 100 }
        }
    } catch {} finally {
        try { $mtx.ReleaseMutex() } catch {}
        $mtx.Dispose()
    }
}

$Script:ShowDebug        = $false   # toggled by ChkDebug checkbox on Create tab
$Script:CancelRequested  = $false   # set by Cancel Sync Wait button during STEP 3

function DBG {
    # Always writes to debug CSV. Shows in Creation Log only when ShowDebug is on.
    param($Step, $Status, $Detail)
    Write-DebugLog $Step $Status $Detail
    if ($Script:ShowDebug -and $RichLog) {
        Write-Log $RichLog "  [DEBUG] ${Step}: $Detail" ([System.Drawing.Color]::Cyan)
    }
}

# ============================================================
#  MAIN FORM
# ============================================================

# ---- YEARLY LOG ROTATION -----------------------------------
# If it's a new year and the current log exists, archive it.
function Invoke-LogRotation {
    param([string]$LogPath)
    if (-not (Test-Path $LogPath)) { return }
    $created = (Get-Item $LogPath).LastWriteTime
    if ($created.Year -lt (Get-Date).Year) {
        $archive = $LogPath -replace '\.csv$', "-$($created.Year).csv"
        try { Rename-Item $LogPath $archive -ErrorAction Stop } catch {}
    }
}
Invoke-LogRotation $AuditLogPath
Invoke-LogRotation $DebugLogPath

$Form                  = New-Object System.Windows.Forms.Form
$Form.Text             = "AD Provisioning Tool  v2.0"
$Form.Size             = (New-Object System.Drawing.Size -Property @{Width=980; Height=740})
$Form.StartPosition    = "CenterScreen"
$Form.BackColor        = $ColorGray
$Form.FormBorderStyle  = "FixedSingle"
$Form.MaximizeBox      = $false
$Form.Font             = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Regular))

# Header
$Header           = New-Object System.Windows.Forms.Panel
$Header.Size      = (New-Object System.Drawing.Size -Property @{Width=980; Height=60})
$Header.Location  = (New-Object System.Drawing.Point -Property @{X=0; Y=0})
$Header.BackColor = $ColorDark
$Form.Controls.Add($Header)

$HeaderLbl           = New-Object System.Windows.Forms.Label
$HeaderLbl.Text      = "AD Provisioning Tool  v2.0"
$HeaderLbl.ForeColor = $ColorWhite
$HeaderLbl.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]15),([System.Drawing.FontStyle]::Bold))
$HeaderLbl.AutoSize  = $true
$HeaderLbl.Location  = (New-Object System.Drawing.Point -Property @{X=20; Y=15})
$Header.Controls.Add($HeaderLbl)

# Last-created status bar at bottom of form
$StatusBar           = New-Object System.Windows.Forms.Panel
$StatusBar.Size      = (New-Object System.Drawing.Size -Property @{Width=980; Height=24})
$StatusBar.Location  = (New-Object System.Drawing.Point -Property @{X=0; Y=716})
$StatusBar.BackColor = $ColorDark
$Form.Controls.Add($StatusBar)

$LblLastCreated           = New-Object System.Windows.Forms.Label
$LblLastCreated.Text      = "Last created: --"
$LblLastCreated.ForeColor = $ColorRGB_180_255_180
$LblLastCreated.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Regular))
$LblLastCreated.AutoSize  = $true
$LblLastCreated.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=4})
$StatusBar.Controls.Add($LblLastCreated)

# Tabs
$Tabs          = New-Object System.Windows.Forms.TabControl
$Tabs.Size     = (New-Object System.Drawing.Size -Property @{Width=940; Height=635})
$Tabs.Location = (New-Object System.Drawing.Point -Property @{X=20; Y=75})
$Tabs.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$Form.Controls.Add($Tabs)

# ============================================================
#  TAB 1 - SEARCH
# ============================================================
$TabSearch           = New-Object System.Windows.Forms.TabPage
$TabSearch.Text      = "  Search Users  "
$TabSearch.BackColor = $ColorGray
$Tabs.Controls.Add($TabSearch)

$LblSrch          = New-Object System.Windows.Forms.Label
$LblSrch.Text     = "Search (name, SAM, UPN, email) -- supports wildcards e.g. john*"
$LblSrch.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=15})
$LblSrch.AutoSize = $true
$TabSearch.Controls.Add($LblSrch)

$TxtSearch          = New-Object System.Windows.Forms.TextBox
$TxtSearch.Size     = (New-Object System.Drawing.Size -Property @{Width=540; Height=26})
$TxtSearch.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=38})
$TxtSearch.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]10),([System.Drawing.FontStyle]::Regular))
$TabSearch.Controls.Add($TxtSearch)

$BtnSearch           = New-Object System.Windows.Forms.Button
$BtnSearch.Text      = "Search"
$BtnSearch.Size      = (New-Object System.Drawing.Size -Property @{Width=100; Height=28})
$BtnSearch.Location  = (New-Object System.Drawing.Point -Property @{X=558; Y=37})
$BtnSearch.BackColor = $ColorAccent
$BtnSearch.ForeColor = $ColorWhite
$BtnSearch.FlatStyle = "Flat"
$BtnSearch.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabSearch.Controls.Add($BtnSearch)

$LblDomFlt          = New-Object System.Windows.Forms.Label
$LblDomFlt.Text     = "Domain:"
$LblDomFlt.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=76})
$LblDomFlt.AutoSize = $true
$TabSearch.Controls.Add($LblDomFlt)

$CboDomain                  = New-Object System.Windows.Forms.ComboBox
$CboDomain.Size             = (New-Object System.Drawing.Size -Property @{Width=260; Height=26})
$CboDomain.Location         = (New-Object System.Drawing.Point -Property @{X=68; Y=73})
$CboDomain.DropDownStyle    = "DropDownList"
$CboDomain.Items.Add("All Domains") | Out-Null
foreach ($d in $AllDomains) { $CboDomain.Items.Add($d) | Out-Null }
$CboDomain.SelectedIndex    = 0
$TabSearch.Controls.Add($CboDomain)

$Grid                               = New-Object System.Windows.Forms.DataGridView
$Grid.Size                          = (New-Object System.Drawing.Size -Property @{Width=888; Height=430})
$Grid.Location                      = (New-Object System.Drawing.Point -Property @{X=10; Y=108})
$Grid.BackgroundColor               = $ColorWhite
$Grid.BorderStyle                   = "FixedSingle"
$Grid.RowHeadersVisible             = $false
$Grid.AllowUserToAddRows            = $false
$Grid.ReadOnly                      = $true
$Grid.SelectionMode                 = "FullRowSelect"
$Grid.AutoSizeColumnsMode           = "Fill"
$Grid.ColumnHeadersDefaultCellStyle.BackColor = $ColorDark
$Grid.ColumnHeadersDefaultCellStyle.ForeColor = $ColorWhite
$Grid.ColumnHeadersDefaultCellStyle.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$Grid.EnableHeadersVisualStyles     = $false
$TabSearch.Controls.Add($Grid)

$LblSrchStatus          = New-Object System.Windows.Forms.Label
$LblSrchStatus.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=545})
$LblSrchStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=880; Height=22})
$LblSrchStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Italic))
$TabSearch.Controls.Add($LblSrchStatus)

$BtnSearch.Add_Click({
    $term = $TxtSearch.Text.Trim()
    if (-not $term) { [System.Windows.Forms.MessageBox]::Show("Please enter a search term."); return }

    # Build a smart filter -- handle "First Last" name searches by also trying
    # GivenName and Surname independently so "Jordan Avery" finds the account
    $wildTerm = if ($term.Contains("*")) { $term } else { "*$term*" }

    # Check if the input looks like "FirstName LastName" (two words, no dot/at)
    $nameParts  = $term.Trim().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
    $firstPart  = if ($nameParts.Count -ge 1) { "*$($nameParts[0])*" } else { $null }
    $lastPart   = if ($nameParts.Count -ge 2) { "*$($nameParts[-1])*" } else { $null }

    $LblSrchStatus.Text      = "Searching..."
    $failedDomains = @()
    $LblSrchStatus.ForeColor = $ColorGold
    $Form.Refresh()

    $searchDomains = if ($CboDomain.SelectedItem -eq "All Domains") { $AllDomains } else { @($CboDomain.SelectedItem) }
    $results = @()
    foreach ($d in $searchDomains) {
        try {
            # Primary filter -- SAM, DisplayName, Email, UPN (string form for reliable variable expansion)
            # A single quote ends the -Filter string literal, so O'Neil is doubled to O''Neil.
            $qTerm = $wildTerm -replace "'", "''"
            $users = Get-ADUser -Filter "SamAccountName -like '$qTerm' -or DisplayName -like '$qTerm' -or EmailAddress -like '$qTerm' -or UserPrincipalName -like '$qTerm'" `
                -Server $d -Properties DisplayName,EmailAddress,UserPrincipalName,Title,Department,Office,Enabled -ErrorAction SilentlyContinue

            # Secondary filter -- GivenName + Surname for "First Last" searches
            # Catches accounts where DisplayName is not plain "First Last"
            if ($lastPart) {
                $nameUsers = Get-ADUser -Filter "GivenName -like '$($firstPart -replace "'", "''")' -and Surname -like '$($lastPart -replace "'", "''")'" `
                    -Server $d -Properties DisplayName,EmailAddress,UserPrincipalName,Title,Department,Office,Enabled -ErrorAction SilentlyContinue
                if ($nameUsers) {
                    # Merge without duplicates
                    $existingSAMs = @($users | ForEach-Object { $_.SamAccountName })
                    foreach ($nu in $nameUsers) {
                        if ($nu.SamAccountName -notin $existingSAMs) {
                            $users = @($users) + $nu
                        }
                    }
                }
            }

            foreach ($u in $users) {
                $results += [PSCustomObject]@{
                    "Display Name" = $u.DisplayName
                    "SAM"          = $u.SamAccountName
                    "UPN"          = $u.UserPrincipalName
                    "Email"        = $u.EmailAddress
                    "Title"        = $u.Title
                    "Department"   = $u.Department
                    "Enabled"      = $u.Enabled
                    "Domain"       = $d
                }
            }
        } catch {
            $failedDomains += $d
            DBG "SEARCH" "WARN" "Search failed in ${d}: $_"
        }
    }

    # Deduplicate across domains by SAM
    $results = $results | Sort-Object SAM -Unique

    $Grid.DataSource = $null
    if ($results.Count -gt 0) {
        $dt = New-Object System.Data.DataTable
        $results[0].PSObject.Properties.Name | ForEach-Object { $dt.Columns.Add($_) | Out-Null }
        foreach ($r in $results) {
            $row = $dt.NewRow()
            $r.PSObject.Properties | ForEach-Object { $row[$_.Name] = $_.Value }
            $dt.Rows.Add($row)
        }
        $Grid.DataSource         = $dt
        $LblSrchStatus.Text      = "$($results.Count) result(s) found."
        $LblSrchStatus.ForeColor = $ColorAccent
    } else {
        $LblSrchStatus.Text      = "No results found."
        $LblSrchStatus.ForeColor = $ColorRed
    }
    if ($failedDomains.Count) {
        $LblSrchStatus.Text     += "  Could not search: $($failedDomains -join ', ')"
        $LblSrchStatus.ForeColor = $ColorGold
    }
})
$TxtSearch.Add_KeyDown({ if ($_.KeyCode -eq "Return") { $BtnSearch.PerformClick() } })

# Click a search result row -> load UPN into Health Check and run it
$Grid.Add_CellDoubleClick({
    $row = $Grid.CurrentRow
    if (-not $row) { return }
    $upnVal = ""
    $colNames = @("UserPrincipalName","UPN","upn")
    foreach ($cn in $colNames) {
        if ($Grid.Columns[$cn]) { $upnVal = $row.Cells[$cn].Value; break }
    }
    if (-not $upnVal) {
        # Try first column that looks like a UPN
        foreach ($cell in $row.Cells) {
            if ($cell.Value -like "*@*") { $upnVal = $cell.Value; break }
        }
    }
    if ($upnVal) {
        $TxtHCUpn.Text = $upnVal
        $Tabs.SelectedTab = $TabHealth
        $BtnHCCheck.PerformClick()
    }
})

# ============================================================
#  TAB 2 - HEALTH CHECK
# ============================================================
$TabHealth           = New-Object System.Windows.Forms.TabPage
$TabHealth.Text      = "  Health Check  "
$TabHealth.BackColor = $ColorGray
$Tabs.Controls.Add($TabHealth)

# Search bar
$LblHC = New-Object System.Windows.Forms.Label
$LblHC.Text     = "UPN (first.last@contoso.com)"
$LblHC.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=15})
$LblHC.AutoSize = $true
$TabHealth.Controls.Add($LblHC)

$TxtHCUpn          = New-Object System.Windows.Forms.TextBox
$TxtHCUpn.Size     = (New-Object System.Drawing.Size -Property @{Width=540; Height=26})
$TxtHCUpn.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=35})
$TxtHCUpn.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]10),([System.Drawing.FontStyle]::Regular))
$TabHealth.Controls.Add($TxtHCUpn)

$BtnHCCheck           = New-Object System.Windows.Forms.Button
$BtnHCCheck.Text      = "Run Check"
$BtnHCCheck.Size      = (New-Object System.Drawing.Size -Property @{Width=110; Height=28})
$BtnHCCheck.Location  = (New-Object System.Drawing.Point -Property @{X=558; Y=34})
$BtnHCCheck.BackColor = $ColorAccent
$BtnHCCheck.ForeColor = $ColorWhite
$BtnHCCheck.FlatStyle = "Flat"
$BtnHCCheck.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabHealth.Controls.Add($BtnHCCheck)

$BtnHCCopy           = New-Object System.Windows.Forms.Button
$BtnHCCopy.Text      = "Copy Report"
$BtnHCCopy.Size      = (New-Object System.Drawing.Size -Property @{Width=110; Height=28})
$BtnHCCopy.Location  = (New-Object System.Drawing.Point -Property @{X=676; Y=34})
$BtnHCCopy.BackColor = $ColorLight
$BtnHCCopy.ForeColor = $ColorDark
$BtnHCCopy.FlatStyle = "Flat"
$BtnHCCopy.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnHCCopy.Enabled   = $false
$TabHealth.Controls.Add($BtnHCCopy)

# Results panel - left column (AD) and right column (Entra)
$GrpAD           = New-Object System.Windows.Forms.GroupBox
$GrpAD.Text      = "Active Directory"
$GrpAD.Size      = (New-Object System.Drawing.Size -Property @{Width=438; Height=510})
$GrpAD.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=72})
$GrpAD.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpAD.ForeColor = $ColorDark
$TabHealth.Controls.Add($GrpAD)

$RichAD               = New-Object System.Windows.Forms.RichTextBox
$RichAD.Size          = (New-Object System.Drawing.Size -Property @{Width=418; Height=480})
$RichAD.Location      = (New-Object System.Drawing.Point -Property @{X=10; Y=22})
$RichAD.BackColor     = [System.Drawing.Color]::Black
$RichAD.ForeColor     = [System.Drawing.Color]::LimeGreen
$RichAD.Font          = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$RichAD.ReadOnly      = $true
$RichAD.BorderStyle   = "None"
$GrpAD.Controls.Add($RichAD)

$GrpEntra           = New-Object System.Windows.Forms.GroupBox
$GrpEntra.Text      = "Entra ID / M365"
$GrpEntra.Size      = (New-Object System.Drawing.Size -Property @{Width=438; Height=510})
$GrpEntra.Location  = (New-Object System.Drawing.Point -Property @{X=458; Y=72})
$GrpEntra.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpEntra.ForeColor = $ColorDark
$TabHealth.Controls.Add($GrpEntra)

$RichEntra               = New-Object System.Windows.Forms.RichTextBox
$RichEntra.Size          = (New-Object System.Drawing.Size -Property @{Width=418; Height=480})
$RichEntra.Location      = (New-Object System.Drawing.Point -Property @{X=10; Y=22})
$RichEntra.BackColor     = [System.Drawing.Color]::Black
$RichEntra.ForeColor     = [System.Drawing.Color]::LimeGreen
$RichEntra.Font          = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$RichEntra.ReadOnly      = $true
$RichEntra.BorderStyle   = "None"
$GrpEntra.Controls.Add($RichEntra)

$LblHCStatus          = New-Object System.Windows.Forms.Label
$LblHCStatus.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=590})
$LblHCStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=880; Height=22})
$LblHCStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Italic))
$TabHealth.Controls.Add($LblHCStatus)

# Helper: append a traffic-light line to a RichTextBox
function Write-HC {
    param($Box, $Label, $Value, $Status)
    # Status: "ok"=green, "warn"=gold, "err"=red, "info"=cyan, "head"=white bold
    $color = switch ($Status) {
        "ok"   { [System.Drawing.Color]::LimeGreen }
        "warn" { [System.Drawing.Color]::Gold }
        "err"  { [System.Drawing.Color]::OrangeRed }
        "info" { [System.Drawing.Color]::Cyan }
        "head" { [System.Drawing.Color]::White }
        default { [System.Drawing.Color]::LimeGreen }
    }
    $line = if ($Label) { "$($Label.PadRight(22)): $Value" } else { $Value }
    $Box.SelectionStart  = $Box.TextLength
    $Box.SelectionLength = 0
    $Box.SelectionColor  = $color
    $Box.AppendText("$line`n")
    [System.Windows.Forms.Application]::DoEvents()
}

# Helper: get a Graph token using the same config the rest of the tool uses
function Get-HCGraphToken {
    $scriptDir = $PSScriptRoot
    if (-not $scriptDir) { $scriptDir = "C:\ADProvisioning" }
    return Get-GraphToken -ScriptDir $scriptDir
}

# ---- HEALTH CHECK CLICK ------------------------------------
$Script:HCReportLines = @()

$BtnHCCheck.Add_Click({
    $upn = $TxtHCUpn.Text.Trim()
    if (-not $upn) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a UPN.","Missing Input"); return
    }

    $RichAD.Clear(); $RichEntra.Clear()
    $Script:HCReportLines = @()
    $BtnHCCopy.Enabled    = $false
    $LblHCStatus.Text     = "Running health check..."
    $LblHCStatus.ForeColor = $ColorGold
    $Form.Refresh()

    # helper to accumulate report lines for the Copy button
    function HC-Line { param($lbl,$value,$status)
        Write-HC $RichAD $lbl $value $status
        $Script:HCReportLines += "$($lbl.PadRight(22)): $value"
    }
    function HE-Line { param($lbl,$value,$status)
        Write-HC $RichEntra $lbl $value $status
        $Script:HCReportLines += "$($lbl.PadRight(22)): $value"
    }

    # ---- AD SECTION ----------------------------------------
    Write-HC $RichAD "" "=== IDENTITY ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== AD: IDENTITY ==="

    # Resolve the account. We cannot reliably derive the AD domain from the UPN
    # (every division uses @contoso.com as primary or alias), so we search every
    # domain. IMPORTANT: use the STRING filter form -- a scriptblock filter
    # { UserPrincipalName -eq $upn } does NOT expand $upn reliably inside an
    # event handler, which is why the previous version found nothing.
    $adUser   = $null
    $adDomain = $null
    $adProps  = @("DisplayName","SamAccountName","UserPrincipalName","EmailAddress",
                  "Enabled","LockedOut","PasswordLastSet","PasswordNeverExpires",
                  "PasswordExpired","Title","Department","targetAddress","proxyAddresses",
                  "msExchRemoteRecipientType","msExchRecipientTypeDetails",
                  "msExchRecipientDisplayType","orgCategoryID","orgAccountType",
                  "Description","whenCreated","DistinguishedName")

    foreach ($dom in $AllDomains) {
        try {
            $found = Get-ADUser -Filter "UserPrincipalName -eq '$($upn -replace "'", "''")'" -Server $dom `
                        -Properties $adProps -ErrorAction SilentlyContinue
            if ($found) { $adUser = $found; $adDomain = $dom; break }
        } catch { DBG "HEALTH_CHECK" "WARN" "UPN lookup failed in ${dom}: $_" }
    }

    # Fallback: if UPN match failed, try matching the SAM (local part of the UPN)
    if (-not $adUser) {
        $samGuess = $upn.Split([char]'@')[0]
        foreach ($dom in $AllDomains) {
            try {
                $found = Get-ADUser -Filter "SamAccountName -eq '$($samGuess -replace "'", "''")'" -Server $dom `
                            -Properties $adProps -ErrorAction SilentlyContinue
                if ($found) { $adUser = $found; $adDomain = $dom; break }
            } catch { DBG "HEALTH_CHECK" "WARN" "SAM lookup failed in ${dom}: $_" }
        }
    }

    if (-not $adUser) {
        Write-HC $RichAD "" "[X] Account not found in any AD domain for UPN: $upn" "err"
        $LblHCStatus.Text = "[X] Not found in AD."; $LblHCStatus.ForeColor = $ColorRed; return
    }

    # Identity
    HC-Line "Display Name"   $adUser.DisplayName                                             "info"
    HC-Line "SAM"            $adUser.SamAccountName                                          "info"
    HC-Line "UPN"            $adUser.UserPrincipalName                                       "info"
    HC-Line "Email"          $(if ($adUser.EmailAddress) { $adUser.EmailAddress } else { "NOT SET" }) $(if ($adUser.EmailAddress) { "ok" } else { "warn" })
    HC-Line "Domain"         $adDomain                                                       "info"
    HC-Line "Created"        $adUser.whenCreated                                             "info"
    HC-Line "Title"          $(if ($adUser.Title) { $adUser.Title } else { "(none)" })       "info"

    Write-HC $RichAD "" "" "info"
    Write-HC $RichAD "" "=== ACCOUNT STATE ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== AD: ACCOUNT STATE ==="

    HC-Line "Enabled"        $adUser.Enabled        $(if ($adUser.Enabled) { "ok" } else { "err" })
    HC-Line "Locked Out"     $adUser.LockedOut       $(if ($adUser.LockedOut) { "err" } else { "ok" })
    HC-Line "Pwd Last Set"   $(if ($adUser.PasswordLastSet) { $adUser.PasswordLastSet.ToString("yyyy-MM-dd HH:mm") } else { "NEVER" }) $(if ($adUser.PasswordLastSet) { "ok" } else { "warn" })
    HC-Line "Pwd Never Exp"  $adUser.PasswordNeverExpires  $(if ($adUser.PasswordNeverExpires) { "warn" } else { "ok" })
    HC-Line "Pwd Expired"    $adUser.PasswordExpired       $(if ($adUser.PasswordExpired) { "err" } else { "ok" })

    # Calculate password expiry date from domain MaxPasswordAge policy
    $pwdExpiryStr = "N/A"
    $pwdExpiryStatus = "info"
    if ($adUser.PasswordNeverExpires) {
        $pwdExpiryStr = "Never expires"
        $pwdExpiryStatus = "warn"
    } elseif (-not $adUser.PasswordLastSet) {
        $pwdExpiryStr = "Unknown (never set)"
        $pwdExpiryStatus = "warn"
    } else {
        try {
            $maxAge = (Get-ADDefaultDomainPasswordPolicy -Server $adDomain -ErrorAction Stop).MaxPasswordAge
            if ($maxAge.TotalDays -eq 0) {
                $pwdExpiryStr = "Never expires (policy)"
                $pwdExpiryStatus = "warn"
            } else {
                $expiry    = $adUser.PasswordLastSet + $maxAge
                $daysLeft  = [math]::Round(($expiry - (Get-Date)).TotalDays)
                $expiryFmt = $expiry.ToString("yyyy-MM-dd")
                if ($daysLeft -lt 0) {
                    $pwdExpiryStr    = "$expiryFmt  (EXPIRED $([math]::Abs($daysLeft)) days ago)"
                    $pwdExpiryStatus = "err"
                } elseif ($daysLeft -le 14) {
                    $pwdExpiryStr    = "$expiryFmt  ($daysLeft days remaining)"
                    $pwdExpiryStatus = "warn"
                } else {
                    $pwdExpiryStr    = "$expiryFmt  ($daysLeft days remaining)"
                    $pwdExpiryStatus = "ok"
                }
            }
        } catch {
            $pwdExpiryStr    = "Could not calculate: $_"
            $pwdExpiryStatus = "warn"
        }
    }
    HC-Line "Pwd Expires"    $pwdExpiryStr   $pwdExpiryStatus

    Write-HC $RichAD "" "" "info"
    Write-HC $RichAD "" "=== CUSTOM ATTRIBUTES ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== AD: CUSTOM ATTRIBUTES ==="

    $catVal  = if ($adUser.orgCategoryID)  { $adUser.orgCategoryID }  else { "NOT SET" }
    $typeVal = if ($adUser.orgAccountType) { $adUser.orgAccountType } else { "NOT SET" }
    HC-Line "orgCategoryID"  $catVal  $(if ($adUser.orgCategoryID)  { "ok" } else { "err" })
    HC-Line "orgAccountType" $typeVal $(if ($adUser.orgAccountType) { "ok" } else { "err" })

    Write-HC $RichAD "" "" "info"
    Write-HC $RichAD "" "=== MAILBOX ATTRIBUTES ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== AD: MAILBOX ATTRIBUTES ==="

    # targetAddress
    $ta = $adUser.targetAddress
    if ($ta -like "SMTP:*@contoso.mail.onmicrosoft.com*") {
        HC-Line "targetAddress" $ta "ok"
    } elseif ($ta) {
        HC-Line "targetAddress" "$ta  [!] unexpected value" "warn"
    } else {
        HC-Line "targetAddress" "NOT SET -- mailbox not enabled" "err"
    }

    # msExch attributes with expected-value checks
    $exchType    = $adUser.msExchRemoteRecipientType
    $exchDetails = $adUser.msExchRecipientTypeDetails
    $exchDisplay = $adUser.msExchRecipientDisplayType

    HC-Line "msExchRemoteRecipientType"   $(if ($exchType)    { $exchType }    else { "NOT SET" }) $(if ($exchType    -eq "1")           { "ok" } elseif ($exchType)    { "warn" } else { "err" })
    HC-Line "msExchRecipientTypeDetails"  $(if ($exchDetails) { $exchDetails } else { "NOT SET" }) $(if ($exchDetails -eq "2147483648")  { "ok" } elseif ($exchDetails) { "warn" } else { "err" })
    HC-Line "msExchRecipientDisplayType"  $(if ($exchDisplay) { $exchDisplay } else { "NOT SET" }) $(if ($exchDisplay -eq "-2147483642") { "ok" } elseif ($exchDisplay) { "warn" } else { "err" })

    # proxyAddresses
    Write-HC $RichAD "" "" "info"
    if ($adUser.proxyAddresses -and $adUser.proxyAddresses.Count -gt 0) {
        Write-HC $RichAD "" "proxyAddresses:" "head"
        $Script:HCReportLines += "proxyAddresses:"
        foreach ($px in ($adUser.proxyAddresses | Sort-Object)) {
            $isPrimary = $px.StartsWith("SMTP:")
            $color     = if ($isPrimary) { "ok" } else { "info" }
            $tag       = if ($isPrimary) { "  [PRIMARY]" } else { "  [alias]" }
            Write-HC $RichAD "" "  $px$tag" $color
            $Script:HCReportLines += "  $px$tag"
        }
    } else {
        HC-Line "proxyAddresses" "NOT SET" "err"
    }

    # ---- OU / PATH
    Write-HC $RichAD "" "" "info"
    Write-HC $RichAD "" "=== LOCATION ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== AD: LOCATION ==="
    HC-Line "OU Path" $adUser.DistinguishedName "info"

    # ---- ENTRA SECTION -------------------------------------
    Write-HC $RichEntra "" "=== ENTRA IDENTITY ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== ENTRA: IDENTITY ==="

    # Get-GraphToken returns $null on failure (the reason is in the debug log); it does not throw.
    $hcHeaders = Get-HCGraphToken
    if (-not $hcHeaders) {
        Write-HC $RichEntra "" "[X] Could not acquire Graph token -- see GRAPH_TOKEN in the debug log." "err"
        $LblHCStatus.Text = "AD check complete. Entra check failed (token)."; $LblHCStatus.ForeColor = $ColorGold
        $BtnHCCopy.Enabled = $true; return
    }

    # Use the RESOLVED AD UPN for Graph -- the tech may have typed a SAM or partial
    # value that AD matched via fallback, but Graph needs the real UPN. If the on-prem
    # UPN somehow doesn't match Entra, fall back to filtering by onPremisesImmutableId.
    $graphUpn = $adUser.UserPrincipalName
    Write-HC $RichEntra "" "(looking up: $graphUpn)" "info"

    # Core user object
    $entraUser = $null
    try {
        $entraUser = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$graphUpn`?`$select=id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesImmutableId,jobTitle,department,assignedLicenses" `
            -Headers $hcHeaders -ErrorAction Stop
    } catch {
        $sc = 0; try { $sc = [int]$_.Exception.Response.StatusCode } catch {}
        if ($sc -eq 404) {
            # Fallback: the Entra UPN may differ from on-prem. Match on the sync anchor
            # (mS-DS-ConsistencyGuid -> onPremisesImmutableId, base64 of the GUID).
            try {
                $guidBytes = ([guid]$adUser.ObjectGUID).ToByteArray()
                $immutable = [Convert]::ToBase64String($guidBytes)
                $flt       = [Uri]::EscapeDataString("onPremisesImmutableId eq '$immutable'")
                $byAnchor  = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users?`$filter=$flt&`$select=id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesImmutableId" `
                    -Headers $hcHeaders -ErrorAction Stop
                if ($byAnchor.value -and $byAnchor.value.Count -gt 0) {
                    $entraUser = $byAnchor.value[0]
                    Write-HC $RichEntra "" "(matched by sync anchor; Entra UPN differs: $($entraUser.userPrincipalName))" "warn"
                }
            } catch {}
        }
        if (-not $entraUser) {
            if ($sc -eq 404) {
                Write-HC $RichEntra "" "[!] Not found in Entra -- not yet synced or sync issue." "err"
            } else {
                Write-HC $RichEntra "" "[X] Graph error ($sc): $_" "err"
            }
            $LblHCStatus.Text = "AD check complete. User not found in Entra."; $LblHCStatus.ForeColor = $ColorGold
            $BtnHCCopy.Enabled = $true; return
        }
    }

    HE-Line "Display Name"    $entraUser.displayName                                                          "info"
    HE-Line "UPN"             $entraUser.userPrincipalName                                                    "info"
    HE-Line "Mail"            $(if ($entraUser.mail) { $entraUser.mail } else { "NOT SET" })                  $(if ($entraUser.mail) { "ok" } else { "warn" })
    HE-Line "Object ID"       $entraUser.id                                                                   "info"

    Write-HC $RichEntra "" "" "info"
    Write-HC $RichEntra "" "=== ACCOUNT STATE ===" "head"
    $Script:HCReportLines += ""; $Script:HCReportLines += "=== ENTRA: ACCOUNT STATE ==="

    HE-Line "Enabled"         $entraUser.accountEnabled   $(if ($entraUser.accountEnabled) { "ok" } else { "err" })

    # Sync status
    $syncEnabled  = $entraUser.onPremisesSyncEnabled
    $lastSync     = $entraUser.onPremisesLastSyncDateTime
    $syncAge      = $null
    if ($lastSync) {
        $syncAge = [math]::Round(((Get-Date) - [datetime]$lastSync).TotalMinutes)
    }
    HE-Line "On-Prem Sync"    $(if ($syncEnabled) { "Yes" } else { "No (cloud-only or broken)" }) $(if ($syncEnabled) { "ok" } else { "warn" })
    HE-Line "Last Sync"       $(if ($lastSync) { "$lastSync  ($syncAge min ago)" } else { "Never" }) $(if ($syncAge -ne $null -and $syncAge -lt 120) { "ok" } elseif ($syncAge -ne $null) { "warn" } else { "err" })
    HE-Line "Immutable ID"    $(if ($entraUser.onPremisesImmutableId) { "Set" } else { "NOT SET" }) $(if ($entraUser.onPremisesImmutableId) { "ok" } else { "err" })

    # Authentication methods (requires UserAuthenticationMethod.Read.All)
    try {
        $authMethods = Invoke-RestMethod `
            -Uri "https://graph.microsoft.com/v1.0/users/$($entraUser.id)/authentication/methods" `
            -Headers $hcHeaders -ErrorAction Stop
        $methodNames = @()
        $mfaCapable  = $false
        foreach ($m in $authMethods.value) {
            $mType = $m.'@odata.type'
            switch -Wildcard ($mType) {
                "*microsoftAuthenticator*" { $methodNames += "Authenticator"; $mfaCapable = $true }
                "*phone*"                  { $methodNames += "Phone"; $mfaCapable = $true }
                "*email*"                  { $methodNames += "Email"; $mfaCapable = $true }
                "*fido2*"                  { $methodNames += "FIDO2 Key"; $mfaCapable = $true }
                "*windowsHello*"           { $methodNames += "Windows Hello"; $mfaCapable = $true }
                "*softwareOath*"           { $methodNames += "TOTP App"; $mfaCapable = $true }
                "*password*"               { $methodNames += "Password" }
                "*temporaryAccessPass*"    { $methodNames += "TAP" }
            }
        }
        $methodStr = if ($methodNames.Count -gt 0) { ($methodNames | Sort-Object -Unique) -join ", " } else { "None" }
        HE-Line "Auth Methods"    $methodStr "info"
        HE-Line "MFA Capable"     $(if ($mfaCapable) { "Yes" } else { "No -- only password" }) $(if ($mfaCapable) { "ok" } else { "warn" })

        # Microsoft Authenticator details (device name + registration date)
        try {
            $authApp = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/v1.0/users/$($entraUser.id)/authentication/microsoftAuthenticatorMethods" `
                -Headers $hcHeaders -ErrorAction Stop
            if ($authApp.value -and $authApp.value.Count -gt 0) {
                foreach ($app in $authApp.value) {
                    $appDevice = if ($app.displayName) { $app.displayName } else { "Unknown device" }
                    $appDate   = "N/A"
                    if ($app.createdDateTime -and $app.createdDateTime -ne "") {
                        try { $appDate = ([datetime]$app.createdDateTime).ToString("yyyy-MM-dd") } catch { $appDate = "N/A" }
                    }
                    $appVer = if ($app.phoneAppVersion) { " v$($app.phoneAppVersion)" } else { "" }
                    HE-Line "Authenticator"   "$appDevice$appVer (registered $appDate)" "ok"
                }
            } else {
                HE-Line "Authenticator"   "Not registered" "warn"
            }
        } catch {
            HE-Line "Authenticator"   "Could not retrieve details" "warn"
        }

        # Windows Hello for Business -- most recent only (sorted by createdDateTime desc)
        try {
            $whfb = Invoke-RestMethod `
                -Uri "https://graph.microsoft.com/v1.0/users/$($entraUser.id)/authentication/windowsHelloForBusinessMethods" `
                -Headers $hcHeaders -ErrorAction Stop
            if ($whfb.value -and $whfb.value.Count -gt 0) {
                # Sort by createdDateTime descending, show only the most recent
                $sorted = $whfb.value | Sort-Object { [datetime]$_.createdDateTime } -Descending
                $latest = $sorted[0]
                $whfbDevice = if ($latest.displayName) { $latest.displayName } else { "Unknown device" }
                $whfbDate   = if ($latest.createdDateTime) { ([datetime]$latest.createdDateTime).ToString("yyyy-MM-dd") } else { "Unknown" }
                $whfbKey    = if ($latest.keyStrength) { $latest.keyStrength } else { "" }
                $whfbTotal  = $whfb.value.Count
                $whfbInfo   = "$whfbDevice (registered $whfbDate$(if($whfbKey){', '+$whfbKey}else{''}))"
                if ($whfbTotal -gt 1) { $whfbInfo += " -- $whfbTotal total registrations" }
                HE-Line "WHfB (latest)"  $whfbInfo "ok"
            } else {
                HE-Line "WHfB"           "Not registered" "info"
            }
        } catch {
            HE-Line "WHfB"           "Could not retrieve details" "warn"
        }
    } catch {
        HE-Line "Auth Methods"    "Could not retrieve" "warn"
    }

    $LblHCStatus.Text      = "Health check complete for $upn"
    $LblHCStatus.ForeColor = $ColorAccent
    $BtnHCCopy.Enabled     = $true
})

# Copy report to clipboard
$BtnHCCopy.Add_Click({
    [System.Windows.Forms.Clipboard]::SetText(($Script:HCReportLines -join "`r`n"))
    $LblHCStatus.Text      = "Report copied to clipboard."
    $LblHCStatus.ForeColor = $ColorAccent
})

$TxtHCUpn.Add_KeyDown({ if ($_.KeyCode -eq "Return") { $BtnHCCheck.PerformClick() } })

# ============================================================
#  TAB 3 - AUDIT LOG
# ============================================================
$TabAudit           = New-Object System.Windows.Forms.TabPage
$TabAudit.Text      = "  Audit Log  "
$TabAudit.BackColor = $ColorGray
$Tabs.Controls.Add($TabAudit)

# Toolbar
$BtnAuditRefresh           = New-Object System.Windows.Forms.Button
$BtnAuditRefresh.Text      = "Refresh"
$BtnAuditRefresh.Size      = (New-Object System.Drawing.Size -Property @{Width=100; Height=28})
$BtnAuditRefresh.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=10})
$BtnAuditRefresh.BackColor = $ColorAccent
$BtnAuditRefresh.ForeColor = $ColorWhite
$BtnAuditRefresh.FlatStyle = "Flat"
$BtnAuditRefresh.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabAudit.Controls.Add($BtnAuditRefresh)

$BtnAuditReport           = New-Object System.Windows.Forms.Button
$BtnAuditReport.Text      = "Generate HTML Report"
$BtnAuditReport.Size      = (New-Object System.Drawing.Size -Property @{Width=160; Height=28})
$BtnAuditReport.Location  = (New-Object System.Drawing.Point -Property @{X=118; Y=10})
$BtnAuditReport.BackColor = $ColorLight
$BtnAuditReport.ForeColor = $ColorDark
$BtnAuditReport.FlatStyle = "Flat"
$BtnAuditReport.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabAudit.Controls.Add($BtnAuditReport)

$LblAuditTime          = New-Object System.Windows.Forms.Label
$LblAuditTime.Text     = "Time window:"
$LblAuditTime.Location = (New-Object System.Drawing.Point -Property @{X=290; Y=15})
$LblAuditTime.AutoSize = $true
$TabAudit.Controls.Add($LblAuditTime)

$CboAuditTime                = New-Object System.Windows.Forms.ComboBox
$CboAuditTime.Size           = (New-Object System.Drawing.Size -Property @{Width=120; Height=24})
$CboAuditTime.Location       = (New-Object System.Drawing.Point -Property @{X=380; Y=12})
$CboAuditTime.DropDownStyle  = "DropDownList"
@("Last 7 Days","Last 30 Days","Last 90 Days","All Time") | ForEach-Object { $CboAuditTime.Items.Add($_) | Out-Null }
$CboAuditTime.SelectedIndex  = 1   # default: Last 30 Days
$TabAudit.Controls.Add($CboAuditTime)

$LblAuditFilter          = New-Object System.Windows.Forms.Label
$LblAuditFilter.Text     = "Operator:"
$LblAuditFilter.Location = (New-Object System.Drawing.Point -Property @{X=512; Y=15})
$LblAuditFilter.AutoSize = $true
$TabAudit.Controls.Add($LblAuditFilter)

$CboAuditOperator                = New-Object System.Windows.Forms.ComboBox
$CboAuditOperator.Size           = (New-Object System.Drawing.Size -Property @{Width=170; Height=24})
$CboAuditOperator.Location       = (New-Object System.Drawing.Point -Property @{X=588; Y=12})
$CboAuditOperator.DropDownStyle  = "DropDownList"
$TabAudit.Controls.Add($CboAuditOperator)

$LblAuditFilter2          = New-Object System.Windows.Forms.Label
$LblAuditFilter2.Text     = "Result:"
$LblAuditFilter2.Location = (New-Object System.Drawing.Point -Property @{X=768; Y=15})
$LblAuditFilter2.AutoSize = $true
$TabAudit.Controls.Add($LblAuditFilter2)

$CboAuditResult                = New-Object System.Windows.Forms.ComboBox
$CboAuditResult.Size           = (New-Object System.Drawing.Size -Property @{Width=130; Height=24})
$CboAuditResult.Location       = (New-Object System.Drawing.Point -Property @{X=810; Y=12})
$CboAuditResult.DropDownStyle  = "DropDownList"
$TabAudit.Controls.Add($CboAuditResult)

$LblAuditStatus          = New-Object System.Windows.Forms.Label
$LblAuditStatus.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=560})
$LblAuditStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=880; Height=22})
$LblAuditStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Italic))
$TabAudit.Controls.Add($LblAuditStatus)

# DataGridView for audit rows
$AuditGrid                        = New-Object System.Windows.Forms.DataGridView
$AuditGrid.Size                   = (New-Object System.Drawing.Size -Property @{Width=880; Height=520})
$AuditGrid.Location               = (New-Object System.Drawing.Point -Property @{X=10; Y=48})
$AuditGrid.BackgroundColor        = $ColorRGB_30_30_30
$AuditGrid.ForeColor              = [System.Drawing.Color]::LimeGreen
$AuditGrid.GridColor              = $ColorRGB_60_60_60
$AuditGrid.DefaultCellStyle.BackColor    = $ColorRGB_20_20_20
$AuditGrid.DefaultCellStyle.ForeColor    = [System.Drawing.Color]::LimeGreen
$AuditGrid.DefaultCellStyle.SelectionBackColor = $ColorAccent
$AuditGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$AuditGrid.ColumnHeadersDefaultCellStyle.BackColor = $ColorDark
$AuditGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$AuditGrid.ColumnHeadersDefaultCellStyle.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Bold))
$AuditGrid.RowHeadersVisible      = $false
$AuditGrid.AllowUserToAddRows     = $false
$AuditGrid.AllowUserToDeleteRows  = $false
$AuditGrid.ReadOnly               = $true
$AuditGrid.SelectionMode          = "FullRowSelect"
$AuditGrid.AutoSizeColumnsMode    = "Fill"
$AuditGrid.ColumnHeadersHeightSizeMode = "DisableResizing"
$AuditGrid.ColumnHeadersHeight    = 28
$AuditGrid.Font                   = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$AuditGrid.BorderStyle            = "None"
$TabAudit.Controls.Add($AuditGrid)

# ---- AUDIT LOG FUNCTIONS -----------------------------------

function Read-AuditCSV {
    # Reads the CSV safely using FileShare.ReadWrite so it never blocks writers.
    # Returns an array of PSCustomObjects, newest first.
    if (-not (Test-Path $AuditLogPath)) { return @() }
    try {
        $fs     = [System.IO.File]::Open($AuditLogPath,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.StreamReader -ArgumentList $fs
        $raw    = $reader.ReadToEnd()
        $reader.Close(); $fs.Dispose()
        $rows   = $raw | ConvertFrom-Csv
        return ($rows | Sort-Object Timestamp -Descending)
    } catch {
        return @()
    }
}

function Load-AuditGrid {
    param($OpFilter = "All", $ResFilter = "All", $TimeWindow = "Last 30 Days")
    $AuditGrid.Rows.Clear()
    $AuditGrid.Columns.Clear()

    $rows = Read-AuditCSV
    if (-not $rows -or $rows.Count -eq 0) {
        $LblAuditStatus.Text = "No audit records found."; return
    }

    # Apply time window
    $cutoff = switch ($TimeWindow) {
        "Last 7 Days"  { (Get-Date).AddDays(-7) }
        "Last 30 Days" { (Get-Date).AddDays(-30) }
        "Last 90 Days" { (Get-Date).AddDays(-90) }
        default        { [datetime]::MinValue }
    }
    if ($cutoff -gt [datetime]::MinValue) {
        $rows = $rows | Where-Object {
            try { [datetime]$_.Timestamp -ge $cutoff } catch { $false }
        }
    }

    # Populate operator/result dropdowns on first load
    if ($CboAuditOperator.Items.Count -le 1) {
        $allRows = Read-AuditCSV
        $CboAuditOperator.Items.Clear()
        $CboAuditOperator.Items.Add("All") | Out-Null
        $allRows | Select-Object -ExpandProperty Operator -Unique | Sort-Object | ForEach-Object {
            $CboAuditOperator.Items.Add($_) | Out-Null
        }
        $CboAuditOperator.SelectedIndex = 0
    }
    if ($CboAuditResult.Items.Count -le 1) {
        $allRows = Read-AuditCSV
        $CboAuditResult.Items.Clear()
        $CboAuditResult.Items.Add("All") | Out-Null
        $allRows | Select-Object -ExpandProperty Result -Unique | Sort-Object | ForEach-Object {
            $CboAuditResult.Items.Add($_) | Out-Null
        }
        $CboAuditResult.SelectedIndex = 0
    }

    # Apply filters
    if ($OpFilter  -ne "All") { $rows = $rows | Where-Object { $_.Operator -eq $OpFilter } }
    if ($ResFilter -ne "All") { $rows = $rows | Where-Object { $_.Result   -eq $ResFilter } }

    if (-not $rows -or $rows.Count -eq 0) {
        $LblAuditStatus.Text = "No records match the current filter."; return
    }

    # Build columns
    $cols = @("Timestamp","Operator","Machine","Action","TargetSAM","TargetUPN","Division","EmpType","Result","Details")
    foreach ($c in $cols) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c; $col.HeaderText = $c
        $AuditGrid.Columns.Add($col) | Out-Null
    }
    $AuditGrid.Columns["Details"].FillWeight = 250
    $AuditGrid.Columns["TargetUPN"].FillWeight = 160
    $AuditGrid.Columns["Timestamp"].FillWeight = 120

    # Populate rows with color coding
    foreach ($r in $rows) {
        $idx = $AuditGrid.Rows.Add(
            $r.Timestamp, $r.Operator, $r.Machine, $r.Action,
            $r.TargetSAM, $r.TargetUPN, $r.Division, $r.EmpType,
            $r.Result, $r.Details
        )
        $row = $AuditGrid.Rows[$idx]
        $res = $r.Result
        if ($res -like "*FAILED*" -or $res -like "*ERROR*") {
            $row.DefaultCellStyle.BackColor = $ColorRGB_80_20_20
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::OrangeRed
        } elseif ($r.Action -eq "PARTIAL_CREATE") {
            $row.DefaultCellStyle.BackColor = $ColorRGB_70_45_5
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Orange
        } elseif ($res -like "*WARN*" -or $res -like "*COLLISION*" -or $res -like "*UNCONFIRMED*") {
            $row.DefaultCellStyle.BackColor = $ColorRGB_60_50_10
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gold
        } elseif ($res -eq "OPENED") {
            $row.DefaultCellStyle.BackColor = $ColorRGB_20_20_50
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::CornflowerBlue
        }
    }
    $LblAuditStatus.Text = "$($rows.Count) record(s) shown  |  Log: $AuditLogPath"
}

function Generate-AuditReport {
    # Generates a self-contained HTML report from the CSV and opens it.
    $rows = Read-AuditCSV
    $generated = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $total     = $rows.Count
    $creates   = ($rows | Where-Object { $_.Action -eq "CREATE" -and $_.Result -eq "AD_OK" }).Count
    $failures  = ($rows | Where-Object { $_.Result -like "*FAILED*" -or $_.Result -like "*ERROR*" }).Count
    $operators = ($rows | Select-Object -ExpandProperty Operator -Unique | Sort-Object) -join ", "

    $rowsHtml = foreach ($r in $rows) {
        $res = $r.Result
        $cls = if     ($res -like "*FAILED*" -or $res -like "*ERROR*") { "fail" }
               elseif ($res -like "*WARN*" -or $res -like "*COLLISION*" -or $res -like "*UNCONFIRMED*") { "warn" }
               elseif ($res -eq "OPENED") { "info" }
               else   { "ok" }
        "<tr class='$cls'><td>$($r.Timestamp)</td><td>$($r.Operator)</td><td>$($r.Machine)</td><td>$($r.Action)</td><td>$($r.TargetSAM)</td><td>$($r.TargetUPN)</td><td>$($r.Division)</td><td>$($r.EmpType)</td><td>$($r.Result)</td><td>$($r.Details)</td></tr>"
    }

    $html = @"
<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>AD Provisioning Tool - Audit Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; background:#1a1a1a; color:#ccc; margin:0; padding:20px; }
  h1 { color:#0B5CAD; font-size:22px; margin-bottom:4px; }
  .subtitle { color:#888; font-size:12px; margin-bottom:20px; }
  .summary { display:flex; gap:20px; margin-bottom:20px; flex-wrap:wrap; }
  .card { background:#2a2a2a; border-left:4px solid #0B5CAD; padding:12px 18px; border-radius:4px; min-width:140px; }
  .card.fail { border-color:#e53935; }
  .card.warn { border-color:#fbc02d; }
  .card.info { border-color:#1e88e5; }
  .card-value { font-size:28px; font-weight:bold; color:#fff; }
  .card-lbl { font-size:11px; color:#888; margin-top:2px; }
  .filters { margin-bottom:12px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
  .filters label { color:#888; font-size:12px; }
  .filters select, .filters input { background:#2a2a2a; color:#ccc; border:1px solid #444; padding:4px 8px; border-radius:3px; font-size:12px; }
  table { width:100%; border-collapse:collapse; font-size:12px; }
  thead th { background:#1F2A44; color:#fff; padding:8px 10px; text-align:left; position:sticky; top:0; cursor:pointer; user-select:none; }
  thead th:hover { background:#2C3E66; }
  tbody tr { border-bottom:1px solid #2a2a2a; }
  tbody tr:hover { background:#252525; }
  tr.ok   td { color:#A9C7E8; }
  tr.fail td { color:#ef9a9a; background:#2a1515; }
  tr.warn td { color:#ffe082; background:#2a2510; }
  tr.info td { color:#90caf9; background:#131a2a; }
  td { padding:6px 10px; vertical-align:top; word-break:break-word; }
  td:last-child { max-width:300px; }
  .badge { display:inline-block; padding:2px 6px; border-radius:3px; font-size:10px; font-weight:bold; }
  .badge-ok   { background:#1F2A44; color:#A9C7E8; }
  .badge-fail { background:#b71c1c; color:#ef9a9a; }
  .badge-warn { background:#f57f17; color:#ffe082; }
  .badge-info { background:#0d47a1; color:#90caf9; }
  .op-filter { margin-right:4px; }
</style>
</head><body>
<h1>AD Provisioning Tool - Audit Report</h1>
<div class="subtitle">Generated: $generated &nbsp;|&nbsp; Log: $AuditLogPath &nbsp;|&nbsp; Operators: $operators</div>
<div class="summary">
  <div class="card"><div class="card-value">$total</div><div class="card-lbl">Total Records</div></div>
  <div class="card"><div class="card-value">$creates</div><div class="card-lbl">Accounts Created</div></div>
  <div class="card fail"><div class="card-value">$failures</div><div class="card-lbl">Failures</div></div>
</div>
<div class="filters">
  <label>Operator:</label>
  <select id="opFilter" onchange="applyFilters()"><option>All</option></select>
  <label>Action:</label>
  <select id="actFilter" onchange="applyFilters()"><option>All</option></select>
  <label>Result:</label>
  <select id="resFilter" onchange="applyFilters()"><option>All</option></select>
  <label>Search:</label>
  <input type="text" id="searchBox" oninput="applyFilters()" placeholder="SAM, UPN, details...">
</div>
<table id="auditTable">
<thead><tr>
  <th onclick="sortTable(0)">Timestamp &#9651;</th>
  <th onclick="sortTable(1)">Operator</th>
  <th onclick="sortTable(2)">Machine</th>
  <th onclick="sortTable(3)">Action</th>
  <th onclick="sortTable(4)">SAM</th>
  <th onclick="sortTable(5)">UPN</th>
  <th onclick="sortTable(6)">Division</th>
  <th onclick="sortTable(7)">EmpType</th>
  <th onclick="sortTable(8)">Result</th>
  <th onclick="sortTable(9)">Details</th>
</tr></thead>
<tbody id="tableBody">
$($rowsHtml -join "`n")
</tbody></table>
<script>
// Populate filter dropdowns
(function(){
  var rows=document.querySelectorAll('#tableBody tr');
  var ops={},acts={},ress={};
  rows.forEach(function(r){
    var c=r.cells; if(!c||c.length<9) return;
    ops[c[1].textContent]=1; acts[c[3].textContent]=1; ress[c[8].textContent]=1;
  });
  function fill(id,obj){ var s=document.getElementById(id); Object.keys(obj).sort().forEach(function(v){ var o=document.createElement('option'); o.value=v; o.text=v; s.add(o); }); }
  fill('opFilter',ops); fill('actFilter',acts); fill('resFilter',ress);
})();

function applyFilters(){
  var op=document.getElementById('opFilter').value;
  var act=document.getElementById('actFilter').value;
  var res=document.getElementById('resFilter').value;
  var srch=document.getElementById('searchBox').value.toLowerCase();
  document.querySelectorAll('#tableBody tr').forEach(function(r){
    var c=r.cells; if(!c||c.length<9){r.style.display='';return;}
    var show=(op==='All'||c[1].textContent===op)&&
             (act==='All'||c[3].textContent===act)&&
             (res==='All'||c[8].textContent===res)&&
             (!srch||r.textContent.toLowerCase().includes(srch));
    r.style.display=show?'':'none';
  });
}

var sortDir=1;
function sortTable(col){
  var tb=document.getElementById('tableBody');
  var rows=Array.from(tb.querySelectorAll('tr'));
  rows.sort(function(a,b){
    var av=a.cells[col]?a.cells[col].textContent:'';
    var bv=b.cells[col]?b.cells[col].textContent:'';
    return av.localeCompare(bv)*sortDir;
  });
  sortDir*=-1;
  rows.forEach(function(r){tb.appendChild(r);});
}
</script>
</body></html>
"@

    $reportPath = Join-Path (Split-Path -Parent $AuditLogPath) "ADProvisioningTool-Report.html"
    [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
    Start-Process $reportPath
    return $reportPath
}

# ---- AUDIT TAB EVENTS --------------------------------------
$BtnAuditRefresh.Add_Click({
    $op  = if ($CboAuditOperator.SelectedItem) { $CboAuditOperator.SelectedItem.ToString() } else { "All" }
    $res = if ($CboAuditResult.SelectedItem)   { $CboAuditResult.SelectedItem.ToString()   } else { "All" }
    $tw  = if ($CboAuditTime.SelectedItem)     { $CboAuditTime.SelectedItem.ToString()     } else { "Last 30 Days" }
    Load-AuditGrid -OpFilter $op -ResFilter $res -TimeWindow $tw
})

$BtnAuditReport.Add_Click({
    $LblAuditStatus.Text = "Generating report..."; $Form.Refresh()
    $path = Generate-AuditReport
    $LblAuditStatus.Text = "Report generated: $path"
})

$CboAuditTime.Add_SelectedIndexChanged({
    $op  = if ($CboAuditOperator.SelectedItem) { $CboAuditOperator.SelectedItem.ToString() } else { "All" }
    $res = if ($CboAuditResult.SelectedItem)   { $CboAuditResult.SelectedItem.ToString()   } else { "All" }
    $tw  = $CboAuditTime.SelectedItem.ToString()
    Load-AuditGrid -OpFilter $op -ResFilter $res -TimeWindow $tw
})

$CboAuditOperator.Add_SelectedIndexChanged({
    $op  = $CboAuditOperator.SelectedItem.ToString()
    $res = if ($CboAuditResult.SelectedItem) { $CboAuditResult.SelectedItem.ToString() } else { "All" }
    $tw  = if ($CboAuditTime.SelectedItem)   { $CboAuditTime.SelectedItem.ToString()   } else { "Last 30 Days" }
    Load-AuditGrid -OpFilter $op -ResFilter $res -TimeWindow $tw
})

$CboAuditResult.Add_SelectedIndexChanged({
    $op  = if ($CboAuditOperator.SelectedItem) { $CboAuditOperator.SelectedItem.ToString() } else { "All" }
    $res = $CboAuditResult.SelectedItem.ToString()
    $tw  = if ($CboAuditTime.SelectedItem)     { $CboAuditTime.SelectedItem.ToString()     } else { "Last 30 Days" }
    Load-AuditGrid -OpFilter $op -ResFilter $res -TimeWindow $tw
})

# Load on tab switch
$Tabs.Add_SelectedIndexChanged({
    if ($Tabs.SelectedTab -eq $TabAudit) {
        $tw = if ($CboAuditTime.SelectedItem) { $CboAuditTime.SelectedItem.ToString() } else { "Last 30 Days" }
        Load-AuditGrid -TimeWindow $tw
    }
    if ($Tabs.SelectedTab -eq $TabDebug) {
        Load-DebugGrid
    }
})

# ============================================================
#  TAB 4 - DEBUG LOG
# ============================================================
$TabDebug           = New-Object System.Windows.Forms.TabPage
$TabDebug.Text      = "  Debug Log  "
$TabDebug.BackColor = $ColorGray
$Tabs.Controls.Add($TabDebug)

# Toolbar
$BtnDebugRefresh           = New-Object System.Windows.Forms.Button
$BtnDebugRefresh.Text      = "Refresh"
$BtnDebugRefresh.Size      = (New-Object System.Drawing.Size -Property @{Width=100; Height=28})
$BtnDebugRefresh.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=10})
$BtnDebugRefresh.BackColor = $ColorAccent
$BtnDebugRefresh.ForeColor = $ColorWhite
$BtnDebugRefresh.FlatStyle = "Flat"
$BtnDebugRefresh.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabDebug.Controls.Add($BtnDebugRefresh)

$LblDebugTime          = New-Object System.Windows.Forms.Label
$LblDebugTime.Text     = "Time window:"
$LblDebugTime.Location = (New-Object System.Drawing.Point -Property @{X=120; Y=15})
$LblDebugTime.AutoSize = $true
$TabDebug.Controls.Add($LblDebugTime)

$CboDebugTime                = New-Object System.Windows.Forms.ComboBox
$CboDebugTime.Size           = (New-Object System.Drawing.Size -Property @{Width=120; Height=24})
$CboDebugTime.Location       = (New-Object System.Drawing.Point -Property @{X=210; Y=12})
$CboDebugTime.DropDownStyle  = "DropDownList"
@("Last 7 Days","Last 30 Days","Last 90 Days","All Time") | ForEach-Object { $CboDebugTime.Items.Add($_) | Out-Null }
$CboDebugTime.SelectedIndex  = 1
$TabDebug.Controls.Add($CboDebugTime)

$LblDebugStatus          = New-Object System.Windows.Forms.Label
$LblDebugStatus.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=590})
$LblDebugStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=900; Height=22})
$LblDebugStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Italic))
$TabDebug.Controls.Add($LblDebugStatus)

# Top grid -- one row per creation session
$DebugGrid                        = New-Object System.Windows.Forms.DataGridView
$DebugGrid.Size                   = (New-Object System.Drawing.Size -Property @{Width=910; Height=270})
$DebugGrid.Location               = (New-Object System.Drawing.Point -Property @{X=10; Y=48})
$DebugGrid.BackgroundColor        = $ColorRGB_20_20_20
$DebugGrid.ForeColor              = [System.Drawing.Color]::LimeGreen
$DebugGrid.GridColor              = $ColorRGB_50_50_50
$DebugGrid.DefaultCellStyle.BackColor    = $ColorRGB_15_15_15
$DebugGrid.DefaultCellStyle.ForeColor    = [System.Drawing.Color]::LimeGreen
$DebugGrid.DefaultCellStyle.SelectionBackColor = $ColorAccent
$DebugGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$DebugGrid.ColumnHeadersDefaultCellStyle.BackColor = $ColorDark
$DebugGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$DebugGrid.ColumnHeadersDefaultCellStyle.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Bold))
$DebugGrid.RowHeadersVisible      = $false
$DebugGrid.AllowUserToAddRows     = $false
$DebugGrid.AllowUserToDeleteRows  = $false
$DebugGrid.ReadOnly               = $true
$DebugGrid.SelectionMode          = "FullRowSelect"
$DebugGrid.AutoSizeColumnsMode    = "Fill"
$DebugGrid.ColumnHeadersHeightSizeMode = "DisableResizing"
$DebugGrid.ColumnHeadersHeight    = 26
$DebugGrid.Font                   = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$DebugGrid.BorderStyle            = "None"
$TabDebug.Controls.Add($DebugGrid)

# Detail panel -- shows full trace for selected session
$GrpDebugDetail           = New-Object System.Windows.Forms.GroupBox
$GrpDebugDetail.Text      = "Debug Trace  (double-click a row above to load)"
$GrpDebugDetail.Size      = (New-Object System.Drawing.Size -Property @{Width=910; Height=290})
$GrpDebugDetail.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=326})
$GrpDebugDetail.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpDebugDetail.ForeColor = $ColorDark
$TabDebug.Controls.Add($GrpDebugDetail)

$BtnDebugCopy           = New-Object System.Windows.Forms.Button
$BtnDebugCopy.Text      = "Copy Trace"
$BtnDebugCopy.Size      = (New-Object System.Drawing.Size -Property @{Width=100; Height=24})
$BtnDebugCopy.Location  = (New-Object System.Drawing.Point -Property @{X=800; Y=18})
$BtnDebugCopy.BackColor = $ColorLight
$BtnDebugCopy.ForeColor = $ColorDark
$BtnDebugCopy.FlatStyle = "Flat"
$BtnDebugCopy.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Bold))
$BtnDebugCopy.Enabled   = $false
$GrpDebugDetail.Controls.Add($BtnDebugCopy)

$RichDebugDetail               = New-Object System.Windows.Forms.RichTextBox
$RichDebugDetail.Size          = (New-Object System.Drawing.Size -Property @{Width=890; Height=248})
$RichDebugDetail.Location      = (New-Object System.Drawing.Point -Property @{X=10; Y=30})
$RichDebugDetail.BackColor     = [System.Drawing.Color]::Black
$RichDebugDetail.ForeColor     = [System.Drawing.Color]::Cyan
$RichDebugDetail.Font          = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$RichDebugDetail.ReadOnly      = $true
$RichDebugDetail.BorderStyle   = "None"
$GrpDebugDetail.Controls.Add($RichDebugDetail)

$Script:DebugTraceText = ""

# ---- DEBUG LOG FUNCTIONS -----------------------------------
function Read-DebugCSV {
    param($TimeWindow = "Last 30 Days")
    if (-not (Test-Path $DebugLogPath)) { return @() }
    try {
        $fs     = [System.IO.File]::Open($DebugLogPath,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite)
        $reader = New-Object System.IO.StreamReader -ArgumentList $fs
        $raw    = $reader.ReadToEnd()
        $reader.Close(); $fs.Dispose()
        $rows   = $raw | ConvertFrom-Csv
        $cutoff = switch ($TimeWindow) {
            "Last 7 Days"  { (Get-Date).AddDays(-7) }
            "Last 30 Days" { (Get-Date).AddDays(-30) }
            "Last 90 Days" { (Get-Date).AddDays(-90) }
            default        { [datetime]::MinValue }
        }
        if ($cutoff -gt [datetime]::MinValue) {
            $rows = $rows | Where-Object { try { [datetime]$_.Timestamp -ge $cutoff } catch { $false } }
        }
        return $rows
    } catch { return @() }
}

function Load-DebugGrid {
    param($TimeWindow = "Last 30 Days")
    $DebugGrid.Rows.Clear()
    $DebugGrid.Columns.Clear()
    $RichDebugDetail.Clear()
    $BtnDebugCopy.Enabled = $false

    $rows = Read-DebugCSV -TimeWindow $TimeWindow
    if (-not $rows -or $rows.Count -eq 0) {
        $LblDebugStatus.Text = "No debug records found."; return
    }

    # Group by SAM + Timestamp-session (first timestamp for each SAM group)
    $sessions = $rows | Group-Object SAM | ForEach-Object {
        $grp     = $_.Group | Sort-Object Timestamp
        $first   = $grp[0]
        $hasErr  = ($grp | Where-Object { $_.Status -like "*FAIL*" -or $_.Status -like "*ERROR*" }).Count -gt 0
        $hasWarn = ($grp | Where-Object { $_.Status -like "*WARN*" -or $_.Status -like "*not set*" }).Count -gt 0
        $overall = if ($hasErr) { "FAILED" } elseif ($hasWarn) { "WARNING" } else { "OK" }
        [PSCustomObject]@{
            Timestamp   = $first.Timestamp
            SAM         = $first.SAM
            DisplayName = $first.DisplayName
            Operator    = $first.Operator
            Steps       = $grp.Count
            Result      = $overall
        }
    } | Sort-Object Timestamp -Descending

    # Build columns
    foreach ($c in @("Timestamp","SAM","DisplayName","Operator","Steps","Result")) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c; $col.HeaderText = $c
        $DebugGrid.Columns.Add($col) | Out-Null
    }

    foreach ($s in $sessions) {
        $idx = $DebugGrid.Rows.Add($s.Timestamp,$s.SAM,$s.DisplayName,$s.Operator,$s.Steps,$s.Result)
        $row = $DebugGrid.Rows[$idx]
        if     ($s.Result -eq "FAILED")  { $row.DefaultCellStyle.BackColor=$ColorRGB_70_15_15; $row.DefaultCellStyle.ForeColor=[System.Drawing.Color]::OrangeRed }
        elseif ($s.Result -eq "WARNING") { $row.DefaultCellStyle.BackColor=$ColorRGB_55_45_5;  $row.DefaultCellStyle.ForeColor=[System.Drawing.Color]::Gold }
    }
    $LblDebugStatus.Text = "$($sessions.Count) creation session(s) shown. Double-click a row to view trace."
}

function Load-DebugDetail {
    param($Sam, $Timestamp)
    $rows = Read-DebugCSV -TimeWindow "All Time"
    # Match by SAM; use timestamp as a tiebreaker if multiple sessions for same SAM
    $session = $rows | Where-Object { $_.SAM -eq $Sam } | Sort-Object Timestamp
    if (-not $session) { $RichDebugDetail.Text = "No trace found for $Sam"; return }

    $RichDebugDetail.Clear()
    $Script:DebugTraceText = ""
    foreach ($r in $session) {
        $color = switch -Wildcard ($r.Status) {
            "*FAIL*"  { [System.Drawing.Color]::OrangeRed }
            "*ERROR*" { [System.Drawing.Color]::OrangeRed }
            "*WARN*"  { [System.Drawing.Color]::Gold }
            "*OK*"    { [System.Drawing.Color]::LimeGreen }
            default   { [System.Drawing.Color]::Cyan }
        }
        $line = "[$($r.Timestamp)] [$($r.Status.PadRight(12))] $($r.Step): $($r.Detail)"
        $RichDebugDetail.SelectionStart  = $RichDebugDetail.TextLength
        $RichDebugDetail.SelectionLength = 0
        $RichDebugDetail.SelectionColor  = $color
        $RichDebugDetail.AppendText("$line`n")
        $Script:DebugTraceText += "$line`r`n"
    }
    $BtnDebugCopy.Enabled = $true
    $GrpDebugDetail.Text  = "Debug Trace  --  $Sam  ($($session.Count) steps)"
}

# ---- DEBUG TAB EVENTS --------------------------------------
$DebugGrid.Add_CellDoubleClick({
    $row = $DebugGrid.CurrentRow
    if (-not $row) { return }
    $sam = $row.Cells["SAM"].Value
    $ts  = $row.Cells["Timestamp"].Value
    if ($sam) { Load-DebugDetail -Sam $sam -Timestamp $ts }
})

$BtnDebugRefresh.Add_Click({
    $tw = if ($CboDebugTime.SelectedItem) { $CboDebugTime.SelectedItem.ToString() } else { "Last 30 Days" }
    Load-DebugGrid -TimeWindow $tw
})

$CboDebugTime.Add_SelectedIndexChanged({
    $tw = $CboDebugTime.SelectedItem.ToString()
    Load-DebugGrid -TimeWindow $tw
})

$BtnDebugCopy.Add_Click({
    if ($Script:DebugTraceText) {
        [System.Windows.Forms.Clipboard]::SetText($Script:DebugTraceText)
        $LblDebugStatus.Text = "Debug trace copied to clipboard."
    }
})

# ============================================================
#  TAB 5 - BULK CREATE
# ============================================================
$TabBulk           = New-Object System.Windows.Forms.TabPage
$TabBulk.Text      = "  Bulk Create  "
$TabBulk.BackColor = $ColorGray
$Tabs.Controls.Add($TabBulk)

# ---- BULK INPUT GRID ------------------------------------
$bulkCols = @(
    @{Label="First *";     X=8;   W=100},
    @{Label="Mid";         X=112; W=40},
    @{Label="Last *";      X=156; W=100},
    @{Label="Division *";  X=260; W=155},
    @{Label="Emp Type *";  X=419; W=120},
    @{Label="Category *";  X=543; W=80},
    @{Label="Acct Type *"; X=627; W=100},
    @{Label="Ticket #";    X=731; W=80}
)

foreach ($col in $bulkCols) {
    $lbl           = New-Object System.Windows.Forms.Label
    $lbl.Text      = $col.Label
    $lbl.Location  = (New-Object System.Drawing.Point -Property @{X=$col.X; Y=8})
    $lbl.Size      = (New-Object System.Drawing.Size -Property @{Width=$col.W; Height=16})
    $lbl.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Bold))
    $lbl.ForeColor = $ColorDark
    $TabBulk.Controls.Add($lbl)
}

# Fill Down buttons
$bulkFillButtons = @{}
$fillY = 26
$fillCols = @(
    @{Key="Division"; X=260; W=155},
    @{Key="EmpType";  X=419; W=120},
    @{Key="Category"; X=543; W=80},
    @{Key="AcctType"; X=627; W=100}
)
foreach ($fc in $fillCols) {
    $btn           = New-Object System.Windows.Forms.Button
    $btn.Text      = "Fill Down"
    $btn.Size      = (New-Object System.Drawing.Size -Property @{Width=$fc.W; Height=20})
    $btn.Location  = (New-Object System.Drawing.Point -Property @{X=$fc.X; Y=$fillY})
    $btn.BackColor = $ColorAccent
    $btn.ForeColor = $ColorWhite
    $btn.FlatStyle = "Flat"
    $btn.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Bold))
    $btn.Tag       = $fc.Key
    $TabBulk.Controls.Add($btn)
    $bulkFillButtons[$fc.Key] = $btn
}

# 5 account rows
$bulkRows = @()
$rowY = 48
for ($r = 0; $r -lt 5; $r++) {
    $row = @{}
    $row.Index = $r

    $numLbl           = New-Object System.Windows.Forms.Label
    $numLbl.Text      = "$($r+1)"
    $numLbl.Location  = (New-Object System.Drawing.Point -Property @{X=0; Y=($rowY+4)})
    $numLbl.Size      = (New-Object System.Drawing.Size -Property @{Width=8; Height=20})
    $numLbl.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Regular))
    $numLbl.ForeColor = $ColorDark
    $TabBulk.Controls.Add($numLbl)
    $row.NumLbl = $numLbl

    $txF = New-Object System.Windows.Forms.TextBox
    $txF.Size = (New-Object System.Drawing.Size -Property @{Width=100; Height=22})
    $txF.Location = (New-Object System.Drawing.Point -Property @{X=8; Y=$rowY})
    $txF.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Regular))
    $TabBulk.Controls.Add($txF); $row.First = $txF

    $txM = New-Object System.Windows.Forms.TextBox
    $txM.Size = (New-Object System.Drawing.Size -Property @{Width=40; Height=22})
    $txM.Location = (New-Object System.Drawing.Point -Property @{X=112; Y=$rowY})
    $txM.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Regular))
    $txM.MaxLength = 1
    $TabBulk.Controls.Add($txM); $row.Mid = $txM

    $txL = New-Object System.Windows.Forms.TextBox
    $txL.Size = (New-Object System.Drawing.Size -Property @{Width=100; Height=22})
    $txL.Location = (New-Object System.Drawing.Point -Property @{X=156; Y=$rowY})
    $txL.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Regular))
    $TabBulk.Controls.Add($txL); $row.Last = $txL

    $cbDiv = New-Object System.Windows.Forms.ComboBox
    $cbDiv.Size = (New-Object System.Drawing.Size -Property @{Width=155; Height=22})
    $cbDiv.Location = (New-Object System.Drawing.Point -Property @{X=260; Y=$rowY})
    $cbDiv.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Regular))
    $cbDiv.DropDownStyle = "DropDown"
    $cbDiv.AutoCompleteMode = "SuggestAppend"
    $cbDiv.AutoCompleteSource = "ListItems"
    $cbDiv.Items.Add("-- Select --") | Out-Null
    foreach ($div in $DivisionMap) { $cbDiv.Items.Add("$($div.Code) - $($div.Name)") | Out-Null }
    $cbDiv.SelectedIndex = 0
    $TabBulk.Controls.Add($cbDiv); $row.Division = $cbDiv

    $cbEmp = New-Object System.Windows.Forms.ComboBox
    $cbEmp.Size = (New-Object System.Drawing.Size -Property @{Width=120; Height=22})
    $cbEmp.Location = (New-Object System.Drawing.Point -Property @{X=419; Y=$rowY})
    $cbEmp.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Regular))
    $cbEmp.DropDownStyle = "DropDownList"
    $cbEmp.Items.Add("-- Select --") | Out-Null
    @("Knowledge Worker","Frontline Worker") | ForEach-Object { $cbEmp.Items.Add($_) | Out-Null }
    $cbEmp.SelectedIndex = 0
    $TabBulk.Controls.Add($cbEmp); $row.EmpType = $cbEmp

    $cbCat = New-Object System.Windows.Forms.ComboBox
    $cbCat.Size = (New-Object System.Drawing.Size -Property @{Width=80; Height=22})
    $cbCat.Location = (New-Object System.Drawing.Point -Property @{X=543; Y=$rowY})
    $cbCat.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Regular))
    $cbCat.DropDownStyle = "DropDownList"
    foreach ($c in $CategoryList) { $cbCat.Items.Add($c) | Out-Null }
    $cbCat.SelectedIndex = 0
    $TabBulk.Controls.Add($cbCat); $row.Category = $cbCat

    $cbAT = New-Object System.Windows.Forms.ComboBox
    $cbAT.Size = (New-Object System.Drawing.Size -Property @{Width=100; Height=22})
    $cbAT.Location = (New-Object System.Drawing.Point -Property @{X=627; Y=$rowY})
    $cbAT.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Regular))
    $cbAT.DropDownStyle = "DropDownList"
    foreach ($a in $AccountTypeList) { $cbAT.Items.Add($a) | Out-Null }
    $cbAT.SelectedIndex = 0
    $TabBulk.Controls.Add($cbAT); $row.AcctType = $cbAT

    $txTk = New-Object System.Windows.Forms.TextBox
    $txTk.Size = (New-Object System.Drawing.Size -Property @{Width=80; Height=22})
    $txTk.Location = (New-Object System.Drawing.Point -Property @{X=731; Y=$rowY})
    $txTk.Font = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Regular))
    $TabBulk.Controls.Add($txTk); $row.Ticket = $txTk

    $bulkRows += $row
    $rowY += 26
}

# Create All works from the Preview All results, so any edit to a row's inputs sends the tech
# back to Preview All. The ticket is read again at create time and does not need this.
foreach ($r in $bulkRows) {
    foreach ($ctl in @($r.First, $r.Mid, $r.Last)) { $ctl.Add_TextChanged({ if (-not $Script:BulkRunning) { $BtnBulkCreate.Enabled = $false } }) }
    foreach ($ctl in @($r.Division, $r.EmpType, $r.Category, $r.AcctType)) { $ctl.Add_SelectedIndexChanged({ if (-not $Script:BulkRunning) { $BtnBulkCreate.Enabled = $false } }) }
}

# Buttons row
$btnRowY = $rowY + 4
$BtnBulkPreview           = New-Object System.Windows.Forms.Button
$BtnBulkPreview.Text      = "Preview All"
$BtnBulkPreview.Size      = (New-Object System.Drawing.Size -Property @{Width=110; Height=28})
$BtnBulkPreview.Location  = (New-Object System.Drawing.Point -Property @{X=8; Y=$btnRowY})
$BtnBulkPreview.BackColor = $ColorLight
$BtnBulkPreview.ForeColor = $ColorDark
$BtnBulkPreview.FlatStyle = "Flat"
$BtnBulkPreview.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabBulk.Controls.Add($BtnBulkPreview)

$BtnBulkCreate           = New-Object System.Windows.Forms.Button
$BtnBulkCreate.Text      = "Create All"
$BtnBulkCreate.Size      = (New-Object System.Drawing.Size -Property @{Width=110; Height=28})
$BtnBulkCreate.Location  = (New-Object System.Drawing.Point -Property @{X=126; Y=$btnRowY})
$BtnBulkCreate.BackColor = $ColorAccent
$BtnBulkCreate.ForeColor = $ColorWhite
$BtnBulkCreate.FlatStyle = "Flat"
$BtnBulkCreate.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnBulkCreate.Enabled   = $false
$TabBulk.Controls.Add($BtnBulkCreate)

$BtnBulkClear           = New-Object System.Windows.Forms.Button
$BtnBulkClear.Text      = "Clear All"
$BtnBulkClear.Size      = (New-Object System.Drawing.Size -Property @{Width=90; Height=28})
$BtnBulkClear.Location  = (New-Object System.Drawing.Point -Property @{X=244; Y=$btnRowY})
$BtnBulkClear.BackColor = $ColorGray
$BtnBulkClear.ForeColor = $ColorDark
$BtnBulkClear.FlatStyle = "Flat"
$BtnBulkClear.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Regular))
$TabBulk.Controls.Add($BtnBulkClear)

$BtnBulkCopyAll           = New-Object System.Windows.Forms.Button
$BtnBulkCopyAll.Text      = "Copy All Results"
$BtnBulkCopyAll.Size      = (New-Object System.Drawing.Size -Property @{Width=130; Height=28})
$BtnBulkCopyAll.Location  = (New-Object System.Drawing.Point -Property @{X=342; Y=$btnRowY})
$BtnBulkCopyAll.BackColor = $ColorLight
$BtnBulkCopyAll.ForeColor = $ColorDark
$BtnBulkCopyAll.FlatStyle = "Flat"
$BtnBulkCopyAll.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnBulkCopyAll.Enabled   = $false
$TabBulk.Controls.Add($BtnBulkCopyAll)

$LblBulkStatus          = New-Object System.Windows.Forms.Label
$LblBulkStatus.Location = (New-Object System.Drawing.Point -Property @{X=480; Y=($btnRowY+6)})
$LblBulkStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=350; Height=18})
$LblBulkStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Italic))
$TabBulk.Controls.Add($LblBulkStatus)

# ---- STATUS BOARD (Mini consoles per account) --------
$statusY = $btnRowY + 34
$GrpBulkStatus           = New-Object System.Windows.Forms.GroupBox
$GrpBulkStatus.Text      = "Creation Status"
$GrpBulkStatus.Size      = (New-Object System.Drawing.Size -Property @{Width=900; Height=395})
$GrpBulkStatus.Location  = (New-Object System.Drawing.Point -Property @{X=8; Y=$statusY})
$GrpBulkStatus.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpBulkStatus.ForeColor = $ColorDark
$TabBulk.Controls.Add($GrpBulkStatus)

$Script:BulkStatusRows = @()
for ($r = 0; $r -lt 5; $r++) {
    $sy = 18 + ($r * 74)
    $srow = @{}

    # Header line: account number + name + status + buttons
    $nl           = New-Object System.Windows.Forms.Label
    $nl.Text      = "$($r+1).  (empty)"
    $nl.Location  = (New-Object System.Drawing.Point -Property @{X=8; Y=$sy})
    $nl.Size      = (New-Object System.Drawing.Size -Property @{Width=500; Height=16})
    $nl.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Bold))
    $nl.ForeColor = $ColorRGB_150_150_150
    $GrpBulkStatus.Controls.Add($nl); $srow.NameLbl = $nl

    $sl           = New-Object System.Windows.Forms.Label
    $sl.Text      = "Queued"
    $sl.Location  = (New-Object System.Drawing.Point -Property @{X=510; Y=$sy})
    $sl.Size      = (New-Object System.Drawing.Size -Property @{Width=100; Height=16})
    $sl.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]8),([System.Drawing.FontStyle]::Bold))
    $sl.ForeColor = $ColorRGB_150_150_150
    $GrpBulkStatus.Controls.Add($sl); $srow.StatusLbl = $sl

    # Skip / Retry / Skip Sync buttons (right of header)
    $skipBtn           = New-Object System.Windows.Forms.Button
    $skipBtn.Text      = "Skip"
    $skipBtn.Size      = (New-Object System.Drawing.Size -Property @{Width=65; Height=18})
    $skipBtn.Location  = (New-Object System.Drawing.Point -Property @{X=620; Y=$sy})
    $skipBtn.FlatStyle = "Flat"
    $skipBtn.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Bold))
    $skipBtn.Visible   = $false
    $skipBtn.Tag       = $r
    try { $skipBtn.BackColor = [System.Drawing.Color]::DarkRed } catch {}
    try { $skipBtn.ForeColor = $ColorWhite } catch {}
    $GrpBulkStatus.Controls.Add($skipBtn); $srow.SkipBtn = $skipBtn

    $retryBtn           = New-Object System.Windows.Forms.Button
    $retryBtn.Text      = "Retry"
    $retryBtn.Size      = (New-Object System.Drawing.Size -Property @{Width=65; Height=18})
    $retryBtn.Location  = (New-Object System.Drawing.Point -Property @{X=690; Y=$sy})
    $retryBtn.FlatStyle = "Flat"
    $retryBtn.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Bold))
    $retryBtn.Visible   = $false
    $retryBtn.Tag       = $r
    try { $retryBtn.BackColor = $ColorAccent } catch {}
    try { $retryBtn.ForeColor = $ColorWhite } catch {}
    $GrpBulkStatus.Controls.Add($retryBtn); $srow.RetryBtn = $retryBtn

    $syncSkipBtn           = New-Object System.Windows.Forms.Button
    $syncSkipBtn.Text      = "Skip Sync"
    $syncSkipBtn.Size      = (New-Object System.Drawing.Size -Property @{Width=85; Height=18})
    $syncSkipBtn.Location  = (New-Object System.Drawing.Point -Property @{X=760; Y=$sy})
    $syncSkipBtn.FlatStyle = "Flat"
    $syncSkipBtn.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]7),([System.Drawing.FontStyle]::Bold))
    $syncSkipBtn.Visible   = $false
    $syncSkipBtn.Tag       = $r
    try { $syncSkipBtn.BackColor = $ColorRGB_120_80_0 } catch {}
    try { $syncSkipBtn.ForeColor = $ColorWhite } catch {}
    $GrpBulkStatus.Controls.Add($syncSkipBtn); $srow.SyncSkipBtn = $syncSkipBtn

    # Mini console (RichTextBox per account)
    $mc           = New-Object System.Windows.Forms.RichTextBox
    $mc.Size      = (New-Object System.Drawing.Size -Property @{Width=882; Height=52})
    $mc.Location  = (New-Object System.Drawing.Point -Property @{X=8; Y=($sy+18)})
    $mc.Font      = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]7),([System.Drawing.FontStyle]::Regular))
    $mc.ReadOnly  = $true
    $mc.ScrollBars = "Vertical"
    try { $mc.BackColor = [System.Drawing.Color]::FromArgb(-15263969) } catch {} # dark gray/black
    try { $mc.ForeColor = [System.Drawing.Color]::LimeGreen } catch {}
    $GrpBulkStatus.Controls.Add($mc); $srow.Console = $mc

    $srow.Phase         = "queued"
    $srow.SkipRequested = $false
    $srow.RetryNeeded   = $false
    $srow.AccountData   = $null
    $srow.Created       = $false
    $Script:BulkStatusRows += $srow
}

# ---- BULK HELPER FUNCTIONS ----------------------------
$Script:BulkResults   = @()
$Script:BulkRunning   = $false

function BulkLog {
    # Writes a colored line to the per-account mini console
    param($Idx, $Message, $Color)
    try {
        $mc = $Script:BulkStatusRows[$Idx].Console
        if (-not $Color) { $Color = [System.Drawing.Color]::LimeGreen }
        $mc.SelectionStart  = $mc.TextLength
        $mc.SelectionLength = 0
        $mc.SelectionColor  = $Color
        $ts = (Get-Date).ToString("HH:mm:ss")
        $mc.AppendText("[$ts] $Message`n")
        $mc.ScrollToCaret()
    } catch {}
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-BulkRow {
    param($Idx, $Phase, $LogMsg="", $Color=$null)
    try {
        $srow = $Script:BulkStatusRows[$Idx]
        try { $srow.StatusLbl.Text = $Phase } catch {}
        switch -Wildcard ($Phase) {
            "COMPLETE" {
                try { $srow.NameLbl.ForeColor   = [System.Drawing.Color]::LimeGreen } catch {}
                try { $srow.StatusLbl.ForeColor = [System.Drawing.Color]::LimeGreen } catch {}
            }
            "FAILED*" {
                try { $srow.NameLbl.ForeColor   = [System.Drawing.Color]::OrangeRed } catch {}
                try { $srow.StatusLbl.ForeColor = [System.Drawing.Color]::OrangeRed } catch {}
            }
            "PARTIAL*" {
                try { $srow.NameLbl.ForeColor   = [System.Drawing.Color]::Orange } catch {}
                try { $srow.StatusLbl.ForeColor = [System.Drawing.Color]::Orange } catch {}
            }
            "SKIPPED" {
                try { $srow.NameLbl.ForeColor   = [System.Drawing.Color]::Gray } catch {}
                try { $srow.StatusLbl.ForeColor = [System.Drawing.Color]::Gray } catch {}
            }
            "Queued" {
                try { $srow.NameLbl.ForeColor   = $ColorRGB_150_150_150 } catch {}
                try { $srow.StatusLbl.ForeColor = $ColorRGB_150_150_150 } catch {}
            }
            default {
                try { $srow.NameLbl.ForeColor   = $ColorAccent } catch {}
                try { $srow.StatusLbl.ForeColor = $ColorGold } catch {}
            }
        }
        if ($LogMsg) { BulkLog $Idx $LogMsg $Color }
    } catch {}
    [System.Windows.Forms.Application]::DoEvents()
}

# ---- FILL DOWN BUTTON EVENTS -------------------------
foreach ($fc in $fillCols) {
    $bulkFillButtons[$fc.Key].Add_Click({
        $key = $this.Tag
        $srcRow = $bulkRows[0]
        for ($i = 1; $i -lt 5; $i++) {
            $hasName = ($bulkRows[$i].First.Text.Trim() -or $bulkRows[$i].Last.Text.Trim())
            if (-not $hasName) { continue }
            switch ($key) {
                "Division" { $bulkRows[$i].Division.SelectedIndex = $srcRow.Division.SelectedIndex }
                "EmpType"  { $bulkRows[$i].EmpType.SelectedIndex  = $srcRow.EmpType.SelectedIndex }
                "Category" { $bulkRows[$i].Category.SelectedIndex = $srcRow.Category.SelectedIndex }
                "AcctType" { $bulkRows[$i].AcctType.SelectedIndex = $srcRow.AcctType.SelectedIndex }
            }
        }
    }.GetNewClosure())
}

# Skip / Retry / SyncSkip event wiring
for ($r = 0; $r -lt 5; $r++) {
    $srow = $Script:BulkStatusRows[$r]
    $srow.SkipBtn.Add_Click({
        $idx = [int]$this.Tag
        $Script:BulkStatusRows[$idx].SkipRequested = $true
        $Script:BulkStatusRows[$idx].SkipBtn.Enabled  = $false
        $Script:BulkStatusRows[$idx].RetryBtn.Enabled = $false
    }.GetNewClosure())
    $srow.RetryBtn.Add_Click({
        $idx = [int]$this.Tag
        $Script:BulkStatusRows[$idx].RetryNeeded = $true
        $Script:BulkStatusRows[$idx].SkipBtn.Enabled  = $false
        $Script:BulkStatusRows[$idx].RetryBtn.Enabled = $false
    }.GetNewClosure())
    $srow.SyncSkipBtn.Add_Click({
        $idx = [int]$this.Tag
        $Script:BulkStatusRows[$idx].SkipRequested = $true
        $Script:BulkStatusRows[$idx].SyncSkipBtn.Enabled = $false
    }.GetNewClosure())
}

# ---- BULK CLEAR ---------------------------------------
$BtnBulkClear.Add_Click({
    # Clearing mid-run would reset the Skip/Retry flags a waiting row depends on and empty its ticket.
    if ($Script:BulkRunning -or $Script:SingleRunning) { return }
    for ($i = 0; $i -lt 5; $i++) {
        $bulkRows[$i].First.Clear(); $bulkRows[$i].Mid.Clear()
        $bulkRows[$i].Last.Clear();  $bulkRows[$i].Ticket.Clear()
        $bulkRows[$i].Division.SelectedIndex = 0
        $bulkRows[$i].EmpType.SelectedIndex  = 0
        $bulkRows[$i].Category.SelectedIndex = 0
        $bulkRows[$i].AcctType.SelectedIndex = 0
        $Script:BulkStatusRows[$i].NameLbl.Text   = "$($i+1).  (empty)"
        $Script:BulkStatusRows[$i].StatusLbl.Text = "Queued"
        try { $Script:BulkStatusRows[$i].NameLbl.ForeColor = $ColorRGB_150_150_150 } catch {}
        try { $Script:BulkStatusRows[$i].StatusLbl.ForeColor = $ColorRGB_150_150_150 } catch {}
        $Script:BulkStatusRows[$i].Console.Clear()
        $Script:BulkStatusRows[$i].SkipBtn.Visible     = $false
        $Script:BulkStatusRows[$i].RetryBtn.Visible    = $false
        $Script:BulkStatusRows[$i].SyncSkipBtn.Visible = $false
        $Script:BulkStatusRows[$i].Phase           = "queued"
        $Script:BulkStatusRows[$i].SkipRequested   = $false
        $Script:BulkStatusRows[$i].RetryNeeded     = $false
        $Script:BulkStatusRows[$i].AccountData     = $null
        $Script:BulkStatusRows[$i].Created         = $false
    }
    $Script:BulkResults     = @()
    $BtnBulkCreate.Enabled  = $false
    $BtnBulkCopyAll.Enabled = $false
    $LblBulkStatus.Text     = ""
})

# ---- BULK PREVIEW ------------------------------------
$BtnBulkPreview.Add_Click({
    if ($Script:BulkRunning) { return }
    if ($Script:SingleRunning) {
        [System.Windows.Forms.MessageBox]::Show("A single-account run is in progress. Wait for it to finish first.", "Run in progress") | Out-Null
        return
    }
    $LblBulkStatus.Text = "Running pre-flight checks..."; $Form.Refresh()
    $anyFilled = $false
    $allPassed = $true

    $BtnBulkCreate.Enabled = $false
    $pvwDir  = $PSScriptRoot; if (-not $pvwDir) { $pvwDir = "C:\ADProvisioning" }
    $pvwHdrs = Get-GraphToken -ScriptDir $pvwDir

    for ($i = 0; $i -lt 5; $i++) {
        $row    = $bulkRows[$i]
        $srow   = $Script:BulkStatusRows[$i]
        $srow.Console.Clear()
        # A row whose account was written (or may have been) is never queued again until Bulk Clear:
        # pre-flight would see that account, pick first.last2 and create a duplicate.
        if ($srow.Created) {
            BulkLog $i "Already created in this run -- use Clear before reusing this row" ([System.Drawing.Color]::Gray)
            continue
        }
        # Drop the previous Preview's result so an emptied or failed row cannot be created.
        $srow.AccountData = $null
        $srow.Phase = "queued"
        $first  = Format-ProperCase $row.First.Text.Trim()
        $last   = Format-ProperCase $row.Last.Text.Trim()
        if (-not $first -or -not $last) {
            $srow.NameLbl.Text = "$($i+1).  (empty)"
            try { $srow.NameLbl.ForeColor = $ColorRGB_150_150_150 } catch {}
            $srow.StatusLbl.Text = "Queued"
            continue
        }
        $anyFilled = $true
        $mid = Format-ProperCase $row.Mid.Text.Trim()

        # Validate required fields
        $sel = $row.Division.SelectedItem
        if (-not $sel -or $sel -eq "-- Select --") {
            Update-BulkRow $i "FAILED" "Select a division" ([System.Drawing.Color]::OrangeRed)
            $allPassed = $false; continue
        }
        $empSel = $row.EmpType.SelectedItem
        if (-not $empSel -or $empSel -eq "-- Select --") {
            Update-BulkRow $i "FAILED" "Select an employee type" ([System.Drawing.Color]::OrangeRed)
            $allPassed = $false; continue
        }
        $nameProblems = @(Get-NameInputProblem -First $first -Middle $mid -Last $last)
        if ($nameProblems.Count) {
            BulkLog $i "FAILED: $($nameProblems -join '; ')" ([System.Drawing.Color]::OrangeRed)
            Update-BulkRow $i "FAILED" "Check the name" ([System.Drawing.Color]::OrangeRed)
            $allPassed = $false; continue
        }

        $divCode    = Get-DivCodeFromItem $sel
        $divObj     = $DivisionMap | Where-Object { $_.Code -eq $divCode }
        $domain     = $divObj.Domain
        $emailDomain= Get-EmailDomain $divCode
        $dispName   = "$first $last - $divCode"
        $srow.NameLbl.Text = "$($i+1).  $dispName"
        BulkLog $i "Pre-flight: checking $first $last..." ([System.Drawing.Color]::Yellow)
        Update-BulkRow $i "Checking..."

        $resolved = Resolve-AccountCandidate `
            -First $first -Middle $mid -Last $last `
            -Domain $domain -EmailDomain $emailDomain `
            -GraphHeaders $pvwHdrs

        if ($resolved.Status -eq 'Resolved') {
            # Check if CN already exists in target OU (catches same-name people across batches)
            # cnName = CN for AD (must be unique per OU, may include (sam) suffix)
            # dispName = DisplayName for Entra/GAL (always clean, can duplicate)
            $cnName = $dispName
            $ou = Get-OUPath -Domain $domain
            try {
                $existingCN = Get-ADObject -LDAPFilter "(name=$(ConvertTo-LdapFilterValue $dispName))" -SearchBase $ou -Server $domain -ErrorAction Stop
                if ($existingCN) {
                    $cnName = "$first $last - $divCode ($($resolved.SAM))"
                    BulkLog $i "CN exists in OU -- CN will be: $cnName (display name stays: $dispName)" ([System.Drawing.Color]::Yellow)
                }
            } catch {
                BulkLog $i "[!] Could not validate OU: $ou" ([System.Drawing.Color]::Orange)
            }
            $lengthProblems = @(Get-NameLengthProblem -CnName $cnName -LocalPart $resolved.LocalPart)
            if ($lengthProblems.Count) {
                BulkLog $i "FAILED: $($lengthProblems -join '; ') -- shorten the name" ([System.Drawing.Color]::OrangeRed)
                Update-BulkRow $i "FAILED"
                $allPassed = $false; continue
            }
            foreach ($tp in @(Get-EntraTargetProblem -AUNames (Get-AUCodes $divCode) -GroupName $LicenseGroups[$empSel.ToString()] -Headers $pvwHdrs)) {
                BulkLog $i "[!] $tp -- assign by hand until the division map is fixed" ([System.Drawing.Color]::Orange)
                DBG "PREVIEW_TARGETS" "WARN" $tp
            }

            $srow.NameLbl.Text = "$($i+1).  $dispName"
            $srow.AccountData = @{
                First=$first; Mid=$mid; Last=$last; DivCode=$divCode
                Domain=$domain; EmailDomain=$emailDomain
                DispName=$dispName; CnName=$cnName
                EmpType=$empSel.ToString()
                CatId=$row.Category.SelectedItem.ToString().Split(" ")[0]
                AcctType=$row.AcctType.SelectedItem.ToString()
                Ticket=$row.Ticket.Text.Trim()
                Resolved=$resolved
            }
            BulkLog $i "SAM: $($resolved.SAM) | UPN: $($resolved.UPN)" ([System.Drawing.Color]::Cyan)
            if ($resolved.Truncated) { BulkLog $i "SAM shortened to the 20 character limit; UPN and email keep the full name" ([System.Drawing.Color]::Yellow) }
            if ($resolved.TenantAlias) { BulkLog $i "Alias: $($resolved.TenantAlias)" ([System.Drawing.Color]::Cyan) }
            BulkLog $i "All checks passed -- ready to create" ([System.Drawing.Color]::LimeGreen)
            Update-BulkRow $i "Ready"
            try { $srow.NameLbl.ForeColor = $ColorAccent } catch {}
        } else {
            BulkLog $i "FAILED ($($resolved.Status)): $($resolved.Reason)" ([System.Drawing.Color]::OrangeRed)
            Update-BulkRow $i "FAILED"
            $allPassed = $false
        }
        [System.Windows.Forms.Application]::DoEvents()
    }

    # ---- Cross-row validation with auto-resolution ----
    # Local parts are claimed batch-wide because the routing address and the
    # contoso.com alias are shared across every email domain. SAMs only need to be
    # unique within one AD domain. A row that clashes with an earlier row is
    # re-resolved with the earlier rows' claims excluded.
    $claimedLocal = @{}
    $claimedSam   = @{}
    for ($i = 0; $i -lt 5; $i++) {
        $sd = $Script:BulkStatusRows[$i]
        if (-not $sd.AccountData) { continue }
        if ($sd.Phase -eq "failed" -or $sd.Phase -eq "skipped") { continue }
        $rowDomain = $sd.AccountData.Domain
        $rowLocal  = $sd.AccountData.Resolved.LocalPart
        $rowSam    = $sd.AccountData.Resolved.SAM
        $clashRow  = Get-BatchClash -ClaimedLocal $claimedLocal -ClaimedSam $claimedSam -LocalPart $rowLocal -Domain $rowDomain -Sam $rowSam

        if ($null -ne $clashRow) {
            BulkLog $i "Name clashes with row $($clashRow+1) -- re-resolving..." ([System.Drawing.Color]::Yellow)
            $exclude = Get-BatchExclusion -ClaimedLocal $claimedLocal -ClaimedSam $claimedSam -Domain $rowDomain

            $reResolved = Resolve-AccountCandidate `
                -First $sd.AccountData.First `
                -Middle $sd.AccountData.Mid `
                -Last $sd.AccountData.Last `
                -Domain $rowDomain `
                -EmailDomain $sd.AccountData.EmailDomain `
                -GraphHeaders $pvwHdrs `
                -ExcludeLocalParts $exclude.LocalParts `
                -ExcludeSAMs $exclude.SAMs

            if ($reResolved.Status -eq 'Resolved') {
                $sd.AccountData.Resolved = $reResolved
                $newDispName = "$($sd.AccountData.First) $($sd.AccountData.Last) - $($sd.AccountData.DivCode)"
                $newCnName   = $newDispName
                $ouCheck = Get-OUPath -Domain $rowDomain
                try {
                    $existCN = Get-ADObject -LDAPFilter "(name=$(ConvertTo-LdapFilterValue $newDispName))" -SearchBase $ouCheck -Server $rowDomain -ErrorAction Stop
                    if ($existCN) { $newCnName = "$newDispName ($($reResolved.SAM))" }
                } catch { BulkLog $i "[!] Could not check for a name collision in $ouCheck : $_" ([System.Drawing.Color]::Orange) }
                $sd.AccountData.DispName = $newDispName
                $sd.AccountData.CnName   = $newCnName
                try { $sd.NameLbl.Text = "$($i+1).  $newDispName" } catch { DBG "BULK_UI" "WARN" "Row label update failed: $_" }
                BulkLog $i "Re-resolved: SAM=$($reResolved.SAM) UPN=$($reResolved.UPN)" ([System.Drawing.Color]::Cyan)
                BulkLog $i "Ready to create" ([System.Drawing.Color]::LimeGreen)
                Update-BulkRow $i "Ready"
                $claimedLocal[$reResolved.LocalPart] = $i
                $claimedSam["$rowDomain|$($reResolved.SAM)"] = $i
            } else {
                BulkLog $i "FAILED ($($reResolved.Status)) after excluding batch claims: $($reResolved.Reason)" ([System.Drawing.Color]::OrangeRed)
                Update-BulkRow $i "FAILED"
                $allPassed = $false
            }
        } else {
            $claimedLocal[$rowLocal]          = $i
            $claimedSam["$rowDomain|$rowSam"] = $i
        }
    }

    # CN duplicate check -- auto-resolve by appending (SAM) to CN only (not DisplayName)
    $seenCNs = @{}
    for ($i = 0; $i -lt 5; $i++) {
        $sd = $Script:BulkStatusRows[$i]
        if (-not $sd.AccountData) { continue }
        if ($sd.Phase -eq "failed" -or $sd.Phase -eq "skipped") { continue }
        $rowCN = "$($sd.AccountData.CnName)|$($sd.AccountData.Domain)"
        if ($seenCNs.ContainsKey($rowCN)) {
            $other = $seenCNs[$rowCN]
            $otherData = $Script:BulkStatusRows[$other].AccountData
            if (-not $otherData.CnName.EndsWith(")")) {
                $otherData.CnName = "$($otherData.DispName) ($($otherData.Resolved.SAM))"
                BulkLog $other "CN collision with row $($i+1): CN=$($otherData.CnName) (display name unchanged)" ([System.Drawing.Color]::Yellow)
            }
            $sd.AccountData.CnName = "$($sd.AccountData.DispName) ($($sd.AccountData.Resolved.SAM))"
            BulkLog $i "CN collision with row $($other+1): CN=$($sd.AccountData.CnName) (display name unchanged)" ([System.Drawing.Color]::Yellow)
        } else { $seenCNs[$rowCN] = $i }
    }

    if (-not $anyFilled) {
        $LblBulkStatus.Text = "Enter at least one account before previewing."
        $LblBulkStatus.ForeColor = $ColorRed
        return
    }
    if ($allPassed) {
        $LblBulkStatus.Text      = "Pre-flight passed. Click Create All to proceed."
        $LblBulkStatus.ForeColor = $ColorAccent
        $BtnBulkCreate.Enabled   = $true
    } else {
        $LblBulkStatus.Text      = "Fix red rows and re-preview."
        $LblBulkStatus.ForeColor = $ColorRed
        $BtnBulkCreate.Enabled   = $false
    }
})

# ---- COPY ALL RESULTS --------------------------------
$BtnBulkCopyAll.Add_Click({
    if ($Script:BulkResults.Count -eq 0) { return }
    $clip = $Script:BulkResults | ForEach-Object { $_ + "`r`n---`r`n" }
    [System.Windows.Forms.Clipboard]::SetText($clip -join "")
    $LblBulkStatus.Text      = "All results copied to clipboard."
    $LblBulkStatus.ForeColor = $ColorAccent
})

# ---- BULK CREATE MAIN FLOW ---------------------------
$BtnBulkCreate.Add_Click({
    if ($Script:BulkRunning) { return }
    if ($Script:SingleRunning) {
        [System.Windows.Forms.MessageBox]::Show("A single-account run is in progress. Wait for it to finish first.", "Run in progress") | Out-Null
        return
    }

    # Count queued accounts
    $queueCount = 0
    for ($qi = 0; $qi -lt 5; $qi++) {
        $qd = $Script:BulkStatusRows[$qi]
        if ($qd.AccountData -and $qd.Phase -eq "queued") { $queueCount++ }
    }
    if ($queueCount -eq 0) {
        $LblBulkStatus.Text = "No accounts queued. Run Preview All first."
        $LblBulkStatus.ForeColor = $ColorRed
        return
    }

    # Confirmation dialog
    $confirmMsg = "Create $queueCount account(s) now?`n`nThis will:`n- Create AD accounts`n- Enable Exchange mailboxes`n- Wait for Entra sync`n- Assign M365 licenses"
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        $confirmMsg,
        "Confirm Bulk Creation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne "Yes") { return }

    $Script:BulkRunning    = $true
    $BtnBulkCreate.Enabled = $false
    $BtnBulkPreview.Enabled= $false
    $Script:BulkResults    = @()

    $queue = @()
    for ($i = 0; $i -lt 5; $i++) {
        $sd = $Script:BulkStatusRows[$i]
        if ($sd.AccountData -and $sd.Phase -eq "queued") {
            $sd.Phase = "queued"
            $queue += @{ Idx=$i; Data=$sd.AccountData }
        }
    }
    if ($queue.Count -eq 0) {
        $LblBulkStatus.Text = "No accounts queued. Run Preview All first."
        $LblBulkStatus.ForeColor = $ColorRed
        $Script:BulkRunning = $false; $BtnBulkPreview.Enabled = $true; return
    }

    $LblBulkStatus.Text = "PHASE 1: Creating accounts..."
    $LblBulkStatus.ForeColor = $ColorAccent

    $bDir  = $PSScriptRoot; if (-not $bDir) { $bDir = "C:\ADProvisioning" }
    $bHdrs = Get-GraphToken -ScriptDir $bDir
    $createdAccounts = @()

    # ---- PHASE 1: Create all accounts ----
    foreach ($item in $queue) {
        $i    = $item.Idx
        $d    = $item.Data
        $srow = $Script:BulkStatusRows[$i]
        # The ticket box stays editable after Preview All, so read it now.
        $d.Ticket = $bulkRows[$i].Ticket.Text.Trim()
        $Script:AuditTicket = $d.Ticket
        $Script:DebugSessionSAM  = $d.Resolved.SAM
        $Script:DebugSessionName = $d.DispName

        BulkLog $i "Starting creation..." ([System.Drawing.Color]::Yellow)
        Update-BulkRow $i "P1: Connecting..."

        # Open Exchange session: 3 attempts, then Skip or Retry until one works or the tech skips.
        $bSession = $null
        $connected = $false
        $skipped = $false
        while (-not $connected -and -not $skipped) {
            for ($attempt = 0; $attempt -lt 3 -and -not $connected; $attempt++) {
                try {
                    $bSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                        -ConnectionUri $ExchangeServer -Authentication Kerberos -ErrorAction Stop
                    Import-PSSession $bSession -DisableNameChecking -AllowClobber | Out-Null
                    $connected = $true
                    BulkLog $i "Exchange session established" ([System.Drawing.Color]::LimeGreen)
                    DBG "BULK_EXCH" "OK" "Exchange session opened for account $($i+1)"
                } catch {
                    if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue; $bSession = $null }
                    if ($attempt -lt 2) {
                        BulkLog $i "Exchange connect attempt $($attempt+1) failed, retrying..." ([System.Drawing.Color]::Yellow)
                        Wait-UiDelay -Seconds 5
                    } else {
                        BulkLog $i "Exchange session failed after 3 attempts" ([System.Drawing.Color]::OrangeRed)
                        DBG "BULK_EXCH" "FAIL" "Exchange session failed: $_"
                    }
                }
            }
            if (-not $connected) {
                Update-BulkRow $i "FAILED" "Exchange session failed -- Skip or Retry" ([System.Drawing.Color]::OrangeRed)
                $srow.SkipBtn.Visible = $true; $srow.RetryBtn.Visible = $true
                $srow.SkipBtn.Enabled = $true; $srow.RetryBtn.Enabled = $true
                $srow.SkipRequested = $false; $srow.RetryNeeded = $false
                while (-not $srow.SkipRequested -and -not $srow.RetryNeeded) {
                    Start-Sleep -Milliseconds 200; [System.Windows.Forms.Application]::DoEvents()
                }
                $srow.SkipBtn.Visible = $false; $srow.RetryBtn.Visible = $false
                if ($srow.SkipRequested) {
                    $skipped = $true
                    $srow.Phase = "skipped"
                    Update-BulkRow $i "SKIPPED" "Skipped by tech" ([System.Drawing.Color]::Gray)
                    DBG "BULK_EXCH" "WARN" "Row $($i+1) skipped by tech after Exchange session failure"
                    Write-AuditLog "CREATE" $d.Resolved.SAM $d.Resolved.UPN $d.DivCode $d.EmpType "SKIPPED" "Bulk | Exchange session failed, nothing created"
                }
                if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue; $bSession = $null }
            }
        }
        if ($skipped) { continue }

        # Names can be taken between Preview All and Create All, so check this row again,
        # now with Exchange, before writing anything.
        $fresh = Get-GraphToken -ScriptDir $bDir
        if ($fresh) { $bHdrs = $fresh }
        $others  = @($queue | Where-Object { $_.Idx -ne $i })
        $exLocal = @($others | ForEach-Object { $_.Data.Resolved.LocalPart })
        $exSam   = @($others | Where-Object { $_.Data.Domain -eq $d.Domain } | ForEach-Object { $_.Data.Resolved.SAM })
        $recheck = Resolve-AccountCandidate -First $d.First -Middle $d.Mid -Last $d.Last -Domain $d.Domain `
            -EmailDomain $d.EmailDomain -GraphHeaders $bHdrs -ExchSessionActive -ExcludeLocalParts $exLocal -ExcludeSAMs $exSam
        if ($recheck.Status -ne 'Resolved') {
            $why = "$($recheck.Status): $($recheck.Reason)"
            BulkLog $i "FAILED: name check before create ($why)" ([System.Drawing.Color]::OrangeRed)
            Write-AuditLog "CREATE" $d.Resolved.SAM $d.Resolved.UPN $d.DivCode $d.EmpType "PREFLIGHT_BLOCKED" "Bulk recheck: $why"
            DBG "BULK_RECHECK" "FAIL" $why
            $srow.Phase = "failed"
            Update-BulkRow $i "FAILED"
            if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue }
            continue
        }
        if ($recheck.LocalPart -ne $d.Resolved.LocalPart -or $recheck.SAM -ne $d.Resolved.SAM) {
            # Preview All has no Exchange session, so an alias held only in Exchange first shows up here.
            # The new values already exclude every other row, so they are used instead of failing the row.
            BulkLog $i "Names changed since Preview All: now $($recheck.UPN) / $($recheck.SAM) (was $($d.Resolved.UPN) / $($d.Resolved.SAM))" ([System.Drawing.Color]::Yellow)
            DBG "BULK_RECHECK" "WARN" "Re-resolved from $($d.Resolved.UPN)/$($d.Resolved.SAM) to $($recheck.UPN)/$($recheck.SAM)"
            if ($d.CnName -ne $d.DispName) { $d.CnName = "$($d.DispName) ($($recheck.SAM))" }
            $d.Resolved = $recheck
        }
        if ($d.CnName -eq $d.DispName) {
            try {
                $cnHit = Get-ADObject -LDAPFilter "(name=$(ConvertTo-LdapFilterValue $d.DispName))" -SearchBase (Get-OUPath -Domain $d.Domain) -Server $d.Domain -ErrorAction Stop
                if ($cnHit) {
                    $d.CnName = "$($d.DispName) ($($d.Resolved.SAM))"
                    BulkLog $i "CN now exists in OU -- CN will be: $($d.CnName)" ([System.Drawing.Color]::Yellow)
                }
            } catch { DBG "BULK_RECHECK" "WARN" "CN check failed: $_" }
        }
        $lenNow = @(Get-NameLengthProblem -CnName $d.CnName -LocalPart $d.Resolved.LocalPart)
        if ($lenNow.Count) {
            BulkLog $i "FAILED: $($lenNow -join '; ') -- shorten the name" ([System.Drawing.Color]::OrangeRed)
            Write-AuditLog "CREATE" $d.Resolved.SAM $d.Resolved.UPN $d.DivCode $d.EmpType "PREFLIGHT_BLOCKED" "Bulk recheck: $($lenNow -join '; ')"
            $srow.Phase = "failed"
            Update-BulkRow $i "FAILED"
            if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue }
            continue
        }

        # Create account
        Update-BulkRow $i "P1: Creating..."
        $resolved    = $d.Resolved
        $sam         = $resolved.SAM
        $upn         = $resolved.UPN
        $routing     = $resolved.Routing
        $primarySmtp = $resolved.PrimarySmtp
        $tenantAlias  = $resolved.TenantAlias
        $ou          = Get-OUPath -Domain $d.Domain
        $plainPass   = New-TempPassword
        $secPass     = ConvertTo-SecureString $plainPass -AsPlainText -Force
        $targetPDC   = $null
        try { $targetPDC = (Get-ADDomain -Identity $d.Domain -Server $d.Domain).PDCEmulator }
        catch { $targetPDC = $d.Domain }

        $maxRetries = 2
        $created = $false
        for ($tryNum = 0; $tryNum -le $maxRetries -and -not $created; $tryNum++) {
            if ($tryNum -gt 0) {
                BulkLog $i "Retry attempt $tryNum..." ([System.Drawing.Color]::Yellow)
            }
            # Only New-RemoteMailbox is inside this try: once it succeeds the account exists,
            # and a later failure must not trigger a retry or drop the row from Phase 2.
            try {
                $nrmP = Get-NewRemoteMailboxParameter -Resolved $resolved -CnName $d.CnName -DispName $d.DispName `
                    -First $d.First -Last $d.Last -Ou $ou -Pdc $targetPDC -Password $secPass
                BulkLog $i "New-RemoteMailbox: $sam ..." ([System.Drawing.Color]::Yellow)
                $newMb = New-RemoteMailbox @nrmP -ErrorAction Stop
                $created = $true
                $srow.Created = $true
                BulkLog $i "Account created: $sam" ([System.Drawing.Color]::LimeGreen)
                DBG "BULK_CREATE" "OK" "New-RemoteMailbox succeeded: $sam"
                Write-AuditLog "CREATE" $sam $upn $d.DivCode $d.EmpType "AD_OK" "Bulk | New-RemoteMailbox | Ticket=$($d.Ticket)"
            } catch {
                DBG "BULK_CREATE" "FAIL" "$_"
                BulkLog $i "FAILED: $_" ([System.Drawing.Color]::OrangeRed)
                $state = Get-CreatedObjectState -Sam $sam -Pdc $targetPDC
                if ($state.State -ne 'Absent') {
                    $srow.Created = $true
                    # A retry would fail with "already exists" or, after a new Preview, create a numbered duplicate.
                    $what = if ($state.State -eq 'Exists') { "$sam exists in AD anyway: $($state.Detail)" } else { "could not check whether $sam was created: $($state.Detail)" }
                    BulkLog $i "[!] $what -- do NOT create again, finish by hand" ([System.Drawing.Color]::OrangeRed)
                    Write-AuditLog "PARTIAL_CREATE" $sam $upn $d.DivCode $d.EmpType $(if ($state.State -eq 'Exists') { "OBJECT_EXISTS" } else { "STOPPED_STEP2" }) "Bulk | $what | $_"
                    $srow.Phase = "failed"
                    Update-BulkRow $i "FAILED" "Check account -- finish by hand" ([System.Drawing.Color]::OrangeRed)
                    break
                }
                Write-AuditLog "CREATE" $sam $upn $d.DivCode $d.EmpType "FAILED" "Bulk: $_"

                if ($tryNum -lt $maxRetries) {
                    Update-BulkRow $i "FAILED" "Creation failed -- Skip or Retry" ([System.Drawing.Color]::OrangeRed)
                    $srow.SkipBtn.Visible = $true; $srow.RetryBtn.Visible = $true
                    $srow.SkipBtn.Enabled = $true; $srow.RetryBtn.Enabled = $true
                    $srow.SkipRequested = $false; $srow.RetryNeeded = $false
                    while (-not $srow.SkipRequested -and -not $srow.RetryNeeded) {
                        Start-Sleep -Milliseconds 200; [System.Windows.Forms.Application]::DoEvents()
                    }
                    $srow.SkipBtn.Visible = $false; $srow.RetryBtn.Visible = $false
                    if ($srow.SkipRequested) {
                        $srow.Phase = "skipped"
                        Update-BulkRow $i "SKIPPED" "Skipped by tech" ([System.Drawing.Color]::Gray)
                        DBG "BULK_CREATE" "WARN" "$sam skipped by tech after a failed create"
                        break
                    }
                    $srow.RetryNeeded = $false
                    # The tech may have waited a long time before Retry; Exchange drops idle sessions.
                    if (-not $bSession -or $bSession.State -ne 'Opened') {
                        if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue; $bSession = $null }
                        try {
                            $bSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                -ConnectionUri $ExchangeServer -Authentication Kerberos -ErrorAction Stop
                            Import-PSSession $bSession -DisableNameChecking -AllowClobber | Out-Null
                            DBG "BULK_EXCH" "OK" "Exchange session reopened before retry"
                        } catch {
                            DBG "BULK_EXCH" "FAIL" "Exchange session reopen failed: $_"
                            BulkLog $i "Exchange session could not be reopened: $_" ([System.Drawing.Color]::OrangeRed)
                        }
                    }
                } else {
                    $srow.Phase = "failed"
                    Update-BulkRow $i "FAILED"
                }
            }
        }

        if ($created) {
            $warnings = @()
            $emailAddr = "$($newMb.PrimarySmtpAddress)"
            if ($emailAddr -ine $primarySmtp) {
                try {
                    Set-RemoteMailbox -Identity $sam -PrimarySmtpAddress $primarySmtp -EmailAddressPolicyEnabled $false `
                        -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                    BulkLog $i "  Primary email corrected: $emailAddr -> $primarySmtp" ([System.Drawing.Color]::Yellow)
                    Write-AuditLog "MAILBOX" $sam $upn $d.DivCode $d.EmpType "OK" "Bulk | Primary corrected to $primarySmtp"
                    $emailAddr = $primarySmtp
                } catch {
                    BulkLog $i "  [!] Primary email is $emailAddr, expected $primarySmtp -- fix manually" ([System.Drawing.Color]::Orange)
                    Write-AuditLog "MAILBOX" $sam $upn $d.DivCode $d.EmpType "WARN" "Bulk | Primary is $emailAddr, expected $primarySmtp : $_"
                    DBG "BULK_PRIMARY" "WARN" "$sam primary $emailAddr, expected $primarySmtp : $_"
                    $warnings += "primary email is $emailAddr"
                }
            }

            BulkLog $i "Setting proxy addresses..." ([System.Drawing.Color]::Yellow)
            # If the mailbox cannot be read, every address is added anyway: a duplicate add is harmless,
            # a missing alias is not.
            $current = @()
            try { $current = @((Get-RemoteMailbox -Identity $sam -DomainController $targetPDC -ErrorAction Stop).EmailAddresses) }
            catch { DBG "BULK_ALIAS" "WARN" "$sam could not read current addresses, adding all: $_" }
            $adds = @()
            if ($tenantAlias -and -not (Test-ProxyAddressPresent $current $tenantAlias)) { $adds += "smtp:$tenantAlias" }
            if (-not (Test-ProxyAddressPresent $current $routing)) { $adds += "smtp:$routing" }
            Wait-UiDelay -Seconds 2
            try {
                $legDN = (Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC -Properties legacyExchangeDN -ErrorAction Stop).legacyExchangeDN
                if ($legDN) { $adds += "X500:$legDN" }
                else { BulkLog $i "  [!] legacyExchangeDN empty -- X500 not added" ([System.Drawing.Color]::Orange); $warnings += "X500 not added" }
            } catch {
                BulkLog $i "  [!] Could not read legacyExchangeDN -- X500 not added: $_" ([System.Drawing.Color]::Orange)
                DBG "BULK_ALIAS" "WARN" "$sam legacyExchangeDN read failed: $_"
                $warnings += "X500 not added"
            }
            foreach ($addr in $adds) {
                try {
                    Set-RemoteMailbox -Identity $sam -EmailAddresses @{ add = $addr } -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                    BulkLog $i "  Added $addr" ([System.Drawing.Color]::LimeGreen)
                    Write-AuditLog "ALIAS" $sam $upn $d.DivCode $d.EmpType "OK" "Bulk | Added $addr"
                } catch {
                    BulkLog $i "  [!] Could not add $addr : $_" ([System.Drawing.Color]::Orange)
                    Write-AuditLog "ALIAS" $sam $upn $d.DivCode $d.EmpType "WARN" "Bulk | $addr failed: $_"
                    DBG "BULK_ALIAS" "WARN" "$sam $addr failed: $_"
                    $warnings += "$addr not added"
                }
            }

            Wait-UiDelay -Seconds 2
            try {
                # ADSI instead of Set-ADUser: Set-ADUser on these schema attributes raises a
                # warning that ps2exe turns into a blocking pop-up.
                $adObj = Get-ADUser -Identity $sam -Server $targetPDC -ErrorAction Stop
                $adsi = [ADSI]"LDAP://$targetPDC/$($adObj.DistinguishedName)"
                $adsi.Put("orgCategoryID", $d.CatId)
                $adsi.Put("orgAccountType", $d.AcctType)
                $adsi.SetInfo()
                BulkLog $i "  Attributes: Cat=$($d.CatId) Type=$($d.AcctType)" ([System.Drawing.Color]::LimeGreen)
                Write-AuditLog "ATTRS" $sam $upn $d.DivCode $d.EmpType "OK" "Bulk | orgCategoryID=$($d.CatId) orgAccountType=$($d.AcctType)"
            } catch {
                BulkLog $i "  [!] Custom attrs failed: $_" ([System.Drawing.Color]::Orange)
                Write-AuditLog "ATTRS" $sam $upn $d.DivCode $d.EmpType "FAILED" "Bulk | $_"
                DBG "BULK_ATTRS" "WARN" "$sam custom attributes failed: $_"
                $warnings += "custom attributes not set"
            }

            $immutableId = $null
            try {
                $adObj = Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC -Properties ObjectGUID -ErrorAction Stop
                if ($adObj) { $immutableId = [Convert]::ToBase64String(([guid]$adObj.ObjectGUID).ToByteArray()) }
            } catch { DBG "BULK_GUID" "WARN" "$sam ObjectGUID read failed, sync will match on UPN only: $_" }

            Write-AuditLog "MAILBOX" $sam $upn $d.DivCode $d.EmpType $(if ($warnings.Count) { "WARN" } else { "OK" }) "Bulk | Primary=$emailAddr | Routing=$routing$(if ($warnings.Count) { ' | ' + ($warnings -join '; ') })"
            BulkLog $i "Phase 1 complete. Temp pwd: $plainPass" ([System.Drawing.Color]::Cyan)
            Update-BulkRow $i "P1: Done"

            $domainShort = $d.Domain.Split(".")[0].ToUpper()
            $createdAccounts += @{
                Idx=$i; SAM=$sam; UPN=$upn; Email=$emailAddr; DispName=$d.DispName
                EmpType=$d.EmpType; DivCode=$d.DivCode; DomainShort=$domainShort
                PlainPass=$plainPass; ImmutableId=$immutableId
                Synced=$false; Licensed=$false; SkipSync=$false
                Warnings=$warnings; UpnNote=$null; AuthFails=0; WrongWarned=$false; Ticket=$d.Ticket
            }
            $srow.Phase = "syncing"
        }
        if ($bSession) { Remove-PSSession $bSession -ErrorAction SilentlyContinue }
    }

    if ($createdAccounts.Count -eq 0) {
        $LblBulkStatus.Text = "No accounts were created: $($queue.Count) not created. See each row."
        $LblBulkStatus.ForeColor = $ColorRed
        $Script:BulkRunning = $false; $Script:AuditTicket = $null; $BtnBulkPreview.Enabled = $true
        try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
        return
    }

    # ---- PHASE 2: Sync + license as they appear ----
    $Script:AuditTicket = $null
    $LblBulkStatus.Text = "PHASE 2: Waiting for Entra sync..."
    $LblBulkStatus.ForeColor = $ColorGold

    foreach ($acc in $createdAccounts) {
        $Script:BulkStatusRows[$acc.Idx].SyncSkipBtn.Visible = $true
        $Script:BulkStatusRows[$acc.Idx].SyncSkipBtn.Enabled = $true
        $Script:BulkStatusRows[$acc.Idx].SkipRequested       = $false
        BulkLog $acc.Idx "Waiting for Entra sync..." ([System.Drawing.Color]::Yellow)
        Update-BulkRow $acc.Idx "P2: Syncing..."
    }

    $maxWait = 2700
    $elapsed = 0
    $interval= 120

    # Initial 2-min wait
    for ($s = 120; $s -ge 0; $s--) {
        $LblBulkStatus.Text = "Phase 2: First sync check in $s sec..."
        Start-Sleep -Seconds 1
        [System.Windows.Forms.Application]::DoEvents()
        # Allow skip during countdown
        foreach ($acc in ($createdAccounts | Where-Object { -not $_.Synced -and -not $_.SkipSync })) {
            if ($Script:BulkStatusRows[$acc.Idx].SkipRequested) {
                $acc.SkipSync = $true
                $Script:BulkStatusRows[$acc.Idx].SyncSkipBtn.Visible = $false
                BulkLog $acc.Idx "Sync skipped -- license not assigned" ([System.Drawing.Color]::Orange)
                Update-BulkRow $acc.Idx "PARTIAL"
                Write-AuditLog "PARTIAL_CREATE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "SYNC_CANCELLED" "Bulk | Ticket=$($acc.Ticket)"; DBG "BULK_SYNC" "WARN" "$($acc.SAM): sync wait skipped by tech"
                $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: NOT ASSIGNED -- sync skipped"
            }
        }
        $pending = $createdAccounts | Where-Object { -not $_.Synced -and -not $_.SkipSync }
        if ($pending.Count -eq 0) { break }
    }

    while ($elapsed -lt $maxWait) {
        $pending = $createdAccounts | Where-Object { -not $_.Synced -and -not $_.SkipSync }
        if ($pending.Count -eq 0) { break }

        foreach ($acc in $pending) {
            $sidx = $acc.Idx
            if ($Script:BulkStatusRows[$sidx].SkipRequested) {
                $acc.SkipSync = $true
                $Script:BulkStatusRows[$sidx].SyncSkipBtn.Visible = $false
                BulkLog $sidx "Sync skipped" ([System.Drawing.Color]::Orange)
                Update-BulkRow $sidx "PARTIAL"
                Write-AuditLog "PARTIAL_CREATE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "SYNC_CANCELLED" "Bulk | Ticket=$($acc.Ticket)"; DBG "BULK_SYNC" "WARN" "$($acc.SAM): sync wait skipped by tech"
                $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: NOT ASSIGNED"
                continue
            }
            $Script:DebugSessionSAM = $acc.SAM
            $Script:AuditTicket = $acc.Ticket
            $fresh = Get-GraphToken -ScriptDir $bDir
            if ($fresh) { $bHdrs = $fresh }
            try {
                $poll = Get-SyncedEntraUser -Upn $acc.UPN -ImmutableId $acc.ImmutableId -Headers $bHdrs
            } catch {
                # Phase 1 plus a 45 minute sync wait can outlive the Graph token.
                $pollStatus = 0
                try { $pollStatus = [int]$_.Exception.Response.StatusCode } catch { $pollStatus = 0 }
                if ($pollStatus -eq 401) { $acc.AuthFails = 1 + [int]$acc.AuthFails }
                if ($pollStatus -eq 403 -or [int]$acc.AuthFails -ge 3) {
                    # A refused permission does not fix itself, so stop polling this row instead of waiting out 45 minutes.
                    $acc.SkipSync = $true
                    $Script:BulkStatusRows[$sidx].SyncSkipBtn.Visible = $false
                    BulkLog $sidx "Stopped: Graph refused the sync check (HTTP $pollStatus) -- license not assigned" ([System.Drawing.Color]::OrangeRed)
                    Update-BulkRow $sidx "PARTIAL"
                    Write-AuditLog "PARTIAL_CREATE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "SYNC_AUTH_FAILED" "Bulk | HTTP $pollStatus"
                    DBG "BULK_SYNC" "FAIL" "$($acc.SAM): Graph refused the sync check (HTTP $pollStatus): $(Get-GraphErrorDetail $_)"
                    $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: NOT ASSIGNED -- Graph refused the sync check"
                } elseif ($pollStatus -eq 401) {
                    DBG "BULK_SYNC" "WARN" "$($acc.SAM): Graph token rejected (401), requesting a new token"
                    $bHdrs = Get-GraphToken -ScriptDir $bDir -Force
                } else {
                    DBG "BULK_SYNC" "WARN" "$($acc.SAM): sync check failed (HTTP $pollStatus): $(Get-GraphErrorDetail $_)"
                }
                continue
            }
            $acc.AuthFails = 0
            if ($poll.State -eq 'NotYet') { continue }
            if ($poll.State -eq 'WrongObject') {
                if (-not $acc.WrongWarned) {
                    BulkLog $sidx "[!] $($poll.Detail) -- still waiting for our account" ([System.Drawing.Color]::Orange)
                    DBG "BULK_SYNC" "WARN" "$($acc.SAM): $($poll.Detail)"
                    $acc.WrongWarned = $true
                }
                continue
            }
            $candidate = $poll.User
            if ($poll.State -eq 'UpnChanged') {
                $acc.UpnNote = $poll.Detail
                BulkLog $sidx "[!] $($poll.Detail) -- UPN already taken in Entra, fix it after this run" ([System.Drawing.Color]::Orange)
                Write-AuditLog "SYNC" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "UPN_CHANGED" "Bulk | $($poll.Detail)"
                DBG "BULK_SYNC" "WARN" "$($acc.SAM): $($poll.Detail)"
            }
            $acc.Synced = $true
            $Script:BulkStatusRows[$sidx].SyncSkipBtn.Visible = $false
            BulkLog $sidx "Synced! Assigning license..." ([System.Drawing.Color]::LimeGreen)
            Update-BulkRow $sidx "P2: Licensing..."
            $fresh = Get-GraphToken -ScriptDir $bDir
            if ($fresh) { $bHdrs = $fresh }
            try {
                $groupName = $LicenseGroups[$acc.EmpType]
                $gFilter   = [Uri]::EscapeDataString("displayName eq '$groupName'")
                $gResp     = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$gFilter" -Headers $bHdrs -ErrorAction Stop
                $groupId   = @($gResp.value)[0].id
                if (-not $groupId) { throw "Group '$groupName' not found in Entra." }
                $memberBody= '{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/' + $candidate.id + '"}'
                try {
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" -Method POST -Headers $bHdrs -Body $memberBody -ErrorAction Stop
                } catch {
                    if ((Get-GraphErrorDetail $_) -notmatch 'already exist') { throw }
                }
                $acc.Licensed = $true
                Write-AuditLog "LICENSE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "OK" "Bulk | Group=$groupName"
                BulkLog $sidx "Licensed: $groupName" ([System.Drawing.Color]::LimeGreen)
            } catch {
                Write-AuditLog "LICENSE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "FAILED" "Bulk: $_"
                BulkLog $sidx "License failed: $_ -- assign manually" ([System.Drawing.Color]::OrangeRed)
                DBG "BULK_LICENSE" "FAIL" "$($acc.SAM): $_"
            }

            # AU assignment (after license, regardless of license outcome)
            $auResult = "Not assigned"
            $auOK = $false
            $bulkAuCodes = Get-AUCodes $acc.DivCode
            DBG "BULK_AU" "INFO" "Assigning AU(s): $($bulkAuCodes -join ', ') for $($acc.SAM) (user $($candidate.id))"
            $auRes = @(Add-UserToAUs -UserId $candidate.id -AUNames $bulkAuCodes -Headers $bHdrs)
            foreach ($ar in $auRes) {
                if ($ar.OK) {
                    BulkLog $sidx "AU assigned: $($ar.AU)" ([System.Drawing.Color]::LimeGreen)
                    Write-AuditLog "AU" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "OK" "Bulk | AU=$($ar.AU)"
                    DBG "BULK_AU" "OK" "$($ar.AU): $($ar.Detail)"
                } else {
                    BulkLog $sidx "AU '$($ar.AU)' failed -- assign manually: $($ar.Detail)" ([System.Drawing.Color]::Orange)
                    Write-AuditLog "AU" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "FAILED" "Bulk | AU=$($ar.AU) | $($ar.Detail)"
                    DBG "BULK_AU" "FAIL" "$($ar.AU): $($ar.Detail)"
                }
            }
            $auDone = @($auRes | Where-Object { $_.OK } | ForEach-Object { $_.AU })
            $auMiss = @($auRes | Where-Object { -not $_.OK } | ForEach-Object { $_.AU })
            $auOK   = ($auRes.Count -gt 0 -and $auMiss.Count -eq 0)
            if ($auDone.Count) { $auResult = $auDone -join ', ' }
            if ($auMiss.Count) { $auResult += " (NOT ASSIGNED: $($auMiss -join ', '))" }

            # COMPLETE needs the license, every AU, the UPN kept by sync, and no Phase 1 warnings.
            $rowNotes = @()
            if ($acc.UpnNote) { $rowNotes += $acc.UpnNote }
            $rowNotes += @($acc.Warnings)
            if ($acc.Licensed -and $auOK -and $rowNotes.Count -eq 0) {
                BulkLog $sidx "COMPLETE | SAM: $($acc.SAM) | Pwd: $($acc.PlainPass)" ([System.Drawing.Color]::Cyan)
                Update-BulkRow $sidx "COMPLETE"
            } elseif ($acc.Licensed -and $auOK) {
                BulkLog $sidx "Licensed, follow-up needed: $($rowNotes -join '; ') | SAM: $($acc.SAM) | Pwd: $($acc.PlainPass)" ([System.Drawing.Color]::Cyan)
                Update-BulkRow $sidx "PARTIAL"
            } elseif ($acc.Licensed) {
                BulkLog $sidx "Licensed but AU not assigned | SAM: $($acc.SAM) | Pwd: $($acc.PlainPass)" ([System.Drawing.Color]::Cyan)
                Update-BulkRow $sidx "PARTIAL"
            } else {
                Update-BulkRow $sidx "PARTIAL"
            }
            $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: $(if($acc.Licensed){$LicenseGroups[$acc.EmpType]}else{'NOT ASSIGNED'})`r`nAdmin Unit: $auResult$(if ($rowNotes.Count) { "`r`nFollow-up: $($rowNotes -join '; ')" })"
            DBG "BULK_COMPLETE" "OK" "$($acc.SAM) done (Licensed=$($acc.Licensed) AU=$auOK)"
        }

        $stillPending = $createdAccounts | Where-Object { -not $_.Synced -and -not $_.SkipSync }
        if ($stillPending.Count -eq 0) { break }

        $elapsed += $interval
        if ($elapsed -ge $maxWait) {
            foreach ($acc in $stillPending) {
                BulkLog $acc.Idx "Sync timed out -- license not assigned" ([System.Drawing.Color]::OrangeRed)
                Update-BulkRow $acc.Idx "PARTIAL"
                Write-AuditLog "PARTIAL_CREATE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "SYNC_TIMEOUT" "Bulk | Ticket=$($acc.Ticket)"
                DBG "BULK_SYNC" "FAIL" "$($acc.SAM): not found in Entra after 45 min"
                $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: NOT ASSIGNED -- sync timed out"
            }
            break
        }

        for ($s = $interval; $s -ge 0; $s--) {
            $LblBulkStatus.Text = "Phase 2: Next check in $s sec ($([math]::Round($elapsed/60))m elapsed)..."
            Start-Sleep -Seconds 1
            [System.Windows.Forms.Application]::DoEvents()
            foreach ($acc in ($createdAccounts | Where-Object { -not $_.Synced -and -not $_.SkipSync })) {
                if ($Script:BulkStatusRows[$acc.Idx].SkipRequested) {
                    $acc.SkipSync = $true
                    $Script:BulkStatusRows[$acc.Idx].SyncSkipBtn.Visible = $false
                    BulkLog $acc.Idx "Sync skipped" ([System.Drawing.Color]::Orange)
                    Update-BulkRow $acc.Idx "PARTIAL"
                    Write-AuditLog "PARTIAL_CREATE" $acc.SAM $acc.UPN $acc.DivCode $acc.EmpType "SYNC_CANCELLED" "Bulk | Ticket=$($acc.Ticket)"; DBG "BULK_SYNC" "WARN" "$($acc.SAM): sync wait skipped by tech"
                    $Script:BulkResults += "SAM Account Name: $($acc.DomainShort)\$($acc.SAM)`r`nUPN: $($acc.UPN)`r`nEmail: $($acc.Email)`r`nTemp Password: $($acc.PlainPass)`r`nLicense Assigned: NOT ASSIGNED"
                }
            }
        }
    }

    # Count final statuses from the status board (reflects both license AND AU)
    $done = 0; $partial = 0
    foreach ($acc in $createdAccounts) {
        $phase = $Script:BulkStatusRows[$acc.Idx].StatusLbl.Text
        if ($phase -eq "COMPLETE") { $done++ } else { $partial++ }
    }
    $notCreated = $queue.Count - @($createdAccounts).Count
    $LblBulkStatus.Text      = "Complete. $done fully provisioned, $partial need follow-up, $notCreated not created."
    $LblBulkStatus.ForeColor = if ($partial -eq 0 -and $notCreated -eq 0) { $ColorAccent } else { $ColorGold }
    $BtnBulkCopyAll.Enabled  = $true
    $BtnBulkPreview.Enabled  = $true
    $Script:BulkRunning      = $false
    $Script:AuditTicket      = $null
    try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
})

# ============================================================
#  TAB 6 - CREATE ACCOUNT
# ============================================================
$TabCreate           = New-Object System.Windows.Forms.TabPage
$TabCreate.Text      = "  Create Account  "
$TabCreate.BackColor = $ColorGray
$Tabs.Controls.Add($TabCreate)

function New-Lbl($t,$x,$y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $t; $l.Location = (New-Object System.Drawing.Point -Property @{X=$x; Y=$y}); $l.AutoSize = $true
    return $l
}
function New-Txt($x,$y,$w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Size = (New-Object System.Drawing.Size -Property @{Width=$w; Height=24}); $t.Location = (New-Object System.Drawing.Point -Property @{X=$x; Y=$y})
    return $t
}

# Left form panel - Row 1: Name fields
$TabCreate.Controls.Add((New-Lbl "First Name *"    10  14))
$TxtFirst = New-Txt 10 32 170; $TabCreate.Controls.Add($TxtFirst)

$TabCreate.Controls.Add((New-Lbl "Mid. Initial"   188  14))
$TxtMid = New-Txt 188 32 60; $TxtMid.MaxLength = 1; $TabCreate.Controls.Add($TxtMid)

$TabCreate.Controls.Add((New-Lbl "Last Name *"    256  14))
$TxtLast = New-Txt 256 32 195; $TabCreate.Controls.Add($TxtLast)

# Row 2: Division
$TabCreate.Controls.Add((New-Lbl "Division *"      10  66))
$CboDivision                 = New-Object System.Windows.Forms.ComboBox
$CboDivision.Size            = (New-Object System.Drawing.Size -Property @{Width=440; Height=24})
$CboDivision.Location        = (New-Object System.Drawing.Point -Property @{X=10; Y=84})
$CboDivision.DropDownStyle   = "DropDown"
$CboDivision.AutoCompleteMode   = "SuggestAppend"
$CboDivision.AutoCompleteSource = "ListItems"
$CboDivision.Items.Add("-- Select Division --") | Out-Null
foreach ($div in $DivisionMap) { $CboDivision.Items.Add("$($div.Code) - $($div.Name)") | Out-Null }
$CboDivision.SelectedIndex   = 0
$TabCreate.Controls.Add($CboDivision)

# Row 3: Target Domain (auto-populated from division, read-only)
$TabCreate.Controls.Add((New-Lbl "Target Domain (auto-filled from division)"   10  116))
$TxtDomain          = New-Txt 10 134 440
$TxtDomain.ReadOnly = $true
$TxtDomain.BackColor = $ColorPale
$TabCreate.Controls.Add($TxtDomain)

# Auto-populate domain when division changes
$CboDivision.Add_SelectedIndexChanged({
    $sel = $CboDivision.SelectedItem
    if (-not $sel -or $sel -eq "-- Select Division --") { $TxtDomain.Text = ""; return }
    $code = Get-DivCodeFromItem $sel
    $div  = $DivisionMap | Where-Object { $_.Code -eq $code }
    if ($div) { $TxtDomain.Text = $div.Domain }
})

# Row 4: Employee Type + Job Title
$TabCreate.Controls.Add((New-Lbl "Employee Type *"  10  166))
$CboEmpType                = New-Object System.Windows.Forms.ComboBox
$CboEmpType.Size           = (New-Object System.Drawing.Size -Property @{Width=210; Height=24})
$CboEmpType.Location       = (New-Object System.Drawing.Point -Property @{X=10; Y=184})
$CboEmpType.DropDownStyle  = "DropDownList"
$CboEmpType.Items.Add("-- Select Type --")    | Out-Null
$CboEmpType.Items.Add("Frontline Worker")     | Out-Null
$CboEmpType.Items.Add("Knowledge Worker")     | Out-Null
$CboEmpType.SelectedIndex  = 0
$TabCreate.Controls.Add($CboEmpType)

$TabCreate.Controls.Add((New-Lbl "Job Title (optional)" 228 166))
$TxtTitle = New-Txt 228 184 222; $TabCreate.Controls.Add($TxtTitle)
$TxtTitle.Add_Leave({
    if ($TxtTitle.Text.Trim()) {
        $TxtTitle.Text = (Get-Culture).TextInfo.ToTitleCase($TxtTitle.Text.Trim().ToLower())
    }
})

# Row 5: Category ID + Account Type (custom AD attributes, both mandatory)
$TabCreate.Controls.Add((New-Lbl "Category *" 10 216))
$CboCategory                = New-Object System.Windows.Forms.ComboBox
$CboCategory.Size           = (New-Object System.Drawing.Size -Property @{Width=210; Height=24})
$CboCategory.Location       = (New-Object System.Drawing.Point -Property @{X=10; Y=234})
$CboCategory.DropDownStyle  = "DropDownList"
foreach ($c in $CategoryList) { $CboCategory.Items.Add($c) | Out-Null }
$CboCategory.SelectedIndex  = 0   # defaults to "1 - General"
$TabCreate.Controls.Add($CboCategory)

$TabCreate.Controls.Add((New-Lbl "Account Type *" 228 216))
$CboAcctType                = New-Object System.Windows.Forms.ComboBox
$CboAcctType.Size           = (New-Object System.Drawing.Size -Property @{Width=222; Height=24})
$CboAcctType.Location       = (New-Object System.Drawing.Point -Property @{X=228; Y=234})
$CboAcctType.DropDownStyle  = "DropDownList"
foreach ($a in $AccountTypeList) { $CboAcctType.Items.Add($a) | Out-Null }
$CboAcctType.SelectedIndex  = 0   # defaults to "Primary"
$TabCreate.Controls.Add($CboAcctType)

# Row 6: Ticket Number (optional)
$TabCreate.Controls.Add((New-Lbl "Ticket # (optional)" 10 258))
$TxtTicket = New-Txt 10 276 210
$TabCreate.Controls.Add($TxtTicket)

# Debug toggle checkbox
$ChkDebug           = New-Object System.Windows.Forms.CheckBox
$ChkDebug.Text      = "Show debug in Creation Log"
$ChkDebug.Location  = (New-Object System.Drawing.Point -Property @{X=228; Y=278})
$ChkDebug.AutoSize  = $true
$ChkDebug.ForeColor = $ColorDark
$ChkDebug.Checked   = $false
$TabCreate.Controls.Add($ChkDebug)

# Progress Log
$GrpLog           = New-Object System.Windows.Forms.GroupBox
$GrpLog.Text      = "Creation Log"
$GrpLog.Size      = (New-Object System.Drawing.Size -Property @{Width=440; Height=370})
$GrpLog.Location  = (New-Object System.Drawing.Point -Property @{X=462; Y=12})
$GrpLog.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpLog.ForeColor = $ColorDark
$TabCreate.Controls.Add($GrpLog)

$RichLog               = New-Object System.Windows.Forms.RichTextBox
$RichLog.Size          = (New-Object System.Drawing.Size -Property @{Width=420; Height=340})
$RichLog.Location      = (New-Object System.Drawing.Point -Property @{X=10; Y=22})
$RichLog.BackColor     = [System.Drawing.Color]::Black
$RichLog.ForeColor     = [System.Drawing.Color]::LimeGreen
$RichLog.Font          = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$RichLog.ReadOnly      = $true
$RichLog.BorderStyle   = "None"
$GrpLog.Controls.Add($RichLog)

# Copy Account Information button - sits below the log, grayed until creation succeeds
$BtnCopyAcct           = New-Object System.Windows.Forms.Button
$BtnCopyAcct.Text      = "Copy Account Information for Ticket"
$BtnCopyAcct.Size      = (New-Object System.Drawing.Size -Property @{Width=280; Height=28})
$BtnCopyAcct.Location  = (New-Object System.Drawing.Point -Property @{X=462; Y=386})
$BtnCopyAcct.BackColor = $ColorLight
$BtnCopyAcct.ForeColor = $ColorDark
$BtnCopyAcct.FlatStyle = "Flat"
$BtnCopyAcct.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnCopyAcct.Enabled   = $false
$TabCreate.Controls.Add($BtnCopyAcct)

$BtnCancelSync           = New-Object System.Windows.Forms.Button
$BtnCancelSync.Text      = "Cancel Sync Wait"
$BtnCancelSync.Size      = (New-Object System.Drawing.Size -Property @{Width=152; Height=28})
$BtnCancelSync.Location  = (New-Object System.Drawing.Point -Property @{X=750; Y=386})
$BtnCancelSync.BackColor = $ColorRGB_180_60_60
$BtnCancelSync.ForeColor = [System.Drawing.Color]::White
$BtnCancelSync.FlatStyle = "Flat"
$BtnCancelSync.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnCancelSync.Enabled   = $false
$TabCreate.Controls.Add($BtnCancelSync)

$Script:LastAccountDetails = $null

# Account Preview
$GrpPreview           = New-Object System.Windows.Forms.GroupBox
$GrpPreview.Text      = "Account Preview"
$GrpPreview.Size      = (New-Object System.Drawing.Size -Property @{Width=440; Height=158})
$GrpPreview.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=326})
$GrpPreview.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$GrpPreview.ForeColor = $ColorDark
$TabCreate.Controls.Add($GrpPreview)

$TxtPreview            = New-Object System.Windows.Forms.RichTextBox
$TxtPreview.Size       = (New-Object System.Drawing.Size -Property @{Width=420; Height=130})
$TxtPreview.Location   = (New-Object System.Drawing.Point -Property @{X=10; Y=20})
$TxtPreview.BackColor  = $ColorPale
$TxtPreview.ReadOnly   = $true
$TxtPreview.Font       = (New-Object System.Drawing.Font -ArgumentList "Consolas",([single]8),([System.Drawing.FontStyle]::Regular))
$TxtPreview.BorderStyle = "None"
$GrpPreview.Controls.Add($TxtPreview)

# Buttons
$BtnPreview           = New-Object System.Windows.Forms.Button
$BtnPreview.Text      = "Preview"
$BtnPreview.Size      = (New-Object System.Drawing.Size -Property @{Width=110; Height=32})
$BtnPreview.Location  = (New-Object System.Drawing.Point -Property @{X=10; Y=496})
$BtnPreview.BackColor = $ColorLight
$BtnPreview.ForeColor = $ColorDark
$BtnPreview.FlatStyle = "Flat"
$BtnPreview.Font      = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$TabCreate.Controls.Add($BtnPreview)

$BtnCreate            = New-Object System.Windows.Forms.Button
$BtnCreate.Text       = "Create Account"
$BtnCreate.Size       = (New-Object System.Drawing.Size -Property @{Width=140; Height=32})
$BtnCreate.Location   = (New-Object System.Drawing.Point -Property @{X=128; Y=496})
$BtnCreate.BackColor  = $ColorAccent
$BtnCreate.ForeColor  = $ColorWhite
$BtnCreate.FlatStyle  = "Flat"
$BtnCreate.Font       = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Bold))
$BtnCreate.Enabled    = $false
$TabCreate.Controls.Add($BtnCreate)

$BtnClear             = New-Object System.Windows.Forms.Button
$BtnClear.Text        = "Clear"
$BtnClear.Size        = (New-Object System.Drawing.Size -Property @{Width=80; Height=32})
$BtnClear.Location    = (New-Object System.Drawing.Point -Property @{X=276; Y=496})
$BtnClear.BackColor   = $ColorGray
$BtnClear.ForeColor   = $ColorDark
$BtnClear.FlatStyle   = "Flat"
$TabCreate.Controls.Add($BtnClear)

$LblCreateStatus          = New-Object System.Windows.Forms.Label
$LblCreateStatus.Location = (New-Object System.Drawing.Point -Property @{X=10; Y=538})
$LblCreateStatus.Size     = (New-Object System.Drawing.Size -Property @{Width=440; Height=50})
$LblCreateStatus.Font     = (New-Object System.Drawing.Font -ArgumentList "Segoe UI",([single]9),([System.Drawing.FontStyle]::Regular))
$TabCreate.Controls.Add($LblCreateStatus)

$Script:ResolvedSAM       = $null
$Script:ResolvedCandidate = $null

# Any edit after Preview invalidates its result, so Create waits for a new Preview.
foreach ($ctl in @($TxtFirst, $TxtMid, $TxtLast)) { $ctl.Add_TextChanged({ $BtnCreate.Enabled = $false }) }
foreach ($ctl in @($CboDivision, $CboEmpType)) { $ctl.Add_SelectedIndexChanged({ $BtnCreate.Enabled = $false }) }

# ---- PREVIEW CLICK -----------------------------------------
$BtnPreview.Add_Click({
    if ($Script:BulkRunning) {
        [System.Windows.Forms.MessageBox]::Show("A bulk run is in progress. Wait for it to finish first.", "Run in progress") | Out-Null
        return
    }
    $Script:DebugSessionSAM  = "(preview)"
    $Script:DebugSessionName = "$($TxtFirst.Text.Trim()) $($TxtLast.Text.Trim())"
    $first  = Format-ProperCase $TxtFirst.Text.Trim()
    $last   = Format-ProperCase $TxtLast.Text.Trim()
    $mid    = Format-ProperCase $TxtMid.Text.Trim()
    $domain = $TxtDomain.Text.Trim()
    $emp    = $CboEmpType.SelectedItem

    if (-not $first -or -not $last) {
        [System.Windows.Forms.MessageBox]::Show("First Name and Last Name are required.","Missing Fields")
        return
    }
    if ($CboDivision.SelectedIndex -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Division.","Missing Fields")
        return
    }
    if ($CboEmpType.SelectedIndex -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select an Employee Type.","Missing Fields")
        return
    }
    $nameProblems = @(Get-NameInputProblem -First $first -Middle $mid -Last $last)
    if ($nameProblems.Count) {
        [System.Windows.Forms.MessageBox]::Show(($nameProblems -join "`n"), "Check the name") | Out-Null
        $BtnCreate.Enabled = $false
        return
    }

    $LblCreateStatus.Text      = "Checking availability (AD + Entra)..."
    $LblCreateStatus.ForeColor = $ColorGold
    $Form.Refresh()

    $divCode     = Get-DivCodeFromItem $CboDivision.SelectedItem
    $emailDomain = Get-EmailDomain $divCode
    $ou          = Get-OUPath -Domain $domain
    $catId       = $CboCategory.SelectedItem.ToString().Split(" ")[0]
    $acctType    = $CboAcctType.SelectedItem.ToString()

    # Acquire Graph token for preview checks (Graph-only -- no Exchange session at preview)
    $pvwScriptDir = $PSScriptRoot
    if (-not $pvwScriptDir) { $pvwScriptDir = "C:\ADProvisioning" }
    $pvwHeaders = Get-GraphToken -ScriptDir $pvwScriptDir

    # No Exchange session at preview; the Exchange alias check runs again at create time.
    $resolved = Resolve-AccountCandidate `
        -First $first -Middle $mid -Last $last `
        -Domain $domain -EmailDomain $emailDomain `
        -GraphHeaders $pvwHeaders

    if ($resolved.Status -ne 'Resolved') {
        $LblCreateStatus.Text      = if ($resolved.Status -eq 'Unverified') { "[X] Could not complete duplicate checks: $($resolved.Reason)" } else { "[X] $($resolved.Reason)" }
        $LblCreateStatus.ForeColor = $ColorRed
        $BtnCreate.Enabled         = $false
        return
    }

    $sam      = $resolved.SAM
    $upn      = $resolved.UPN
    $primary  = $resolved.PrimarySmtp
    $dispName = "$first $last - $divCode"
    $cnName   = $dispName

    # Check if CN already exists in target OU (same name, same division)
    # CN may get (sam) suffix but DisplayName stays clean
    $cnCollision = $false
    try {
        $existingCN = Get-ADObject -LDAPFilter "(name=$(ConvertTo-LdapFilterValue $dispName))" -SearchBase $ou -Server $domain -ErrorAction Stop
        if ($existingCN) {
            $cnName = "$first $last - $divCode ($sam)"
            $cnCollision = $true
        }
    } catch { DBG "CN_CHECK" "WARN" "Could not check for a name collision in ${ou}: $_" }

    $lengthProblems = @(Get-NameLengthProblem -CnName $cnName -LocalPart $resolved.LocalPart)
    if ($lengthProblems.Count) {
        $LblCreateStatus.Text      = "[X] $($lengthProblems -join '; '). Shorten the name in the ticket and Preview again."
        $LblCreateStatus.ForeColor = $ColorRed
        $BtnCreate.Enabled         = $false
        return
    }
    $targetProblems = @(Get-EntraTargetProblem -AUNames (Get-AUCodes $divCode) -GroupName $LicenseGroups[$emp] -Headers $pvwHeaders)
    foreach ($tp in $targetProblems) { DBG "PREVIEW_TARGETS" "WARN" $tp }

    $Script:ResolvedSAM        = $sam
    $Script:ResolvedCandidate  = $resolved

    # Build preview text with full resolved picture and any conflict warnings
    $aliasLine   = ""
    $aliasStatus = ""
    if ($emailDomain -ne "contoso.com" -and $resolved.TenantAlias) {
        $defaultAlias = "$($resolved.LocalPart)@contoso.com"
        if ($resolved.TenantAlias -ne $defaultAlias) {
            $aliasLine   = "`ntenant alias  : $($resolved.TenantAlias)"
            $aliasStatus = "`n[!] $defaultAlias was taken -- resolved to $($resolved.TenantAlias)"
        } else {
            $aliasLine = "`ntenant alias  : $($resolved.TenantAlias)  [free]"
        }
    }
    $truncNote  = if ($resolved.Truncated)  { "`n[!] SAM shortened to the 20 character limit; UPN and email keep the full name" } else { "" }
    $attemptNote = if ($resolved.Attempt -gt 1) { "`n[!] Resolved on attempt $($resolved.Attempt) -- earlier candidates were taken" } else { "" }
    $cnNote     = if ($cnCollision) { "`n[!] CN collision in OU -- ADUC name: $cnName  (display name stays: $dispName)" } else { "" }
    $targetNote = if ($targetProblems.Count) { "`n[!] " + ($targetProblems -join "`n[!] ") + "`n    Fix the name in the tool's division map; assign by hand until then." } else { "" }

    $TxtPreview.Text = "Display Name : $dispName`nSAM          : $sam`nUPN / Login  : $upn`nPrimary SMTP : $primary$aliasLine`nRouting      : $($resolved.Routing)`nDomain       : $domain`nOU           : $ou`nCategory     : $catId`nAccount Type : $acctType`nLicense      : $($LicenseGroups[$emp])$aliasStatus$truncNote$attemptNote$cnNote$targetNote"

    if ($targetProblems.Count) {
        $LblCreateStatus.Text      = "[+] '$upn' available, but see the warnings in the preview before you create."
        $LblCreateStatus.ForeColor = $ColorGold
    } else {
        $LblCreateStatus.Text      = "[+] '$upn' available -- all checks passed. Click Create Account to proceed."
        $LblCreateStatus.ForeColor = $ColorAccent
    }
    $BtnCreate.Enabled         = $true
})

# ---- CREATE CLICK ------------------------------------------
# The body is a separate scriptblock so one try/finally clears the run flag on every early return.
$Script:SingleRunning = $false
$Script:SingleCreateBody = {
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Create account '$($Script:ResolvedSAM)' now?`n`nThis will:`n- Create the AD account`n- Enable the O365 mailbox`n- Assign the M365 license group",
        "Confirm Creation",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne "Yes") { return }

    $BtnCreate.Enabled  = $false
    $BtnPreview.Enabled = $false
    $RichLog.Clear()

    $first    = Format-ProperCase $TxtFirst.Text.Trim()
    $last     = Format-ProperCase $TxtLast.Text.Trim()
    $mid      = Format-ProperCase $TxtMid.Text.Trim()
    $sam      = $Script:ResolvedSAM
    $domain   = $TxtDomain.Text.Trim()
    $ticket   = $TxtTicket.Text.Trim()
    $divCode  = Get-DivCodeFromItem $CboDivision.SelectedItem
    $emailDomain = Get-EmailDomain $divCode
    $upn      = $Script:ResolvedCandidate.UPN
    $dispName = "$first $last - $divCode"
    $empType  = $CboEmpType.SelectedItem.ToString()
    $catId    = $CboCategory.SelectedItem.ToString().Split(" ")[0]
    $acctType = $CboAcctType.SelectedItem.ToString()
    $Script:ShowDebug = $ChkDebug.Checked
    if (-not $domain) {
        Write-Log $RichLog "[X] No division selected -- choose a division and run Preview again." $ColorRed
        DBG "DOMAIN_VALIDATE" "FAILED" "No division selected"
        $BtnPreview.Enabled = $true; return
    }
    $ou       = Get-OUPath -Domain $domain

    # Same-name object in the OU: the CN gets " (sam)" after pre-flight has fixed the SAM.
    # cnName = AD object name (unique per OU), dispName = display name (clean)
    $cnName = $dispName
    $cnCollision = $false
    try {
        $existingCN = Get-ADObject -LDAPFilter "(name=$(ConvertTo-LdapFilterValue $dispName))" -SearchBase $ou -Server $domain -ErrorAction Stop
        if ($existingCN) { $cnCollision = $true }
    } catch {
        Write-Log $RichLog "  [!] Could not check for a name collision in $ou -- creation may fail if the name already exists: $_" $ColorGold
        DBG "CN_CHECK" "WARN" "$_"
    }

    $plainPass = New-TempPassword
    $secPass   = ConvertTo-SecureString $plainPass -AsPlainText -Force
    $emailAddr = $upn
    $upnNote   = $null
    $followUpItems = @()

    # Init debug session tracking
    $Script:DebugSessionSAM    = $sam
    $Script:DebugSessionName   = $dispName
    $aliasGraphHeaders         = $null
    $Script:NewAcctImmutableId = $null

    # B. Domain validation
    Write-Log $RichLog "Validating domain $domain..." $ColorLight
    try {
        $null = Get-ADDomain -Identity $domain -Server $domain -ErrorAction Stop
        Write-Log $RichLog "  Domain validated: $domain" ([System.Drawing.Color]::LimeGreen)
        DBG "DOMAIN_VALIDATE" "OK" $domain
    } catch {
        Write-Log $RichLog "[X] Domain '$domain' is not reachable: $_" $ColorRed
        DBG "DOMAIN_VALIDATE" "FAILED" "$_"
        Write-AuditLog "CREATE" $sam $upn $divCode $empType "DOMAIN_FAILED" "Domain unreachable: $domain"
        $BtnPreview.Enabled = $true; return
    }

    # ---- PRE-FLIGHT: Full availability check before ANY AD write ----
    # Opens a brief Exchange session for Get-Recipient, acquires Graph token,
    # runs Resolve-AccountCandidate against all sources. Hard blocks if any
    # candidate is unavailable. Pre-resolved values used throughout creation.
    Write-Log $RichLog "" $null
    Write-Log $RichLog "PRE-FLIGHT: Verifying availability across AD, Exchange, and Entra..." $ColorLight

    # Acquire Graph token
    $pfScriptDir = $PSScriptRoot
    if (-not $pfScriptDir) { $pfScriptDir = "C:\ADProvisioning" }
    $pfHeaders = Get-GraphToken -ScriptDir $pfScriptDir
    if ($pfHeaders) {
        Write-Log $RichLog "  Graph token acquired." ([System.Drawing.Color]::LimeGreen)
        DBG "PREFLIGHT_TOKEN" "OK" "Graph token acquired"
    } else {
        Write-Log $RichLog "  [!] Graph token unavailable -- Entra cannot be checked, pre-flight will stop." $ColorGold
        DBG "PREFLIGHT_TOKEN" "WARN" "Graph token unavailable"
    }

    # Open brief Exchange session for Get-Recipient check
    $pfSession  = $null
    $pfExchOK   = $false
    try {
        $pfSession = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ExchangeServer -Authentication Kerberos -ErrorAction Stop
        Import-PSSession $pfSession -DisableNameChecking -AllowClobber | Out-Null
        $pfExchOK = $true
        Write-Log $RichLog "  Exchange session opened for pre-flight check." ([System.Drawing.Color]::LimeGreen)
        DBG "PREFLIGHT_EXCH" "OK" "Exchange session opened"
    } catch {
        Write-Log $RichLog "  [!] Exchange session unavailable for pre-flight -- alias still checked in AD (Global Catalog) and Entra." $ColorGold
        DBG "PREFLIGHT_EXCH" "WARN" "Exchange session failed: $_"
    }

    # Run full resolution
    $pfResolved = Resolve-AccountCandidate `
        -First $first -Middle $mid -Last $last `
        -Domain $domain -EmailDomain $emailDomain `
        -GraphHeaders $pfHeaders `
        -ExchSessionActive:$pfExchOK

    # Close pre-flight Exchange session immediately -- STEP 2 opens its own
    if ($pfSession) { Remove-PSSession $pfSession -ErrorAction SilentlyContinue }

    if ($pfResolved.Status -ne 'Resolved') {
        if ($pfResolved.Status -eq 'Unverified') {
            Write-Log $RichLog "[X] PRE-FLIGHT STOPPED: a duplicate check could not be completed." $ColorRed
            Write-Log $RichLog "  $($pfResolved.Reason)" $ColorRed
            Write-Log $RichLog "  Nothing was created. Fix the connection and run Preview again." $ColorRed
        } else {
            Write-Log $RichLog "[X] PRE-FLIGHT FAILED: $($pfResolved.Reason)" $ColorRed
            Write-Log $RichLog "  Cannot create account -- resolve conflicts manually before proceeding." $ColorRed
        }
        Write-AuditLog "CREATE" $sam $upn $divCode $empType "PREFLIGHT_BLOCKED" "$($pfResolved.Status): $($pfResolved.Reason)"
        DBG "PREFLIGHT" "FAIL" "Hard block -- $($pfResolved.Status): $($pfResolved.Reason)"
        $BtnPreview.Enabled = $true; return
    }

    # Use pre-flight resolved values for the rest of the flow
    $sam         = $pfResolved.SAM
    $upn         = $pfResolved.UPN
    $tenantAlias  = $pfResolved.TenantAlias
    $routing     = $pfResolved.Routing
    $Script:DebugSessionSAM   = $sam
    $Script:DebugSessionName  = "$first $last - $divCode"
    $aliasGraphHeaders        = $pfHeaders   # reuse in STEP 3
    if ($cnCollision) {
        $cnName = "$first $last - $divCode ($sam)"
        Write-Log $RichLog "  [!] CN collision in OU -- ADUC name: $cnName (display name stays clean)" $ColorGold
    }

    if ($pfResolved.Attempt -gt 1) {
        Write-Log $RichLog "  [!] Original candidate taken -- resolved to attempt $($pfResolved.Attempt): UPN=$upn" $ColorGold
    }
    if ($pfResolved.Truncated) {
        Write-Log $RichLog "  [!] SAM shortened to the 20 character limit: $sam (UPN and email keep the full name)" $ColorGold
    }
    Write-Log $RichLog "  Pre-flight passed: SAM=$sam UPN=$upn$(if($tenantAlias){' Alias='+$tenantAlias}else{''})" ([System.Drawing.Color]::LimeGreen)
    Write-Log $RichLog "  Routing: $routing" ([System.Drawing.Color]::LimeGreen)
    DBG "PREFLIGHT" "OK" "SAM=$sam UPN=$upn Alias=$tenantAlias Routing=$routing Attempt=$($pfResolved.Attempt)"
    # ---- END PRE-FLIGHT ----

    # Resolve the PDC -- pin everything to it so Exchange writes and AD reads
    # are always on the same DC with no replication gap.
    try {
        $targetPDC = (Get-ADDomain -Identity $domain -Server $domain).PDCEmulator
    } catch {
        $targetPDC = $domain
    }
    Write-Log $RichLog "  Target DC (pinned): $targetPDC" ([System.Drawing.Color]::Yellow)
    # Read by the click handler if the run stops on an unexpected error.
    $Script:RunSam = $sam; $Script:RunUpn = $upn; $Script:RunPdc = $targetPDC; $Script:RunDiv = $divCode; $Script:RunEmp = $empType

    # STEP 1+2 - Create account AND enable mailbox atomically via New-RemoteMailbox.
    # This runs inside an Exchange PSSession and creates the AD object AND stamps all
    # mailbox attributes (targetAddress, proxyAddresses, msExch* values, legacyExchangeDN)
    # in one transaction -- the same path as creating via the Exchange Admin Center manually.
    # Eliminates the two-step create-then-enable race condition entirely.
    Write-Log $RichLog "" $null
    Write-Log $RichLog "STEP 1+2: Creating account and enabling mailbox via Exchange..." $ColorLight
    $exchSession = $null
    $exchReady   = $false
    $nothingCreated = $false
    $newMailbox  = $null
    $step2OK     = $false
    try {
        $exchSession = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri $ExchangeServer -Authentication Kerberos -ErrorAction Stop
        Import-PSSession $exchSession -DisableNameChecking -AllowClobber | Out-Null
        $exchReady = $true
        Write-Log $RichLog "  Exchange session established." ([System.Drawing.Color]::LimeGreen)
        DBG "EXCH_SESSION" "OK" "Exchange session established"

        $routing     = $pfResolved.Routing
        $primarySmtp = $pfResolved.PrimarySmtp
        $tenantAlias  = $pfResolved.TenantAlias

        # Final alias race-condition check
        if ($emailDomain -ne "contoso.com" -and $tenantAlias) {
            $finalCheck = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$tenantAlias'" -ErrorAction SilentlyContinue
            if ($finalCheck) {
                Write-Log $RichLog "  [!] Race condition: $tenantAlias taken since pre-flight -- assign manually." $ColorGold
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "RACE_CONDITION" "$tenantAlias taken post-preflight"
                $tenantAlias = $null
            } else {
                DBG "ALIAS_FINAL_CHECK" "OK" "$tenantAlias still free"
            }
        }

        $nrmParams = Get-NewRemoteMailboxParameter -Resolved $pfResolved -CnName $cnName -DispName $dispName `
            -First $first -Last $last -Ou $ou -Pdc $targetPDC -Password $secPass

        DBG "NEW_REMOTE_MAILBOX" "INFO" "Calling New-RemoteMailbox: SAM=$sam UPN=$upn Routing=$routing OU=$ou"
        $newMailbox = New-RemoteMailbox @nrmParams -ErrorAction Stop

        Write-Log $RichLog "  Account and mailbox created: $dispName" ([System.Drawing.Color]::LimeGreen)
        Write-Log $RichLog "  SAM: $sam  |  UPN: $upn" ([System.Drawing.Color]::LimeGreen)
        Write-Log $RichLog "  Primary: $($newMailbox.PrimarySmtpAddress)  |  Routing: $routing" ([System.Drawing.Color]::LimeGreen)
        DBG "NEW_REMOTE_MAILBOX" "OK" "Created. PrimarySmtp=$($newMailbox.PrimarySmtpAddress) RecipientType=$($newMailbox.RecipientTypeDetails)"
        Write-AuditLog "CREATE" $sam $upn $divCode $empType "AD_OK" "New-RemoteMailbox on $targetPDC | OU=$ou | Ticket=$ticket"
        Write-AuditLog "MAILBOX" $sam $upn $divCode $empType "OK" "Primary=$($newMailbox.PrimarySmtpAddress) | Routing=$routing"
        $emailAddr = "$($newMailbox.PrimarySmtpAddress)"
        if ($emailAddr -ine $primarySmtp) {
            Write-Log $RichLog "  [!] Primary email is $emailAddr, expected $primarySmtp -- correcting..." $ColorGold
            DBG "PRIMARY_SMTP" "WARN" "Primary $emailAddr does not match $primarySmtp"
            try {
                Set-RemoteMailbox -Identity $sam -PrimarySmtpAddress $primarySmtp -EmailAddressPolicyEnabled $false `
                    -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                $emailAddr = $primarySmtp
                Write-AuditLog "MAILBOX" $sam $upn $divCode $empType "OK" "Primary corrected to $primarySmtp"
            } catch {
                Write-Log $RichLog "  [!] Could not correct primary email: $_" $ColorGold
                Write-AuditLog "MAILBOX" $sam $upn $divCode $empType "WARN" "Primary is $emailAddr, expected $primarySmtp : $_"
                $followUpItems += "Primary email is $emailAddr, should be $primarySmtp"
            }
        }

        # Stamp custom AD attributes using ADSI (direct LDAP).
        # ADSI avoids the "no settings modified" warning popup that ps2exe
        # generates from Set-ADUser on custom schema attributes.
        # Title uses Set-ADUser since it's a standard attribute (no warning).
        Wait-UiDelay -Seconds 3
        $titleVal = $TxtTitle.Text.Trim()
        try {
            $adObj = Get-ADUser -Identity $sam -Server $targetPDC -ErrorAction Stop
            $adsi = [ADSI]"LDAP://$targetPDC/$($adObj.DistinguishedName)"
            $adsi.Put("orgCategoryID", $catId)
            $adsi.Put("orgAccountType", $acctType)
            $adsi.SetInfo()
            Write-Log $RichLog "  Custom attributes set: Category=$catId  Type=$acctType" ([System.Drawing.Color]::LimeGreen)
            DBG "CUSTOM_ATTRS" "OK" "orgCategoryID=$catId orgAccountType=$acctType (ADSI)"
            Write-AuditLog "ATTRS" $sam $upn $divCode $empType "OK" "orgCategoryID=$catId orgAccountType=$acctType"
            if ($titleVal) {
                try {
                    ActiveDirectory\Set-ADUser -Identity $sam -Server $targetPDC -Title $titleVal -ErrorAction Stop
                    Write-Log $RichLog "  Title set: $titleVal" ([System.Drawing.Color]::LimeGreen)
                } catch {
                    Write-Log $RichLog "  [!] Title not set: $_" $ColorGold
                    DBG "CUSTOM_ATTRS" "WARN" "Title not set: $_"
                    $followUpItems += "Title not set"
                }
            }

            # Verify attributes were actually written
            $attrCheck = ActiveDirectory\Get-ADUser -Identity $sam -Server $targetPDC `
                -Properties orgCategoryID, orgAccountType, Title -ErrorAction SilentlyContinue
            DBG "CUSTOM_ATTRS_VERIFY" "INFO" "orgCategoryID=$(if($attrCheck.orgCategoryID){$attrCheck.orgCategoryID}else{'<not set>'}) orgAccountType=$(if($attrCheck.orgAccountType){$attrCheck.orgAccountType}else{'<not set>'}) Title=$(if($attrCheck.Title){$attrCheck.Title}else{'<not set>'})"
            if (-not $attrCheck.orgCategoryID -or -not $attrCheck.orgAccountType) {
                Write-Log $RichLog "  [!] Custom attributes not verified after write -- check ADUC." $ColorGold
                DBG "CUSTOM_ATTRS_VERIFY" "WARN" "Attributes not confirmed -- may be replication lag or write failed"
                $followUpItems += "Check orgCategoryID / orgAccountType in ADUC"
            }
        } catch {
            Write-Log $RichLog "  [!] Custom attribute set failed: $_ -- set manually via ADUC." $ColorGold
            DBG "CUSTOM_ATTRS" "WARN" "$_"
            Write-AuditLog "ATTRS" $sam $upn $divCode $empType "FAILED" "$_"
            $followUpItems += "Set orgCategoryID / orgAccountType in ADUC"
        }

        # Capture ObjectGUID for ImmutableId verification in STEP 3
        $newADObj = Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC `
            -Properties ObjectGUID -ErrorAction SilentlyContinue
        if ($newADObj) {
            $guidBytes                 = ([guid]$newADObj.ObjectGUID).ToByteArray()
            $Script:NewAcctImmutableId = [Convert]::ToBase64String($guidBytes)
            DBG "OBJECT_GUID" "OK" "ObjectGUID=$($newADObj.ObjectGUID) ImmutableId=$($Script:NewAcctImmutableId)"
        } else {
            $Script:NewAcctImmutableId = $null
            DBG "OBJECT_GUID" "WARN" "Could not read ObjectGUID -- STEP 3 will fall back to UPN match"
        }

        # ---- DEBUG: AD state after New-RemoteMailbox ----
        DBG "DEBUG" "INFO" "Reading AD object from $targetPDC after New-RemoteMailbox..."
        try {
            $dbgAD = Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC `
                -Properties targetAddress, proxyAddresses, msExchRemoteRecipientType,
                             msExchRecipientTypeDetails, msExchRecipientDisplayType `
                -ErrorAction SilentlyContinue
            if ($dbgAD) {
                DBG "Object found on PDC" "INFO" "Object found on PDC: $($dbgAD.DistinguishedName)"
                DBG "targetAddress" "INFO" "targetAddress       : $(if($dbgAD.targetAddress){'SET: '+$dbgAD.targetAddress}else{'<not set>'})"
                DBG "proxyAddresses" "INFO" "proxyAddresses      : $(if($dbgAD.proxyAddresses){($dbgAD.proxyAddresses -join ' | ')}else{'<not set>'})"
                DBG "msExchRemoteRecipientType" "INFO" "msExchRemoteRecipientType: $(if($dbgAD.msExchRemoteRecipientType){$dbgAD.msExchRemoteRecipientType}else{'<not set>'})"
            }
        } catch { DBG "AD read error" "INFO" "AD read error: $_" }
        # ---- END DEBUG ----

        # Add contoso.com alias for vanity domain accounts
        if ($emailDomain -ne "contoso.com" -and $tenantAlias) {
            $aliasInUse = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$tenantAlias'" -ErrorAction SilentlyContinue
            if ($aliasInUse) {
                Write-Log $RichLog "  [!] contoso.com alias $tenantAlias in use (race condition) -- assign manually." $ColorGold
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "COLLISION" "$tenantAlias race condition"
                $followUpItems += "Add a contoso.com alias"
            } else {
                try {
                    Set-RemoteMailbox -Identity $sam -EmailAddresses @{ add = "smtp:$tenantAlias" } -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                    Write-Log $RichLog "  contoso.com alias added: $tenantAlias" ([System.Drawing.Color]::LimeGreen)
                    Write-AuditLog "ALIAS" $sam $upn $divCode $empType "OK" "Added smtp:$tenantAlias"
                } catch {
                    Write-Log $RichLog "  [!] contoso.com alias add failed: $_" $ColorGold
                    Write-AuditLog "ALIAS" $sam $upn $divCode $empType "WARN" "smtp:$tenantAlias failed: $_"
                    $followUpItems += "Add contoso.com alias $tenantAlias"
                }
            }
        } elseif (-not $tenantAlias -and $emailDomain -ne "contoso.com") {
            Write-Log $RichLog "  [!] No free contoso.com alias -- assign manually in Exchange." $ColorGold
            $followUpItems += "Add a contoso.com alias"
        }

        # New-RemoteMailbox usually stamps the routing address already. Re-adding it changes
        # nothing and makes Exchange warn, which ps2exe shows as a blocking pop-up.
        try {
            $rmNow = Get-RemoteMailbox -Identity $sam -DomainController $targetPDC -ErrorAction Stop
            if (Test-ProxyAddressPresent $rmNow.EmailAddresses $routing) {
                Write-Log $RichLog "  onmicrosoft routing alias already present: $routing" ([System.Drawing.Color]::LimeGreen)
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "OK" "smtp:$routing already present"
            } else {
                Set-RemoteMailbox -Identity $sam -EmailAddresses @{ add = "smtp:$routing" } -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                Write-Log $RichLog "  onmicrosoft alias added: $routing" ([System.Drawing.Color]::LimeGreen)
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "OK" "Added smtp:$routing"
            }
        } catch {
            Write-Log $RichLog "  [!] onmicrosoft alias add failed: $_" $ColorGold
            Write-AuditLog "ALIAS" $sam $upn $divCode $empType "WARN" "onmicrosoft alias failed: $_"
            $followUpItems += "Add routing address $routing"
        }
        try {
            $dbgOnms   = Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC -Properties proxyAddresses -ErrorAction SilentlyContinue
            $onmsFound = $dbgOnms.proxyAddresses | Where-Object { $_ -like "*contoso.mail.onmicrosoft.com*" }
            DBG "onmicrosoft in AD proxyAddresses" "INFO" "onmicrosoft in AD proxyAddresses: $(if($onmsFound){$onmsFound}else{'NOT YET VISIBLE -- may be replication lag'})"
        } catch { DBG "onmicrosoft alias AD read error" "INFO" "$_" }

        # Add X500 from legacyExchangeDN
        Write-Log $RichLog "  Reading legacyExchangeDN for X500 proxy entry..." ([System.Drawing.Color]::LimeGreen)
        try {
            $legDN = (Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC `
                -Properties legacyExchangeDN -ErrorAction Stop).legacyExchangeDN
            DBG "legacyExchangeDN read" "INFO" "legacyExchangeDN read: $(if($legDN){$legDN}else{'<empty>'})"
            if ($legDN) {
                Set-RemoteMailbox -Identity $sam -EmailAddresses @{ add = "X500:$legDN" } -DomainController $targetPDC -ErrorAction Stop -WarningAction SilentlyContinue
                Write-Log $RichLog "  X500 proxy entry added." ([System.Drawing.Color]::LimeGreen)
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "OK" "Added X500:$legDN"
                Wait-UiDelay -Seconds 3
                $dbgX500   = Get-ADUser -Filter "SamAccountName -eq '$sam'" -Server $targetPDC -Properties proxyAddresses -ErrorAction SilentlyContinue
                $x500Found = $dbgX500.proxyAddresses | Where-Object { $_ -like "X500:*" }
                DBG "X500 in AD proxyAddresses" "INFO" "X500 in AD proxyAddresses: $(if($x500Found){$x500Found}else{'NOT YET VISIBLE -- may be replication lag'})"
            } else {
                Write-Log $RichLog "  [!] legacyExchangeDN empty -- X500 not added." $ColorGold
                Write-AuditLog "ALIAS" $sam $upn $divCode $empType "WARN" "legacyExchangeDN empty -- X500 not added"
                $followUpItems += "Add the X500 address"
            }
        } catch {
            Write-Log $RichLog "  [!] X500 proxy add failed: $_" $ColorGold
            Write-AuditLog "ALIAS" $sam $upn $divCode $empType "WARN" "X500 add failed: $_"
            $followUpItems += "Add the X500 address"
        }

        # Verify -- with New-RemoteMailbox attributes are stamped atomically so this
        # should confirm immediately without a retry loop.
        Write-Log $RichLog "  Verifying mailbox attributes..." ([System.Drawing.Color]::Yellow)
        $rmVerify = Get-RemoteMailbox -Identity $sam -DomainController $targetPDC -ErrorAction SilentlyContinue
        DBG "VERIFY" "INFO" "RemoteRoutingAddress: $(if($rmVerify.RemoteRoutingAddress){$rmVerify.RemoteRoutingAddress}else{'<empty>'})"
        DBG "VERIFY" "INFO" "PrimarySmtpAddress: $(if($rmVerify.PrimarySmtpAddress){$rmVerify.PrimarySmtpAddress}else{'<empty>'})"
        DBG "VERIFY" "INFO" "EmailAddresses: $(if($rmVerify.EmailAddresses){($rmVerify.EmailAddresses -join ' | ')}else{'<empty>'})"
        DBG "VERIFY" "INFO" "RecipientTypeDetails: $($rmVerify.RecipientTypeDetails)"
        if ($rmVerify -and $rmVerify.RemoteRoutingAddress -like "*@contoso.mail.onmicrosoft.com*") {
            $step2OK = $true
            Write-Log $RichLog "  [OK] Mailbox verified." ([System.Drawing.Color]::LimeGreen)
            Write-AuditLog "MAILBOX_VERIFY" $sam $upn $divCode $empType "OK" "Verified: $($rmVerify.RemoteRoutingAddress)"
        } else {
            # Attributes not visible yet -- brief wait and one retry (replication can still lag)
            Write-Log $RichLog "  Attributes not yet visible -- waiting 30s and retrying..." $ColorGold
            Wait-UiDelay -Seconds 30
            [System.Windows.Forms.Application]::DoEvents()
            $rmRetry = Get-RemoteMailbox -Identity $sam -DomainController $targetPDC -ErrorAction SilentlyContinue
            if ($rmRetry -and $rmRetry.RemoteRoutingAddress -like "*@contoso.mail.onmicrosoft.com*") {
                $step2OK = $true
                Write-Log $RichLog "  [OK] Mailbox verified on retry." ([System.Drawing.Color]::LimeGreen)
                Write-AuditLog "MAILBOX_VERIFY" $sam $upn $divCode $empType "OK" "Verified on retry: $($rmRetry.RemoteRoutingAddress)"
            } else {
                Write-Log $RichLog "  [!] Mailbox attributes not confirmed -- check Exchange manually." $ColorGold
                Write-Log $RichLog "    Get-RemoteMailbox $sam | fl RemoteRoutingAddress,EmailAddresses" $ColorGold
                Write-AuditLog "MAILBOX_VERIFY" $sam $upn $divCode $empType "UNCONFIRMED" "Not visible after creation + 30s retry"
                # With New-RemoteMailbox the account IS created -- this is almost certainly
                # replication lag, not a failure. Prompt tech rather than hard-blocking.
                $promptResult = [System.Windows.Forms.MessageBox]::Show(
                    "Mailbox attributes not visible after 30 seconds.`n`n" +
                    "The account was created via New-RemoteMailbox and almost certainly succeeded.`n" +
                    "This is likely replication lag on the PDC.`n`n" +
                    "YES = Proceed to license assignment (recommended)`n" +
                    "NO = Stop here -- investigate before licensing",
                    "Mailbox Verify",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($promptResult -eq "Yes") {
                    $step2OK = $true
                    Write-Log $RichLog "  Tech confirmed -- proceeding to license." $ColorGold
                    Write-AuditLog "MAILBOX_VERIFY" $sam $upn $divCode $empType "MANUAL_PROCEED" "Tech confirmed after replication lag warning"
                } else {
                    $step2OK = $false
                    Write-Log $RichLog "  [STOP] Tech chose to investigate. License NOT assigned." $ColorRed
                    Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "STOPPED_STEP2" "Tech halted at verify -- license not assigned"
                }
            }
        }

    } catch {
        $step2OK = $false
        DBG "STEP2_FAIL" "FAIL" "$_"
        if (-not $exchReady) {
            Write-Log $RichLog "[X] Exchange session failed: $_" $ColorRed
            Write-Log $RichLog "  Nothing was created. Fix the connection and run Preview again." $ColorRed
            $nothingCreated = $true
            Write-AuditLog "CREATE" $sam $upn $divCode $empType "FAILED" "Exchange session failed, nothing created: $_"
        } elseif ($newMailbox) {
            Write-Log $RichLog "[X] The account was created, but a later step failed: $_" $ColorRed
            Write-Log $RichLog "  Do NOT run Create again. Finish the account by hand." $ColorRed
            Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "STOPPED_STEP2" "After create: $_"
        } else {
            Write-Log $RichLog "[X] Account/mailbox creation failed: $_" $ColorRed
            $state = Get-CreatedObjectState -Sam $sam -Pdc $targetPDC
            if ($state.State -eq 'Absent') {
                Write-Log $RichLog "  Confirmed nothing was created ($sam is not in AD). Fix the cause and run Preview again." $ColorRed
                $nothingCreated = $true
                Write-AuditLog "CREATE" $sam $upn $divCode $empType "FAILED" "$_"
            } elseif ($state.State -eq 'Exists') {
                Write-Log $RichLog "  [!] $sam exists in AD anyway: $($state.Detail)" $ColorGold
                Write-Log $RichLog "  Do NOT run Create again. Finish this account by hand." $ColorGold
                Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "OBJECT_EXISTS" "New-RemoteMailbox error but object exists: $($state.Detail) | $_"
            } else {
                Write-Log $RichLog "  [!] Could not check whether $sam was created ($($state.Detail)). Search for it before running Create again." $ColorGold
                Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "STOPPED_STEP2" "New-RemoteMailbox failed, existence unknown: $_"
            }
        }
    } finally {
        if ($exchSession) { Remove-PSSession $exchSession -ErrorAction SilentlyContinue }
    }

    # Gate: only continue to STEP 3 if mailbox is confirmed
    if (-not $step2OK -and $nothingCreated) {
        # No account exists, so there are no details or temp password to hand out.
        $Script:LastAccountDetails = $null
        $BtnCopyAcct.Enabled = $false
        $LblCreateStatus.Text = "Nothing created -- see the log, fix the cause and Preview again."; $LblCreateStatus.ForeColor = $ColorRed
        $BtnPreview.Enabled = $true
        return
    }
    if (-not $step2OK) {
        Write-Log $RichLog "" $null
        Write-Log $RichLog "  Mailbox not confirmed -- skipping Entra sync and license assignment." $ColorRed
        Write-Log $RichLog "  Once resolved, assign license manually -- Group: $($LicenseGroups[$empType])" $ColorGold
        # Still activate Copy button with partial info so tech can paste into ticket
        $divShort = $domain.Split(".")[0].ToUpper()
        $Script:LastAccountDetails = "SAM Account Name: $divShort\$sam`r`nUPN: $upn`r`nEmail: $emailAddr`r`nTemp Password: $plainPass`r`nLicense Assigned: NOT ASSIGNED -- mailbox verify failed, assign manually`r`n"
        $BtnCopyAcct.Enabled = $true
        $LblCreateStatus.Text = "Stopped at STEP 2 -- license not assigned."; $LblCreateStatus.ForeColor = $ColorRed
        $BtnPreview.Enabled = $true
        try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
        return
    }

    # STEP 3 - Smart Entra Sync Wait
    Write-Log $RichLog "" $null
    Write-Log $RichLog "STEP 3: Waiting for Entra sync (up to 45 min)..." $ColorLight

    # Reuse the pre-flight Graph token. Pre-flight stops without a token,
    # so the fallback below is defensive only.
    $headers = $null
    $tokenRefreshed = $false
    $step3ScriptDir = $PSScriptRoot
    if (-not $step3ScriptDir) { $step3ScriptDir = "C:\ADProvisioning" }
    if ($aliasGraphHeaders) {
        $headers = $aliasGraphHeaders
        Write-Log $RichLog "  Using Graph token from STEP 2." ([System.Drawing.Color]::LimeGreen)
    } else {
        Write-Log $RichLog "  Getting Graph token..." ([System.Drawing.Color]::LimeGreen)
        $headers = Get-GraphToken -ScriptDir $step3ScriptDir
        if ($headers) {
            Write-Log $RichLog "  Token acquired successfully." ([System.Drawing.Color]::LimeGreen)
        } else {
            Write-Log $RichLog "[X] Failed to get Graph token." $ColorRed
            Write-Log $RichLog "  Assign license manually in Entra -- Group: $($LicenseGroups[$empType])" $ColorGold
            Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "STOPPED_STEP3" "Graph token failed"
            $BtnPreview.Enabled = $true
            return
        }
    }

    # Smart sync wait -- poll every 2 min up to 45 min
    $maxWait  = 2700
    $interval = 120
    $elapsed  = 0
    $synced   = $false
    $syncUser = $null
    $step3OK  = $false

    # Enable Cancel button now -- it stays active until STEP 3 ends
    $Script:CancelRequested  = $false
    $BtnCancelSync.Enabled   = $true
    $BtnCancelSync.BackColor = $ColorRGB_180_60_60

    Write-Log $RichLog "  Waiting 2 minutes before first check..." ([System.Drawing.Color]::Yellow)
    Write-Log $RichLog "  (Click 'Cancel Sync Wait' at any time to skip and assign license manually)" $ColorGold
    for ($s = 120; $s -ge 0; $s--) {
        if ($Script:CancelRequested) { break }
        $LblCreateStatus.Text      = "Waiting for Entra sync -- checking in $s seconds..."
        $LblCreateStatus.ForeColor = $ColorGold
        Start-Sleep -Seconds 1
        [System.Windows.Forms.Application]::DoEvents()
    }

    $wrongWarned = $false
    $authFailed  = $false
    while (-not $synced -and -not $Script:CancelRequested -and $elapsed -lt $maxWait) {
        Write-Log $RichLog "  Checking if user has synced to Entra..." ([System.Drawing.Color]::Yellow)
        $poll = $null
        $fresh = Get-GraphToken -ScriptDir $step3ScriptDir
        if ($fresh) { $headers = $fresh }
        try {
            $poll = Get-SyncedEntraUser -Upn $upn -ImmutableId $Script:NewAcctImmutableId -Headers $headers
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { $statusCode = 0 }
            if ($statusCode -eq 401 -and -not $tokenRefreshed) {
                # Long sync waits can outlive the Graph token; retry once with a fresh one.
                DBG "SYNC_POLL" "WARN" "Graph token rejected (401), requesting a new token"
                $headers = Get-GraphToken -ScriptDir $step3ScriptDir -Force
                $tokenRefreshed = $true
                continue
            }
            if ($statusCode -eq 401 -or $statusCode -eq 403) {
                Write-Log $RichLog "[X] Graph refused the sync check (HTTP $statusCode) -- check the app's Graph permissions." $ColorRed
                DBG "SYNC_POLL" "FAIL" "Auth error ${statusCode}: $(Get-GraphErrorDetail $_)"
                $authFailed = $true
                break
            }
            DBG "SYNC_POLL" "WARN" "Sync check failed (HTTP $statusCode): $(Get-GraphErrorDetail $_)"
        }

        if ($poll -and ($poll.State -eq 'Synced' -or $poll.State -eq 'UpnChanged')) {
            $synced   = $true
            $step3OK  = $true
            $syncUser = $poll.User
            DBG "SYNC_VERIFY" "OK" "$($poll.Detail) (ID=$($syncUser.id))"
            Write-Log $RichLog "  Correct account confirmed in Entra! ($([math]::Round(($elapsed+120)/60)) min)" ([System.Drawing.Color]::LimeGreen)
            Write-Log $RichLog "  Entra Object ID: $($syncUser.id)" ([System.Drawing.Color]::LimeGreen)
            if ($poll.State -eq 'UpnChanged') {
                # Entra Connect renamed the UPN because it is already taken in the tenant.
                $upnNote = $poll.Detail
                Write-Log $RichLog "  [!] $($poll.Detail). The UPN is already used in Entra -- fix the UPN after this run." $ColorGold
                Write-AuditLog "SYNC" $sam $upn $divCode $empType "UPN_CHANGED" $poll.Detail
                DBG "SYNC_VERIFY" "WARN" $poll.Detail
            }
            break
        }
        if ($poll -and $poll.State -eq 'WrongObject' -and -not $wrongWarned) {
            Write-Log $RichLog "  [!] $($poll.Detail) -- waiting for our account..." $ColorGold
            DBG "SYNC_VERIFY" "WARN" $poll.Detail
            $wrongWarned = $true
        }

        $elapsed += $interval
        if ($elapsed -lt $maxWait -and -not $Script:CancelRequested) {
            Write-Log $RichLog "  Not synced yet -- rechecking in 2 min ($([math]::Round($elapsed/60)) min elapsed)" ([System.Drawing.Color]::Yellow)
            for ($s = $interval; $s -ge 0; $s--) {
                if ($Script:CancelRequested) { break }
                $LblCreateStatus.Text      = "Not synced yet -- rechecking in $s seconds..."
                $LblCreateStatus.ForeColor = $ColorGold
                Start-Sleep -Seconds 1
                [System.Windows.Forms.Application]::DoEvents()
            }
        }
    }

    # Disable Cancel button -- STEP 3 is over regardless of how it ended
    $BtnCancelSync.Enabled   = $false
    $BtnCancelSync.BackColor = $ColorRGB_100_40_40

    # Handle cancellation
    if ($Script:CancelRequested -and -not $synced) {
        $BtnCancelSync.Enabled   = $false
        $BtnCancelSync.Text      = "Cancel Sync Wait"
        $BtnCancelSync.BackColor = $ColorRGB_180_60_60
        Write-Log $RichLog "" $null
        Write-Log $RichLog "  Sync wait cancelled by tech -- license NOT assigned." $ColorGold
        Write-Log $RichLog "  Once the account appears in Entra, assign license manually -- Group: $($LicenseGroups[$empType])" $ColorGold
        Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "SYNC_CANCELLED" "Tech cancelled sync wait -- license not assigned"
        DBG "STEP3_CANCELLED" "WARN" "Tech cancelled sync wait"
        $divShort = $domain.Split(".")[0].ToUpper()
        $Script:LastAccountDetails = "SAM Account Name: $divShort\$sam`r`nUPN: $upn`r`nEmail: $emailAddr`r`nTemp Password: $plainPass`r`nLicense Assigned: NOT ASSIGNED -- sync wait cancelled, assign manually`r`n"
        $BtnCopyAcct.Enabled = $true
        $LblCreateStatus.Text = "Sync wait cancelled -- license not assigned."
        $LblCreateStatus.ForeColor = $ColorGold
        $BtnPreview.Enabled = $true
        try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
        return
    }

    # Gate: hard stop if sync was not confirmed -- never license an unsynced account
    if (-not $step3OK) {
        Write-Log $RichLog "" $null
        if ($authFailed) {
            Write-Log $RichLog "  [!] Stopped: Graph refused the sync check -- license NOT assigned." $ColorRed
            Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "SYNC_AUTH_FAILED" "Graph refused the sync check -- license not assigned"
            DBG "STEP3_AUTH" "FAIL" "Graph refused the sync check"
        } else {
            Write-Log $RichLog "  [!] Entra sync not confirmed after 45 min -- license NOT assigned." $ColorRed
            Write-AuditLog "PARTIAL_CREATE" $sam $upn $divCode $empType "SYNC_TIMEOUT" "Entra sync timed out after 45 min -- license not assigned"
            DBG "STEP3_TIMEOUT" "FAIL" "Entra sync not confirmed after 45 min"
        }
        Write-Log $RichLog "  Once the account appears in Entra, assign manually -- Group: $($LicenseGroups[$empType])" $ColorGold
        $divShort = $domain.Split(".")[0].ToUpper()
        $Script:LastAccountDetails = "SAM Account Name: $divShort\$sam`r`nUPN: $upn`r`nEmail: $emailAddr`r`nTemp Password: $plainPass`r`nLicense Assigned: NOT ASSIGNED -- Entra sync not confirmed, assign manually`r`n"
        $BtnCopyAcct.Enabled = $true
        $LblCreateStatus.Text = "Stopped at STEP 3 -- sync not confirmed, license not assigned."; $LblCreateStatus.ForeColor = $ColorRed
        $BtnPreview.Enabled = $true
        try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
        return
    }

    # STEP 4 - Assign License Group via REST (only if step2OK AND step3OK)
    $licenseOK = $false
    Write-Log $RichLog "" $null
    Write-Log $RichLog "STEP 4: Assigning M365 license group..." $ColorLight
    $fresh = Get-GraphToken -ScriptDir $step3ScriptDir
    if ($fresh) { $headers = $fresh }
    try {
        $groupName   = $LicenseGroups[$empType]
        $userId      = $syncUser.id
        $groupFilter = [Uri]::EscapeDataString("displayName eq '$groupName'")
        $groupResp   = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups?`$filter=$groupFilter" -Headers $headers -ErrorAction Stop
        $groupId     = $groupResp.value[0].id
        if (-not $groupId) { throw "Group '$groupName' not found in Entra." }
        $memberBody  = '{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/' + $userId + '"}'
        try {
            Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/groups/$groupId/members/`$ref" -Method POST -Headers $headers -Body $memberBody -ErrorAction Stop
        } catch {
            if ((Get-GraphErrorDetail $_) -notmatch 'already exist') { throw }
        }
        Write-Log $RichLog "  License group assigned: $groupName" ([System.Drawing.Color]::LimeGreen)
        Write-AuditLog "LICENSE" $sam $upn $divCode $empType "OK" "Group=$groupName"
        $licenseOK = $true
    } catch {
        Write-Log $RichLog "[X] License assignment error: $_" $ColorRed
        Write-Log $RichLog "  Assign manually in Entra -- Group: $($LicenseGroups[$empType])" $ColorGold
        Write-AuditLog "LICENSE" $sam $upn $divCode $empType "FAILED" "$_"
    }

    # STEP 5 - Assign Administrative Unit(s) via Graph
    $auAssigned = "Not assigned"
    $auCodes = Get-AUCodes $divCode
    Write-Log $RichLog "" $null
    Write-Log $RichLog "STEP 5: Assigning Administrative Unit(s) ($($auCodes -join ', '))..." $ColorLight
    DBG "AU_ASSIGN" "INFO" "Assigning AU(s): $($auCodes -join ', ') to user $($syncUser.id)"
    $auRes = @(Add-UserToAUs -UserId $syncUser.id -AUNames $auCodes -Headers $headers)
    foreach ($ar in $auRes) {
        if ($ar.OK) {
            Write-Log $RichLog "  AU assigned: $($ar.AU)" ([System.Drawing.Color]::LimeGreen)
            Write-AuditLog "AU" $sam $upn $divCode $empType "OK" "AU=$($ar.AU)"
            DBG "AU_ASSIGN" "OK" "$($ar.AU): $($ar.Detail)"
        } else {
            Write-Log $RichLog "  [!] AU '$($ar.AU)' failed -- assign manually in Entra: $($ar.Detail)" $ColorGold
            Write-AuditLog "AU" $sam $upn $divCode $empType "FAILED" "AU=$($ar.AU) | $($ar.Detail)"
            DBG "AU_ASSIGN" "WARN" "$($ar.AU): $($ar.Detail)"
        }
    }
    $auDone = @($auRes | Where-Object { $_.OK } | ForEach-Object { $_.AU })
    $auMiss = @($auRes | Where-Object { -not $_.OK } | ForEach-Object { $_.AU })
    if ($auDone.Count) { $auAssigned = $auDone -join ', ' }
    if ($auMiss.Count) { $auAssigned += " (NOT ASSIGNED: $($auMiss -join ', '))" }

    # DONE
    $licenseAssigned = if ($licenseOK) { $LicenseGroups[$empType] } else { "NOT ASSIGNED -- assign manually: $($LicenseGroups[$empType])" }
    $followUp = (-not $licenseOK) -or ($auMiss.Count -gt 0) -or $upnNote -or ($followUpItems.Count -gt 0)
    Write-Log $RichLog "" $null
    Write-Log $RichLog "============================================" ([System.Drawing.Color]::Cyan)
    if ($followUp) {
        Write-Log $RichLog "ACCOUNT CREATED -- FOLLOW-UP NEEDED" $ColorGold
    } else {
        Write-Log $RichLog "ACCOUNT CREATION COMPLETE" ([System.Drawing.Color]::Cyan)
    }
    Write-Log $RichLog "  Display Name : $dispName" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  SAM          : $sam" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  UPN          : $upn" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  Email        : $emailAddr" ([System.Drawing.Color]::Cyan)
    if ($upnNote) { Write-Log $RichLog "  [!] $upnNote" $ColorGold }
    foreach ($item in $followUpItems) { Write-Log $RichLog "  [!] Follow-up: $item" $ColorGold }
    Write-Log $RichLog "  Domain       : $domain" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  Temp Password: $plainPass" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  License      : $licenseAssigned" ([System.Drawing.Color]::Cyan)
    Write-Log $RichLog "  Admin Unit   : $auAssigned" ([System.Drawing.Color]::Cyan)
    if ($ticket) { Write-Log $RichLog "  Ticket #     : $ticket" ([System.Drawing.Color]::Cyan) }
    Write-Log $RichLog "============================================" ([System.Drawing.Color]::Cyan)

    # Update last-created status bar
    $LblLastCreated.Text = "Last created: $sam  |  $upn  |  $((Get-Date).ToString('HH:mm'))"

    # Store formatted account details for clipboard copy
    $divShort = $domain.Split(".")[0].ToUpper()
    $Script:LastAccountDetails = @"
SAM Account Name: $divShort\$sam
UPN: $upn
Email: $emailAddr
Temp Password: $plainPass
License Assigned: $licenseAssigned
Admin Unit: $auAssigned
"@
    $BtnCopyAcct.Enabled = $true

    if ($followUp) {
        $LblCreateStatus.Text      = "Account created -- follow-up needed, see the log."
        $LblCreateStatus.ForeColor = $ColorGold
    } else {
        $LblCreateStatus.Text      = "Done! Account created successfully."
        $LblCreateStatus.ForeColor = $ColorAccent
    }
    $BtnPreview.Enabled        = $true
    $BtnCancelSync.Enabled     = $false
    $BtnCancelSync.Text        = "Cancel Sync Wait"
    $BtnCancelSync.BackColor   = $ColorRGB_180_60_60

    # Auto-generate audit HTML report after every successful creation
    try { Generate-AuditReport | Out-Null } catch { DBG "AUDIT_REPORT" "WARN" "HTML report generation failed: $_" }
}
$BtnCreate.Add_Click({
    if ($Script:BulkRunning) {
        [System.Windows.Forms.MessageBox]::Show("A bulk run is in progress. Wait for it to finish before creating a single account.", "Run in progress") | Out-Null
        return
    }
    $Script:SingleRunning = $true
    $Script:AuditTicket = $TxtTicket.Text.Trim()
    $Script:RunSam = $null; $Script:RunUpn = $null; $Script:RunPdc = $null; $Script:RunDiv = $null; $Script:RunEmp = $null
    try {
        & $Script:SingleCreateBody
    } catch {
        # Inside a try block any error the body does not handle ends the run. Say so, and say
        # whether the account exists, so a half-built account is never left without a trace.
        $where = if ($_.InvocationInfo -and $_.InvocationInfo.ScriptLineNumber) { " (line $($_.InvocationInfo.ScriptLineNumber))" } else { "" }
        Write-Log $RichLog "" $null
        Write-Log $RichLog "[X] The run stopped on an unexpected error${where}: $($_.Exception.Message)" $ColorRed
        DBG "RUN_ABORTED" "FAIL" "$($_.Exception.Message)$where"
        $state = $null
        if ($Script:RunSam -and $Script:RunPdc) { $state = Get-CreatedObjectState -Sam $Script:RunSam -Pdc $Script:RunPdc }
        if ($state -and $state.State -eq 'Exists') {
            Write-Log $RichLog "  The account $($Script:RunSam) EXISTS in AD. Do NOT run Create again." $ColorRed
            Write-Log $RichLog "  Finish it by hand: check the mailbox, then add the license group and AU." $ColorRed
            Write-AuditLog "PARTIAL_CREATE" $Script:RunSam $Script:RunUpn $Script:RunDiv $Script:RunEmp "RUN_ABORTED" "Account exists; run stopped on: $($_.Exception.Message)$where"
            $LblCreateStatus.Text = "Run stopped -- account exists, finish by hand. See the log."
        } elseif ($state -and $state.State -eq 'Absent') {
            Write-Log $RichLog "  Nothing was created ($($Script:RunSam) is not in AD)." $ColorRed
            Write-AuditLog "CREATE" $Script:RunSam $Script:RunUpn $Script:RunDiv $Script:RunEmp "FAILED" "Run stopped before create on: $($_.Exception.Message)$where"
            $LblCreateStatus.Text = "Run stopped -- nothing created. See the log."
        } else {
            Write-Log $RichLog "  Search for the account before running Create again: it may or may not exist." $ColorRed
            Write-AuditLog "PARTIAL_CREATE" "$($Script:RunSam)" "$($Script:RunUpn)" "$($Script:RunDiv)" "$($Script:RunEmp)" "RUN_ABORTED" "Existence unknown; run stopped on: $($_.Exception.Message)$where"
            $LblCreateStatus.Text = "Run stopped -- check whether the account exists. See the log."
        }
        $LblCreateStatus.ForeColor = $ColorRed
        $BtnPreview.Enabled    = $true
        $BtnCancelSync.Enabled = $false
    } finally { $Script:SingleRunning = $false; $Script:AuditTicket = $null }
})

# ---- CANCEL SYNC WAIT CLICK --------------------------------
$BtnCancelSync.Add_Click({
    $Script:CancelRequested  = $true
    $BtnCancelSync.Enabled   = $false
    $BtnCancelSync.Text      = "Cancelling..."
    $LblCreateStatus.Text    = "Cancelling sync wait..."
    $LblCreateStatus.ForeColor = $ColorGold
})
$BtnCopyAcct.Add_Click({
    if ($Script:LastAccountDetails) {
        [System.Windows.Forms.Clipboard]::SetText($Script:LastAccountDetails)
        # Write to RichLog so tech sees confirmation next to the Creation Log
        # LblCreateStatus is on the left panel -- not visible when looking at the log
        if ($RichLog) {
            Write-Log $RichLog "  [OK] Account details copied to clipboard." ([System.Drawing.Color]::LimeGreen)
        }
        $LblCreateStatus.Text      = "Account details copied to clipboard."
        $LblCreateStatus.ForeColor = $ColorAccent
    }
})

# ---- CLEAR CLICK -------------------------------------------
$BtnClear.Add_Click({
    # Clearing mid-run would disable Cancel Sync Wait while the run still needs it.
    if ($Script:BulkRunning -or $Script:SingleRunning) { return }
    $TxtFirst.Clear(); $TxtMid.Clear(); $TxtLast.Clear()
    $TxtTitle.Clear(); $TxtDomain.Clear(); $TxtPreview.Clear()
    $TxtTicket.Clear(); $RichLog.Clear()
    $LblCreateStatus.Text        = ""
    $CboDivision.SelectedIndex   = 0
    $CboEmpType.SelectedIndex    = 0
    $BtnCreate.Enabled           = $false
    $BtnCopyAcct.Enabled         = $false
    $BtnCancelSync.Enabled       = $false
    $BtnCancelSync.Text          = "Cancel Sync Wait"
    $BtnCancelSync.BackColor     = $ColorRGB_180_60_60
    $Script:ResolvedSAM          = $null
    $Script:LastAccountDetails   = $null
    $Script:CancelRequested      = $false
})

# Closing mid-run would leave the run's loops working against a closed window, with an
# account half provisioned and no audit row saying so. The tech stops the run first.
$Form.Add_FormClosing({
    # In a WinForms event scriptblock $_ is the event args (FormClosingEventArgs here).
    $closeArgs = $_
    if ($Script:SingleRunning -or $Script:BulkRunning) {
        [System.Windows.Forms.MessageBox]::Show(
            "A run is still in progress. Use Cancel Sync Wait or the row Skip buttons, wait for the run to finish, then close the tool.",
            "Run in progress") | Out-Null
        $closeArgs.Cancel = $true
        return
    }
    Get-PSSession -ErrorAction SilentlyContinue | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' } |
        Remove-PSSession -ErrorAction SilentlyContinue
    Write-AuditLog "SESSION_END" "" "" "" "" "CLOSED" "Tool closed"
})

Write-AuditLog "SESSION_START" "" "" "" "" "OPENED" "Tool launched"
[System.Windows.Forms.Application]::Run($Form)
