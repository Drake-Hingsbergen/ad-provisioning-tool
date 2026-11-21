<#
.SYNOPSIS
    Builds the fake world inside the smoke-test runspace. Run by Invoke-HandlerSmoke.ps1, not by hand.
.DESCRIPTION
    Runs at the global scope of a new, empty runspace, so nothing here touches the caller's session.
    It loads the tool's functions and button handlers, defines stand-in window controls, and defines
    fakes for every AD, Exchange and Graph command. Before any handler runs, a gate walks all the tool
    code the handlers can reach and stops if that code calls a command that is not a tool function, a
    fake, or a reviewed built-in. That gate is what keeps a future edit to the tool from reaching a
    real domain through this harness.
.NOTES
    Windows PowerShell 5.1 and PowerShell 7. The SmokeUi.* types are compiled by Invoke-HandlerSmoke.ps1.
#>
param([hashtable]$Opt, [string]$ToolPath, [string]$WorkDir)

$ErrorActionPreference = 'Continue'
# Global, because the engine only reads the global value. No module (for example a real ActiveDirectory
# module installed on this machine) can load itself in this runspace.
$global:PSModuleAutoLoadingPreference = 'None'

#region World state
$global:SmokeWorld = @{
    Opt = $Opt
    Ad = New-Object System.Collections.Generic.List[object]
    Entra = New-Object System.Collections.Generic.List[object]
    Pending = New-Object System.Collections.Generic.List[object]
    Calls = New-Object System.Collections.Generic.List[object]
    Unexpected = New-Object System.Collections.Generic.List[string]
    SessionsOpened = 0; SessionsRemoved = 0; OpenSessions = 0
    CreateCalls = 0; MailboxReads = 0; Sleeps = 0; TokenForce = 0
    AUs = @(); Groups = @(); MidRun = $null; CloseCancelled = $null
}
function Add-SmokeCall { param([string]$Command, [hashtable]$Data) $global:SmokeWorld.Calls.Add([PSCustomObject]@{ Command = $Command; Data = $Data }) }
function New-SmokeAdObject {
    param([string]$Sam, [string]$Upn, [string]$Domain, [string]$Name, [string]$Ou, [string]$Primary, [string]$Routing)
    if (-not $Primary) { $Primary = $Upn }
    $proxies = New-Object System.Collections.ArrayList
    [void]$proxies.Add("SMTP:$Primary")
    if ($Routing) { [void]$proxies.Add("smtp:$Routing") }
    [PSCustomObject]@{
        SamAccountName = $Sam; UserPrincipalName = $Upn; Name = $Name; Domain = $Domain
        DistinguishedName = "CN=$Name,$Ou"; ObjectGUID = [guid]::NewGuid(); proxyAddresses = $proxies; mail = $Primary
        legacyExchangeDN = "/o=Contoso/ou=Exchange Administrative Group/cn=Recipients/cn=$Sam"
        targetAddress = $(if ($Routing) { "SMTP:$Routing" }); Routing = $Routing
        Title = $null; orgCategoryID = $null; orgAccountType = $null; Enabled = $true
    }
}
function Add-SmokeExistingAd {
    param($Items)
    foreach ($e in @($Items)) {
        $ou = if ($e.Ou) { $e.Ou } else { 'OU=Users,' + (($e.Domain.Split('.') | ForEach-Object { "DC=$_" }) -join ',') }
        $name = if ($e.Name) { $e.Name } else { $e.Sam }
        $global:SmokeWorld.Ad.Add((New-SmokeAdObject -Sam $e.Sam -Upn $e.Upn -Domain $e.Domain -Name $name -Ou $ou))
    }
}
Add-SmokeExistingAd $Opt.ExistingAd
foreach ($e in @($Opt.ExistingEntra)) {
    $global:SmokeWorld.Entra.Add([PSCustomObject]@{ id = [guid]::NewGuid().ToString(); userPrincipalName = $e.Upn; mail = $e.Upn
            onPremisesImmutableId = $e.ImmutableId; proxyAddresses = @("SMTP:$($e.Upn)"); IntendedUpn = $e.Upn })
}
#endregion

#region Fake ActiveDirectory module
# Module-qualified calls (ActiveDirectory\Get-ADUser) resolve by module name, so the AD fakes live in an
# in-memory module of that name.
$adModule = New-Module -Name ActiveDirectory -ScriptBlock {
    $script:DefaultUserProperties = 'SamAccountName', 'UserPrincipalName', 'Name', 'DistinguishedName', 'Enabled', 'ObjectGUID', 'Domain'
    function ConvertFrom-SmokeLdapValue { param([string]$Value) [regex]::Replace($Value, '\\([0-9a-fA-F]{2})', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) }) }
    function Get-SmokeDomainOfServer { param([string]$Server) if ($Server -like 'pdc.*') { return $Server.Substring(4) } if ($Server -like '*:3268') { return $null } return $Server }
    function Update-SmokeAdsi {
        foreach ($c in [SmokeUi.Adsi]::Committed) {
            $dn = $c[0] -replace '^LDAP://[^/]+/', ''
            $o = $global:SmokeWorld.Ad | Where-Object { $_.DistinguishedName -eq $dn } | Select-Object -First 1
            if ($o) { $o.($c[1]) = $c[2] }
        }
    }
    function Get-ADDomainController {
        [CmdletBinding()] param([switch]$Discover, $Service)
        if ($global:SmokeWorld.Opt.GcUnavailable) { Write-Error 'Unable to find a suitable global catalog.'; return }
        [PSCustomObject]@{ HostName = @('gc1.contoso.local') }
    }
    function Get-ADDomain {
        [CmdletBinding()] param($Identity, $Server)
        [PSCustomObject]@{ PDCEmulator = "pdc.$Identity"; DNSRoot = $Identity }
    }
    function Get-ADObject {
        [CmdletBinding()] param($LDAPFilter, $Server, $SearchBase, $Properties, $Filter)
        $w = $global:SmokeWorld
        $domain = Get-SmokeDomainOfServer "$Server"
        if ("$LDAPFilter" -match '^\(sAMAccountName=(.+)\)$') {
            $v = ConvertFrom-SmokeLdapValue $Matches[1]
            return @($w.Ad | Where-Object { $_.SamAccountName -eq $v -and (-not $domain -or $_.Domain -eq $domain) })
        }
        if ("$LDAPFilter" -match '^\(name=(.+)\)$') {
            $v = ConvertFrom-SmokeLdapValue $Matches[1]
            return @($w.Ad | Where-Object { $_.Name -eq $v -and (-not $SearchBase -or $_.DistinguishedName -like "*,$SearchBase") })
        }
        if ("$LDAPFilter" -like '(|*') {
            $values = @([regex]::Matches("$LDAPFilter", '\((userPrincipalName|mail|proxyAddresses)=([^)]+)\)') | ForEach-Object { $_.Groups[2].Value })
            return @($w.Ad | Where-Object {
                    $o = $_
                    @($values | Where-Object { $o.UserPrincipalName -eq $_ -or $o.mail -eq $_ -or $o.proxyAddresses -contains $_ }).Count -gt 0
                })
        }
        $w.Unexpected.Add("Get-ADObject filter not understood: $LDAPFilter $Filter")
    }
    function Get-ADUser {
        # Like the real cmdlet, returns only the default properties plus those named in -Properties.
        [CmdletBinding()] param($Identity, $Filter, $Server, $Properties)
        $w = $global:SmokeWorld
        Update-SmokeAdsi
        $domain = Get-SmokeDomainOfServer "$Server"
        $sam = $null
        if ($Identity) { $sam = "$Identity" }
        elseif ("$Filter" -match "^SamAccountName -eq '(.+)'$") { $sam = $Matches[1] -replace "''", "'" }
        else { $w.Unexpected.Add("Get-ADUser filter not understood: $Filter"); return }
        $hit = $w.Ad | Where-Object { $_.SamAccountName -eq $sam -and (-not $domain -or $_.Domain -eq $domain) } | Select-Object -First 1
        if (-not $hit) { if ($Identity) { Write-Error "Cannot find an object with identity: '$Identity'" }; return }
        $wanted = @($script:DefaultUserProperties) + @($Properties)
        if (@($Properties) -contains '*') { return $hit }
        $copy = [ordered]@{}
        foreach ($p in $hit.PSObject.Properties) { if ($wanted -contains $p.Name) { $copy[$p.Name] = $p.Value } }
        return [PSCustomObject]$copy
    }
    function Set-ADUser {
        [CmdletBinding()] param($Identity, $Server, $Title)
        $w = $global:SmokeWorld
        Add-SmokeCall 'Set-ADUser' @{ Identity = "$Identity"; Title = $Title }
        $hit = $w.Ad | Where-Object { $_.SamAccountName -eq "$Identity" } | Select-Object -First 1
        if (-not $hit) { Write-Error "Cannot find an object with identity: '$Identity'"; return }
        if ($w.Opt.TitleFails) { Write-Error 'Insufficient access rights to perform the operation'; return }
        $hit.Title = $Title
    }
    Export-ModuleMember -Function Get-ADDomainController, Get-ADDomain, Get-ADObject, Get-ADUser, Set-ADUser
}
Import-Module $adModule -Global -Force -DisableNameChecking
#endregion

#region Load the tool
$source = [System.IO.File]::ReadAllText($ToolPath)
if ($Opt.SourceTransform) {
    if (-not $source.Contains($Opt.SourceTransform.Find)) { throw 'SourceTransform: text to replace was not found in the tool source.' }
    $source = $source.Replace($Opt.SourceTransform.Find, $Opt.SourceTransform.Replace)
}
$typeMap = [ordered]@{
    'System.Windows.Forms.MessageBoxButtons' = 'SmokeUi.Buttons'
    'System.Windows.Forms.MessageBoxIcon'    = 'SmokeUi.Icon'
    'System.Windows.Forms.MessageBox'        = 'SmokeUi.MessageBox'
    'System.Windows.Forms.Application'       = 'SmokeUi.Application'
    'System.Windows.Forms.Clipboard'         = 'SmokeUi.Clipboard'
    'ADSI'                                   = 'SmokeUi.Adsi'
}
foreach ($k in $typeMap.Keys) { $source = [regex]::Replace($source, '\[\s*' + [regex]::Escape($k) + '\s*\]', "[$($typeMap[$k])]", 'IgnoreCase') }

$parseTokens = $null; $parseErrors = $null
$toolAst = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$parseTokens, [ref]$parseErrors)
if ($parseErrors) { throw "Tool source does not parse: $($parseErrors[0].Message)" }

$toolFunctions = @{}
foreach ($fn in $toolAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $null -eq $n.Parent.Parent.Parent }, $false)) {
    $toolFunctions[$fn.Name] = $fn
    . ([scriptblock]::Create($fn.Extent.Text))
}
function Get-SmokeAssignmentAst {
    param([string]$Left)
    $a = $toolAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq $Left }.GetNewClosure(), $true) | Select-Object -First 1
    if (-not $a) { throw "Assignment to $Left not found in the tool source." }
    return $a
}
function Get-SmokeHandlerAst {
    param([string]$Control, [string]$EventMember = 'Add_Click')
    $call = $toolAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
            $n.Member.Value -eq $EventMember -and $n.Expression.Extent.Text -eq $Control }.GetNewClosure(), $true) | Select-Object -First 1
    if (-not $call) { throw "Handler $Control.$EventMember not found in the tool source." }
    $sb = $call.Arguments[0].Find({ param($n) $n -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true)
    if (-not $sb) { throw "Handler $Control.$EventMember has no scriptblock." }
    return $sb.ScriptBlock
}
function ConvertTo-SmokeScriptBlock {
    # $this (the clicked control) does not exist outside WinForms, so row handlers read $smokeSender instead.
    param($ScriptBlockAst)
    $text = $ScriptBlockAst.Extent.Text
    $text = $text.Substring(1, $text.Length - 2) -replace '\$this\b', '$smokeSender'
    return [scriptblock]::Create($text)
}
$handlerAsts = [ordered]@{
    Preview    = Get-SmokeHandlerAst '$BtnPreview'
    Create     = Get-SmokeHandlerAst '$BtnCreate'
    PreviewAll = Get-SmokeHandlerAst '$BtnBulkPreview'
    CreateAll  = Get-SmokeHandlerAst '$BtnBulkCreate'
    Copy       = Get-SmokeHandlerAst '$BtnCopyAcct'
    CopyAll    = Get-SmokeHandlerAst '$BtnBulkCopyAll'
    Clear      = Get-SmokeHandlerAst '$BtnClear'
    BulkClear  = Get-SmokeHandlerAst '$BtnBulkClear'
    CancelSync = Get-SmokeHandlerAst '$BtnCancelSync'
    Close      = Get-SmokeHandlerAst '$Form' 'Add_FormClosing'
    Division   = Get-SmokeHandlerAst '$CboDivision' 'Add_SelectedIndexChanged'
    RowSkip    = Get-SmokeHandlerAst '$srow.SkipBtn'
    RowRetry   = Get-SmokeHandlerAst '$srow.RetryBtn'
    RowSyncSkip = Get-SmokeHandlerAst '$srow.SyncSkipBtn'
}
$createBodyAst = (Get-SmokeAssignmentAst '$Script:SingleCreateBody')
#endregion

#region Gate: nothing reachable may call an unknown command
$fakeCommands = @('Get-ADDomainController', 'Get-ADDomain', 'Get-ADObject', 'Get-ADUser', 'Set-ADUser',
    'New-PSSession', 'Import-PSSession', 'Remove-PSSession', 'Get-PSSession', 'Get-Recipient', 'New-RemoteMailbox', 'Get-RemoteMailbox', 'Set-RemoteMailbox',
    'Invoke-RestMethod', 'Start-Sleep', 'Get-GraphToken', 'Generate-AuditReport', 'Wait-UiDelay', 'New-TempPassword')
# Built-ins the reachable tool code may use. They only touch memory or the temp log folder.
$allowedBuiltins = @('Where-Object', 'ForEach-Object', 'Select-Object', 'Sort-Object', 'Out-Null', 'Get-Date', 'New-Object',
    'ConvertTo-SecureString', 'ConvertTo-Json', 'ConvertTo-Csv', 'Test-Path', 'New-Item', 'Split-Path', 'Join-Path')
$allowedDynamic = @('unverified', 'Script:SingleCreateBody')
$forbiddenTypes = 'Windows\.Forms|DirectoryServices|^adsi$|Net\.WebClient|Net\.Http|Runspaces|Diagnostics\.Process'

$gateProblems = New-Object System.Collections.Generic.List[string]
$visited = @{}
$queue = New-Object System.Collections.Generic.Queue[object]
foreach ($k in $handlerAsts.Keys) { $queue.Enqueue(@{ Name = "handler $k"; Ast = $handlerAsts[$k] }) }
$queue.Enqueue(@{ Name = 'SingleCreateBody'; Ast = $createBodyAst })
while ($queue.Count -gt 0) {
    $item = $queue.Dequeue()
    foreach ($cmd in $item.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $cmd.GetCommandName()
        if (-not $name) {
            $first = $cmd.CommandElements[0]
            if ($first -is [System.Management.Automation.Language.VariableExpressionAst] -and $allowedDynamic -contains $first.VariablePath.UserPath) { continue }
            $gateProblems.Add("$($item.Name): dynamic command call '$($cmd.Extent.Text.Split("`n")[0])'"); continue
        }
        $bare = $name -replace '^ActiveDirectory\\', ''
        if ($fakeCommands -contains $bare) { continue }
        if ($toolFunctions.ContainsKey($bare)) {
            if (-not $visited.ContainsKey($bare)) { $visited[$bare] = $true; $queue.Enqueue(@{ Name = "function $bare"; Ast = $toolFunctions[$bare] }) }
            continue
        }
        if ($allowedBuiltins -contains $bare) { continue }
        $gateProblems.Add("$($item.Name): calls '$name', which is not a tool function, a fake, or an allowed built-in")
    }
    foreach ($t in $item.Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TypeExpressionAst] -or $n -is [System.Management.Automation.Language.TypeConstraintAst] }, $true)) {
        if ($t.TypeName.FullName -match $forbiddenTypes) { $gateProblems.Add("$($item.Name): uses type [$($t.TypeName.FullName)], which has no stand-in") }
    }
}
if ($gateProblems.Count) {
    throw ("Smoke harness gate: the tool code reachable from the button handlers uses something the harness does not fake. " +
        "Add a fake (or review and allow it) in Tests\Smoke\SmokeWorld.ps1 before running. " + (($gateProblems | Select-Object -Unique) -join ' | '))
}
#endregion

#region Tool data and handler scriptblocks
. ([scriptblock]::Create((Get-SmokeAssignmentAst '$DivisionMap').Extent.Text))
. ([scriptblock]::Create((Get-SmokeAssignmentAst '$LicenseGroups').Extent.Text))
. ([scriptblock]::Create((Get-SmokeAssignmentAst '$ExchangeServer').Extent.Text))
. ([scriptblock]::Create($createBodyAst.Extent.Text))
$handlers = @{}
foreach ($k in $handlerAsts.Keys) { $handlers[$k] = ConvertTo-SmokeScriptBlock $handlerAsts[$k] }

$AuditLogPath = Join-Path $WorkDir 'ADProvisioningTool-Audit.csv'
$DebugLogPath = Join-Path $WorkDir 'ADProvisioningTool-Debug.csv'
$global:SmokeWorld.AUs = @($DivisionMap | ForEach-Object { "$($_.AU)".Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique |
        Where-Object { $Opt.MissingAUs -notcontains $_ } | ForEach-Object { [PSCustomObject]@{ id = "au-$_"; displayName = $_ } })
$global:SmokeWorld.Groups = @($LicenseGroups.Values | Where-Object { $Opt.MissingGroups -notcontains $_ } | ForEach-Object { [PSCustomObject]@{ id = "grp-$($_ -replace '\W', '')"; displayName = $_ } })
#endregion

#region Fakes that replace tool functions and timing
function Get-GraphToken {
    param([string]$ScriptDir, [switch]$Force)
    if ($Force) { $global:SmokeWorld.TokenForce++ }
    if ($global:SmokeWorld.Opt.TokenFails) { return $null }
    return @{ Authorization = 'Bearer smoke'; 'Content-Type' = 'application/json' }
}
function Generate-AuditReport { Add-SmokeCall 'Generate-AuditReport' @{} }
# A fixed fake, so test output is stable and no generated password reaches CI logs.
function New-TempPassword { 'Smoke-Fake-Pass-1!' }
function Wait-UiDelay { param([int]$Seconds) Add-SmokeCall 'Wait-UiDelay' @{ Seconds = $Seconds } }
function Invoke-SmokeRowButton {
    param([string]$Handler, [int]$Row, [string]$Button)
    $smokeSender = $Script:BulkStatusRows[$Row].$Button
    & $handlers[$Handler]
}
function Start-Sleep {
    # Counts instead of waiting, and plays the tech: answers Skip/Retry prompts and clicks mid-run buttons
    # through the tool's own handlers.
    param($Seconds, $Milliseconds)
    $w = $global:SmokeWorld
    $w.Sleeps++
    if ($w.Sleeps -gt 400000) { throw 'Smoke harness: a wait loop did not end.' }
    if ($Milliseconds) { [System.Threading.Thread]::Sleep([Math]::Min([int]$Milliseconds, 20)) }
    for ($row = 0; $row -lt 5; $row++) {
        $sr = $Script:BulkStatusRows[$row]
        if ($sr.RetryBtn.Visible -and $sr.RetryBtn.Enabled -and -not $sr.SkipRequested -and -not $sr.RetryNeeded) {
            if ($w.Opt.BulkOnPrompt -eq 'Skip') { Invoke-SmokeRowButton 'RowSkip' $row 'SkipBtn' } else { Invoke-SmokeRowButton 'RowRetry' $row 'RetryBtn' }
        }
        if ($w.Opt.BulkSyncSkipRows -contains $row -and $w.Sleeps -ge [int]$w.Opt.BulkSyncSkipAtSleep -and $sr.SyncSkipBtn.Visible -and $sr.SyncSkipBtn.Enabled -and -not $sr.SkipRequested) {
            Invoke-SmokeRowButton 'RowSyncSkip' $row 'SyncSkipBtn'
        }
    }
    if ($w.Opt.CancelSyncAtSleep -and $w.Sleeps -eq [int]$w.Opt.CancelSyncAtSleep -and $BtnCancelSync.Enabled) { & $handlers.CancelSync }
    if ($w.Opt.MidRunProbe -and -not $w.MidRun -and ($Script:SingleRunning -or $Script:BulkRunning)) {
        # A tech pressing Clear, Bulk Clear and the window X while a run is active.
        $before = "$($TxtFirst.Text)|$($bulkRows[0].First.Text)"
        & $handlers.Clear
        & $handlers.BulkClear
        $closeArgs = [PSCustomObject]@{ Cancel = $false }
        $closeArgs | ForEach-Object -Process $handlers.Close
        $w.MidRun = @{ CloseCancelled = $closeArgs.Cancel; InputsKept = ($before -eq "$($TxtFirst.Text)|$($bulkRows[0].First.Text)") }
    }
}
#endregion

#region Fake Exchange
function Test-SmokeExchangeSession {
    param([string]$Command)
    if ($global:SmokeWorld.OpenSessions -gt 0) { return $true }
    $global:SmokeWorld.Unexpected.Add("$Command was called with no Exchange session open")
    return $false
}
function New-PSSession {
    [CmdletBinding()] param($ConfigurationName, $ConnectionUri, $Authentication)
    $w = $global:SmokeWorld
    if ($w.Opt.SessionFailures -gt 0) { $w.Opt.SessionFailures--; Write-Error 'Connecting to remote server EXCH01 failed: WinRM cannot complete the operation.'; return }
    $w.SessionsOpened++; $w.OpenSessions++
    [PSCustomObject]@{ State = 'Opened'; ConfigurationName = $ConfigurationName; Id = $w.SessionsOpened }
}
function Import-PSSession { [CmdletBinding()] param([Parameter(Position = 0)]$Session, [switch]$DisableNameChecking, [switch]$AllowClobber) 'imported' }
function Remove-PSSession {
    [CmdletBinding()] param([Parameter(Position = 0, ValueFromPipeline = $true)]$Session)
    process { if ($Session -and $Session.State -eq 'Opened') { $Session.State = 'Closed'; $global:SmokeWorld.SessionsRemoved++; $global:SmokeWorld.OpenSessions-- } }
}
function Get-PSSession { [CmdletBinding()] param() }
function Find-SmokeMailbox {
    param([string]$Identity)
    $global:SmokeWorld.Ad | Where-Object { $_.Routing -and ($_.SamAccountName -eq $Identity -or $_.UserPrincipalName -eq $Identity -or $_.DistinguishedName -eq $Identity) } | Select-Object -First 1
}
function Get-Recipient {
    [CmdletBinding()] param($Filter)
    $w = $global:SmokeWorld
    if (-not (Test-SmokeExchangeSession 'Get-Recipient')) { Write-Error 'The term Get-Recipient is not recognized.'; return }
    if ("$Filter" -notmatch "^EmailAddresses -eq 'smtp:(.+)'$") { $w.Unexpected.Add("Get-Recipient filter not understood: $Filter"); return }
    $addr = $Matches[1]
    if ($w.Opt.ExchangeOnlyAddresses -contains $addr) { return [PSCustomObject]@{ Name = "ExchangeOnly $addr"; DistinguishedName = "CN=$addr,OU=Contacts" } }
    $w.Ad | Where-Object { $_.proxyAddresses -contains "smtp:$addr" } | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; DistinguishedName = $_.DistinguishedName } }
}
function New-RemoteMailbox {
    [CmdletBinding()] param($Name, $DisplayName, $FirstName, $LastName, $UserPrincipalName, $SamAccountName, $Password,
        $ResetPasswordOnNextLogon, $OnPremisesOrganizationalUnit, $RemoteRoutingAddress, $PrimarySmtpAddress, $DomainController)
    $w = $global:SmokeWorld
    $bound = @{}; foreach ($k in $PSBoundParameters.Keys) { $bound[$k] = $PSBoundParameters[$k] }
    Add-SmokeCall 'New-RemoteMailbox' $bound
    if (-not (Test-SmokeExchangeSession 'New-RemoteMailbox')) { Write-Error 'The term New-RemoteMailbox is not recognized.'; return }
    if ($Password -isnot [securestring] -or $Password.Length -eq 0) { Write-Error 'Cannot process argument transformation on parameter Password: a non-empty SecureString is required.'; return }
    foreach ($required in 'Name', 'UserPrincipalName', 'SamAccountName', 'OnPremisesOrganizationalUnit', 'RemoteRoutingAddress') {
        if (-not $PSBoundParameters[$required]) { Write-Error "Cannot bind argument to parameter '$required' because it is empty."; return }
    }
    if ("$SamAccountName".Length -gt 20) { Write-Error "The SamAccountName '$SamAccountName' is longer than 20 characters."; return }
    if ("$Name".Length -gt 64) { Write-Error "The name '$Name' is longer than 64 characters."; return }
    $behavior = $null
    if ($w.CreateCalls -lt @($w.Opt.CreateFailures).Count) { $behavior = @($w.Opt.CreateFailures)[$w.CreateCalls] }
    $w.CreateCalls++
    if ($behavior -eq 'Absent') { Write-Error 'The operation could not be performed because the domain controller is not available.'; return }
    $domain = if ("$DomainController" -like 'pdc.*') { "$DomainController".Substring(4) } else { "$DomainController" }
    if ($w.Ad | Where-Object { $_.SamAccountName -eq $SamAccountName -and $_.Domain -eq $domain }) { Write-Error "The object $SamAccountName already exists."; return }
    if ($w.Ad | Where-Object { $_.DistinguishedName -eq "CN=$Name,$OnPremisesOrganizationalUnit" }) { Write-Error "The object $Name already exists."; return }
    if ($w.Ad | Where-Object { $_.UserPrincipalName -eq $UserPrincipalName }) { Write-Error "The user principal name $UserPrincipalName is already in use."; return }
    $primary = "$PrimarySmtpAddress"
    if (-not $primary -or $w.Opt.PolicyNumbersPrimary) { $primary = ("$UserPrincipalName" -replace '@', '2@') }
    $obj = New-SmokeAdObject -Sam $SamAccountName -Upn $UserPrincipalName -Domain $domain -Name $Name -Ou $OnPremisesOrganizationalUnit -Primary $primary -Routing $RemoteRoutingAddress
    $w.Ad.Add($obj)
    if (-not $w.Opt.NeverSync) {
        $entraUpn = if ($w.Opt.SyncRenamesUpn) { ("$UserPrincipalName".Split('@')[0]) + '4821@contoso.onmicrosoft.com' } else { "$UserPrincipalName" }
        $w.Pending.Add([PSCustomObject]@{
                Polls = [int]$w.Opt.SyncAfterPolls
                User  = [PSCustomObject]@{ id = [guid]::NewGuid().ToString(); displayName = $DisplayName; userPrincipalName = $entraUpn; mail = $primary
                    onPremisesImmutableId = [Convert]::ToBase64String($obj.ObjectGUID.ToByteArray()); onPremisesSyncEnabled = $true
                    proxyAddresses = @($obj.proxyAddresses); IntendedUpn = "$UserPrincipalName" }
            })
    }
    if ($behavior -eq 'Exists') { Write-Error 'The WinRM client cannot process the request. The connection was closed.'; return }
    [PSCustomObject]@{ PrimarySmtpAddress = $primary; RecipientTypeDetails = 'RemoteUserMailbox'; Name = $Name }
}
function Get-RemoteMailbox {
    [CmdletBinding()] param($Identity, $DomainController)
    $w = $global:SmokeWorld
    if (-not (Test-SmokeExchangeSession 'Get-RemoteMailbox')) { Write-Error 'The term Get-RemoteMailbox is not recognized.'; return }
    $mb = Find-SmokeMailbox "$Identity"
    if (-not $mb) { Write-Error "The operation couldn't be performed because object '$Identity' couldn't be found."; return }
    $w.MailboxReads++
    $routing = "SMTP:$($mb.Routing)"
    if (@($w.Opt.BlankRoutingOnReads) -contains $w.MailboxReads) { $routing = $null }
    $primary = (@($mb.proxyAddresses | Where-Object { $_ -clike 'SMTP:*' })[0]) -replace '^SMTP:', ''
    [PSCustomObject]@{ PrimarySmtpAddress = $primary; EmailAddresses = @($mb.proxyAddresses); RemoteRoutingAddress = $routing
        RecipientTypeDetails = 'RemoteUserMailbox'; DistinguishedName = $mb.DistinguishedName; SamAccountName = $mb.SamAccountName }
}
function Set-RemoteMailbox {
    [CmdletBinding()] param($Identity, $EmailAddresses, $PrimarySmtpAddress, $EmailAddressPolicyEnabled, $DomainController)
    $w = $global:SmokeWorld
    $bound = @{}; foreach ($k in $PSBoundParameters.Keys) { $bound[$k] = $PSBoundParameters[$k] }
    Add-SmokeCall 'Set-RemoteMailbox' $bound
    if (-not (Test-SmokeExchangeSession 'Set-RemoteMailbox')) { Write-Error 'The term Set-RemoteMailbox is not recognized.'; return }
    $mb = Find-SmokeMailbox "$Identity"
    if (-not $mb) { Write-Error "The operation couldn't be performed because object '$Identity' couldn't be found."; return }
    $target = if ($PrimarySmtpAddress) { "$PrimarySmtpAddress" } else { "$($EmailAddresses.add)" }
    if ($w.Opt.SetMailboxFailPattern -and $target -match $w.Opt.SetMailboxFailPattern) { Write-Error "Set-RemoteMailbox failed for $target : Active Directory operation failed."; return }
    if ($PrimarySmtpAddress) {
        $kept = @($mb.proxyAddresses | ForEach-Object { if ($_ -clike 'SMTP:*') { 'smtp:' + $_.Substring(5) } else { $_ } } | Where-Object { $_ -ne "smtp:$PrimarySmtpAddress" })
        $mb.proxyAddresses.Clear(); [void]$mb.proxyAddresses.Add("SMTP:$PrimarySmtpAddress"); foreach ($k in $kept) { [void]$mb.proxyAddresses.Add($k) }
        $mb.mail = "$PrimarySmtpAddress"
    }
    if ($EmailAddresses -and $EmailAddresses.add) {
        # Real Exchange warns when nothing changes; ps2exe shows a warning as a blocking pop-up.
        if ($mb.proxyAddresses -contains "$($EmailAddresses.add)") { Write-Warning "The command completed successfully but no settings of '$Identity' have been modified." }
        else { [void]$mb.proxyAddresses.Add("$($EmailAddresses.add)") }
    }
}
#endregion

#region Fake Graph
function Invoke-RestMethod {
    [CmdletBinding()] param($Uri, $Method = 'GET', $Headers, $Body, $ContentType)
    $w = $global:SmokeWorld
    $u = [Uri]::UnescapeDataString("$Uri")
    Add-SmokeCall 'Invoke-RestMethod' @{ Uri = $u; Method = "$Method"; Body = $Body }
    if ($u -notlike 'https://graph.microsoft.com/v1.0/*') { $w.Unexpected.Add("Web request outside Graph: $u"); throw (New-Object SmokeUi.HttpException -ArgumentList 400, 'not graph') }
    if (-not $Headers -or -not $Headers.Authorization) { throw (New-Object SmokeUi.HttpException -ArgumentList 401, '{"error":{"code":"InvalidAuthenticationToken"}}') }

    if ($u -match '/v1\.0/users/([^/?]+)\?') {
        $upn = $Matches[1]
        if ($w.Opt.SyncHttpError) { throw (New-Object SmokeUi.HttpException -ArgumentList ([int]$w.Opt.SyncHttpError), '{"error":{"code":"Authorization_RequestDenied"}}') }
        foreach ($p in @($w.Pending | Where-Object { $_.User.IntendedUpn -eq $upn })) {
            $p.Polls--
            if ($p.Polls -le 0) { $w.Entra.Add($p.User); [void]$w.Pending.Remove($p) }
        }
        $hit = $w.Entra | Where-Object { $_.userPrincipalName -eq $upn } | Select-Object -First 1
        if ($hit) { return $hit }
        throw (New-Object SmokeUi.HttpException -ArgumentList 404, '{"error":{"code":"Request_ResourceNotFound"}}')
    }
    if ($u -match "/v1\.0/users\?\`$filter=onPremisesImmutableId eq '([^']+)'") {
        $id = $Matches[1]
        return [PSCustomObject]@{ value = @($w.Entra | Where-Object { $_.onPremisesImmutableId -eq $id }) }
    }
    if ($u -match "/v1\.0/users\?\`$filter=userPrincipalName eq '([^']+)'") {
        $upn = $Matches[1]
        return [PSCustomObject]@{ value = @($w.Entra | Where-Object { $_.userPrincipalName -eq $upn }) }
    }
    if ($u -match "/v1\.0/users\?\`$filter=proxyAddresses/any\(x:x eq 'smtp:([^']+)'\)") {
        $addr = $Matches[1]
        return [PSCustomObject]@{ value = @($w.Entra | Where-Object { $_.proxyAddresses -contains "smtp:$addr" -or $_.mail -eq $addr }) }
    }
    if ($u -match "/v1\.0/groups\?\`$filter=displayName eq '(.+?)'(&|$)") {
        $name = $Matches[1] -replace "''", "'"
        return [PSCustomObject]@{ value = @($w.Groups | Where-Object { $_.displayName -eq $name }) }
    }
    if ($u -match '/v1\.0/groups/([^/]+)/members/\$ref$' -and "$Method" -eq 'POST') {
        $groupId = $Matches[1]
        if ($w.Opt.LicensePostFails) { throw (New-Object SmokeUi.HttpException -ArgumentList 403, '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}') }
        if (-not ($w.Groups | Where-Object { $_.id -eq $groupId })) { throw (New-Object SmokeUi.HttpException -ArgumentList 404, '{"error":{"code":"Request_ResourceNotFound"}}') }
        $ref = "$((ConvertFrom-Json "$Body").'@odata.id')"
        $member = $w.Entra | Where-Object { $ref -eq "https://graph.microsoft.com/v1.0/directoryObjects/$($_.id)" } | Select-Object -First 1
        if (-not $member) { throw (New-Object SmokeUi.HttpException -ArgumentList 400, '{"error":{"code":"Request_BadRequest","message":"Invalid object reference in @odata.id"}}') }
        if ($w.Opt.AlreadyLicensed) { throw (New-Object SmokeUi.HttpException -ArgumentList 400, '{"error":{"code":"Request_BadRequest","message":"One or more added object references already exist for the following modified properties: members."}}') }
        return
    }
    if ($u -match '/v1\.0/directory/administrativeUnits\?') { return [PSCustomObject]@{ value = @($w.AUs) } }
    if ($u -match '/v1\.0/directory/administrativeUnits/([^/]+)/members/\$ref$' -and "$Method" -eq 'POST') {
        $auId = $Matches[1]
        if (-not ($w.AUs | Where-Object { $_.id -eq $auId })) { throw (New-Object SmokeUi.HttpException -ArgumentList 404, '{"error":{"code":"Request_ResourceNotFound"}}') }
        $ref = "$((ConvertFrom-Json "$Body").'@odata.id')"
        $member = $w.Entra | Where-Object { $ref -eq "https://graph.microsoft.com/v1.0/users/$($_.id)" } | Select-Object -First 1
        if (-not $member) { throw (New-Object SmokeUi.HttpException -ArgumentList 400, '{"error":{"code":"Request_BadRequest","message":"Unsupported resource type for AU member reference"}}') }
        if ($w.Opt.AuNotFoundTimes -gt 0) { $w.Opt.AuNotFoundTimes--; throw (New-Object SmokeUi.HttpException -ArgumentList 404, '{"error":{"code":"Request_ResourceNotFound","message":"Resource does not exist"}}') }
        return
    }
    $w.Unexpected.Add("Graph call not understood: $Method $u")
    throw (New-Object SmokeUi.HttpException -ArgumentList 400, '{"error":{"code":"BadRequest"}}')
}
#endregion

#region Tripwires: commands that must never run in a test
foreach ($danger in 'Invoke-WebRequest', 'Invoke-Command', 'Start-Process', 'Enter-PSSession', 'Start-Job', 'Send-MailMessage', 'Invoke-Expression', 'Import-Module', 'New-ADUser', 'Remove-ADUser', 'Set-ADAccountPassword') {
    Set-Item -Path "function:global:$danger" -Value ([scriptblock]::Create("`$global:SmokeWorld.Unexpected.Add('$danger was called'); throw 'Smoke harness: $danger is not allowed in a test run.'"))
}
# Every faked name must resolve to a fake: a function from this runspace, never an alias or a cmdlet.
foreach ($name in $fakeCommands) {
    $resolved = Get-Command -Name $name -ErrorAction SilentlyContinue | Select-Object -First 1
    $ok = $resolved -and $resolved.CommandType -eq 'Function' -and (-not $resolved.Module -or $resolved.Module.Name -eq 'ActiveDirectory')
    if ($ok -and $resolved.Module) { $ok = ($resolved.Module.Path -notlike '*.ps*' -and $resolved.Module.Path -notlike '*.dll') }
    if (-not $ok) { throw "Smoke harness: '$name' does not resolve to the harness fake (found: $($resolved.CommandType) $($resolved.Source)). Not running." }
}
#endregion

#region Stand-in window
function New-SmokeControl { param([string]$Text = '', [bool]$Enabled = $true) $c = New-Object SmokeUi.Control; $c.Text = $Text; $c.Enabled = $Enabled; return $c }
$ColorGold = 'Gold'; $ColorRed = 'Red'; $ColorLight = 'Light'; $ColorAccent = 'Accent'
$ColorRGB_100_40_40 = 'c1'; $ColorRGB_150_150_150 = 'c2'; $ColorRGB_180_60_60 = 'c3'
$Form = New-SmokeControl
$RichLog = New-SmokeControl
$TxtFirst = New-SmokeControl; $TxtMid = New-SmokeControl; $TxtLast = New-SmokeControl; $TxtDomain = New-SmokeControl
$TxtTicket = New-SmokeControl; $TxtTitle = New-SmokeControl; $TxtPreview = New-SmokeControl
$CboDivision = New-SmokeControl; $CboEmpType = New-SmokeControl; $CboCategory = New-SmokeControl; $CboAcctType = New-SmokeControl
$ChkDebug = New-SmokeControl
$BtnPreview = New-SmokeControl; $BtnCreate = New-SmokeControl -Enabled $false; $BtnCopyAcct = New-SmokeControl -Enabled $false
$BtnCancelSync = New-SmokeControl -Enabled $false
$LblCreateStatus = New-SmokeControl; $LblLastCreated = New-SmokeControl
$BtnBulkPreview = New-SmokeControl; $BtnBulkCreate = New-SmokeControl -Enabled $false; $BtnBulkCopyAll = New-SmokeControl -Enabled $false
$LblBulkStatus = New-SmokeControl
$Script:BulkRunning = $false; $Script:SingleRunning = $false; $Script:BulkResults = @(); $Script:ShowDebug = $false
$Script:CancelRequested = $false; $Script:AuditTicket = $null; $Script:LastAccountDetails = $null

function Get-SmokeDivisionItem { param([string]$Code) $d = $DivisionMap | Where-Object { $_.Code -eq $Code } | Select-Object -First 1; if (-not $d) { throw "Division $Code is not in the map." } "$($d.Code) - $($d.Name)" }
$bulkRows = @()
$Script:BulkStatusRows = @()
for ($n = 0; $n -lt 5; $n++) {
    $r = @{ First = (New-SmokeControl); Mid = (New-SmokeControl); Last = (New-SmokeControl); Ticket = (New-SmokeControl)
        Division = (New-SmokeControl); EmpType = (New-SmokeControl); Category = (New-SmokeControl); AcctType = (New-SmokeControl) }
    $r.Division.SelectedItem = '-- Select --'; $r.EmpType.SelectedItem = '-- Select --'
    $r.Category.SelectedItem = '1 - General'; $r.AcctType.SelectedItem = 'Primary'
    $bulkRows += $r
    $Script:BulkStatusRows += @{ NameLbl = (New-SmokeControl); StatusLbl = (New-SmokeControl 'Queued'); Console = (New-SmokeControl)
        SkipBtn = (New-SmokeControl); RetryBtn = (New-SmokeControl); SyncSkipBtn = (New-SmokeControl)
        Phase = 'queued'; SkipRequested = $false; RetryNeeded = $false; AccountData = $null; Created = $false }
    foreach ($b in 'SkipBtn', 'RetryBtn', 'SyncSkipBtn') { $Script:BulkStatusRows[$n].$b.Visible = $false; $Script:BulkStatusRows[$n].$b.Tag = $n }
}
if ($Opt.Mode -eq 'Bulk') {
    for ($n = 0; $n -lt @($Opt.People).Count; $n++) {
        $p = @($Opt.People)[$n]
        $bulkRows[$n].First.Text = $p.First; $bulkRows[$n].Mid.Text = "$($p.Mid)"; $bulkRows[$n].Last.Text = $p.Last; $bulkRows[$n].Ticket.Text = "$($p.Ticket)"
        $bulkRows[$n].Division.SelectedItem = Get-SmokeDivisionItem $p.Division; $bulkRows[$n].Division.SelectedIndex = 1
        $bulkRows[$n].EmpType.SelectedItem = $p.EmpType; $bulkRows[$n].EmpType.SelectedIndex = 1
    }
} else {
    $p = @($Opt.People)[0]
    $TxtFirst.Text = $p.First; $TxtMid.Text = "$($p.Mid)"; $TxtLast.Text = $p.Last; $TxtTicket.Text = "$($p.Ticket)"; $TxtTitle.Text = "$($p.Title)"
    $CboDivision.SelectedItem = Get-SmokeDivisionItem $p.Division; $CboDivision.SelectedIndex = 1
    # The tool's own division handler fills the Target Domain box.
    & $handlers.Division
    $CboEmpType.SelectedItem = $p.EmpType; $CboEmpType.SelectedIndex = 1
    $CboCategory.SelectedItem = '1 - General'; $CboAcctType.SelectedItem = 'Primary'
}
#endregion

#region Step runner and result
function Invoke-SmokeStep {
    # No try/catch here on purpose: an error must behave as it does in the exe, where a handler without
    # its own try block carries on after a failed statement.
    param($Step)
    if ($Step -is [hashtable]) {
        if ($Step.AddAd) { Add-SmokeExistingAd $Step.AddAd }
        if ($Step.SetOpt) { foreach ($k in $Step.SetOpt.Keys) { $global:SmokeWorld.Opt[$k] = $Step.SetOpt[$k] } }
        return 'ran'
    }
    $enabled = switch ($Step) {
        'Create' { $BtnCreate.Enabled } 'CreateAll' { $BtnBulkCreate.Enabled } 'Copy' { $BtnCopyAcct.Enabled } 'CopyAll' { $BtnBulkCopyAll.Enabled } default { $true }
    }
    if (-not $enabled) { return 'disabled' }
    if ($Step -eq 'Close') {
        $closeArgs = [PSCustomObject]@{ Cancel = $false }
        $closeArgs | ForEach-Object -Process $handlers.Close | Out-Null
        $global:SmokeWorld.CloseCancelled = $closeArgs.Cancel
    } else {
        & $handlers[$Step] | Out-Null
    }
    return 'ran'
}
function Get-SmokeStepState {
    [PSCustomObject]@{ CreateEnabled = $BtnCreate.Enabled; BulkCreateEnabled = $BtnBulkCreate.Enabled; CreateStatus = $LblCreateStatus.Text; BulkStatus = $LblBulkStatus.Text }
}
function Get-SmokeResult {
    $w = $global:SmokeWorld
    $audit = @(); if (Test-Path $AuditLogPath) { $audit = @(Import-Csv $AuditLogPath) }
    $debug = @(); if (Test-Path $DebugLogPath) { $debug = @(Import-Csv $DebugLogPath) }
    foreach ($c in [SmokeUi.Adsi]::Committed) {
        $dn = $c[0] -replace '^LDAP://[^/]+/', ''
        $o = $w.Ad | Where-Object { $_.DistinguishedName -eq $dn } | Select-Object -First 1
        if ($o) { $o.($c[1]) = $c[2] }
    }
    [PSCustomObject]@{
        Unexpected      = $w.Unexpected.ToArray()
        Audit           = $audit
        Debug           = $debug
        Log             = @($RichLog.Text -split "`n" | Where-Object { $_ })
        Preview         = $TxtPreview.Text
        CreateStatus    = $LblCreateStatus.Text
        CopyEnabled     = $BtnCopyAcct.Enabled
        CancelEnabled   = $BtnCancelSync.Enabled
        AccountDetails  = $Script:LastAccountDetails
        Clipboard       = [SmokeUi.Clipboard]::Text
        SingleRunning   = $Script:SingleRunning
        BulkRunning     = $Script:BulkRunning
        BulkStatus      = $LblBulkStatus.Text
        BulkRows        = @($Script:BulkStatusRows | ForEach-Object { [PSCustomObject]@{ Status = $_.StatusLbl.Text; Phase = $_.Phase; Created = $_.Created; Console = $_.Console.Text } })
        BulkResults     = @($Script:BulkResults)
        FirstNameBox    = $TxtFirst.Text
        Ad              = $w.Ad.ToArray()
        Entra           = $w.Entra.ToArray()
        Calls           = $w.Calls.ToArray()
        Dialogs         = [SmokeUi.MessageBox]::Shown.ToArray()
        SessionsOpened  = $w.SessionsOpened
        SessionsRemoved = $w.SessionsRemoved
        TokenForce      = $w.TokenForce
        MidRun          = $w.MidRun
        CloseCancelled  = $w.CloseCancelled
        SwallowedErrors = @($global:Error | ForEach-Object { "$_" })
    }
}
#endregion
