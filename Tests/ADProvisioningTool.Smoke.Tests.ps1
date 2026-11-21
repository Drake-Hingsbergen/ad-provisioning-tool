<#
    End-to-end smoke tests. Each test runs the real Preview / Create Account / Preview All / Create All
    handlers from ADProvisioningTool.ps1 through Tests\Smoke\Invoke-HandlerSmoke.ps1, which stands in for the
    window, AD, Exchange and Graph. The unit tests in ADProvisioningTool.Tests.ps1 cover the logic functions;
    these cover the handler code between them.

    Errors: errors that reach the error stream uncaught. In the compiled exe each one is a pop-up.
    Unexpected: calls the fake world did not understand (a sign the tool changed a query).
#>

BeforeAll {
    $script:toolPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'ADProvisioningTool.ps1'
    $script:harness  = Join-Path $PSScriptRoot 'Smoke\Invoke-HandlerSmoke.ps1'
    if (-not (Test-Path $harness)) { $script:harness = Join-Path $PSScriptRoot 'Smoke/Invoke-HandlerSmoke.ps1' }

    function Invoke-Smoke { param([hashtable]$Scenario = @{}) & $script:harness -Scenario $Scenario }
    function Get-AuditRow { param($Result, [string]$Action, [string]$Outcome) @($Result.Audit | Where-Object { $_.Action -eq $Action -and (-not $Outcome -or $_.Result -eq $Outcome) }) }
    function Get-Call { param($Result, [string]$Command) @($Result.Calls | Where-Object { $_.Command -eq $Command }) }
    function Assert-Clean {
        # -Strict also requires that nothing was caught and swallowed along the way. Only for scenarios
        # where the fake world raises no errors at all (no 404 while waiting, no failed lookups).
        param($Result, [switch]$Strict)
        if ($Strict) { $Result.SwallowedErrors | Should -BeNullOrEmpty -Because 'a clean run should not raise errors, even ones the tool catches' }
        $Result.Errors     | Should -BeNullOrEmpty -Because 'an uncaught error is a blocking pop-up in the exe'
        $Result.Unexpected | Should -BeNullOrEmpty -Because 'the fake world must understand every AD, Exchange and Graph call'
        $Result.SingleRunning | Should -BeFalse
        $Result.BulkRunning   | Should -BeFalse
        $Result.SessionsRemoved | Should -Be $Result.SessionsOpened -Because 'every Exchange session must be closed'
        foreach ($call in @(Get-Call $Result 'Set-RemoteMailbox')) {
            "$($call.Data.WarningAction)" | Should -Be 'SilentlyContinue' -Because 'an Exchange warning is a blocking pop-up in the exe'
        }
    }
    $script:twoPeople = @(
        @{ First = 'Ann'; Last = 'Lee';  Division = 'OPS'; EmpType = 'Knowledge Worker'; Ticket = 'T-1' }
        @{ First = 'Raj'; Last = 'Shah'; Division = 'FIN';  EmpType = 'Frontline Worker'; Ticket = 'T-2' }
    )
}

Describe 'Handler code guard' {
    It 'never asks where it is running from ($MyInvocation / $PSCommandPath)' {
        # Handler code runs inside scriptblocks, where these are empty. A Split-Path on one of them
        # raised "Cannot bind argument to parameter Path because it is null" at STEP 3.
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($toolPath, [ref]$null, [ref]$null)
        $hits = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $n.VariablePath.UserPath -in 'MyInvocation', 'PSCommandPath' }, $true) | ForEach-Object { "line $($_.Extent.StartLineNumber)" })
        $hits | Should -BeNullOrEmpty
    }

    It 'notices that bug when it is put back (the run stops before licensing and says so)' {
        $r = Invoke-Smoke @{ SourceTransform = @{
                Find    = '    if (-not $step3ScriptDir) { $step3ScriptDir = "C:\ADProvisioning" }'
                Replace = '    if (-not $step3ScriptDir) { $step3ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path }' } }
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'RUN_ABORTED').Count | Should -Be 1
        @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
        ($r.Log -join "`n") | Should -BeLike "*parameter 'Path' because it is null*"
    }
}

Describe 'Single create, end to end' {
    Context 'a normal account' {
        BeforeAll { $script:r = Invoke-Smoke }

        It 'runs Preview and Create without an uncaught error, or a swallowed one' { Assert-Clean $r -Strict }
        It 'passes a real password and forces a change at first logon' {
            $create = @(Get-Call $r 'New-RemoteMailbox')[0].Data
            $create.Password | Should -BeOfType [securestring]
            $create.ResetPasswordOnNextLogon | Should -BeTrue
            $create.OnPremisesOrganizationalUnit | Should -Be 'OU=New Accounts,DC=west,DC=contoso,DC=local'
            $create.RemoteRoutingAddress | Should -Be 'morgan.reyes@contoso.mail.onmicrosoft.com'
            $create.DomainController | Should -Be 'pdc.west.contoso.local'
        }
        It 'writes the debug log through the whole run' {
            @($r.Debug | Where-Object { $_.Step -eq 'PREFLIGHT_RESOLVED' }).Count | Should -BeGreaterThan 0
            @($r.Debug | Where-Object { $_.Step -eq 'NEW_REMOTE_MAILBOX' -and $_.Status -eq 'OK' }).Count | Should -Be 1
            @($r.Debug | Where-Object { $_.Step -eq 'SYNC_VERIFY' -and $_.Status -eq 'OK' }).Count | Should -Be 1
            @($r.Debug | Where-Object { $_.SAM -eq 'morgan.reyes' }).Count | Should -BeGreaterThan 5
        }
        It 'creates the mailbox with the primary address equal to the UPN' {
            $create = Get-Call $r 'New-RemoteMailbox'
            $create.Count | Should -Be 1
            $create[0].Data.PrimarySmtpAddress | Should -Be 'morgan.reyes@contoso.com'
            $create[0].Data.UserPrincipalName  | Should -Be 'morgan.reyes@contoso.com'
            $create[0].Data.SamAccountName     | Should -Be 'morgan.reyes'
            $create[0].Data.Name               | Should -Be 'Morgan Reyes - OPS'
        }
        It 'audits every step as OK' {
            @(Get-AuditRow $r 'CREATE' 'AD_OK').Count          | Should -Be 1
            @(Get-AuditRow $r 'ATTRS' 'OK').Count              | Should -Be 1
            @(Get-AuditRow $r 'MAILBOX_VERIFY' 'OK').Count     | Should -Be 1
            @(Get-AuditRow $r 'LICENSE' 'OK').Count            | Should -Be 1
            @(Get-AuditRow $r 'AU' 'OK').Count                 | Should -Be 1
            @($r.Audit | Where-Object { $_.Result -in 'FAILED', 'WARN' }) | Should -BeNullOrEmpty
        }
        It 'puts the ticket on every audit row' {
            @($r.Audit | Where-Object { $_.Details -notlike '*Ticket=INC-1001*' }) | Should -BeNullOrEmpty
        }
        It 'reports COMPLETE and enables Copy' {
            ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATION COMPLETE*'
            ($r.Log -join "`n") | Should -Not -BeLike '*FOLLOW-UP*'
            $r.CopyEnabled      | Should -BeTrue
            $r.AccountDetails   | Should -BeLike '*UPN: morgan.reyes@contoso.com*'
        }
        It 'leaves the account with routing, X500 and the custom attributes' {
            $acct = $r.Ad | Where-Object { $_.SamAccountName -eq 'morgan.reyes' }
            $acct.proxyAddresses | Should -Contain 'smtp:morgan.reyes@contoso.mail.onmicrosoft.com'
            @($acct.proxyAddresses | Where-Object { $_ -like 'X500:*' }).Count | Should -Be 1
            $acct.orgCategoryID  | Should -Be '1'
            $acct.orgAccountType | Should -Be 'Primary'
        }
        It 'does not re-add the routing address New-RemoteMailbox already set' {
            @(Get-Call $r 'Set-RemoteMailbox' | Where-Object { "$($_.Data.EmailAddresses.add)" -like 'smtp:*onmicrosoft.com' }) | Should -BeNullOrEmpty
        }
    }

    It 'takes the next name when first.last is held in another domain' {
        $r = Invoke-Smoke @{ ExistingAd = @(@{ Sam = 'morgan.reyes'; Upn = 'morgan.reyes@contoso.com'; Domain = 'central.contoso.local' }) }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.UserPrincipalName  | Should -Be 'morgan.reyes1@contoso.com'
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.PrimarySmtpAddress | Should -Be 'morgan.reyes1@contoso.com'
        @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
    }

    It 'adds the contoso.com alias for a vanity division' {
        $r = Invoke-Smoke @{ People = @(@{ First = 'Dana'; Last = 'Cole'; Division = 'SALES'; EmpType = 'Frontline Worker'; Ticket = 'INC-7' }) }
        Assert-Clean $r
        $acct = $r.Ad | Where-Object { $_.SamAccountName -eq 'dana.cole' }
        $acct.UserPrincipalName | Should -Be 'dana.cole@northwindtraders.com'
        $acct.proxyAddresses    | Should -Contain 'smtp:dana.cole@contoso.com'
        $acct.proxyAddresses    | Should -Contain 'SMTP:dana.cole@northwindtraders.com'
    }

    It 'corrects the primary address when the address policy numbered it' {
        $r = Invoke-Smoke @{ PolicyNumbersPrimary = $true }
        Assert-Clean $r
        @(Get-AuditRow $r 'MAILBOX' 'OK' | Where-Object { $_.Details -like '*Primary corrected to morgan.reyes@contoso.com*' }).Count | Should -Be 1
        ($r.Ad | Where-Object { $_.SamAccountName -eq 'morgan.reyes' }).proxyAddresses | Should -Contain 'SMTP:morgan.reyes@contoso.com'
        ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATION COMPLETE*'
    }

    It 'handles an apostrophe in the name' {
        $r = Invoke-Smoke @{ People = @(@{ First = 'Pat'; Last = "O'Neil"; Division = 'ITS'; EmpType = 'Knowledge Worker'; Ticket = 'INC-9' }) }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.SamAccountName | Should -Be 'pat.oneil'
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.Name           | Should -Be "Pat O'Neil - ITS"
    }

    It 'adds (sam) to the AD name when the same display name is already in the OU' {
        $r = Invoke-Smoke @{ ExistingAd = @(@{ Sam = 'morgan.reyes'; Upn = 'morgan.reyes@contoso.com'; Domain = 'west.contoso.local'
                    Name = 'Morgan Reyes - OPS'; Ou = 'OU=New Accounts,DC=west,DC=contoso,DC=local' }) }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.Name        | Should -Be 'Morgan Reyes - OPS (morgan.reyes1)'
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.DisplayName | Should -Be 'Morgan Reyes - OPS'
    }

    It 'sets the job title when one is entered' {
        $r = Invoke-Smoke @{ People = @(@{ First = 'Morgan'; Last = 'Reyes'; Division = 'OPS'; EmpType = 'Knowledge Worker'; Ticket = 'INC-1'; Title = 'Analyst' }) }
        Assert-Clean $r -Strict
        ($r.Ad | Where-Object { $_.SamAccountName -eq 'morgan.reyes' }).Title | Should -Be 'Analyst'
        ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATION COMPLETE*'
    }

    It 'takes the next name when the first one is taken between Preview and Create' {
        $r = Invoke-Smoke @{ Steps = @('Preview', @{ AddAd = @(@{ Sam = 'morgan.reyes'; Upn = 'morgan.reyes@contoso.com'; Domain = 'central.contoso.local' }) }, 'Create') }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.UserPrincipalName | Should -Be 'morgan.reyes1@contoso.com'
        @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
    }

    It 'counts a user who is already in the license group as licensed' {
        $r = Invoke-Smoke @{ AlreadyLicensed = $true }
        Assert-Clean $r
        @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
        ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATION COMPLETE*'
    }

    It 'retries the AU add while Graph has not caught up with the new user' {
        $r = Invoke-Smoke @{ AuNotFoundTimes = 2 }
        Assert-Clean $r
        @(Get-AuditRow $r 'AU' 'OK').Count | Should -Be 1
        @(Get-Call $r 'Wait-UiDelay' | Where-Object { $_.Data.Seconds -eq 15 }).Count | Should -Be 2
    }

    It 'verifies the mailbox on the retry when the routing address shows up late' {
        $r = Invoke-Smoke @{ BlankRoutingOnReads = @(2) }
        Assert-Clean $r
        ($r.Log -join "`n") | Should -BeLike '*Mailbox verified on retry*'
        @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
    }

    It 'copies the account details, clears the form and closes cleanly after a run' {
        $r = Invoke-Smoke @{ Steps = @('Preview', 'Create', 'Copy', 'Clear', 'Close') }
        Assert-Clean $r
        $r.Clipboard      | Should -BeLike '*UPN: morgan.reyes@contoso.com*'
        $r.FirstNameBox   | Should -BeNullOrEmpty
        $r.CloseCancelled | Should -BeFalse
        @(Get-AuditRow $r 'SESSION_END').Count | Should -Be 1
    }

    It 'ignores Clear and refuses to close while a run is active' {
        $r = Invoke-Smoke @{ MidRunProbe = $true }
        Assert-Clean $r
        $r.MidRun.CloseCancelled | Should -BeTrue
        $r.MidRun.InputsKept     | Should -BeTrue
        ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATION COMPLETE*'
    }

    Context 'when something goes wrong' {
        It 'reports a stop on an unexpected error after create: account exists, finish by hand' {
            $r = Invoke-Smoke @{ SourceTransform = @{
                    Find    = '    Write-Log $RichLog "STEP 4: Assigning M365 license group..." $ColorLight'
                    Replace = '    Split-Path -Parent $null' } }
            $r.Errors | Should -BeNullOrEmpty -Because 'the stop is caught and reported, not left as a pop-up'
            $r.SingleRunning | Should -BeFalse
            $row = @(Get-AuditRow $r 'PARTIAL_CREATE' 'RUN_ABORTED')
            $row.Count | Should -Be 1
            $row[0].Details | Should -BeLike 'Account exists*Ticket=INC-1001*'
            ($r.Log -join "`n") | Should -BeLike '*EXISTS in AD. Do NOT run Create again*'
            $r.CreateStatus | Should -BeLike 'Run stopped -- account exists*'
            $r.Steps[1].CreateEnabled | Should -BeFalse
        }
        It 'reports a stop on an unexpected error before create: nothing created' {
            $r = Invoke-Smoke @{ SourceTransform = @{
                    Find    = '    Write-Log $RichLog "STEP 1+2: Creating account and enabling mailbox via Exchange..." $ColorLight'
                    Replace = '    Split-Path -Parent $null' } }
            $r.Errors | Should -BeNullOrEmpty
            $r.Ad     | Should -BeNullOrEmpty
            @(Get-AuditRow $r 'CREATE' 'FAILED' | Where-Object { $_.Details -like 'Run stopped before create*' }).Count | Should -Be 1
            $r.CreateStatus | Should -BeLike 'Run stopped -- nothing created*'
        }
        It 'blocks Create when the Graph token is lost after Preview' {
            $r = Invoke-Smoke @{ Steps = @('Preview', @{ SetOpt = @{ TokenFails = $true } }, 'Create') }
            $r.Errors | Should -BeNullOrEmpty
            $r.Ad     | Should -BeNullOrEmpty
            @(Get-AuditRow $r 'CREATE' 'PREFLIGHT_BLOCKED').Count | Should -Be 1
        }
        It 'blocks Preview when no Global Catalog answers' {
            $r = Invoke-Smoke @{ GcUnavailable = $true }
            $r.Errors | Should -BeNullOrEmpty
            $r.CreateStatus | Should -BeLike '*No Global Catalog reachable*'
            ($r.Steps | Where-Object Step -eq 'Create').Ran | Should -BeFalse
        }
        It 'blocks Preview when all five name candidates are taken' {
            $taken = 'morgan.reyes', 'morgan.reyes1', 'morgan.reyes2', 'morgan.reyes3', 'morgan.reyes4' |
                ForEach-Object { @{ Sam = $_; Upn = "$_@contoso.com"; Domain = 'central.contoso.local' } }
            $r = Invoke-Smoke @{ ExistingAd = @($taken) }
            $r.Errors | Should -BeNullOrEmpty
            $r.CreateStatus | Should -BeLike '*All 5 UPN/email candidates are taken*'
            ($r.Steps | Where-Object Step -eq 'Create').Ran | Should -BeFalse
        }
        It 'flags follow-up when the job title cannot be set' {
            $r = Invoke-Smoke @{ TitleFails = $true; People = @(@{ First = 'Morgan'; Last = 'Reyes'; Division = 'OPS'; EmpType = 'Knowledge Worker'; Ticket = 'INC-1'; Title = 'Analyst' }) }
            Assert-Clean $r
            ($r.Log -join "`n") | Should -BeLike '*Follow-up: Title not set*'
            @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
        }
        It 'asks the tech when the mailbox still does not verify, and continues on Yes' {
            $r = Invoke-Smoke @{ BlankRoutingOnReads = @(2, 3); DialogAnswers = @('Yes', 'Yes') }
            Assert-Clean $r
            @(Get-AuditRow $r 'MAILBOX_VERIFY' 'MANUAL_PROCEED').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
        }
        It 'stops before licensing when the tech answers No to that question' {
            $r = Invoke-Smoke @{ BlankRoutingOnReads = @(2, 3); DialogAnswers = @('Yes', 'No') }
            Assert-Clean $r
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'STOPPED_STEP2').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
        }
        It 'stops the sync wait when the tech clicks Cancel Sync Wait' {
            $r = Invoke-Smoke @{ NeverSync = $true; CancelSyncAtSleep = 300 }
            Assert-Clean $r
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_CANCELLED').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
            $r.CancelEnabled | Should -BeFalse
        }
        It 'gives up on the AU after three not-found answers and flags follow-up' {
            $r = Invoke-Smoke @{ AuNotFoundTimes = 9 }
            Assert-Clean $r
            @(Get-AuditRow $r 'AU' 'FAILED').Count | Should -Be 1
            ($r.Log -join "`n") | Should -BeLike '*FOLLOW-UP NEEDED*'
        }
        It 'blocks Preview without a Graph token and creates nothing' {
            $r = Invoke-Smoke @{ TokenFails = $true }
            $r.Errors | Should -BeNullOrEmpty
            ($r.Steps | Where-Object Step -eq 'Create').Ran | Should -BeFalse
            $r.Ad | Should -BeNullOrEmpty
            $r.CreateStatus | Should -BeLike '*Graph token unavailable*'
        }
        It 'rejects an accented name at Preview' {
            $r = Invoke-Smoke @{ People = @(@{ First = ('Jos' + [char]0x00E9); Last = 'Lee'; Division = 'ITS'; EmpType = 'Knowledge Worker' }) }
            $r.Errors | Should -BeNullOrEmpty
            "$($r.Dialogs)" | Should -BeLike '*accent*'
            ($r.Steps | Where-Object Step -eq 'Create').Ran | Should -BeFalse
            $r.Ad | Should -BeNullOrEmpty
        }
        It 'says nothing was created, and hands out no details, when New-RemoteMailbox fails cleanly' {
            $r = Invoke-Smoke @{ CreateFailures = @('Absent') }
            Assert-Clean $r
            $r.Ad | Should -BeNullOrEmpty
            @(Get-AuditRow $r 'CREATE' 'FAILED').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE').Count         | Should -Be 0
            $r.CopyEnabled    | Should -BeFalse
            $r.AccountDetails | Should -BeNullOrEmpty
            $r.CreateStatus   | Should -BeLike 'Nothing created*'
        }
        It 'audits OBJECT_EXISTS when New-RemoteMailbox errors after writing the account' {
            $r = Invoke-Smoke @{ CreateFailures = @('Exists') }
            Assert-Clean $r
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'OBJECT_EXISTS').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
            ($r.Log -join "`n") | Should -BeLike '*Do NOT run Create again*'
            @($r.Ad).Count | Should -Be 1
        }
        It 'says nothing was created when the Exchange session cannot open' {
            $r = Invoke-Smoke @{ SessionFailures = 99 }
            $r.Errors | Should -BeNullOrEmpty
            $r.Ad     | Should -BeNullOrEmpty
            $r.CopyEnabled | Should -BeFalse
        }
        It 'continues and flags follow-up when sync renamed the UPN' {
            $r = Invoke-Smoke @{ SyncRenamesUpn = $true }
            Assert-Clean $r
            @(Get-AuditRow $r 'SYNC' 'UPN_CHANGED').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE' 'OK').Count       | Should -Be 1
            @(Get-AuditRow $r 'AU' 'OK').Count            | Should -Be 1
            ($r.Log -join "`n") | Should -BeLike '*ACCOUNT CREATED -- FOLLOW-UP NEEDED*'
        }
        It 'stops with SYNC_AUTH_FAILED, not a timeout, when Graph answers 403' {
            $r = Invoke-Smoke @{ SyncHttpError = 403 }
            Assert-Clean $r
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_AUTH_FAILED').Count | Should -Be 1
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_TIMEOUT').Count     | Should -Be 0
            @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
        }
        It 'asks for a new token once on 401, then stops with SYNC_AUTH_FAILED' {
            $r = Invoke-Smoke @{ SyncHttpError = 401 }
            Assert-Clean $r
            $r.TokenForce | Should -Be 1
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_AUTH_FAILED').Count | Should -Be 1
        }
        It 'times out after 45 minutes of polls when the account never syncs' {
            $r = Invoke-Smoke @{ NeverSync = $true }
            Assert-Clean $r
            @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_TIMEOUT').Count | Should -Be 1
            @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
        }
        It 'warns at Preview and flags follow-up when the AU name is not in Entra' {
            $r = Invoke-Smoke @{ MissingAUs = @('OPS') }
            Assert-Clean $r
            $r.Preview | Should -BeLike "*AU 'OPS' not found in Entra*"
            @(Get-AuditRow $r 'AU' 'FAILED').Count  | Should -Be 1
            @(Get-AuditRow $r 'LICENSE' 'OK').Count | Should -Be 1
            ($r.Log -join "`n") | Should -BeLike '*FOLLOW-UP NEEDED*'
        }
        It 'flags follow-up when the license group add is refused' {
            $r = Invoke-Smoke @{ LicensePostFails = $true }
            Assert-Clean $r
            @(Get-AuditRow $r 'LICENSE' 'FAILED').Count | Should -Be 1
            @(Get-AuditRow $r 'AU' 'OK').Count          | Should -Be 1
            ($r.Log -join "`n") | Should -BeLike '*FOLLOW-UP NEEDED*'
        }
        It 'flags follow-up when the custom attributes cannot be written' {
            $r = Invoke-Smoke @{ AdsiFails = $true }
            Assert-Clean $r
            @(Get-AuditRow $r 'ATTRS' 'FAILED').Count | Should -Be 1
            ($r.Log -join "`n") | Should -BeLike '*FOLLOW-UP NEEDED*'
        }
        It 'flags follow-up, without a pop-up, when the X500 add fails' {
            $r = Invoke-Smoke @{ SetMailboxFailPattern = '^X500:' }
            Assert-Clean $r
            @(Get-AuditRow $r 'ALIAS' 'WARN').Count | Should -BeGreaterThan 0
            ($r.Log -join "`n") | Should -BeLike '*Follow-up: Add the X500 address*'
        }
        It 'flags follow-up, without a pop-up, when the contoso.com alias add fails' {
            $r = Invoke-Smoke @{ People = @(@{ First = 'Dana'; Last = 'Cole'; Division = 'SALES'; EmpType = 'Frontline Worker'; Ticket = 'INC-7' })
                SetMailboxFailPattern = '^smtp:dana\.cole@contoso\.com$' }
            Assert-Clean $r
            ($r.Log -join "`n") | Should -BeLike '*Follow-up: Add contoso.com alias dana.cole@contoso.com*'
        }
    }
}

Describe 'Bulk create, end to end' {
    Context 'two normal rows' {
        BeforeAll { $script:b = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople } }

        It 'runs Preview All and Create All without an uncaught error, or a swallowed one' { Assert-Clean $b -Strict }
        It 'completes both rows and says so' {
            $b.BulkRows[0].Status | Should -Be 'COMPLETE'
            $b.BulkRows[1].Status | Should -Be 'COMPLETE'
            $b.BulkStatus | Should -Be 'Complete. 2 fully provisioned, 0 need follow-up, 0 not created.'
        }
        It 'sets each primary address to the UPN' {
            foreach ($c in @(Get-Call $b 'New-RemoteMailbox')) { $c.Data.PrimarySmtpAddress | Should -Be $c.Data.UserPrincipalName }
        }
        It 'keeps each row on its own ticket in the audit log' {
            @($b.Audit | Where-Object { $_.TargetSAM -eq 'ann.lee'  -and $_.Details -notlike '*Ticket=T-1*' }) | Should -BeNullOrEmpty
            @($b.Audit | Where-Object { $_.TargetSAM -eq 'raj.shah' -and $_.Details -notlike '*Ticket=T-2*' }) | Should -BeNullOrEmpty
        }
        It 'licenses each row with the group for its employee type and adds its AU' {
            @(Get-AuditRow $b 'LICENSE' 'OK' | Where-Object { $_.TargetSAM -eq 'ann.lee'  }).Details | Should -BeLike '*Knowledge Worker*'
            @(Get-AuditRow $b 'LICENSE' 'OK' | Where-Object { $_.TargetSAM -eq 'raj.shah' }).Details | Should -BeLike '*Frontline Worker*'
            @(Get-AuditRow $b 'AU' 'OK').Count | Should -Be 2
        }
    }

    It 'gives two rows with the same name different addresses' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @(
                @{ First = 'Ann'; Last = 'Lee'; Division = 'OPS'; EmpType = 'Knowledge Worker'; Ticket = 'T-1' }
                @{ First = 'Ann'; Last = 'Lee'; Division = 'FIN';  EmpType = 'Frontline Worker'; Ticket = 'T-2' }) }
        Assert-Clean $r
        @($r.Ad | ForEach-Object UserPrincipalName | Sort-Object) | Should -Be @('ann.lee@contoso.com', 'ann.lee1@contoso.com')
        $r.BulkStatus | Should -BeLike 'Complete. 2 fully provisioned*'
    }

    It 'retries a failed create when the tech clicks Retry, and makes one account' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); CreateFailures = @('Absent'); BulkOnPrompt = 'Retry' }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox').Count | Should -Be 2
        @($r.Ad).Count | Should -Be 1
        $r.BulkRows[0].Status | Should -Be 'COMPLETE'
    }

    It 'skips a failed create when the tech clicks Skip, and counts it as not created' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; CreateFailures = @('Absent'); BulkOnPrompt = 'Skip' }
        Assert-Clean $r
        $r.BulkRows[0].Status | Should -Be 'SKIPPED'
        $r.BulkRows[1].Status | Should -Be 'COMPLETE'
        $r.BulkStatus | Should -Be 'Complete. 1 fully provisioned, 0 need follow-up, 1 not created.'
    }

    It 'never re-creates a row whose account was written, even after Preview All runs again' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); CreateFailures = @('Exists'); Steps = @('PreviewAll', 'CreateAll', 'PreviewAll', 'CreateAll') }
        Assert-Clean $r
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'OBJECT_EXISTS').Count | Should -Be 1
        @(Get-Call $r 'New-RemoteMailbox').Count | Should -Be 1 -Because 'a second create would make ann.lee1'
        @($r.Ad).Count | Should -Be 1
        $r.BulkRows[0].Created | Should -BeTrue
        $r.BulkRows[0].Console | Should -BeLike '*Already created in this run*'
    }

    It 'does not re-create completed rows when Preview All and Create All are run a second time' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; Steps = @('PreviewAll', 'CreateAll', 'PreviewAll', 'CreateAll') }
        Assert-Clean $r
        @(Get-Call $r 'New-RemoteMailbox').Count | Should -Be 2
        @($r.Ad).Count | Should -Be 2
    }

    It 'picks a new name at create time when Exchange alone holds the contoso.com alias' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @(@{ First = 'Dana'; Last = 'Cole'; Division = 'SALES'; EmpType = 'Frontline Worker'; Ticket = 'T-9' })
            ExchangeOnlyAddresses = @('dana.cole@contoso.com') }
        Assert-Clean $r
        $r.BulkRows[0].Console | Should -BeLike '*Names changed since Preview All*'
        @(Get-Call $r 'New-RemoteMailbox')[0].Data.UserPrincipalName | Should -Not -Be 'dana.cole@northwindtraders.com'
        $r.BulkRows[0].Status | Should -Be 'COMPLETE'
    }

    It 'marks a row PARTIAL with a note when sync renamed its UPN' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); SyncRenamesUpn = $true }
        Assert-Clean $r
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
        @(Get-AuditRow $r 'SYNC' 'UPN_CHANGED').Count | Should -Be 1
        @(Get-AuditRow $r 'LICENSE' 'OK').Count       | Should -Be 1
        "$($r.BulkResults)" | Should -BeLike '*Follow-up:*'
    }

    It 'stops polling a row on 403 instead of waiting 45 minutes' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); SyncHttpError = 403 }
        Assert-Clean $r
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_AUTH_FAILED').Count | Should -Be 1
        @(Get-Call $r 'Invoke-RestMethod' | Where-Object { $_.Data.Uri -like '*/users/ann.lee@contoso.com?*' }).Count | Should -Be 1
    }

    It 'marks a row PARTIAL when its AU is not in Entra' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); MissingAUs = @('OPS') }
        Assert-Clean $r
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
        @(Get-AuditRow $r 'AU' 'FAILED').Count | Should -Be 1
    }

    It 'marks a row PARTIAL when a Phase 1 address add fails' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); SetMailboxFailPattern = '^X500:' }
        Assert-Clean $r
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
        "$($r.BulkResults)" | Should -BeLike '*not added*'
    }

    It 'stops polling a row after three 401 answers in a row' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); SyncHttpError = 401 }
        Assert-Clean $r
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_AUTH_FAILED').Count | Should -Be 1
        $r.TokenForce | Should -Be 2
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
    }

    It 'skips the sync wait for a row when the tech clicks its Skip Sync button' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; NeverSync = $true; BulkSyncSkipRows = @(0, 1) }
        Assert-Clean $r
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_CANCELLED').Count | Should -Be 2
        @(Get-AuditRow $r 'LICENSE').Count | Should -Be 0
        $r.BulkRows[0].Status | Should -Be 'PARTIAL'
    }

    It 'skips the sync wait when Skip Sync is clicked between two checks' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = @($twoPeople[0]); NeverSync = $true; BulkSyncSkipRows = @(0); BulkSyncSkipAtSleep = 200 }
        Assert-Clean $r
        @(Get-AuditRow $r 'PARTIAL_CREATE' 'SYNC_CANCELLED').Count | Should -Be 1
        @(Get-Call $r 'Invoke-RestMethod' | Where-Object { $_.Data.Uri -like '*/users/ann.lee@contoso.com?*' }).Count | Should -Be 1
    }

    It 'copies all results, then Bulk Clear unlocks the rows' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; Steps = @('PreviewAll', 'CreateAll', 'CopyAll', 'BulkClear') }
        Assert-Clean $r
        $r.Clipboard | Should -BeLike '*ann.lee@contoso.com*raj.shah@contoso.com*'
        $r.BulkRows[0].Created | Should -BeFalse
        $r.BulkRows[0].Phase   | Should -Be 'queued'
    }

    It 'ignores Bulk Clear and refuses to close while a bulk run is active' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; MidRunProbe = $true }
        Assert-Clean $r
        $r.MidRun.CloseCancelled | Should -BeTrue
        $r.MidRun.InputsKept     | Should -BeTrue
        $r.BulkStatus | Should -BeLike 'Complete. 2 fully provisioned*'
    }

    It 'reports every row as not created when Exchange never connects and the tech skips' {
        $r = Invoke-Smoke @{ Mode = 'Bulk'; People = $twoPeople; SessionFailures = 999; BulkOnPrompt = 'Skip' }
        $r.Errors | Should -BeNullOrEmpty
        $r.Ad     | Should -BeNullOrEmpty
        $r.BulkRunning | Should -BeFalse
        $r.BulkStatus  | Should -BeLike 'No accounts were created: 2 not created*'
    }
}
