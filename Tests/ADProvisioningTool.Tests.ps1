#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
.SYNOPSIS
    Pester tests for the provisioning logic in ADProvisioningTool.ps1.
.DESCRIPTION
    ADProvisioningTool.ps1 builds its WinForms GUI as soon as it runs, so it cannot be dot-sourced.
    These tests parse the script, load only the logic functions and the division map, and mock
    AD, Exchange and Graph. Nothing here touches a real directory.
.EXAMPLE
    Invoke-Pester -Path .\Tests\ADProvisioningTool.Tests.ps1 -Output Detailed
.NOTES
    Requires Pester 5 or later. Runs on Windows PowerShell 5.1 and PowerShell 7.
#>

BeforeAll {
    $scriptPath = Join-Path $PSScriptRoot '..\ADProvisioningTool.ps1'
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors) { throw "ADProvisioningTool.ps1 has parse errors: $($parseErrors[0].Message)" }

    $logicFunctions = @(
        'Get-AUCodes', 'Get-DivCodeFromItem', 'Get-GraphErrorDetail', 'Wait-UiDelay', 'Get-AdministrativeUnitList', 'Get-EntraTargetProblem', 'Add-UserToAUs',
        'Get-OUPath', 'Get-EmailDomain', 'Format-ProperCase', 'Get-NameCandidate', 'Truncate-SAM',
        'Get-SamCandidate', 'Get-ForestGlobalCatalog', 'Test-GraphAddressTaken', 'Get-BatchClash', 'Get-BatchExclusion',
        'New-TempPassword', 'Resolve-AccountCandidate',
        'ConvertTo-LdapFilterValue', 'Test-ProxyAddressPresent', 'Get-NewRemoteMailboxParameter', 'Get-SyncedEntraUser',
        'Get-CreatedObjectState', 'Get-NameInputProblem', 'Get-NameLengthProblem',
        'Get-ArcChallengePath', 'Get-ArcManagedIdentityToken', 'Get-GraphToken'
    )
    $definitions = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) |
        Where-Object { $logicFunctions -contains $_.Name }
    $missing = $logicFunctions | Where-Object { ($definitions | ForEach-Object Name) -notcontains $_ }
    if ($missing) { throw "Functions not found in ADProvisioningTool.ps1: $($missing -join ', ')" }
    foreach ($definition in $definitions) { . ([scriptblock]::Create($definition.Extent.Text)) }

    # The first assignment is the literal map; a later one only re-sorts it.
    $mapAssignment = $ast.FindAll({ param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$DivisionMap' }, $true) | Select-Object -First 1
    . ([scriptblock]::Create($mapAssignment.Extent.Text))

    # Stand-ins for commands that only exist on the management server, so they can be mocked.
    function DBG { param($Step, $Status, $Detail) }
    function Get-ADUser { param($Filter, $Server, $Properties, $ErrorAction) }
    function Get-Recipient { param($Filter, $ErrorAction) }
    function Get-ADObject { param($Server, $SearchBase, $LDAPFilter, $Filter, $Properties, $ErrorAction) }
    function Get-ADDomain { param($Identity, $Server, $ErrorAction) }
    function Get-ADDomainController { param([switch]$Discover, $Service, $ErrorAction) }

    $script:ValidAdDomains = 'east', 'west', 'central' |
        ForEach-Object { "$_.contoso.local" }
}

Describe 'Division map' {
    It 'has no duplicate codes' {
        $duplicates = $DivisionMap | Group-Object Code | Where-Object Count -gt 1 | ForEach-Object Name
        $duplicates | Should -BeNullOrEmpty
    }

    It 'gives every division a valid AD domain, a name and an AU' {
        foreach ($division in $DivisionMap) {
            $division.Domain | Should -BeIn $ValidAdDomains -Because "$($division.Code) must use a known AD domain"
            $division.Name   | Should -Not -BeNullOrEmpty -Because "$($division.Code) needs a display name"
            $division.AU     | Should -Not -BeNullOrEmpty -Because "$($division.Code) needs an AU"
        }
    }

    It 'only sets EmailDomain for vanity domains' {
        foreach ($division in ($DivisionMap | Where-Object { $_.EmailDomain })) {
            $division.EmailDomain | Should -Not -Be 'contoso.com' -Because "$($division.Code) should omit EmailDomain for contoso.com"
            $division.EmailDomain | Should -Match '^[a-z0-9.-]+\.[a-z]{2,}$'
        }
    }

    It 'contains only ASCII so ps2exe can compile it' {
        foreach ($division in $DivisionMap) {
            "$($division.Code)$($division.Name)$($division.AU)" | Should -Not -Match '[^\x00-\x7F]'
        }
    }

    It '<Code> uses AD domain <Domain>, email <Email> and AU <AU>' -TestCases @(
        @{ Code = 'FIN'; Domain = 'east.contoso.local'; Email = 'contoso.com'; AU = 'FIN' }
        @{ Code = 'ITS'; Domain = 'central.contoso.local'; Email = 'contoso.com'; AU = 'ITS' }
        @{ Code = 'OPS'; Domain = 'west.contoso.local'; Email = 'contoso.com'; AU = 'OPS' }
        @{ Code = 'MKT'; Domain = 'west.contoso.local'; Email = 'fabrikam.com'; AU = 'MKT' }
        @{ Code = 'SALES'; Domain = 'east.contoso.local'; Email = 'northwindtraders.com'; AU = 'SALES' }
        @{ Code = 'FIELD SVC'; Domain = 'east.contoso.local'; Email = 'contoso.com'; AU = 'FIELD,OPS' }
    ) {
        param($Code, $Domain, $Email, $AU)
        $division = $DivisionMap | Where-Object Code -eq $Code
        $division             | Should -Not -BeNullOrEmpty
        $division.Domain      | Should -Be $Domain
        Get-EmailDomain $Code | Should -Be $Email
        $division.AU          | Should -Be $AU
    }

    It 'does not offer an AU name as a division code: <Code>' -TestCases @(
        @{ Code = 'RD' }; @{ Code = 'FIELD' }
    ) {
        param($Code)
        $DivisionMap | Where-Object Code -eq $Code | Should -BeNullOrEmpty
    }
}

Describe 'Get-AUCodes' {
    It 'returns <Expected> for <Code>' -TestCases @(
        @{ Code = 'FIELD SVC'; Expected = @('FIELD', 'OPS') }
        @{ Code = 'R-D'; Expected = @('RD') }
        @{ Code = 'SALES'; Expected = @('SALES') }
    ) {
        param($Code, $Expected)
        @(Get-AUCodes $Code) | Should -Be $Expected
    }

    It 'falls back to the division code when the code is not in the map' {
        @(Get-AUCodes 'NOTREAL') | Should -Be @('NOTREAL')
    }
}

Describe 'Get-DivCodeFromItem' {
    It 'reads the code back from every dropdown item in the map' {
        foreach ($division in $DivisionMap) {
            Get-DivCodeFromItem "$($division.Code) - $($division.Name)" | Should -Be $division.Code
        }
    }

    It 'keeps codes that contain spaces or hyphens intact' {
        Get-DivCodeFromItem 'FIELD SVC - Field Service' | Should -Be 'FIELD SVC'
        Get-DivCodeFromItem 'R-D - Research and Development' | Should -Be 'R-D'
    }
}

Describe 'Get-NameCandidate' {
    It 'matches the agreed order with no middle initial' {
        $c = 1..5 | ForEach-Object { Get-NameCandidate -First 'Harper' -Middle '' -Last 'Winslow' -Attempt $_ }
        $c | Should -Be @('harper.winslow', 'harper.winslow1', 'harper.winslow2', 'harper.winslow3', 'harper.winslow4')
    }

    It 'matches the agreed order with a middle initial' {
        $c = 1..5 | ForEach-Object { Get-NameCandidate -First 'Harper' -Middle 'J' -Last 'Winslow' -Attempt $_ }
        $c | Should -Be @('harper.winslow', 'harper.j.winslow', 'harper.winslow1', 'harper.winslow2', 'harper.winslow3')
    }

    It 'ignores a middle field that is not a single letter (<Middle>)' -TestCases @(
        @{ Middle = 'None' }, @{ Middle = 'na' }, @{ Middle = 'Robert' }
    ) {
        param($Middle)
        Get-NameCandidate -First 'John' -Middle $Middle -Last 'Smith' -Attempt 2 | Should -Be 'john.smith1'
    }

    It 'accepts an initial typed with a period' {
        Get-NameCandidate -First 'John' -Middle 'Q.' -Last 'Smith' -Attempt 2 | Should -Be 'john.q.smith'
    }

    It 'strips punctuation and spaces from names' {
        Get-NameCandidate -First "Mary Ann" -Middle '' -Last "O'Brien-Lee" -Attempt 1 | Should -Be 'maryann.obrienlee'
    }

    It 'keeps long names at full length (the SAM is cut separately)' {
        Get-NameCandidate -First 'Christopher' -Middle '' -Last 'Lancaster' -Attempt 3 | Should -Be 'christopher.lancaster2'
    }
}

Describe 'Truncate-SAM' {
    It 'leaves usernames of 20 characters or fewer alone' {
        Truncate-SAM 'john.smith' | Should -Be 'john.smith'
    }

    It 'cuts to 20 characters like the HR onboarding workflow' {
        Truncate-SAM 'christopher.lancaster' | Should -Be 'christopher.lancaste'
    }

    It 'never ends a SAM with a period' {
        Truncate-SAM 'abcdefghijklmnopqrs.tuv' | Should -Be 'abcdefghijklmnopqrs'
    }
}

Describe 'Get-SamCandidate (documented examples)' {
    It '<LocalPart> index <Index> gives <Expected>' -TestCases @(
        @{ LocalPart = 'harper.winslow';        Index = 0; Expected = 'harper.winslow' }
        @{ LocalPart = 'christopher.lancaster';   Index = 0; Expected = 'christopher.lancaste' }
        @{ LocalPart = 'christopher.lancaster';   Index = 1; Expected = 'christopher.lancast1' }
        @{ LocalPart = 'christopher.j.lancaster'; Index = 0; Expected = 'christopher.j.lancas' }
        @{ LocalPart = 'christopher.lancaster2';  Index = 0; Expected = 'christopher.lancaste' }
        @{ LocalPart = 'john.smith';              Index = 3; Expected = 'john.smith3' }
    ) {
        param($LocalPart, $Index, $Expected)
        Get-SamCandidate -LocalPart $LocalPart -Index $Index | Should -Be $Expected
    }

    It 'gives 10 distinct SAMs for a long name' {
        $sams = 0..9 | ForEach-Object { Get-SamCandidate -LocalPart 'christopher.lancaster' -Index $_ }
        @($sams | Select-Object -Unique).Count | Should -Be 10
        $sams | ForEach-Object { $_.Length | Should -BeLessOrEqual 20 }
    }
}

Describe 'Format-ProperCase' {
    It 'formats <InputName> as <Expected>' -TestCases @(
        @{ InputName = 'john';        Expected = 'John' }
        @{ InputName = "o'brien";     Expected = "O'Brien" }
        @{ InputName = 'SMITH-REYES'; Expected = 'Smith-Reyes' }
        @{ InputName = 'mary ann';    Expected = 'Mary Ann' }
        @{ InputName = '';            Expected = '' }
        @{ InputName = 'McDonald';    Expected = 'McDonald' }
        @{ InputName = 'mcdonald';    Expected = 'McDonald' }
        @{ InputName = 'MCDONALD';    Expected = 'McDonald' }
        @{ InputName = 'MacDonald';   Expected = 'MacDonald' }
        @{ InputName = 'DeLuca';      Expected = 'DeLuca' }
        @{ InputName = 'van der Berg'; Expected = 'van der Berg' }
        @{ InputName = 'smith iii';   Expected = 'Smith III' }
        @{ InputName = 'PAUL IV';     Expected = 'Paul IV' }
        @{ InputName = "O'NEIL";      Expected = "O'Neil" }
        @{ InputName = "d'angelo";    Expected = "D'Angelo" }
        @{ InputName = "'t hooft";    Expected = "'T Hooft" }
        @{ InputName = '-smith';      Expected = '-Smith' }
        @{ InputName = '  mary   ann '; Expected = 'Mary Ann' }
        @{ InputName = 'SANTA CRUZ'; Expected = 'Santa Cruz' }
    ) {
        param($InputName, $Expected)
        Format-ProperCase $InputName | Should -BeExactly $Expected
    }
}

Describe 'ConvertTo-LdapFilterValue' {
    It 'escapes the characters LDAP filters reserve' {
        ConvertTo-LdapFilterValue "Pat O'Neil (Jr)" | Should -BeExactly "Pat O'Neil \28Jr\29"
        ConvertTo-LdapFilterValue 'a*b\c' | Should -BeExactly 'a\2ab\5cc'
    }
}

Describe 'Test-ProxyAddressPresent' {
    It 'finds an address regardless of prefix case and address case' {
        Test-ProxyAddressPresent @('SMTP:Pat.Lee@contoso.com', 'smtp:pat.lee@contoso.mail.onmicrosoft.com') 'Pat.Lee@Contoso.Mail.OnMicrosoft.com' | Should -BeTrue
    }
    It 'ignores X500 entries and missing addresses' {
        Test-ProxyAddressPresent @('X500:/o=Contoso/cn=pat.lee@contoso.com', 'SMTP:other@contoso.com') 'pat.lee@contoso.com' | Should -BeFalse
        Test-ProxyAddressPresent $null 'pat.lee@contoso.com' | Should -BeFalse
    }
}

Describe 'Get-NewRemoteMailboxParameter' {
    BeforeAll {
        $script:r = @{ UPN = 'harper.winslow1@contoso.com'; SAM = 'harper.winslow1'; PrimarySmtp = 'harper.winslow1@contoso.com'; Routing = 'harper.winslow1@contoso.mail.onmicrosoft.com' }
    }
    It 'pins the primary email to the UPN for a contoso.com division (address policy would pick First.Last2)' {
        $p = Get-NewRemoteMailboxParameter -Resolved $r -CnName 'Harper Winslow - OPS' -DispName 'Harper Winslow - OPS' -First 'Harper' -Last 'Winslow' -Ou 'OU=x' -Pdc 'pdc01' -Password $null
        $p.PrimarySmtpAddress   | Should -Be 'harper.winslow1@contoso.com'
        $p.UserPrincipalName    | Should -Be 'harper.winslow1@contoso.com'
        $p.RemoteRoutingAddress | Should -Be 'harper.winslow1@contoso.mail.onmicrosoft.com'
        $p.SamAccountName       | Should -Be 'harper.winslow1'
        $p.DomainController     | Should -Be 'pdc01'
        $p.ResetPasswordOnNextLogon | Should -BeTrue
    }
    It 'uses the vanity address as primary for vanity divisions' {
        $v = @{ UPN = 'pat.lee@fabrikam.com'; SAM = 'pat.lee'; PrimarySmtp = 'pat.lee@fabrikam.com'; Routing = 'pat.lee@contoso.mail.onmicrosoft.com' }
        (Get-NewRemoteMailboxParameter -Resolved $v -CnName 'c' -DispName 'd' -First 'Pat' -Last 'Lee' -Ou 'o' -Pdc 'p' -Password $null).PrimarySmtpAddress | Should -Be 'pat.lee@fabrikam.com'
    }
}

Describe 'Get-SyncedEntraUser' {
    BeforeAll {
        function New-HttpError([int]$Code) {
            $ex = New-Object System.Exception "HTTP $Code"
            $ex | Add-Member -NotePropertyName Response -NotePropertyValue ([PSCustomObject]@{ StatusCode = $Code })
            return $ex
        }
    }
    It 'reports Synced when the UPN object carries our ImmutableId' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 'u1'; userPrincipalName = 'pat.lee@contoso.com'; onPremisesImmutableId = 'IMM' } }
        $r = Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{}
        $r.State   | Should -Be 'Synced'
        $r.User.id | Should -Be 'u1'
    }
    It 'reports WrongObject when the UPN belongs to someone else' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ id = 'u9'; userPrincipalName = 'pat.lee@contoso.com'; onPremisesImmutableId = $null } }
        $r = Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{}
        $r.State  | Should -Be 'WrongObject'
        $r.Detail | Should -BeLike '*cloud-only*'
    }
    It 'finds the account by ImmutableId when sync gave it another UPN' {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/users/pat.lee%40contoso.com?*' } -MockWith { throw (New-HttpError 404) }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*onPremisesImmutableId*' } -MockWith {
            [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'u1'; userPrincipalName = 'pat.lee1234@contoso.onmicrosoft.com'; onPremisesImmutableId = 'IMM' }) }
        }
        $r = Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{}
        $r.State  | Should -Be 'UpnChanged'
        $r.Detail | Should -BeLike '*pat.lee1234@contoso.onmicrosoft.com*'
    }
    It 'finds the renamed account when another object holds the UPN (how Entra Connect renames)' {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/users/pat.lee%40contoso.com?*' } -MockWith {
            [PSCustomObject]@{ id = 'u9'; userPrincipalName = 'pat.lee@contoso.com'; onPremisesImmutableId = $null }
        }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*onPremisesImmutableId*' } -MockWith {
            [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'u1'; userPrincipalName = 'pat.lee1234@contoso.onmicrosoft.com'; onPremisesImmutableId = 'IMM' }) }
        }
        $r = Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{}
        $r.State   | Should -Be 'UpnChanged'
        $r.User.id | Should -Be 'u1'
    }
    It 'reports NotYet when neither lookup finds it' {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/users/pat.lee%40contoso.com?*' } -MockWith { throw (New-HttpError 404) }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*onPremisesImmutableId*' } -MockWith { [PSCustomObject]@{ value = @() } }
        (Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{}).State | Should -Be 'NotYet'
    }
    It 'passes a 401 up so the caller can refresh the token' {
        Mock Invoke-RestMethod { throw (New-HttpError 401) }
        { Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'IMM' -Headers @{} } | Should -Throw '*401*'
    }
    It 'URL-encodes the ImmutableId (base64 can hold + / =)' {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/users/pat.lee%40contoso.com?*' } -MockWith { throw (New-HttpError 404) }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*onPremisesImmutableId*' } -MockWith { [PSCustomObject]@{ value = @() } }
        Get-SyncedEntraUser -Upn 'pat.lee@contoso.com' -ImmutableId 'a+b/c==' -Headers @{} | Out-Null
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Uri -like '*a%2Bb%2Fc%3D%3D*' } -Times 1
    }
}

Describe 'Get-OUPath' {
    It 'builds the staging OU for a child domain' {
        Get-OUPath 'west.contoso.local' | Should -Be 'OU=New Accounts,DC=west,DC=contoso,DC=local'
    }
}

Describe 'Add-UserToAUs' {
    BeforeEach {
        Mock Invoke-RestMethod -ParameterFilter { $Method -ne 'POST' } -MockWith {
            [PSCustomObject]@{ value = @(
                    [PSCustomObject]@{ id = 'id-field';   displayName = 'FIELD' }
                    [PSCustomObject]@{ id = 'id-ops'; displayName = 'OPS' }
                ) }
        }
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith { }
    }

    It 'adds the user to every AU for a dual-AU division' {
        $results = @(Add-UserToAUs -UserId 'user-1' -AUNames @('FIELD', 'OPS') -Headers @{})
        $results.Count                    | Should -Be 2
        ($results | Where-Object { -not $_.OK }) | Should -BeNullOrEmpty
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -Times 2 -Exactly
    }

    It 'references the user through the users/ path that Graph requires' {
        Add-UserToAUs -UserId 'user-1' -AUNames @('FIELD') -Headers @{} | Out-Null
        Should -Invoke Invoke-RestMethod -ParameterFilter {
            $Method -eq 'POST' -and $Uri -like '*/directory/administrativeUnits/id-field/members/$ref' -and
            $Body -match 'graph.microsoft.com/v1.0/users/user-1'
        } -Times 1 -Exactly
    }

    It 'matches AU names regardless of case' {
        $results = @(Add-UserToAUs -UserId 'user-1' -AUNames @('ops') -Headers @{})
        $results[0].OK | Should -BeTrue
    }

    It 'reports an AU that does not exist in Entra without stopping the others' {
        $results = @(Add-UserToAUs -UserId 'user-1' -AUNames @('MISSING', 'FIELD') -Headers @{})
        ($results | Where-Object AU -eq 'MISSING').OK | Should -BeFalse
        ($results | Where-Object AU -eq 'FIELD').OK     | Should -BeTrue
    }

    It 'treats an existing membership as success so a retry does not fail' {
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith {
            throw 'One or more added object references already exist for the following modified properties'
        }
        $results = @(Add-UserToAUs -UserId 'user-1' -AUNames @('FIELD') -Headers @{})
        $results[0].OK     | Should -BeTrue
        $results[0].Detail | Should -Be 'already a member'
    }

    It 'marks every AU as failed when the AU list cannot be read' {
        Mock Invoke-RestMethod -ParameterFilter { $Method -ne 'POST' } -MockWith { throw 'Graph unavailable' }
        $results = @(Add-UserToAUs -UserId 'user-1' -AUNames @('FIELD', 'OPS') -Headers @{})
        $results.Count | Should -Be 2
        ($results | Where-Object OK) | Should -BeNullOrEmpty
    }
}

Describe 'Add-UserToAUs retry and paging' {
    BeforeEach {
        Mock Start-Sleep { }
        Mock Invoke-RestMethod -ParameterFilter { $Method -ne 'POST' -and $Uri -notlike 'https://next*' } -MockWith {
            [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'id-a'; displayName = 'AUA' }); '@odata.nextLink' = 'https://next/page2' }
        }
        Mock Invoke-RestMethod -ParameterFilter { $Method -ne 'POST' -and $Uri -like 'https://next*' } -MockWith {
            [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'id-b'; displayName = 'AUB' }) }
        }
    }
    It 'finds an AU on the second page of the AU list' {
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith { }
        (Add-UserToAUs -UserId 'u1' -AUNames @('AUB') -Headers @{})[0].OK | Should -BeTrue
    }
    It 'retries when the new user is not visible to the AU service yet' {
        $script:posts = 0
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith {
            $script:posts++
            if ($script:posts -lt 3) { throw 'Request_ResourceNotFound: Resource does not exist' }
        }
        $r = Add-UserToAUs -UserId 'u1' -AUNames @('AUA') -Headers @{} -RetryDelaySeconds 0
        $r[0].OK | Should -BeTrue
        $script:posts | Should -Be 3
    }
    It 'gives up after the retry limit and reports the error' {
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith { throw 'Request_ResourceNotFound: Resource does not exist' }
        $r = Add-UserToAUs -UserId 'u1' -AUNames @('AUA') -Headers @{} -RetryCount 2 -RetryDelaySeconds 0
        $r[0].OK     | Should -BeFalse
        $r[0].Detail | Should -BeLike '*Request_ResourceNotFound*'
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -Times 2 -Exactly
    }
    It 'does not retry other errors' {
        Mock Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -MockWith { throw 'Authorization_RequestDenied' }
        $r = Add-UserToAUs -UserId 'u1' -AUNames @('AUA') -Headers @{} -RetryDelaySeconds 0
        $r[0].OK | Should -BeFalse
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Method -eq 'POST' } -Times 1 -Exactly
    }
}

Describe 'Resolve-AccountCandidate' {
    BeforeEach {
        $script:ForestGC = $null
        Mock DBG { }
        Mock Get-ADDomainController { [PSCustomObject]@{ HostName = @('gc01.contoso.local') } }
        Mock Get-ADDomain { [PSCustomObject]@{ PDCEmulator = 'pdc01.east.contoso.local' } }
        Mock Get-ADObject { }
        Mock Get-Recipient { }
        Mock Invoke-RestMethod { [PSCustomObject]@{ value = @() } }
        $script:p = @{ First = 'John'; Middle = ''; Last = 'Smith'; Domain = 'east.contoso.local'; EmailDomain = 'contoso.com'; GraphHeaders = @{ Authorization = 'Bearer t' } }
    }

    Context 'resolving' {
        It 'returns the first candidate when nothing is taken' {
            $r = Resolve-AccountCandidate @p -EmailDomain 'fabrikam.com'
            $r.Status     | Should -Be 'Resolved'
            $r.SAM        | Should -Be 'john.smith'
            $r.UPN        | Should -Be 'john.smith@fabrikam.com'
            $r.TenantAlias | Should -Be 'john.smith@contoso.com'
            $r.Routing    | Should -Be 'john.smith@contoso.mail.onmicrosoft.com'
            $r.PrimarySmtp | Should -Be $r.UPN
            $r.Attempt    | Should -Be 1
            $r.PDC        | Should -Be 'pdc01.east.contoso.local'
        }

        It 'does not add a contoso.com alias for contoso.com divisions' {
            (Resolve-AccountCandidate @p).TenantAlias | Should -BeNullOrEmpty
        }

        It 'searches the whole forest through the Global Catalog' {
            Resolve-AccountCandidate @p | Out-Null
            Should -Invoke Get-ADObject -ParameterFilter { $Server -eq 'gc01.contoso.local:3268' -and $SearchBase -eq '' -and $LDAPFilter -like '*userPrincipalName=john.smith@contoso.com*' }
        }

        It 'checks the SAM on the PDC across all object types' {
            Resolve-AccountCandidate @p | Out-Null
            Should -Invoke Get-ADObject -ParameterFilter { $Server -eq 'pdc01.east.contoso.local' -and $LDAPFilter -eq '(sAMAccountName=john.smith)' }
        }

        It 'keeps the full UPN for a long name and cuts only the SAM' {
            $r = Resolve-AccountCandidate @p -First 'Christopher' -Last 'Lancaster'
            $r.UPN       | Should -Be 'christopher.lancaster@contoso.com'
            $r.Routing   | Should -Be 'christopher.lancaster@contoso.mail.onmicrosoft.com'
            $r.SAM       | Should -Be 'christopher.lancaste'
            $r.Truncated | Should -BeTrue
        }

        It 'moves to a numbered SAM when only the SAM is taken in the domain, keeping the UPN' {
            Mock Get-ADObject -ParameterFilter { $LDAPFilter -eq '(sAMAccountName=christopher.lancaste)' } -MockWith { [PSCustomObject]@{ DistinguishedName = 'CN=Existing,DC=x' } }
            $r = Resolve-AccountCandidate @p -First 'Christopher' -Last 'Lancaster'
            $r.UPN | Should -Be 'christopher.lancaster@contoso.com'
            $r.SAM | Should -Be 'christopher.lancast1'
        }

        It 'moves to the next UPN when the forest already has it (another domain, synced or not)' {
            Mock Get-ADObject -ParameterFilter { $LDAPFilter -like '*userPrincipalName=john.smith@contoso.com)*' } -MockWith { [PSCustomObject]@{ DistinguishedName = 'CN=Existing,DC=x' } }
            $r = Resolve-AccountCandidate @p
            $r.UPN     | Should -Be 'john.smith1@contoso.com'
            $r.SAM     | Should -Be 'john.smith1'
            $r.Attempt | Should -Be 2
        }

        It 'moves to the next UPN when Entra has it (cloud-only account)' {
            Mock Invoke-RestMethod -ParameterFilter { $Uri -like "*userPrincipalName eq 'john.smith%40contoso.com'*" } -MockWith {
                [PSCustomObject]@{ value = @([PSCustomObject]@{ id = '1'; userPrincipalName = 'john.smith@contoso.com' }) }
            }
            (Resolve-AccountCandidate @p).UPN | Should -Be 'john.smith1@contoso.com'
        }

        It 'checks routing addresses in both cases (SMTP: and smtp:)' {
            Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*SMTP%3Ajohn.smith%40contoso.mail.onmicrosoft.com*' } -MockWith {
                [PSCustomObject]@{ value = @([PSCustomObject]@{ id = '1'; userPrincipalName = 'someone@contoso.com' }) }
            }
            (Resolve-AccountCandidate @p).Routing | Should -Be 'john.smith1@contoso.mail.onmicrosoft.com'
        }

        It 'moves on when on-prem Exchange holds the contoso.com alias' {
            Mock Get-Recipient -ParameterFilter { $Filter -like '*john.smith@contoso.com*' } -MockWith { [PSCustomObject]@{ Name = 'Existing' } }
            (Resolve-AccountCandidate @p -EmailDomain 'fabrikam.com' -ExchSessionActive).UPN | Should -Be 'john.smith1@fabrikam.com'
        }

        It 'skips local parts and SAMs claimed by other rows in a bulk batch' {
            $r = Resolve-AccountCandidate @p -ExcludeLocalParts @('john.smith') -ExcludeSAMs @('john.smith1')
            $r.UPN | Should -Be 'john.smith1@contoso.com'
            $r.SAM | Should -Be 'john.smith11'
        }
    }

    Context 'exhausted' {
        It 'reports Exhausted when all 5 UPN candidates are taken, naming who holds each' {
            Mock Get-ADObject -ParameterFilter { $Server -like '*:3268' } -MockWith { [PSCustomObject]@{ DistinguishedName = 'CN=Existing,DC=x' } }
            $r = Resolve-AccountCandidate @p
            $r.Status | Should -Be 'Exhausted'
            $r.SAM    | Should -BeNullOrEmpty
            $r.Reason | Should -BeLike '*john.smith (AD: CN=Existing,DC=x)*john.smith4 (AD: CN=Existing,DC=x)*'
        }

        It 'never checks the same SAM twice (20-character names ending in a digit)' {
            Mock Get-ADObject -ParameterFilter { $LDAPFilter -like '(sAMAccountName=*' } -MockWith { [PSCustomObject]@{ DistinguishedName = 'CN=x' } }
            Resolve-AccountCandidate @p -First 'Amelia' -Last 'Featherstone1' | Out-Null
            Should -Invoke Get-ADObject -ParameterFilter { $LDAPFilter -eq '(sAMAccountName=amelia.featherstone1)' } -Times 1 -Exactly
        }

        It 'reports Exhausted when the SAM and all 9 numbered variants are taken' {
            Mock Get-ADObject -ParameterFilter { $LDAPFilter -like '(sAMAccountName=*' } -MockWith { [PSCustomObject]@{ DistinguishedName = 'CN=Existing,DC=x' } }
            (Resolve-AccountCandidate @p).Status | Should -Be 'Exhausted'
        }
    }

    Context 'fails closed when a check cannot run' {
        It 'stops when there is no Graph token' {
            $r = Resolve-AccountCandidate @p -GraphHeaders $null
            $r.Status | Should -Be 'Unverified'
            $r.Reason | Should -BeLike '*Graph token*'
        }

        It 'stops when no Global Catalog can be found' {
            Mock Get-ADDomainController { throw 'no GC' }
            (Resolve-AccountCandidate @p).Status | Should -Be 'Unverified'
        }

        It 'stops when the PDC cannot be reached' {
            Mock Get-ADDomain { throw 'The server is not operational' }
            (Resolve-AccountCandidate @p).Status | Should -Be 'Unverified'
        }

        It 'stops when the forest query errors (unreachable DC)' {
            Mock Get-ADObject -ParameterFilter { $Server -like '*:3268' } -MockWith { throw 'The server is not operational' }
            $r = Resolve-AccountCandidate @p
            $r.Status | Should -Be 'Unverified'
            $r.Reason | Should -BeLike '*not operational*'
        }

        It 'stops when a Graph lookup errors' {
            Mock Invoke-RestMethod { throw '403 Forbidden' }
            (Resolve-AccountCandidate @p).Status | Should -Be 'Unverified'
        }

        It 'stops when the Exchange lookup errors' {
            Mock Get-Recipient { throw 'session lost' }
            (Resolve-AccountCandidate @p -EmailDomain 'fabrikam.com' -ExchSessionActive).Status | Should -Be 'Unverified'
        }

        It 'stops when the SAM lookup errors' {
            Mock Get-ADObject -ParameterFilter { $LDAPFilter -like '(sAMAccountName=*' } -MockWith { throw 'timeout' }
            (Resolve-AccountCandidate @p).Status | Should -Be 'Unverified'
        }
    }
}

Describe 'Bulk batch claims' {
    BeforeEach {
        $script:local = @{ 'john.smith' = 0 }
        $script:sams  = @{ 'east.contoso.local|john.smith' = 0 }
    }

    It 'flags the same name in a different email domain (routing address is shared)' {
        Get-BatchClash -ClaimedLocal $local -ClaimedSam $sams -LocalPart 'john.smith' -Domain 'west.contoso.local' -Sam 'john.smith' | Should -Be 0
    }

    It 'flags the same SAM in the same AD domain' {
        Get-BatchClash -ClaimedLocal $local -ClaimedSam $sams -LocalPart 'jon.smith' -Domain 'east.contoso.local' -Sam 'john.smith' | Should -Be 0
    }

    It 'allows the same SAM in a different AD domain when the name differs' {
        Get-BatchClash -ClaimedLocal $local -ClaimedSam $sams -LocalPart 'john.smith1' -Domain 'west.contoso.local' -Sam 'john.smith' | Should -BeNullOrEmpty
    }

    It 'excludes every claimed name but only same-domain SAMs' {
        $sams['west.contoso.local|mary.reyes'] = 1
        $x = Get-BatchExclusion -ClaimedLocal $local -ClaimedSam $sams -Domain 'east.contoso.local'
        $x.LocalParts | Should -Be @('john.smith')
        $x.SAMs       | Should -Be @('john.smith')
    }
}

Describe 'New-TempPassword' {
    It 'is 16 characters with upper, lower, digit and symbol' {
        foreach ($n in 1..50) {
            $p = New-TempPassword
            $p.Length | Should -Be 16
            $p | Should -MatchExactly '[A-Z]'
            $p | Should -MatchExactly '[a-z]'
            $p | Should -MatchExactly '[0-9]'
            $p | Should -Match '[!#%*+\-=?@]'
            $p | Should -Not -MatchExactly '[0O1lI]'
        }
    }
    It 'does not repeat' {
        @(1..200 | ForEach-Object { New-TempPassword } | Select-Object -Unique).Count | Should -Be 200
    }
}

Describe 'Get-ArcManagedIdentityToken' {
    BeforeEach {
        $env:IDENTITY_ENDPOINT = $null
        $script:keyFile = Join-Path $TestDrive 'challenge.key'
        Set-Content -Path $keyFile -Value 'key123' -NoNewline
        Mock Get-ArcChallengePath { $keyFile }
    }

    It 'answers the challenge with the key file contents and returns the token' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = 'tok-abc' } }
        Get-ArcManagedIdentityToken | Should -Be 'tok-abc'
        Should -Invoke Invoke-RestMethod -Times 1 -ParameterFilter {
            $Headers.Authorization -eq 'Basic key123' -and $Headers.Metadata -eq 'true' -and
            $Uri -like 'http://localhost:40342/metadata/identity/oauth2/token?api-version=2020-06-01&resource=https%3A%2F%2Fgraph.microsoft.com'
        }
    }

    It 'uses IDENTITY_ENDPOINT when the agent sets it' {
        $env:IDENTITY_ENDPOINT = 'http://localhost:9999/metadata/identity/oauth2/token'
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = 'tok' } }
        Get-ArcManagedIdentityToken | Out-Null
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Uri -like 'http://localhost:9999/*' }
        $env:IDENTITY_ENDPOINT = $null
    }

    It 'explains the group requirement when the key file cannot be read' {
        Mock Get-ArcChallengePath { Join-Path $TestDrive 'missing.key' }
        { Get-ArcManagedIdentityToken } | Should -Throw "*Hybrid agent extension applications*"
    }

    It 'fails when the agent returns no token' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = $null } }
        { Get-ArcManagedIdentityToken } | Should -Throw "*no access token*"
    }
}

Describe 'Get-ArcChallengePath' {
    It 'reports an unreachable agent' {
        Mock Invoke-WebRequest { throw (New-Object System.Net.WebException 'Unable to connect') }
        { Get-ArcChallengePath -Uri 'http://localhost:40342/x' } | Should -Throw "*not reachable*"
    }
}

Describe 'Get-GraphToken' {
    BeforeEach {
        $script:cfgDir = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $cfgDir | Out-Null
        Mock DBG { }
    }

    It 'uses the Arc managed identity when AuthMode is ManagedIdentity' {
        '{"AuthMode":"ManagedIdentity","TenantId":"t1"}' | Set-Content (Join-Path $cfgDir 'ADProvisioningTool.json')
        Mock Get-ArcManagedIdentityToken { 'mi-token' }
        Mock Invoke-RestMethod { throw 'secret flow must not run' }
        $headers = Get-GraphToken -ScriptDir $cfgDir
        $headers.Authorization | Should -Be 'Bearer mi-token'
        Should -Invoke Invoke-RestMethod -Times 0
    }

    It 'uses the client secret when AuthMode is not set' {
        '{"TenantId":"t1","ClientId":"c1","ClientSecret":"s1","IsEncrypted":false}' | Set-Content (Join-Path $cfgDir 'ADProvisioningTool.json')
        Mock Get-ArcManagedIdentityToken { throw 'MI must not run' }
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = 'secret-token' } }
        $headers = Get-GraphToken -ScriptDir $cfgDir
        $headers.Authorization | Should -Be 'Bearer secret-token'
        Should -Invoke Invoke-RestMethod -ParameterFilter { $Body.client_secret -eq 's1' -and $Uri -like '*/t1/oauth2/v2.0/token' }
    }

    It 'returns null and logs the reason when the identity fails' {
        '{"AuthMode":"ManagedIdentity"}' | Set-Content (Join-Path $cfgDir 'ADProvisioningTool.json')
        Mock Get-ArcManagedIdentityToken { throw 'Cannot read Arc challenge file' }
        Get-GraphToken -ScriptDir $cfgDir | Should -BeNullOrEmpty
        Should -Invoke DBG -ParameterFilter { $Step -eq 'GRAPH_TOKEN' -and $Status -eq 'FAIL' -and $Detail -like '*challenge file*' }
    }

    It 'returns null and logs when no config exists' {
        Get-GraphToken -ScriptDir $cfgDir | Should -BeNullOrEmpty
        Should -Invoke DBG -ParameterFilter { $Detail -like 'No config file found*' }
    }
}

Describe 'Get-GraphToken cache' {
    BeforeEach {
        $script:cfgDir = Join-Path $TestDrive ([guid]::NewGuid().ToString())
        New-Item -ItemType Directory -Path $cfgDir | Out-Null
        '{"TenantId":"t1","ClientId":"c1","ClientSecret":"s1","IsEncrypted":false}' | Set-Content (Join-Path $cfgDir 'ADProvisioningTool.json')
        Mock DBG { }
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = 'tok'; expires_in = 3599 } }
        $script:GraphTokenCache = $null
    }
    It 'reuses a token that is not near expiry' {
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly
    }
    It 'gets a new token with -Force (after a 401)' {
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        Get-GraphToken -ScriptDir $cfgDir -Force | Out-Null
        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
    }
    It 'drops the cached token when a forced refresh fails' {
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        Mock Invoke-RestMethod { throw 'AADSTS7000215: Invalid client secret' }
        Get-GraphToken -ScriptDir $cfgDir -Force | Should -BeNullOrEmpty
        $script:GraphTokenCache | Should -BeNullOrEmpty
    }
    It 'gets a new token when the cached one is within 5 minutes of expiry' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ access_token = 'tok'; expires_in = 200 } }
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        $script:GraphTokenCache.RefreshAt = (Get-Date).AddSeconds(-1)
        Get-GraphToken -ScriptDir $cfgDir | Out-Null
        Should -Invoke Invoke-RestMethod -Times 2 -Exactly
    }
}

Describe 'Get-CreatedObjectState' {
    It 'reports Exists with the DN when the SAM is in AD' {
        Mock Get-ADObject { [PSCustomObject]@{ DistinguishedName = 'CN=Pat Smith - ITS,OU=New Accounts,DC=west' } }
        $r = Get-CreatedObjectState -Sam 'pat.smith' -Pdc 'pdc1'
        $r.State  | Should -Be 'Exists'
        $r.Detail | Should -BeLike 'CN=Pat Smith*'
    }
    It 'reports Absent when nothing is found' {
        Mock Get-ADObject { }
        (Get-CreatedObjectState -Sam 'pat.smith' -Pdc 'pdc1').State | Should -Be 'Absent'
    }
    It 'reports Unknown, never Absent, when the lookup fails' {
        Mock Get-ADObject { throw 'server down' }
        $r = Get-CreatedObjectState -Sam 'pat.smith' -Pdc 'pdc1'
        $r.State  | Should -Be 'Unknown'
        $r.Detail | Should -BeLike '*server down*'
    }
    It 'escapes the SAM in the LDAP filter' {
        Mock Get-ADObject { }
        Get-CreatedObjectState -Sam 'a*b' -Pdc 'pdc1' | Out-Null
        Should -Invoke Get-ADObject -ParameterFilter { $LDAPFilter -eq '(sAMAccountName=a\2ab)' -and $Server -eq 'pdc1' }
    }
}

Describe 'Get-NameInputProblem' {
    It 'accepts <First> <Last>' -ForEach @(
        @{ First = 'Pat'; Last = "O'Neil" }
        @{ First = 'Mary-Jo'; Last = 'Smith Jr.' }
        @{ First = 'Ana'; Last = 'De La Cruz' }
    ) {
        @(Get-NameInputProblem -First $First -Middle 'A' -Last $Last).Count | Should -Be 0
    }
    It 'rejects a digit' {
        Get-NameInputProblem -First 'Pat2' -Middle '' -Last 'Lee' | Should -BeLike '*digit*'
    }
    It 'rejects accented letters with a clear message' {
        Get-NameInputProblem -First ('Jos' + [char]0x00E9) -Middle '' -Last 'Lee' | Should -BeLike '*accent*'
    }
    It 'rejects a name with no letters' {
        Get-NameInputProblem -First '-' -Middle '' -Last 'Lee' | Should -Contain 'First name has no letters'
    }
    It 'rejects other punctuation' {
        Get-NameInputProblem -First 'Pat' -Middle '' -Last 'Lee/Smith' | Should -BeLike '*characters other than*'
    }
}

Describe 'Get-NameLengthProblem' {
    It 'passes names within the limits' {
        @(Get-NameLengthProblem -CnName 'Pat Smith - ITS' -LocalPart 'pat.smith').Count | Should -Be 0
    }
    It 'flags a CN over 64 characters' {
        Get-NameLengthProblem -CnName ('x' * 65) -LocalPart 'pat' | Should -BeLike '*65 characters (limit 64)*'
    }
    It 'flags a local part over 64 characters' {
        Get-NameLengthProblem -CnName 'ok' -LocalPart ('y' * 70) | Should -BeLike 'Email name*70 characters*'
    }
}

Describe 'Get-EntraTargetProblem' {
    BeforeEach {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*administrativeUnits*' } -MockWith {
            [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'a1'; displayName = 'FIN' }) }
        }
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/groups?*' } -MockWith { [PSCustomObject]@{ value = @([PSCustomObject]@{ id = 'g1' }) } }
    }
    It 'returns nothing when the AUs and the group exist' {
        @(Get-EntraTargetProblem -AUNames @('FIN') -GroupName 'M365 License - Frontline Worker' -Headers @{}).Count | Should -Be 0
    }
    It 'names an AU that is not in Entra' {
        Get-EntraTargetProblem -AUNames @('NOSUCH') -GroupName 'G' -Headers @{} | Should -BeLike "AU 'NOSUCH' not found*"
    }
    It 'names a license group that is not in Entra' {
        Mock Invoke-RestMethod -ParameterFilter { $Uri -like '*/groups?*' } -MockWith { [PSCustomObject]@{ value = @() } }
        Get-EntraTargetProblem -AUNames @('FIN') -GroupName 'Wrong Group' -Headers @{} | Should -BeLike "License group 'Wrong Group' not found*"
    }
    It 'reports a missing token instead of passing' {
        Get-EntraTargetProblem -AUNames @('FIN') -GroupName 'G' -Headers $null | Should -BeLike '*not checked*'
    }
}
