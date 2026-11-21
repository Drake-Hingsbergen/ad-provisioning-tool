# AD Provisioning Tool

A Windows desktop tool (PowerShell and WinForms, compiled to a single exe) that takes a new hire from a name and a division to a working hybrid account: Active Directory user, Exchange remote mailbox, Microsoft 365 license group, Entra administrative units and custom directory attributes, with an audit trail for every step.

This is a sanitized work sample. I built the original for a multi-domain enterprise environment. Every company-specific value has been replaced: the company is "Contoso", the domains, divisions, groups, servers, people and IDs are invented, and the sample data, logs, screenshots and operations documents were left out. See [What was left out](#what-was-left-out).

## What it does

**Create Account** (one person) and **Bulk Create** (many rows at once) run the same flow:

1. **Preflight.** Builds the sign-in name from the naming rule (`first.last`, then `first.m.last`, then `first.last1`, `first.last2` and so on) and checks each candidate across the whole forest through a Global Catalog, in Entra ID through Microsoft Graph, in Exchange, and against the target domain's PDC. A candidate is only used when every check says it is free. If a check cannot be completed the result is "Unverified" and the tool stops instead of guessing.
2. **Create.** Calls `New-RemoteMailbox` on the on-premises Exchange server, which creates the AD user and the mail routing in one step. The account goes into a staging OU in the division's domain with a random 16-character temporary password.
3. **Wait for sync.** Polls Graph until the account appears in Entra ID. The operator can cancel the wait; the account is then logged as a partial create with the follow-up steps listed.
4. **License.** Adds the user to the license group for the chosen worker type.
5. **Administrative units.** Adds the user to the Entra administrative units mapped to the division.
6. **Attributes.** Writes two custom AD attributes (category and account type) and the job title, then reads them back to confirm.

Other tabs:

- **Search Users** finds accounts across every domain in the forest.
- **Health Check** shows one account's AD and Entra state side by side (attributes, mailbox, license, sign-in methods).
- **Audit Log** shows who created what, when, with which settings and ticket number, with filters.
- **Debug Log** shows the step-by-step trace of each run so a failure can be diagnosed from the log alone.

## How it is built

| Path | Purpose |
| --- | --- |
| `ADProvisioningTool.ps1` | The tool: settings, division map, logic functions, the WinForms UI and the button handlers |
| `ADProvisioningTool-Setup.ps1` | Run once per server. Asks for the tenant ID, client ID and client secret, encrypts the secret with DPAPI and writes `ADProvisioningTool.json` |
| `settings.example.json` | The shape of that settings file, with placeholder values. The real file is git-ignored |
| `Build/Invoke-CIChecks.ps1` | The four checks: parse, ps2exe compatibility, PSScriptAnalyzer against a baseline, Pester |
| `Build/Build-Exe.ps1` | Runs the checks, then compiles the exe with ps2exe |
| `Tests/ADProvisioningTool.Tests.ps1` | Unit tests for the logic functions, with AD, Exchange and Graph mocked |
| `Tests/ADProvisioningTool.Smoke.Tests.ps1` | End-to-end tests that run the real button handlers against a fake directory |
| `Tests/Smoke/` | The smoke harness: an isolated runspace with in-memory fakes for every AD, Exchange and Graph command |
| `.github/workflows/ci.yml` | Runs the checks and a build on every push, on Windows PowerShell 5.1 |

Runs on Windows PowerShell 5.1 with the ActiveDirectory module (RSAT). It talks to Exchange over remote PowerShell and to Microsoft Graph over REST, so it needs no Exchange or Graph modules.

Graph authentication is either an app registration with a client secret (encrypted at rest with DPAPI, machine scope) or the server's Azure Arc managed identity, which needs no stored secret.

## How to deploy

1. Edit the settings block at the top of `ADProvisioningTool.ps1`: the Exchange endpoint, the audit log path, the license group names, the category and account type lists, the domain list and the division map (code, name, AD domain, optional email domain, administrative units).
2. Create an Entra app registration with application permissions that cover the calls the tool makes: read users, read and add group members, read and add administrative unit members, and read authentication methods. Or enable the Arc managed identity and grant it the same permissions.
3. Run the checks and build:

   ```powershell
   Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser
   Install-Module PSScriptAnalyzer, ps2exe -Scope CurrentUser
   .\Build\Invoke-CIChecks.ps1
   .\Build\Build-Exe.ps1
   ```

4. Copy `ADProvisioningTool.exe` and `ADProvisioningTool-Setup.ps1` to a folder on the management server (the sample uses `C:\ADProvisioning`).
5. Run `ADProvisioningTool-Setup.ps1` once on that server to write the encrypted settings file. The secret can only be decrypted on that machine.
6. Give the operators Modify on the `logs` subfolder. Operators run the exe as themselves and need the AD and Exchange rights to create mailboxes in the staging OU.

## Design decisions

- **One exe, no install.** The operators are support staff, not scripters. ps2exe gives them a double-click tool. The cost is a set of rules the code has to follow (ASCII only, no constructor syntax, and every PowerShell warning becomes a blocking pop-up), so a CI check enforces those rules on every push.
- **Never guess on a name collision.** A duplicate sign-in name in a multi-domain forest is expensive to clean up. Preflight treats "could not check" as a stop, not as "free", and checks the forest, Entra and Exchange separately because each can hold a name the others do not know about.
- **Full UPN, shortened SAM.** Long names keep their full sign-in name. Only the 20-character SAM account name is cut, and it gets its own collision check in the target domain.
- **`New-RemoteMailbox` instead of `New-ADUser` plus mailbox enablement.** One call creates the user and the routing address, so there is no window where the account exists without mail attributes.
- **Custom attributes through ADSI.** `Set-ADUser` raised a warning in this flow, which the compiled exe turned into a pop-up that blocked the run. Writing through ADSI avoids it, and the tool reads the values back to confirm.
- **Random temporary passwords.** Generated from a cryptographic random source, 16 characters, with look-alike characters left out because the password is read to the user.
- **Audit and debug logs as first-class features.** Every state-changing step writes an audit row with operator, target, settings and ticket. Writes are guarded by a named mutex so several operators can run the tool at once.
- **Partial creates are reported, not hidden.** If a later step fails or the operator cancels the sync wait, the account is logged as partial with the remaining manual steps listed.
- **A smoke harness that cannot touch a real directory.** The unit tests covered the functions but not the button-handler code between them, and a regression in that code once got past them. The harness loads the real handlers into an empty runspace, replaces the UI types with stand-ins, fakes every directory command, and refuses to run if the handlers can reach any command it has not faked.
- **Analyzer baseline instead of a clean slate.** The script predates the analyzer. Existing findings are recorded in a baseline and only new findings fail the build.

## Known limits

- The tool is one large script (about 4,500 lines). Settings and the division map live in the script, not in a data file. Splitting it into a module, a data file and a UI file is the obvious next step.
- Warnings are suppressed globally because of how ps2exe handles them. That can hide a warning worth seeing; the debug log is the mitigation.
- Some `catch {}` blocks around cosmetic UI calls are empty.
- This public copy has only been run against mocks and fakes. With placeholder settings it cannot connect to anything, and the GUI itself was not launched or visually checked in this form.
- The tests do not prove behavior against a real Exchange server or tenant. Replication timing, throttling and permission errors are simulated.
- Windows only. The tests and the harness also run on PowerShell 7 on Linux, but the tool needs WinForms and RSAT.
- The naming rule, the display name format and the staging OU pattern are fixed in code.

## What was left out

- The real division map, domain names, group names, administrative unit names, attribute names, server names and tenant IDs. The domains and divisions in this copy are a small invented set, as are the category and account type lists.
- The real display name format and staging OU name. The ones in this copy are stand-ins.
- The operations runbook, setup guide and change log, which described a specific environment.
- Audit logs, debug logs, reports, screenshots, the application icon and all built outputs.

## How AI was used

I designed the provisioning flow, the naming and collision rules and the requirements from doing this work by hand, and I ran every version against the real environment, read the logs, found the failures and decided what to change. I used Claude (Anthropic) as a pair programmer for most of the code: it wrote and refactored large parts of the script, the Pester tests, the smoke harness and the CI checks under my direction, and it reviewed changes before I deployed them. It also did the scripted sanitization for this public copy, which I reviewed and approved before publishing.
