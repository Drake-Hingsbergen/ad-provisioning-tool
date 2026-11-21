<#
.SYNOPSIS
    Runs the real button handlers of ADProvisioningTool.ps1 end to end against a fake AD, Exchange and Graph.
.DESCRIPTION
    The unit tests cover the logic functions. This harness covers the code between them: the Preview,
    Create Account, Preview All and Create All handlers (and the smaller buttons), which no unit test runs.

    Each call builds a new, empty runspace and runs Tests\Smoke\SmokeWorld.ps1 in it. That script loads
    the tool's functions and handlers, swaps the WinForms and ADSI types for stand-ins (SmokeUi.*), and
    defines in-memory fakes for every AD, Exchange and Graph command. The caller's session is never
    changed, module auto-loading is off inside the runspace, and a gate refuses to run if the tool code
    the handlers can reach calls any command the harness does not fake. So a run cannot reach a real
    domain, Exchange server or the network, even on a machine with RSAT or an Exchange session loaded.

    Errors and warnings that reach their streams uncaught are returned in Errors. In the compiled exe
    each of those is a blocking pop-up, so a clean run must return none.
.PARAMETER Scenario
    Hashtable. Mode 'Single' or 'Bulk', Steps, the people to create, and switches that make the fake
    world misbehave. Every key is listed in the defaults below.
.EXAMPLE
    $r = & .\Tests\Smoke\Invoke-HandlerSmoke.ps1 -Scenario @{ Mode = 'Single' }
    $r.Errors; $r.Audit | Format-Table Action, Result, Details
.NOTES
    Runs on Windows PowerShell 5.1 and PowerShell 7 (Windows or Linux). Used by Tests\ADProvisioningTool.Smoke.Tests.ps1.
#>
[CmdletBinding()]
param(
    [hashtable]$Scenario = @{},
    [string]$ToolPath = (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'ADProvisioningTool.ps1')
)

#region Scenario defaults
$opt = @{
    Mode                  = 'Single'
    Steps                 = $null       # Preview, Create, Copy, Clear, Close (single); PreviewAll, CreateAll, CopyAll, BulkClear (bulk);
                                        # or @{ AddAd = ... } / @{ SetOpt = @{ ... } } to change the world between steps
    People                = @(@{ First = 'Morgan'; Mid = ''; Last = 'Reyes'; Division = 'OPS'; EmpType = 'Knowledge Worker'; Ticket = 'INC-1001'; Title = '' })
    ExistingAd            = @()         # @{ Sam; Upn; Domain; Name; Ou } objects already in AD
    ExistingEntra         = @()         # @{ Upn; ImmutableId } users already in Entra
    ExchangeOnlyAddresses = @()         # addresses only on-prem Exchange knows about
    CreateFailures        = @()         # per New-RemoteMailbox call: 'Absent' (fails, nothing written) or 'Exists' (fails after writing)
    PolicyNumbersPrimary  = $false      # New-RemoteMailbox returns first.last2@ as primary (the old address policy behavior)
    SessionFailures       = 0           # number of New-PSSession calls that fail first
    SetMailboxFailPattern = $null       # regex: Set-RemoteMailbox fails when the address or primary matches
    BlankRoutingOnReads   = @()         # which Get-RemoteMailbox reads (1-based) show no routing address yet
    AdsiFails             = $false
    TitleFails            = $false
    GcUnavailable         = $false
    SyncAfterPolls        = 1
    NeverSync             = $false
    SyncRenamesUpn        = $false
    SyncHttpError         = 0           # 401 or 403 on every sync poll
    TokenFails            = $false
    MissingAUs            = @()
    MissingGroups         = @()
    LicensePostFails      = $false
    AlreadyLicensed       = $false
    AuNotFoundTimes       = 0           # AU add answers Request_ResourceNotFound this many times first
    BulkOnPrompt          = 'Retry'     # what the "tech" clicks when a row shows Skip/Retry
    BulkSyncSkipRows      = @()         # rows whose Skip Sync button the "tech" clicks
    BulkSyncSkipAtSleep   = 0           # do not click Skip Sync before this wait tick
    CancelSyncAtSleep     = 0           # click Cancel Sync Wait at this wait tick (single)
    MidRunProbe           = $false      # press Clear, Bulk Clear and the window X while a run is active
    DialogAnswers         = @()         # answers for message boxes in order; 'Yes' once these run out
    SourceTransform       = $null       # @{ Find; Replace } applied to the tool source, for mutation checks
}
foreach ($k in $Scenario.Keys) {
    if (-not $opt.ContainsKey($k)) { throw "Unknown scenario key '$k'." }
    $opt[$k] = $Scenario[$k]
}
if (-not $opt.Steps) { $opt.Steps = if ($opt.Mode -eq 'Bulk') { @('PreviewAll', 'CreateAll') } else { @('Preview', 'Create') } }
#endregion

#region Stand-in types (compiled once per process; C# 5 so Windows PowerShell 5.1 can build them)
if (-not ('SmokeUi.Control' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
namespace SmokeUi {
    public enum Buttons { OK, OKCancel, YesNo, YesNoCancel }
    public enum Icon { None, Question, Warning, Error, Information }
    public static class MessageBox {
        public static Queue<string> Answers = new Queue<string>();
        public static List<string> Shown = new List<string>();
        public static string Show(object text) { return Show(text, null, null, null); }
        public static string Show(object text, object caption) { return Show(text, caption, null, null); }
        public static string Show(object text, object caption, object buttons) { return Show(text, caption, buttons, null); }
        public static string Show(object text, object caption, object buttons, object icon) {
            Shown.Add(Convert.ToString(text));
            return Answers.Count > 0 ? Answers.Dequeue() : "Yes";
        }
    }
    public static class Application { public static int Events; public static void DoEvents() { Events++; } }
    public static class Clipboard { public static string Text; public static void SetText(string s) { Text = s; } }
    public class Control {
        private string text = "";
        public string Text { get { return text; } set { text = value ?? ""; } }
        public bool Enabled { get; set; }
        public bool Visible { get; set; }
        public bool Checked { get; set; }
        public object SelectedItem { get; set; }
        public int SelectedIndex { get; set; }
        public object ForeColor { get; set; }
        public object BackColor { get; set; }
        public object SelectionColor { get; set; }
        public int SelectionStart { get; set; }
        public int SelectionLength { get; set; }
        public object Tag { get; set; }
        public int TextLength { get { return text.Length; } }
        public void Clear() { text = ""; }
        public void Refresh() { }
        public void ScrollToCaret() { }
        public void AppendText(string s) { text += s; }
    }
    public class Adsi {
        public static bool FailSetInfo;
        public static List<string[]> Committed = new List<string[]>();
        private readonly string path;
        private readonly List<string[]> pending = new List<string[]>();
        public Adsi(string path) { this.path = path; }
        public void Put(string name, object value) { pending.Add(new[] { path, name, Convert.ToString(value) }); }
        public void SetInfo() { if (FailSetInfo) throw new UnauthorizedAccessException("Access is denied."); Committed.AddRange(pending); pending.Clear(); }
    }
    public class HttpResponse {
        public int StatusCode;
        public string Body;
        public Stream GetResponseStream() { return new MemoryStream(Encoding.UTF8.GetBytes(Body ?? "")); }
    }
    public class HttpException : Exception {
        public HttpResponse Response;
        public HttpException(int status, string body) : base("HTTP " + status + " " + body) { Response = new HttpResponse { StatusCode = status, Body = body }; }
    }
}
'@
}
if (-not ('System.Drawing.Color' -as [type])) { Add-Type -AssemblyName System.Drawing }
[SmokeUi.MessageBox]::Answers.Clear()
foreach ($a in @($opt.DialogAnswers)) { [SmokeUi.MessageBox]::Answers.Enqueue("$a") }
[SmokeUi.MessageBox]::Shown.Clear()
[SmokeUi.Adsi]::Committed.Clear()
[SmokeUi.Adsi]::FailSetInfo = [bool]$opt.AdsiFails
[SmokeUi.Clipboard]::Text = $null
#endregion

$worldScript = [System.IO.File]::ReadAllText((Join-Path $PSScriptRoot 'SmokeWorld.ps1'))
$workDir = Join-Path ([System.IO.Path]::GetTempPath()) ('ContosoSmoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $workDir | Out-Null
$runspace = [runspacefactory]::CreateRunspace([System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault())
$runspace.Open()

function Invoke-InSmokeRunspace {
    # Runs script text at the runspace's global scope. Uncaught errors and warnings come back from the
    # streams, exactly the ones ps2exe would show; a script-terminating error comes back as Terminating.
    param([string]$Script, [hashtable]$Parameters = @{})
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    [void]$ps.AddScript($Script, $false)
    foreach ($k in $Parameters.Keys) { [void]$ps.AddParameter($k, $Parameters[$k]) }
    $result = @{ Output = @(); Problems = @(); Terminating = $null }
    try { $result.Output = @($ps.Invoke()) }
    catch {
        $inner = $_.Exception; while ($inner.InnerException) { $inner = $inner.InnerException }
        $result.Terminating = $inner.Message
    }
    foreach ($e in $ps.Streams.Error) {
        $line = ''; if ($e.InvocationInfo -and $e.InvocationInfo.Line) { $line = " [line: $($e.InvocationInfo.Line.Trim())]" }
        $result.Problems += "ERROR $($e.Exception.Message)$line"
    }
    foreach ($w in $ps.Streams.Warning) { $result.Problems += "WARNING $($w.Message)" }
    $ps.Dispose()
    return $result
}

try {
    $setup = Invoke-InSmokeRunspace -Script $worldScript -Parameters @{ Opt = $opt; ToolPath = $ToolPath; WorkDir = $workDir }
    if ($setup.Terminating) { throw "Smoke harness setup failed: $($setup.Terminating)" }
    if ($setup.Problems) { throw "Smoke harness setup reported: $($setup.Problems -join ' | ')" }

    $errors = New-Object System.Collections.Generic.List[string]
    $steps = New-Object System.Collections.Generic.List[object]
    foreach ($step in $opt.Steps) {
        $label = if ($step -is [hashtable]) { 'ChangeWorld' } else { "$step" }
        $run = Invoke-InSmokeRunspace -Script 'param($Step) Invoke-SmokeStep -Step $Step' -Parameters @{ Step = $step }
        $stepErrors = @($run.Problems | ForEach-Object { "$label : $_" })
        if ($run.Terminating) { $stepErrors += "$label : TERMINATING $($run.Terminating)" }
        foreach ($e in $stepErrors) { $errors.Add($e) }
        $state = (Invoke-InSmokeRunspace -Script 'Get-SmokeStepState').Output | Select-Object -First 1
        $steps.Add([PSCustomObject]@{ Step = $label; Ran = (@($run.Output) -contains 'ran'); Errors = $stepErrors
                CreateEnabled = $state.CreateEnabled; BulkCreateEnabled = $state.BulkCreateEnabled; CreateStatus = $state.CreateStatus; BulkStatus = $state.BulkStatus })
    }

    $final = Invoke-InSmokeRunspace -Script 'Get-SmokeResult'
    if ($final.Terminating -or -not $final.Output) { throw "Smoke harness could not read the result: $($final.Terminating) $($final.Problems -join ' | ')" }
    $result = $final.Output | Select-Object -First 1
    $result | Add-Member -NotePropertyName Errors -NotePropertyValue $errors.ToArray()
    $result | Add-Member -NotePropertyName Steps -NotePropertyValue $steps.ToArray()
    $result
} finally {
    $runspace.Dispose()
    Remove-Item -Path $workDir -Recurse -Force -ErrorAction SilentlyContinue
}
