<#
.SYNOPSIS
    Windows Updates - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that checks the selected device for available Windows
    Updates and, on request, downloads and installs them - reporting live
    per-update progress the whole way through.

    This is the toolkit's first LOCAL action. Every other module talks to
    Microsoft Graph; this one connects to the machine itself over PowerShell
    Remoting (WinRM) and drives the Windows Update Agent COM API on the far
    end. Nothing is installed on the target: the WUA is built into Windows.

      - The device is passed in from the Toolkit (the device selected in the
        search results) and is pre-filled in the window.
      - Connectivity is tested first (DNS, WinRM, then a real remote command)
        so an unreachable or non-domain-joined machine is reported clearly
        rather than as a COM error thirty seconds later.
      - Scanning is read-only. Every update found is listed with its KB,
        severity, size, reboot requirement and whether it is already staged.
      - Download and Install are separate, explicit steps. Each asks for
        confirmation naming the device and the update count, and each
        defaults to No.
      - Updates are processed ONE AT A TIME using the synchronous
        Download()/Install() methods, so every update maps 1:1 to its own
        result. See the note on callbacks in the worker below.
      - Feature updates (Windows 11 24H2 and the like) are not installable
        through IUpdateInstaller. They are detected and routed through the
        Update Orchestrator instead, with progress polled from the setup
        engine. Auto-reboot is suppressed unless Allow reboot is ticked.
      - Results can be exported to CSV.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so WindowsUpdates.xaml
    is not required at runtime; embedded XAML is used by default
    when launched from Toolkit.ps1. Pass -XamlPath to load an external
    WindowsUpdates.xaml instead while iterating on the UI.

    Every function here is Wud-prefixed on purpose. This module,
    CopyDeviceGroups.ps1, RemoveDeviceGroups.ps1, BulkAddToGroup.ps1,
    AppDependencyCheck.ps1 and AssetStatus.ps1 are dot-sourced into the SAME
    session by Toolkit.ps1, so a shared name would mean the last file loaded
    silently wins and could change the behaviour of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\WindowsUpdates\WindowsUpdates.ps1"
        Show-WindowsUpdatesWindow -Device $Script:SelectedDevice `
                                  -Owner  $Window

    Requires: Windows PowerShell 5.1 (-STA)
    Graph scopes: none - this module never calls Graph.
    Needs instead: PS Remoting (WinRM) to the target, and admin rights on it.
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external WindowsUpdates.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# Row type for the updates grid.
#
# A separate type from the other modules' row classes on purpose: Add-Type
# cannot redefine a type that is already loaded in the session, so reusing a
# name would leave whichever module loaded second without its own properties.
#
# A [pscustomobject] cannot raise change notifications, and this grid is
# repainted constantly while downloads and installs run - so every property
# the UI watches raises PropertyChanged.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyUpdateRow' -as [type])) {
    Add-Type -ReferencedAssemblies System.ComponentModel.TypeConverter -TypeDefinition @'
using System.ComponentModel;

public class EndpointguyUpdateRow : INotifyPropertyChanged
{
    private bool   _selected;
    private string _status      = "";
    private string _result      = "NotStarted";
    private int    _downloadPct;
    private int    _installPct;
    private string _note        = "";

    public bool Selected
    {
        get { return _selected; }
        set { if (_selected != value) { _selected = value; OnPropertyChanged("Selected"); } }
    }

    public string Status
    {
        get { return _status; }
        set { if (_status != value) { _status = value; OnPropertyChanged("Status"); } }
    }

    public string Result
    {
        get { return _result; }
        set { if (_result != value) { _result = value; OnPropertyChanged("Result"); } }
    }

    public int DownloadPct
    {
        get { return _downloadPct; }
        set { if (_downloadPct != value) { _downloadPct = value; OnPropertyChanged("DownloadPct"); OnPropertyChanged("DownloadText"); } }
    }

    public int InstallPct
    {
        get { return _installPct; }
        set { if (_installPct != value) { _installPct = value; OnPropertyChanged("InstallPct"); OnPropertyChanged("InstallText"); } }
    }

    public string Note
    {
        get { return _note; }
        set { if (_note != value) { _note = value; OnPropertyChanged("Note"); } }
    }

    // Percentages read better in the grid as text, and a dash is clearer
    // than a hard 0% for work that has not been attempted yet.
    public string DownloadText { get { return _downloadPct > 0 ? _downloadPct + "%" : "-"; } }
    public string InstallText  { get { return _installPct  > 0 ? _installPct  + "%" : "-"; } }

    public int    Index          { get; set; }
    public string Title          { get; set; }
    public string KB             { get; set; }
    public string Severity       { get; set; }
    public double SizeBytes      { get; set; }
    public string SizeText       { get; set; }
    public bool   IsDownloaded   { get; set; }
    public bool   RebootRequired { get; set; }
    public string RebootText     { get; set; }
    public string Categories     { get; set; }
    public bool   IsFeature      { get; set; }
    public int    BundleCount    { get; set; }
    public string UpdateId       { get; set; }
    public string HResult        { get; set; }
    public double DownloadSec    { get; set; }
    public double InstallSec     { get; set; }

    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged(string name)
    {
        PropertyChangedEventHandler h = PropertyChanged;
        if (h != null) { h(this, new PropertyChangedEventArgs(name)); }
    }
}
'@
}

# ---------------------------------------------------------------------------
# Diagnostics log
#
# The window installs a sink here so the helper functions - which live
# outside its scope - can write to the on-screen console. When no window is
# open the sink is null and the lines go to Write-Verbose instead, so the
# module still behaves when dot-sourced or run headless.
# ---------------------------------------------------------------------------
$Script:WudLogSink    = $null
$Script:WudLogVerbose = $false

function Write-WudLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','REMOTE','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:WudLogVerbose) { return }

    $line = '[{0}] {1,-6} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:WudLogSink) {
        try { & $Script:WudLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-WudLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-WudErrorText $ErrorRecord
    Write-WudLog "$Context failed: $detail" 'ERROR'
    if ($Script:WudLogVerbose -and $ErrorRecord) {
        $ex = Get-WudProperty $ErrorRecord 'Exception'
        if ($ex) { Write-WudLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-WudProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-WudLog "At: $((Get-WudProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-WudProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every remote response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-WudProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-WudProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-WudErrorText {
    <#
        Remoting failures are famously unreadable - the useful sentence is
        usually buried under 'Connecting to remote server failed with the
        following error message'. This pulls out the readable part and, for
        the handful of WinRM errors an admin actually hits, replaces it with
        something that says what to do about it.
    #>
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) { return 'Unknown error.' }

    $msg = ''
    $ex  = Get-WudProperty $ErrorRecord 'Exception'
    if ($ex) { $msg = [string]$ex.Message }
    if ([string]::IsNullOrWhiteSpace($msg)) { $msg = [string]$ErrorRecord }

    # WinRM wraps the real cause in a long preamble; keep the tail.
    if ($msg -match 'error message\s*:\s*(.+)$') { $msg = $Matches[1].Trim() }

    switch -Regex ($msg) {
        'cannot find the computer|No such host is known|could not be resolved' {
            return "$msg`n`nThe name could not be resolved. Check the device is on the network and that you are using a name DNS knows about."
        }
        'Access is denied' {
            return "$msg`n`nYour account is not a local administrator on that machine, so the Windows Update Agent will not let it connect."
        }
        'WinRM cannot complete the operation|cannot connect to the destination|The client cannot connect' {
            return "$msg`n`nPS Remoting (WinRM) is not reachable. The machine may be offline, asleep, on a different network, or WinRM may not be enabled on it."
        }
        'Kerberos|negotiate|authentication' {
            return "$msg`n`nAuthentication to the target failed. This is usually a workgroup or cross-domain machine that needs a TrustedHosts entry or explicit credentials."
        }
    }
    return $msg
}

function Convert-WudSize {
    # Matches the console script's Convert-Size, so sizes read identically.
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N2} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N2} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Convert-WudResultCode {
    # Maps the OperationResultCode enum returned by the Windows Update Agent.
    param($Code)
    switch ([int]$Code) {
        0 { 'NotStarted' }
        1 { 'InProgress' }
        2 { 'Succeeded' }
        3 { 'SucceededWithErrors' }
        4 { 'Failed' }
        5 { 'Aborted' }
        default { "Unknown($Code)" }
    }
}

function Get-WudTargetName {
    <#
        Works out what to actually connect to.

        The Intune device name is the right answer almost always, but it is
        not guaranteed to be resolvable - a renamed or off-domain machine
        will not be. The operator can override it in the window, which is
        what the Target box is for.
    #>
    param([Parameter(Mandatory)]$Device)

    $name = [string](Get-WudProperty $Device 'DeviceName' '')
    if ([string]::IsNullOrWhiteSpace($name)) {
        throw 'The selected device has no device name, so there is nothing to connect to.'
    }
    return $name.Trim()
}

function Test-WudReachable {
    <#
        Three escalating checks, because "it did not work" is useless on its
        own and each layer fails for a different reason with a different fix:

          1. DNS       - does the name resolve at all?
          2. WinRM     - is the remoting endpoint listening?
          3. Real call - can we actually run a command and get a result?

        Returns a result object rather than throwing, so the window can show
        exactly which rung it fell off.
    #>
    param(
        [Parameter(Mandatory)][string]$Target,
        [System.Management.Automation.PSCredential]$Credential
    )

    # The window reports on Ok/Stage/Message, so the contract is kept flat and
    # explicit: Stage names the rung that failed, Message says what to do.
    $out = [pscustomobject]@{
        Ok           = $false
        Stage        = 'DNS'
        Message      = ''
        IpAddress    = ''
        ComputerName = ''
        OsCaption    = ''
        IsElevated   = $false
        PsVersion    = ''
    }

    # 1. DNS
    Write-WudLog "Resolving $Target ..." 'REMOTE'
    try {
        $dns = [System.Net.Dns]::GetHostEntry($Target)
        $ip  = @($dns.AddressList | Where-Object { $_.AddressFamily -eq 'InterNetwork' } |
                 Select-Object -First 1)
        if ($ip.Count -gt 0) { $out.IpAddress = $ip[0].IPAddressToString }
        $out.Stage = 'WinRM'
        Write-WudLog "  -> resolved to $($out.IpAddress)" 'DEBUG'
    }
    catch {
        $out.Message = "'$Target' could not be resolved in DNS. Check the name, or type an IP address or FQDN in the Target box."
        Write-WudLog "  -> DNS lookup failed" 'ERROR'
        return $out
    }

    # 2. WinRM endpoint
    Write-WudLog "Testing WinRM on $Target ..." 'REMOTE'
    $wsmanArgs = @{ ComputerName = $Target; ErrorAction = 'Stop' }
    if ($Credential) { $wsmanArgs['Credential'] = $Credential }
    try {
        Test-WSMan @wsmanArgs | Out-Null
        $out.Stage = 'Remote command'
        Write-WudLog '  -> WinRM responded' 'DEBUG'
    }
    catch {
        $out.Message = "WinRM did not answer on '$Target'.`n`n$(Get-WudErrorText $_)`n`nEnable it on the target with:  Enable-PSRemoting -Force"
        Write-WudLog '  -> WinRM did not answer' 'ERROR'
        return $out
    }

    # 3. A real remote command. WinRM answering does not prove a command can
    #    run under this account, which is the thing that actually matters.
    Write-WudLog 'Running a test command on the target ...' 'REMOTE'
    $icmArgs = @{ ComputerName = $Target; ErrorAction = 'Stop' }
    if ($Credential) { $icmArgs['Credential'] = $Credential }
    try {
        $probe = Invoke-Command @icmArgs -ScriptBlock {
            $id = [Security.Principal.WindowsIdentity]::GetCurrent()
            $pr = New-Object Security.Principal.WindowsPrincipal($id)
            [pscustomobject]@{
                Name      = $env:COMPUTERNAME
                OsCaption = (Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue).Caption
                Elevated  = $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                PsVersion = $PSVersionTable.PSVersion.ToString()
            }
        }
        $out.Ok           = $true
        $out.ComputerName = [string](Get-WudProperty $probe 'Name' '')
        $out.OsCaption    = [string](Get-WudProperty $probe 'OsCaption' '')
        $out.IsElevated   = [bool](Get-WudProperty $probe 'Elevated' $false)
        $out.PsVersion    = [string](Get-WudProperty $probe 'PsVersion' '')
        Write-WudLog "  -> ran as $(if ($out.IsElevated) { 'an administrator' } else { 'a standard user' }) on $($out.ComputerName)" 'DEBUG'
        Write-WudLog "  -> remote PowerShell $($out.PsVersion)" 'DEBUG'
    }
    catch {
        $out.Message = Get-WudErrorText $_
        Write-WudLog '  -> the test command failed' 'ERROR'
        return $out
    }

    return $out
}

# ---------------------------------------------------------------------------
# Remote worker - SCAN
#
# Runs on the target machine. Read-only: it searches and reports, and never
# downloads or installs anything.
#
# Everything the UI needs is flattened into plain values here. The IUpdate
# COM objects cannot cross the remoting boundary - they are not serialisable -
# so the update identity is carried back as its UpdateID string and re-found
# on the far end when a download or install is asked for.
# ---------------------------------------------------------------------------
$Script:WudScanWorker = {
    param($Criteria, $IncludeDrivers)

    $ErrorActionPreference = 'Stop'

    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'EndpointGuy-Intune-Toolkit'
    $searcher = $session.CreateUpdateSearcher()

    if ($IncludeDrivers) {
        # Include the driver service so hardware updates are also returned.
        $searcher.ServiceID       = '7971f918-a847-4430-9279-4a52d1efe18d'
        $searcher.SearchScope     = 1
        $searcher.ServerSelection = 3
    }

    $sw     = [Diagnostics.Stopwatch]::StartNew()
    $result = $searcher.Search($Criteria)
    $sw.Stop()

    $list = @()
    for ($i = 0; $i -lt $result.Updates.Count; $i++) {
        $u   = $result.Updates.Item($i)
        $kb  = ($u.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', '
        $cat = (($u.Categories | ForEach-Object { $_.Name }) -join '; ')

        $list += [pscustomobject]@{
            Index          = $i + 1
            Title          = [string]$u.Title
            KB             = if ($kb) { $kb } else { 'n/a' }
            Severity       = if ($u.MsrcSeverity) { [string]$u.MsrcSeverity } else { 'Unspecified' }
            SizeBytes      = [double]$u.MaxDownloadSize
            IsDownloaded   = [bool]$u.IsDownloaded
            RebootRequired = ([int]$u.InstallationBehavior.RebootBehavior -eq 1)
            Categories     = $cat
            # A feature update is an in-place upgrade driven by the setup
            # engine, not by IUpdateInstaller. Catch it by category and by
            # the version string in the title, because neither alone is
            # reliable across releases.
            IsFeature      = ($cat -match 'Upgrade') -or ([string]$u.Title -match 'version\s+\d\dH\d')
            BundleCount    = $(try { [int]$u.BundledUpdates.Count } catch { 0 })
            UpdateId       = [string]$u.Identity.UpdateID
        }
    }

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        ElapsedSec   = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        RebootPending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                        (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
        Updates      = $list
    }
}

# ---------------------------------------------------------------------------
# Remote worker - DOWNLOAD / INSTALL for ONE update
#
# NOTE ON APPROACH (carried over from the console script, and the reason this
# module works at all):
#
# BeginDownload/BeginInstall are asynchronous and require a real COM callback.
# Passing $null for the callback makes the WUA return a job whose GetProgress()
# dereferences a null callback internally -> 'Object reference not set to an
# instance of an object'. PowerShell cannot supply a native
# IDownloadProgressChangedCallback object.
#
# The robust pattern is to process ONE UPDATE AT A TIME using the SYNCHRONOUS
# Download()/Install() methods, each in its own single-item UpdateColl. That
# gives reliable per-update status without any callback, which is exactly what
# is wanted here: the grid shows one row moving at a time, and every result
# maps 1:1 to the update that produced it.
# ---------------------------------------------------------------------------
$Script:WudActionWorker = {
    param($UpdateId, $Mode, $AllowReboot, $IncludeDrivers)

    $ErrorActionPreference = 'Stop'

    # ---- Feature-update installer (orchestrator + setup engine) ----
    # Returns a hashtable: Percent, Result, Status, Note, RebootRequired
    function Install-WudFeatureUpdate {
        param($Title, [switch]$NoReboot)

        $out = @{ Percent = 0; Result = 'Failed'; Status = 'Install Failed'; Note = ''; RebootRequired = $false }

        # Block the orchestrator from auto-restarting while a user may be logged on.
        $auPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
        $savedNoAutoReboot = $null
        if ($NoReboot) {
            try {
                if (-not (Test-Path $auPath)) { New-Item -Path $auPath -Force | Out-Null }
                $existing = Get-ItemProperty -Path $auPath -Name 'NoAutoRebootWithLoggedOnUsers' -ErrorAction SilentlyContinue
                if ($existing) { $savedNoAutoReboot = $existing.NoAutoRebootWithLoggedOnUsers }
                Set-ItemProperty -Path $auPath -Name 'NoAutoRebootWithLoggedOnUsers' -Value 1 -Type DWord -Force
            }
            catch { $out.Note = 'Could not set the no-auto-reboot policy. Proceeding anyway. ' }
        }

        # Kick off the install via the Update Orchestrator.
        $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
        if (-not (Test-Path $uso)) {
            $out.Note += 'UsoClient.exe not found; cannot drive a feature update from this script.'
            return $out
        }

        try { Start-Process -FilePath $uso -ArgumentList 'StartInstall' -WindowStyle Hidden | Out-Null }
        catch { $out.Note += "Failed to start UsoClient: $($_.Exception.Message)"; return $out }

        # Poll for progress. The setup engine reports through the WUA and SetupHost.
        $deadline = (Get-Date).AddMinutes(120)
        $lastPct  = -1
        $seenWork = $false

        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 15

            $pct = 0
            try {
                $hist    = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
                $pending = @($hist.Search('IsInstalled=0 and IsHidden=0').Updates | Where-Object { $_.Title -eq $Title })
                if ($pending.Count -eq 0) { $seenWork = $true; $pct = 100 }
            }
            catch { }

            # SetupHost / setup engine progress lives in the registry during an upgrade.
            try {
                $sp = Get-ItemProperty -Path 'HKLM:\SYSTEM\Setup\MoSetup\Volatile' -ErrorAction SilentlyContinue
                if ($sp -and $null -ne $sp.SetupProgress) { $pct = [int]$sp.SetupProgress; $seenWork = $true }
            }
            catch { }

            $setupRunning = @(Get-Process -Name 'SetupHost','setupprep','WindowsUpdateBox' -ErrorAction SilentlyContinue)
            if ($setupRunning.Count -gt 0) { $seenWork = $true }

            if ($pct -ne $lastPct) {
                $lastPct = $pct
                # NOTE: this is progress-stream only. Invoke-Command does not
                # surface Write-Progress to the caller until the call returns,
                # so the window shows the feature update as a single long-running
                # step rather than a live percentage. The record is still written
                # so it appears in the remote session transcript / F12 console.
                Write-Progress -Id 3 -Activity 'FeaturePct' -Status "$pct" -PercentComplete ([math]::Min([math]::Max($pct,0),100))
            }

            if ($pct -ge 100 -and $setupRunning.Count -eq 0 -and $seenWork) { break }
        }
        Write-Progress -Id 3 -Activity 'FeaturePct' -Completed

        # A staged feature update is 'done' when it is waiting for a reboot.
        $pendingReboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                         (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')

        if ($pendingReboot) {
            $out.Percent        = 100
            $out.Result         = 'Succeeded'
            $out.Status         = 'Installed - reboot pending'
            $out.RebootRequired = $true
            $out.Note          += 'Staged by the setup engine. Finishes on next restart.'
        }
        elseif ($lastPct -ge 100 -and $seenWork) {
            $out.Percent = 100
            $out.Result  = 'Succeeded'
            $out.Status  = 'Installed'
        }
        else {
            $out.Percent = [math]::Max($lastPct, 0)
            $out.Result  = 'Failed'
            $out.Status  = 'Install Failed'
            if (-not $seenWork) {
                $out.Note += 'The orchestrator never started the upgrade. Common causes: compatibility hold, insufficient free disk space, or WSUS/GPO deferral policy blocking the feature update.'
            }
            else {
                $out.Note += 'Setup engine started but did not reach a completed state within the timeout.'
            }
        }

        # Restore the previous auto-reboot policy value.
        if ($NoReboot) {
            try {
                if ($null -eq $savedNoAutoReboot) { Remove-ItemProperty -Path $auPath -Name 'NoAutoRebootWithLoggedOnUsers' -ErrorAction SilentlyContinue }
                else { Set-ItemProperty -Path $auPath -Name 'NoAutoRebootWithLoggedOnUsers' -Value $savedNoAutoReboot -Type DWord -Force }
            }
            catch { }
        }

        return $out
    }

    # ---- Re-find the update by its UpdateID ----
    # COM objects cannot cross the remoting boundary, so the update is looked
    # up again here from the ID the scan returned.
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'EndpointGuy-Intune-Toolkit'
    $searcher = $session.CreateUpdateSearcher()

    if ($IncludeDrivers) {
        $searcher.ServiceID       = '7971f918-a847-4430-9279-4a52d1efe18d'
        $searcher.SearchScope     = 1
        $searcher.ServerSelection = 3
    }

    $found = $null
    $res   = $searcher.Search("UpdateID='$UpdateId'")
    if ($res.Updates.Count -gt 0) { $found = $res.Updates.Item(0) }

    if ($null -eq $found) {
        return [pscustomobject]@{
            Ok = $false; Result = 'Failed'; Status = 'Not found'; Percent = 0
            HResult = ''; Seconds = 0; RebootRequired = $false
            Note = 'The update was no longer offered by the Windows Update Agent when the action ran. It may have been installed or superseded since the scan - rescan to refresh the list.'
        }
    }

    $out = [pscustomobject]@{
        Ok = $false; Result = 'Failed'; Status = ''; Percent = 0
        HResult = ''; Seconds = 0; RebootRequired = $false; Note = ''
    }

    # ---- DOWNLOAD ----
    if ($Mode -eq 'Download') {
        if ($found.IsDownloaded) {
            $out.Ok = $true; $out.Result = 'Succeeded'; $out.Percent = 100
            $out.Status = 'Downloaded - pending install'
            $out.Note   = 'Already staged on the machine; nothing was transferred.'
            return $out
        }

        $oneColl = New-Object -ComObject Microsoft.Update.UpdateColl
        [void]$oneColl.Add($found)

        $dl          = $session.CreateUpdateDownloader()
        $dl.Updates  = $oneColl
        $dl.Priority = 3   # 3 = High

        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            $dlRes = $dl.Download()   # synchronous - blocks until finished
            $rc    = [int]$dlRes.ResultCode
            $hr    = $dlRes.HResult
        }
        catch {
            $rc = 4
            $hr = $_.Exception.HResult
            $out.Note = $_.Exception.Message
        }
        $sw.Stop()

        $ok = ($rc -eq 2)
        $out.Ok      = $ok
        $out.Result  = switch ($rc) {
            0 {'NotStarted'} 1 {'InProgress'} 2 {'Succeeded'}
            3 {'SucceededWithErrors'} 4 {'Failed'} 5 {'Aborted'} default {"Unknown($rc)"} }
        $out.HResult = if ($null -ne $hr) { '0x{0:X8}' -f [int]$hr } else { '' }
        $out.Percent = if ($ok) { 100 } else { 0 }
        $out.Status  = if ($ok) { 'Downloaded - pending install' } else { 'Download failed' }
        $out.Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        return $out
    }

    # ---- INSTALL ----
    if (-not $found.IsDownloaded) {
        $out.Status = 'Not downloaded'
        $out.Note   = 'The update is not staged on the machine yet, so there was nothing to install. Download it first.'
        return $out
    }

    if (-not $found.EulaAccepted) { $found.AcceptEula() }

    $isFeature = ((($found.Categories | ForEach-Object { $_.Name }) -join '; ') -match 'Upgrade') -or
                 ([string]$found.Title -match 'version\s+\d\dH\d')
    $bundled   = $(try { [int]$found.BundledUpdates.Count } catch { 0 })

    $sw2 = [Diagnostics.Stopwatch]::StartNew()

    # IMPORTANT - FEATURE UPDATES (e.g. 'Windows 11, version 24H2'):
    # These are servicing-stack upgrades handled by the Update Orchestrator and
    # its own setup engine (SetupHost.exe), NOT by IUpdateInstaller. Calling
    # Install() on one typically fails fast with 0x80240022
    # (WU_E_ALL_UPDATES_FAILED) because the agent refuses the request. Bundle
    # objects cannot be installed directly either. So they are routed through
    # the orchestrator instead, and real progress is polled.
    if ($isFeature -or $bundled -gt 0) {
        $fu = Install-WudFeatureUpdate -Title ([string]$found.Title) -NoReboot:(-not $AllowReboot)
        $sw2.Stop()
        $out.Ok             = ($fu.Result -eq 'Succeeded')
        $out.Result         = $fu.Result
        $out.Status         = $fu.Status
        $out.Percent        = $fu.Percent
        $out.Note           = $fu.Note
        $out.RebootRequired = $fu.RebootRequired
        $out.Seconds        = [math]::Round($sw2.Elapsed.TotalSeconds, 1)
        return $out
    }

    # ---- Normal path (quality updates and drivers) ----
    $oneColl = New-Object -ComObject Microsoft.Update.UpdateColl
    [void]$oneColl.Add($found)

    $inst         = $session.CreateUpdateInstaller()
    $inst.Updates = $oneColl

    try {
        $inRes = $inst.Install()   # synchronous - blocks until finished
        $rc2   = [int]$inRes.ResultCode
        $hr2   = $inRes.HResult
        if ($inRes.RebootRequired) { $out.RebootRequired = $true }
    }
    catch {
        $rc2 = 4
        $hr2 = $_.Exception.HResult
        $out.Note = $_.Exception.Message
    }
    $sw2.Stop()

    $ok2 = ($rc2 -eq 2)
    $out.Ok      = $ok2
    $out.Result  = switch ($rc2) {
        0 {'NotStarted'} 1 {'InProgress'} 2 {'Succeeded'}
        3 {'SucceededWithErrors'} 4 {'Failed'} 5 {'Aborted'} default {"Unknown($rc2)"} }
    $out.HResult = if ($null -ne $hr2) { '0x{0:X8}' -f [int]$hr2 } else { '' }
    $out.Percent = if ($ok2) { 100 } else { 0 }
    $out.Status  = if ($ok2) { 'Installed' } else { "Install $($out.Result)" }
    $out.Seconds = [math]::Round($sw2.Elapsed.TotalSeconds, 1)
    return $out
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$WudXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows Updates"
        Height="900" Width="1280"
        WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False"
        Background="{DynamicResource WindowBrush}"
        FontFamily="Segoe UI">
    <Window.Resources>

        <!-- Theme brushes. Kept in sync with MainWindow.xaml so the child -->
        <!-- window matches the shell even when opened standalone.         -->
        <SolidColorBrush x:Key="WindowBrush"           Color="#1F1F1F"/>
        <SolidColorBrush x:Key="CardBrush"             Color="#2B2B2B"/>
        <SolidColorBrush x:Key="CardBorderBrush"       Color="#3D3D3D"/>
        <SolidColorBrush x:Key="TextBrush"             Color="#E8E8E8"/>
        <SolidColorBrush x:Key="TitleTextBrush"        Color="#FFFFFF"/>
        <SolidColorBrush x:Key="SubtleTextBrush"       Color="#A6A6A6"/>
        <SolidColorBrush x:Key="AccentBrush"           Color="#2B7CD3"/>
        <SolidColorBrush x:Key="AccentHoverBrush"      Color="#3B8FE8"/>
        <SolidColorBrush x:Key="AccentTextBrush"       Color="#6BB3E8"/>
        <SolidColorBrush x:Key="GreenBrush"            Color="#15803D"/>
        <SolidColorBrush x:Key="GreenHoverBrush"       Color="#1A9D4B"/>
        <SolidColorBrush x:Key="NeutralBrush"          Color="#3A3A3A"/>
        <SolidColorBrush x:Key="NeutralHoverBrush"     Color="#484848"/>
        <SolidColorBrush x:Key="NeutralBorderBrush"    Color="#4D4D4D"/>
        <SolidColorBrush x:Key="NeutralTextBrush"      Color="#F0F0F0"/>
        <SolidColorBrush x:Key="InputBrush"            Color="#2E2E2E"/>
        <SolidColorBrush x:Key="InputBorderBrush"      Color="#4D4D4D"/>
        <SolidColorBrush x:Key="GridHeaderBrush"       Color="#333333"/>
        <SolidColorBrush x:Key="GridRowBrush"          Color="#2B2B2B"/>
        <SolidColorBrush x:Key="GridAltRowBrush"       Color="#303030"/>
        <SolidColorBrush x:Key="GridLineBrush"         Color="#3A3A3A"/>
        <SolidColorBrush x:Key="GridBorderBrush"       Color="#3D3D3D"/>
        <SolidColorBrush x:Key="SelectedRowBrush"      Color="#0F4C7A"/>
        <SolidColorBrush x:Key="SelectedRowTextBrush"  Color="#FFFFFF"/>
        <SolidColorBrush x:Key="InfoPanelBrush"        Color="#16324A"/>
        <SolidColorBrush x:Key="InfoPanelBorderBrush"  Color="#2D5B85"/>
        <SolidColorBrush x:Key="InfoLabelBrush"        Color="#6BB3E8"/>
        <SolidColorBrush x:Key="WarnTextBrush"         Color="#F0B429"/>
        <SolidColorBrush x:Key="ConnBadBrush"          Color="#FF6B6B"/>
        <SolidColorBrush x:Key="ConsoleBrush"          Color="#141414"/>
        <SolidColorBrush x:Key="ConsoleTextBrush"      Color="#D4D4D4"/>
        <SolidColorBrush x:Key="OnAccentTextBrush"     Color="#FFFFFF"/>

        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource CardBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="6"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
        </Style>

        <Style x:Key="CardHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="17"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="{StaticResource TitleTextBrush}"/>
        </Style>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>

        <Style x:Key="NeutralButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource NeutralBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource NeutralTextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource NeutralBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="4"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource NeutralHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource NeutralButton}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource OnAccentTextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="4"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <SolidColorBrush x:Key="DangerBrush"      Color="#B02A2A"/>
        <SolidColorBrush x:Key="DangerHoverBrush" Color="#C93838"/>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="{StaticResource DangerBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource DangerBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="4"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource DangerHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="GreenButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="{StaticResource GreenBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource GreenBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="4"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="1">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource GreenHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,7"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Background" Value="{StaticResource InputBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="CaretBrush" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource InputBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{StaticResource GridHeaderBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource NeutralTextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource GridBorderBrush}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="DataGridCell">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource SelectedRowBrush}"/>
                    <Setter Property="Foreground" Value="{StaticResource SelectedRowTextBrush}"/>
                </Trigger>
            </Style.Triggers>
        </Style>

    </Window.Resources>
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Padding="22,14">
            <StackPanel>
                <TextBlock Text="Windows Updates" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Connects to the device over PowerShell Remoting and drives the Windows Update Agent on the machine itself. This module does not use Microsoft Graph - it needs WinRM and local administrator rights on the target."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- 1. Target and scan options -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="1. Target machine" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>

                <Border Grid.Row="1" CornerRadius="4"
                        Background="{StaticResource InfoPanelBrush}"
                        BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                        Padding="12,10">
                    <StackPanel>
                        <TextBlock Text="SELECTED IN TOOLKIT" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <TextBlock x:Name="TxtDeviceName" Text="No device selected"
                                   FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
                        <TextBlock x:Name="TxtDeviceDetail" Text=""
                                   FontSize="12" Foreground="{StaticResource SubtleTextBrush}"
                                   TextWrapping="Wrap" LineHeight="17" Margin="0,6,0,0"/>
                    </StackPanel>
                </Border>

                <Grid Grid.Row="2" Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <TextBlock Grid.Column="0" Text="Connect to:" VerticalAlignment="Center"
                               FontSize="13" Margin="0,0,10,0"/>
                    <TextBox x:Name="TxtTarget" Grid.Column="1" Height="32"
                             ToolTip="The name used for the remote connection. Defaults to the Intune device name - change it to an FQDN or IP if that name does not resolve."/>
                    <Button x:Name="BtnTestConn" Grid.Column="2" Style="{StaticResource NeutralButton}"
                            Content="Test connection" MinWidth="150" Height="32" Margin="10,0,0,0"/>
                    <Button x:Name="BtnScan" Grid.Column="3" Style="{StaticResource PrimaryButton}"
                            Content="Scan for updates" MinWidth="170" Height="32" Margin="10,0,0,0"/>
                </Grid>

                <WrapPanel Grid.Row="3" Margin="0,12,0,0">
                    <CheckBox x:Name="ChkDrivers" Content="Include driver updates"
                              VerticalAlignment="Center" FontSize="13" Margin="0,0,20,0"
                              ToolTip="Also search the Windows Update driver service for hardware updates."/>
                    <CheckBox x:Name="ChkExcludePreview" Content="Exclude optional and preview updates" IsChecked="True"
                              VerticalAlignment="Center" FontSize="13" Margin="0,0,20,0"
                              ToolTip="Adds BrowseOnly=0 to the search, which leaves out optional and preview quality updates."/>
                    <CheckBox x:Name="ChkAllowReboot" Content="Allow the machine to restart itself"
                              VerticalAlignment="Center" FontSize="13" Margin="0,0,20,0"
                              Foreground="{StaticResource WarnTextBrush}"
                              ToolTip="Off by default. When off, auto-restart is suppressed during a feature update so a logged-on user is not rebooted from under them."/>
                    <CheckBox x:Name="ChkCredential" Content="Use different credentials"
                              VerticalAlignment="Center" FontSize="13"
                              ToolTip="Prompts for a username and password to connect with, instead of your signed-in account."/>
                </WrapPanel>
            </Grid>
        </Border>

        <!-- 2. Updates -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="2. Available updates" Style="{StaticResource CardHeader}"/>
                        <TextBlock x:Name="UpdateCount" Text="" Margin="12,0,0,0"
                                   VerticalAlignment="Bottom" FontSize="13"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="BtnSelectAll"  Style="{StaticResource NeutralButton}" Content="Select all"  Margin="0,0,10,0"/>
                        <Button x:Name="BtnSelectNone" Style="{StaticResource NeutralButton}" Content="Select none" Margin="0,0,10,0"/>
                        <Button x:Name="BtnExportResults" Style="{StaticResource NeutralButton}" Content="Export results"/>
                    </StackPanel>
                </Grid>

                <DataGrid x:Name="GridUpdates" Grid.Row="1"
                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          SelectionMode="Extended"
                          HeadersVisibility="Column"
                          GridLinesVisibility="Horizontal"
                          HorizontalGridLinesBrush="{StaticResource GridLineBrush}"
                          Background="{StaticResource GridRowBrush}"
                          Foreground="{StaticResource TextBrush}"
                          RowBackground="{StaticResource GridRowBrush}"
                          AlternatingRowBackground="{StaticResource GridAltRowBrush}"
                          BorderBrush="{StaticResource GridBorderBrush}"
                          BorderThickness="1" FontSize="12.5" RowHeight="26"
                          CanUserSortColumns="True">
                    <DataGrid.Columns>
                        <DataGridCheckBoxColumn Header="Do" Width="45"
                                                Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            <DataGridCheckBoxColumn.ElementStyle>
                                <Style TargetType="CheckBox">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                    <Setter Property="VerticalAlignment" Value="Center"/>
                                </Style>
                            </DataGridCheckBoxColumn.ElementStyle>
                            <DataGridCheckBoxColumn.EditingElementStyle>
                                <Style TargetType="CheckBox">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                    <Setter Property="VerticalAlignment" Value="Center"/>
                                </Style>
                            </DataGridCheckBoxColumn.EditingElementStyle>
                        </DataGridCheckBoxColumn>
                        <DataGridTextColumn Header="KB"       Binding="{Binding KB}"           Width="110" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Severity" Binding="{Binding Severity}"     Width="95"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="Size"     Binding="{Binding SizeText}"     Width="90"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="Reboot?"  Binding="{Binding RebootText}"   Width="70"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="DL %"     Binding="{Binding DownloadText}" Width="60"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="Inst %"   Binding="{Binding InstallText}"  Width="60"  IsReadOnly="True"/>
                        <DataGridTextColumn Header="Result"   Binding="{Binding Result}"       Width="140" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Status"   Binding="{Binding Status}"       Width="190" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Title"    Binding="{Binding Title}"        Width="*"   IsReadOnly="True"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>

        <!-- Diagnostics console -->
        <Border x:Name="ConsolePanel" Grid.Row="3" Visibility="Collapsed"
                Style="{StaticResource Card}" Padding="18,14" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Diagnostics console" Style="{StaticResource CardHeader}"/>
                        <TextBlock Text="Live remoting calls, update decisions and errors" Margin="12,0,0,0"
                                   VerticalAlignment="Bottom" FontSize="12.5"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <CheckBox x:Name="ChkVerbose" Content="Verbose"
                                  VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"/>
                        <CheckBox x:Name="ChkAutoScroll" Content="Auto-scroll" IsChecked="True"
                                  VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"/>
                        <Button x:Name="BtnCopyLog"  Style="{StaticResource NeutralButton}" Content="Copy log"  Margin="0,0,10,0"/>
                        <Button x:Name="BtnSaveLog"  Style="{StaticResource NeutralButton}" Content="Save log"  Margin="0,0,10,0"/>
                        <Button x:Name="BtnClearLog" Style="{StaticResource NeutralButton}" Content="Clear"/>
                    </StackPanel>
                </Grid>

                <TextBox x:Name="TxtLog" Grid.Row="1" Height="170"
                         IsReadOnly="True" IsReadOnlyCaretVisible="True"
                         AcceptsReturn="True" TextWrapping="NoWrap"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         FontFamily="Consolas, Courier New" FontSize="12"
                         Padding="8,6"
                         Background="{StaticResource ConsoleBrush}"
                         Foreground="{StaticResource ConsoleTextBrush}"
                         BorderBrush="{StaticResource GridBorderBrush}"
                         BorderThickness="1"
                         VerticalContentAlignment="Top"/>
            </Grid>
        </Border>

        <!-- Footer -->
        <Border Grid.Row="4" Style="{StaticResource Card}" Padding="14,10" Margin="0,14,0,0">
            <Grid>
                <StackPanel Orientation="Vertical" VerticalAlignment="Center">
                    <TextBlock x:Name="StatusText" Text="Ready."
                               FontSize="12.5" VerticalAlignment="Center" TextWrapping="NoWrap"/>
                    <ProgressBar x:Name="BarProgress" Height="4" Margin="0,6,0,0" Width="420"
                                 HorizontalAlignment="Left" Minimum="0" Maximum="100" Value="0"
                                 Visibility="Collapsed"
                                 Background="{StaticResource InputBrush}"
                                 Foreground="{StaticResource AccentBrush}"
                                 BorderThickness="0"/>
                </StackPanel>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnDownload" Style="{StaticResource PrimaryButton}"
                            Content="Download selected" MinWidth="180" Height="38" Margin="0,0,12,0"/>
                    <Button x:Name="BtnInstall" Style="{StaticResource GreenButton}"
                            Content="Install selected" MinWidth="170" Height="38" Margin="0,0,12,0"/>
                    <Button x:Name="BtnClose" Style="{StaticResource NeutralButton}"
                            Content="Close" MinWidth="110" Height="38"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# ---------------------------------------------------------------------------
# Show-WindowsUpdatesWindow  -  module entry point
# ---------------------------------------------------------------------------
function Show-WindowsUpdatesWindow {
    [CmdletBinding()]
    param(
        # The device selected in the Toolkit search results.
        [Parameter(Mandatory)] $Device,

        # Toolkit main window, so this opens centred and modal over it.
        $Owner,

        # Dev override - load an external WindowsUpdates.xaml instead of the
        # embedded copy.
        [string]$XamlPath,

        # Open with the diagnostics console already showing.
        [switch]$ShowConsole
    )

    # ---- Load the XAML ----
    $xamlText = $WudXamlString
    if ($XamlPath) {
        if (-not (Test-Path -LiteralPath $XamlPath)) {
            throw "XAML file not found: $XamlPath"
        }
        $xamlText = Get-Content -LiteralPath $XamlPath -Raw
    }

    try {
        [xml]$xamlDoc = $xamlText
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        $win    = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        throw "Failed to parse the Windows Updates window layout: $($_.Exception.Message)"
    }

    # ---- Resolve controls ----
    $ui      = @{}
    $needed  = @(
        'TxtDeviceName','TxtDeviceDetail','TxtTarget','BtnTestConn','BtnScan',
        'ChkDrivers','ChkExcludePreview','ChkAllowReboot','ChkCredential',
        'GridUpdates','UpdateCount','BtnSelectAll','BtnSelectNone','BtnExportResults',
        'ConsolePanel','ChkVerbose','ChkAutoScroll','BtnCopyLog','BtnSaveLog','BtnClearLog','TxtLog',
        'StatusText','BarProgress','BtnDownload','BtnInstall','BtnClose'
    )
    foreach ($n in $needed) {
        $c = $win.FindName($n)
        if ($null -eq $c) { throw "The Windows Updates window is missing the control '$n'." }
        $ui[$n] = $c
    }

    if ($Owner) { $win.Owner = $Owner }

    # ---- Mutable state ----
    # A hashtable, not plain variables: assigning to a plain variable inside an
    # event handler creates a local copy in that scriptblock's scope and the
    # change is lost when it returns.
    $state = @{
        Rows        = New-Object System.Collections.ObjectModel.ObservableCollection[EndpointguyUpdateRow]
        Session     = $null      # persistent PSSession to the target
        Credential  = $null
        Target      = ''
        Busy        = $false
        Scanned     = $false
        RemoteInfo  = $null
        RebootPending = $false
    }
    $ui.GridUpdates.ItemsSource = $state.Rows

    # ---- Logging sink ----
    $Script:WudLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    $appendLog = {
        param([string]$Line)
        $box = $ui.TxtLog
        $box.AppendText($Line + "`r`n")

        # Keep the console from growing without bound over a long install run.
        if ($box.LineCount -gt 2000) {
            $keep = $box.Text -split "`r`n" | Select-Object -Last 1200
            $box.Text = ($keep -join "`r`n")
        }
        if ($ui.ChkAutoScroll.IsChecked) { $box.ScrollToEnd() }
    }
    $Script:WudLogSink = $appendLog

    $ui.ChkVerbose.Add_Checked({   $Script:WudLogVerbose = $true  })
    $ui.ChkVerbose.Add_Unchecked({ $Script:WudLogVerbose = $false })

    # ---- UI helpers ----
    $setStatus = {
        param([string]$Text)
        $ui.StatusText.Text = $Text
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    $setProgress = {
        param($Value)
        if ($null -eq $Value) {
            $ui.BarProgress.Visibility = 'Collapsed'
        }
        else {
            $ui.BarProgress.Visibility = 'Visible'
            $ui.BarProgress.Value      = [math]::Min([math]::Max([double]$Value, 0), 100)
        }
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    $setBusy = {
        param([bool]$On)
        $state.Busy = $On
        foreach ($b in @('BtnTestConn','BtnScan','BtnDownload','BtnInstall','BtnSelectAll','BtnSelectNone','BtnExportResults')) {
            $ui[$b].IsEnabled = -not $On
        }
        $ui.GridUpdates.IsEnabled = -not $On
        $ui.TxtTarget.IsEnabled   = -not $On
        $win.Cursor = if ($On) { [System.Windows.Input.Cursors]::Wait } else { $null }
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # Force the grid to repaint mid-loop. Because the remote calls are
    # synchronous, without this the rows would only update once everything
    # finished - which would defeat the point of showing progress at all.
    $pump = {
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    $updateCount = {
        $total = $state.Rows.Count
        $sel   = @($state.Rows | Where-Object { $_.Selected }).Count
        if ($total -eq 0) { $ui.UpdateCount.Text = '' }
        else { $ui.UpdateCount.Text = "$total found, $sel selected" }
    }

    # ---- Device details ----
    $devName = ''
    if (Test-WudProperty $Device 'DeviceName') { $devName = [string]$Device.DeviceName }
    $ui.TxtDeviceName.Text = if ($devName) { $devName } else { 'Unnamed device' }

    $bits = @()
    foreach ($p in @(
        @{ Label = 'Serial';       Prop = 'SerialNumber' },
        @{ Label = 'Primary user'; Prop = 'UserPrincipalName' },
        @{ Label = 'OS';           Prop = 'OperatingSystem' },
        @{ Label = 'OS version';   Prop = 'OsVersion' },
        @{ Label = 'Last sync';    Prop = 'LastSyncDateTime' }
    )) {
        $v = Get-WudProperty $Device $p.Prop
        if ($v) { $bits += "$($p.Label): $v" }
    }
    $ui.TxtDeviceDetail.Text = ($bits -join '   |   ')

    # Default the connect target to the Intune device name.
    try   { $state.Target = Get-WudTargetName -Device $Device }
    catch { $state.Target = '' }
    $ui.TxtTarget.Text = $state.Target

    # -----------------------------------------------------------------------
    # Session handling
    #
    # One persistent PSSession is opened and reused for the whole visit. Each
    # update is then driven by its own Invoke-Command against that session.
    # That is what makes live progress possible: a single Invoke-Command that
    # looped over every update internally could not report anything until it
    # returned, so the client drives the loop instead and repaints the grid
    # between calls. Reusing the session keeps the cost of that to one
    # round-trip per update rather than a fresh WinRM handshake each time.
    # -----------------------------------------------------------------------
    $ensureSession = {
        param([switch]$Quiet)

        $target = $ui.TxtTarget.Text.Trim()
        if (-not $target) {
            Write-WudLog -Level ERROR -Message 'No target machine name given.'
            return $false
        }

        # A changed target invalidates any session already open.
        if ($state.Session -and $state.Target -ne $target) {
            Write-WudLog -Level INFO -Message "Target changed to '$target'. Closing the previous session."
            try { Remove-PSSession -Session $state.Session -ErrorAction SilentlyContinue } catch { }
            $state.Session = $null
        }
        $state.Target = $target

        if ($state.Session -and $state.Session.State -eq 'Opened') { return $true }

        if ($state.Session) {
            Write-WudLog -Level WARN -Message "The session to $target is $($state.Session.State). Reconnecting."
            try { Remove-PSSession -Session $state.Session -ErrorAction SilentlyContinue } catch { }
            $state.Session = $null
        }

        # Ask for credentials if that was requested and none are held yet.
        if ($ui.ChkCredential.IsChecked -and $null -eq $state.Credential) {
            $state.Credential = Get-Credential -Message "Credentials for $target"
            if ($null -eq $state.Credential) {
                Write-WudLog -Level WARN -Message 'Credential prompt cancelled.'
                return $false
            }
        }
        if (-not $ui.ChkCredential.IsChecked) { $state.Credential = $null }

        if (-not $Quiet) { & $setStatus "Connecting to $target ..." }
        Write-WudLog -Level REMOTE -Message "Opening a PowerShell session to $target ..."

        try {
            $params = @{ ComputerName = $target; ErrorAction = 'Stop' }
            if ($state.Credential) { $params['Credential'] = $state.Credential }
            $state.Session = New-PSSession @params
            Write-WudLog -Level OK -Message "Session open to $target."
            return $true
        }
        catch {
            $state.Session = $null
            Write-WudLog -Level ERROR -Message (Get-WudErrorText $_)
            return $false
        }
    }

    $closeSession = {
        if ($state.Session) {
            try { Remove-PSSession -Session $state.Session -ErrorAction SilentlyContinue } catch { }
            $state.Session = $null
            Write-WudLog -Level INFO -Message 'Remote session closed.'
        }
    }

    # -----------------------------------------------------------------------
    # Test connection
    # -----------------------------------------------------------------------
    $ui.BtnTestConn.Add_Click({
        if ($state.Busy) { return }
        & $setBusy $true
        try {
            $target = $ui.TxtTarget.Text.Trim()
            if (-not $target) {
                [System.Windows.MessageBox]::Show($win,
                    'Enter the name of the machine to connect to first.',
                    'Nothing to test', 'OK', 'Information') | Out-Null
                return
            }

            & $setStatus "Testing $target ..."
            Write-WudLog -Level INFO -Message "Testing reachability of $target ..."

            if ($ui.ChkCredential.IsChecked -and $null -eq $state.Credential) {
                $state.Credential = Get-Credential -Message "Credentials for $target"
            }
            if (-not $ui.ChkCredential.IsChecked) { $state.Credential = $null }

            $res = Test-WudReachable -Target $target -Credential $state.Credential

            if ($res.Ok) {
                $state.RemoteInfo = $res
                $msg = "Connected to $($res.ComputerName)."
                & $setStatus $msg
                Write-WudLog -Level OK -Message "$msg  OS: $($res.OsCaption); PowerShell $($res.PsVersion); elevated: $($res.IsElevated)"

                $extra = ''
                if (-not $res.IsElevated) {
                    $extra = "`n`nWarning: the remote session is NOT running elevated. Windows Update will refuse to install anything without administrator rights on that machine."
                    Write-WudLog -Level WARN -Message 'The remote session is not elevated. Installs will fail.'
                }

                [System.Windows.MessageBox]::Show($win,
                    "Connection succeeded.`n`nMachine: $($res.ComputerName)`nOS: $($res.OsCaption)`nPowerShell: $($res.PsVersion)$extra",
                    'Connection test', 'OK', 'Information') | Out-Null
            }
            else {
                & $setStatus 'Connection failed.'
                Write-WudLog -Level ERROR -Message "Failed at the '$($res.Stage)' stage: $($res.Message)"
                [System.Windows.MessageBox]::Show($win,
                    "Could not reach $target.`n`nFailed at: $($res.Stage)`n`n$($res.Message)",
                    'Connection test', 'OK', 'Warning') | Out-Null
            }
        }
        catch {
            Write-WudLogException $_ 'Connection test'
            & $setStatus 'Connection test failed.'
        }
        finally { & $setBusy $false }
    })

    # -----------------------------------------------------------------------
    # Scan
    # -----------------------------------------------------------------------
    $doScan = {
        if ($state.Busy) { return }
        & $setBusy $true
        try {
            if (-not (& $ensureSession)) {
                & $setStatus 'Not connected.'
                [System.Windows.MessageBox]::Show($win,
                    "Could not open a session to the machine. See the diagnostics console (F12) for the reason.",
                    'Connection failed', 'OK', 'Warning') | Out-Null
                return
            }

            $state.Rows.Clear()
            & $updateCount

            # Build the search criteria.
            $criteria = "IsInstalled=0 and IsHidden=0"
            if ($ui.ChkExcludePreview.IsChecked) { $criteria += " and BrowseOnly=0" }
            $drivers = [bool]$ui.ChkDrivers.IsChecked

            & $setStatus "Scanning $($state.Target) for updates - this can take a couple of minutes ..."
            Write-WudLog -Level REMOTE -Message "Searching with criteria: $criteria (drivers: $drivers)"
            & $setProgress 0

            $sw  = [Diagnostics.Stopwatch]::StartNew()
            $res = Invoke-Command -Session $state.Session -ScriptBlock $Script:WudScanWorker `
                                  -ArgumentList $criteria, $drivers -ErrorAction Stop
            $sw.Stop()

            $state.RebootPending = [bool]$res.RebootPending
            Write-WudLog -Level OK -Message "Scan finished on $($res.ComputerName) in $($res.ElapsedSec)s - $(@($res.Updates).Count) update(s) applicable."

            if ($state.RebootPending) {
                Write-WudLog -Level WARN -Message 'The machine already has a restart pending. Installs may fail until it is restarted.'
            }

            foreach ($u in @($res.Updates)) {
                $row = New-Object EndpointguyUpdateRow
                $row.Index          = [int]$u.Index
                $row.Title          = [string]$u.Title
                $row.KB             = [string]$u.KB
                $row.Severity       = [string]$u.Severity
                $row.SizeBytes      = [double]$u.SizeBytes
                $row.SizeText       = Convert-WudSize ([double]$u.SizeBytes)
                $row.IsDownloaded   = [bool]$u.IsDownloaded
                $row.RebootRequired = [bool]$u.RebootRequired
                $row.RebootText     = if ($u.RebootRequired) { 'Yes' } else { 'No' }
                $row.Categories     = [string]$u.Categories
                $row.IsFeature      = [bool]$u.IsFeature
                $row.BundleCount    = [int]$u.BundleCount
                $row.UpdateId       = [string]$u.UpdateId
                $row.Result         = ''
                $row.Selected       = $true
                $row.Status         = if ($u.IsDownloaded) { 'Downloaded - pending install' } else { 'Available' }
                if ($u.IsDownloaded) { $row.DownloadPct = 100 }
                if ($u.IsFeature) {
                    $row.Note = 'Feature update - installed through the Update Orchestrator.'
                    Write-WudLog -Level WARN -Message "Feature update detected: $($u.Title). This installs via the setup engine and takes considerably longer."
                }
                $state.Rows.Add($row)
                Write-WudLog -Level DEBUG -Message "  [$($u.Index)] $($u.KB)  $($u.Title)"
            }

            $state.Scanned = $true
            & $updateCount
            & $setProgress $null

            if ($state.Rows.Count -eq 0) {
                & $setStatus "No applicable updates - $($res.ComputerName) is up to date."
                [System.Windows.MessageBox]::Show($win,
                    "No applicable updates were found.`n`n$($res.ComputerName) is up to date for the criteria used.",
                    'Up to date', 'OK', 'Information') | Out-Null
            }
            else {
                $needDl = @($state.Rows | Where-Object { -not $_.IsDownloaded }).Count
                & $setStatus "$($state.Rows.Count) update(s) found on $($res.ComputerName); $needDl still to download."
            }
        }
        catch {
            & $setProgress $null
            Write-WudLog -Level ERROR -Message (Get-WudErrorText $_)
            & $setStatus 'Scan failed.'
            [System.Windows.MessageBox]::Show($win,
                "The update scan failed.`n`n$(Get-WudErrorText $_)",
                'Scan failed', 'OK', 'Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    $ui.BtnScan.Add_Click($doScan)

    # -----------------------------------------------------------------------
    # Shared runner for Download / Install
    #
    # Walks the ticked rows one at a time, calling the remote worker for each
    # and repainting the grid between calls so the operator can watch it move.
    # -----------------------------------------------------------------------
    $runAction = {
        param([string]$Mode)   # 'Download' or 'Install'

        if ($state.Busy) { return }

        $picked = @($state.Rows | Where-Object { $_.Selected })
        if ($picked.Count -eq 0) {
            [System.Windows.MessageBox]::Show($win,
                "Tick at least one update first.",
                "Nothing selected", 'OK', 'Information') | Out-Null
            return
        }

        # For an install, only what is already staged can actually go in.
        if ($Mode -eq 'Install') {
            $notReady = @($picked | Where-Object { -not $_.IsDownloaded -and $_.DownloadPct -lt 100 })
            if ($notReady.Count -eq $picked.Count) {
                [System.Windows.MessageBox]::Show($win,
                    "None of the selected updates have been downloaded yet.`n`nDownload them first, then install.",
                    'Nothing to install', 'OK', 'Information') | Out-Null
                return
            }
            if ($notReady.Count -gt 0) {
                Write-WudLog -Level WARN -Message "$($notReady.Count) selected update(s) are not downloaded and will be skipped."
            }
        }

        # ---- Confirmation. Always defaults to No. ----
        $totalSize = ($picked | Measure-Object -Property SizeBytes -Sum).Sum
        $features  = @($picked | Where-Object { $_.IsFeature })
        $reboots   = @($picked | Where-Object { $_.RebootRequired })

        if ($Mode -eq 'Download') {
            $head = "Download $($picked.Count) update(s) to $($state.Target)?"
            $body = "Total size: $(Convert-WudSize $totalSize)`n`nDownloading stages the updates on the machine. Nothing is installed and the machine is not restarted."
        }
        else {
            $head = "INSTALL $($picked.Count) update(s) on $($state.Target)?"
            $body = "This writes to the machine and cannot be undone from here."
            if ($reboots.Count -gt 0) {
                $body += "`n`n$($reboots.Count) of them require a restart to complete."
            }
            if ($features.Count -gt 0) {
                $body += "`n`n$($features.Count) is a FEATURE UPDATE (an in-place Windows upgrade). These take a long time - often an hour or more - and this window must stay open for the duration."
            }
            if ($ui.ChkAllowReboot.IsChecked) {
                $body += "`n`nRESTART IS ALLOWED: the machine may restart itself without warning anyone signed in to it."
            }
            else {
                $body += "`n`nAutomatic restart is suppressed. Updates that need one will finish the next time the machine is restarted."
            }
            if ($state.RebootPending) {
                $body += "`n`nNote: this machine already has a restart pending, which can cause installs to fail."
            }
        }

        $icon = if ($Mode -eq 'Install') { 'Warning' } else { 'Question' }
        $ans  = [System.Windows.MessageBox]::Show($win,
                    "$head`n`n$body",
                    "Confirm $Mode", 'YesNo', $icon, [System.Windows.MessageBoxResult]::No)
        if ($ans -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-WudLog -Level INFO -Message "$Mode cancelled at the confirmation prompt."
            & $setStatus "$Mode cancelled."
            return
        }

        & $setBusy $true
        try {
            if (-not (& $ensureSession)) {
                & $setStatus 'Not connected.'
                return
            }

            $allowReboot = [bool]$ui.ChkAllowReboot.IsChecked
            $drivers     = [bool]$ui.ChkDrivers.IsChecked
            $done = 0; $ok = 0; $failed = 0; $skipped = 0
            $needReboot = $false

            Write-WudLog -Level INFO -Message "---- $Mode starting for $($picked.Count) update(s) on $($state.Target) ----"

            foreach ($row in $picked) {

                # Skip anything not staged when installing.
                if ($Mode -eq 'Install' -and -not $row.IsDownloaded -and $row.DownloadPct -lt 100) {
                    $row.Status = 'Skipped - not downloaded'
                    $row.Result = 'Skipped'
                    $skipped++; $done++
                    & $pump
                    continue
                }

                $done++
                $label = if ($row.KB -and $row.KB -ne 'n/a') { $row.KB } else { $row.Title }

                $row.Status = if ($Mode -eq 'Download') { 'Downloading ...' } else { 'Installing ...' }
                if ($Mode -eq 'Install' -and $row.IsFeature) {
                    $row.Status = 'Upgrading Windows - this takes a while ...'
                }
                & $setStatus "$Mode $done of $($picked.Count): $label"
                & $setProgress ([int](100 * ($done - 1) / $picked.Count))
                Write-WudLog -Level REMOTE -Message "[$done/$($picked.Count)] $Mode - $($row.Title)"
                & $pump

                try {
                    $r = Invoke-Command -Session $state.Session -ScriptBlock $Script:WudActionWorker `
                            -ArgumentList $row.UpdateId, $Mode, $allowReboot, $drivers -ErrorAction Stop

                    $row.Result = [string]$r.Result
                    $row.Status = [string]$r.Status
                    if ($r.HResult) { $row.HResult = [string]$r.HResult }
                    if ($r.Note)    { $row.Note    = [string]$r.Note }

                    if ($Mode -eq 'Download') {
                        $row.DownloadPct = [int]$r.Percent
                        $row.DownloadSec = [double]$r.Seconds
                        if ($r.Ok) { $row.IsDownloaded = $true }
                    }
                    else {
                        $row.InstallPct = [int]$r.Percent
                        $row.InstallSec = [double]$r.Seconds
                        if ($r.RebootRequired) { $needReboot = $true }
                    }

                    if ($r.Ok) {
                        $ok++
                        $row.Selected = $false     # done - untick so a re-run does not repeat it
                        Write-WudLog -Level OK -Message "      $($r.Result) in $($r.Seconds)s - $($r.Status)"
                    }
                    else {
                        $failed++
                        $detail = if ($r.HResult) { " ($($r.HResult))" } else { '' }
                        Write-WudLog -Level ERROR -Message "      $($r.Result)$detail - $($r.Status)"
                        if ($r.Note) { Write-WudLog -Level ERROR -Message "      $($r.Note)" }
                    }
                }
                catch {
                    $failed++
                    $row.Result = 'Failed'
                    $row.Status = 'Remote call failed'
                    $row.Note   = Get-WudErrorText $_
                    Write-WudLog -Level ERROR -Message "      $($row.Note)"

                    # A dropped session is fatal for the rest of the run.
                    if ($state.Session -and $state.Session.State -ne 'Opened') {
                        Write-WudLog -Level ERROR -Message 'The remote session dropped. Stopping here.'
                        break
                    }
                }

                & $setProgress ([int](100 * $done / $picked.Count))
                & $pump
            }

            & $setProgress $null
            & $updateCount

            $summary = "$Mode finished: $ok succeeded, $failed failed" + $(if ($skipped) { ", $skipped skipped" } else { '' }) + "."
            Write-WudLog -Level INFO -Message "---- $summary ----"
            & $setStatus $summary

            $tail = ''
            if ($Mode -eq 'Install' -and $needReboot) {
                $tail = "`n`nOne or more updates need the machine to be restarted before they take effect."
                if (-not $allowReboot) { $tail += " Automatic restart was suppressed, so this will happen the next time it is restarted." }
            }
            if ($failed -gt 0) { $tail += "`n`nSee the diagnostics console (F12) for the failures." }

            [System.Windows.MessageBox]::Show($win,
                "$summary$tail",
                "$Mode complete", 'OK',
                $(if ($failed -gt 0) { 'Warning' } else { 'Information' })) | Out-Null
        }
        catch {
            & $setProgress $null
            Write-WudLogException $_ $Mode
            & $setStatus "$Mode failed."
        }
        finally { & $setBusy $false }
    }

    $ui.BtnDownload.Add_Click({ & $runAction 'Download' })
    $ui.BtnInstall.Add_Click({  & $runAction 'Install'  })

    # -----------------------------------------------------------------------
    # Selection, export and console handlers
    # -----------------------------------------------------------------------
    $ui.BtnSelectAll.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $true }
        & $updateCount
        Write-WudLog -Level DEBUG -Message "Selected all $($state.Rows.Count) update(s)."
    })

    $ui.BtnSelectNone.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $false }
        & $updateCount
        Write-WudLog -Level DEBUG -Message 'Cleared the selection.'
    })

    $ui.GridUpdates.Add_CurrentCellChanged({ & $updateCount })

    $ui.BtnExportResults.Add_Click({
        if ($state.Rows.Count -eq 0) {
            [System.Windows.MessageBox]::Show($win,
                'There is nothing to export yet. Scan for updates first.',
                'Nothing to export', 'OK', 'Information') | Out-Null
            return
        }
        try {
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter     = 'CSV file (*.csv)|*.csv|All files (*.*)|*.*'
            $dlg.FileName   = "WindowsUpdates_$($state.Target)_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $dlg.Title      = 'Export update results'
            if ($dlg.ShowDialog($win) -ne $true) { return }

            $state.Rows | Select-Object `
                @{n='Computer';         e={ $state.Target }},
                @{n='KB';               e={ $_.KB }},
                @{n='Title';            e={ $_.Title }},
                @{n='Severity';         e={ $_.Severity }},
                @{n='Size';             e={ $_.SizeText }},
                @{n='SizeBytes';        e={ $_.SizeBytes }},
                @{n='Categories';       e={ $_.Categories }},
                @{n='FeatureUpdate';    e={ $_.IsFeature }},
                @{n='RequiresReboot';   e={ $_.RebootText }},
                @{n='DownloadPercent';  e={ $_.DownloadPct }},
                @{n='InstallPercent';   e={ $_.InstallPct }},
                @{n='Result';           e={ $_.Result }},
                @{n='Status';           e={ $_.Status }},
                @{n='HResult';          e={ $_.HResult }},
                @{n='DownloadSeconds';  e={ $_.DownloadSec }},
                @{n='InstallSeconds';   e={ $_.InstallSec }},
                @{n='Note';             e={ $_.Note }},
                @{n='UpdateId';         e={ $_.UpdateId }} |
                Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8

            Write-WudLog -Level OK -Message "Exported $($state.Rows.Count) row(s) to $($dlg.FileName)"
            & $setStatus "Exported to $($dlg.FileName)"
        }
        catch {
            Write-WudLogException $_ 'Export'
            [System.Windows.MessageBox]::Show($win,
                "Could not write the CSV.`n`n$($_.Exception.Message)",
                'Export failed', 'OK', 'Error') | Out-Null
        }
    })

    $ui.BtnCopyLog.Add_Click({
        if ($ui.TxtLog.Text) {
            [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
            & $setStatus 'Diagnostics copied to the clipboard.'
        }
    })

    $ui.BtnSaveLog.Add_Click({
        if (-not $ui.TxtLog.Text) { return }
        try {
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter   = 'Log file (*.log)|*.log|Text file (*.txt)|*.txt|All files (*.*)|*.*'
            $dlg.FileName = "WindowsUpdates_$($state.Target)_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
            if ($dlg.ShowDialog($win) -ne $true) { return }
            $ui.TxtLog.Text | Out-File -LiteralPath $dlg.FileName -Encoding UTF8
            & $setStatus "Diagnostics saved to $($dlg.FileName)"
        }
        catch {
            [System.Windows.MessageBox]::Show($win,
                "Could not save the log.`n`n$($_.Exception.Message)",
                'Save failed', 'OK', 'Error') | Out-Null
        }
    })

    $ui.BtnClearLog.Add_Click({ $ui.TxtLog.Clear() })

    $ui.BtnClose.Add_Click({ $win.Close() })

    # F12 toggles the diagnostics console, matching the other modules.
    $win.Add_PreviewKeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::F12) {
            if ($ui.ConsolePanel.Visibility -eq 'Visible') {
                $ui.ConsolePanel.Visibility = 'Collapsed'
                $win.Height = 900
            }
            else {
                $ui.ConsolePanel.Visibility = 'Visible'
                $win.Height = 1080
            }
            $e.Handled = $true
        }
    })

    # Enter in the target box starts a scan.
    $ui.TxtTarget.Add_KeyDown({
        param($sender, $e)
        if ($e.Key -eq [System.Windows.Input.Key]::Return -and -not $state.Busy) {
            $e.Handled = $true
            & $doScan
        }
    })

    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = 'Visible'
        $win.Height = 1080
    }

    # -----------------------------------------------------------------------
    # Lifecycle
    # -----------------------------------------------------------------------
    $win.Add_Closed({
        & $closeSession
        $Script:WudLogSink = $null
    })

    $win.Add_ContentRendered({
        Write-WudLog -Level INFO -Message '=== Windows Updates module opened ==='
        Write-WudLog -Level INFO -Message "Device: $devName; connect target: $($state.Target)"
        Write-WudLog -Level INFO -Message 'This module talks to the machine over WinRM, not Microsoft Graph.'

        if (-not $state.Target) {
            & $setStatus 'The selected device has no name - type the machine name to connect to.'
            Write-WudLog -Level WARN -Message 'No device name available; a target must be entered by hand.'
            $ui.TxtTarget.Focus()
            return
        }

        & $setStatus "Ready. Select Scan for updates to query $($state.Target)."
        $ui.BtnScan.Focus()
    })

    [void]$win.ShowDialog()
}

# ---------------------------------------------------------------------------
# Standalone run
#
# Dot-sourced from Toolkit.ps1 in normal use. Running this file directly opens
# the window against a device name typed at the prompt, which is handy when
# working on the module itself.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    $name = Read-Host 'Device name to connect to'
    if ($name) {
        Show-WindowsUpdatesWindow -Device ([pscustomobject]@{ DeviceName = $name }) `
                                  -XamlPath $XamlPath -ShowConsole
    }
}
