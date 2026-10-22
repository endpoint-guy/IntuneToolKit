<#
.SYNOPSIS
    EndpointGuy Intune Toolkit - WPF front end for Microsoft Graph / Intune administration.

.DESCRIPTION
    Phase 1 build:
      - Connect to Microsoft Graph (interactive) with live connection status indicator
      - Device search by Device Name / Serial Number / Primary Username (contains or exact),
        or by Management Name, which filters on the Asset Status lifecycle labels
      - Advanced filter: stack rules across any result column (all must match)
      - Devices imported into Windows Autopilot but NOT yet enrolled in
        Intune are included in the search, so a brand new machine can be
        found (and added to groups) before it ever enrols
      - Local device cache so repeat searches are instant
      - Export results to CSV
      - Copy Device Groups module (Modules\CopyDeviceGroups) - copies assigned
        security group memberships from the selected device to another device
      - Remove Device Groups module (Modules\RemoveDeviceGroups) - removes the
        selected device from its assigned security groups, with confirmation
      - Bulk Add to Group module (Modules\BulkAddToGroup) - adds every device
        named in a CSV to one assigned security group, with confirmation
      - App Dependency Check module (Modules\AppDependencyCheck) - read-only
        report listing every app that depends on a selected Win32 app,
        directly or indirectly, so the blast radius of changing or
        removing it is visible before the change is made
      - Asset Status module (Modules\AssetStatus) - sets the Intune
        Management name of the selected device to one of seven lifecycle
        statuses (Assigned, In-Stock, Retired, Recycled, Stolen, Legalhold,
        Lost), with confirmation
      - Device Prep Check module (Modules\DevicePrepCheck) - checks a serial
        number is ready for Autopilot provisioning: hardware hash imported,
        enabled in Entra ID, and no existing Intune record, with fixes
      - Modules can be switched on or off per environment from the MODULE
        REGISTRY near the top of this file, or from an optional
        ModuleConfig.psd1 dropped next to Toolkit.ps1. A module that is
        switched off is not loaded, its button is hidden, and its Graph
        scopes are not requested at sign-in.
      - Remaining Device Actions are stubbed with "Coming soon"

.NOTES
    SELF-CONTAINED UI - the WPF XAML is embedded in this file, so there is
    no separate .xaml file to deploy or keep in sync.

    Requires: Windows PowerShell 5.1; Graph authentication module installs automatically
        First launch requires access to the PowerShell Gallery

    Run: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1
    Or double-click Run-Toolkit.bat
#>

[CmdletBinding()]
param(
    [string]$TenantId,
    [switch]$ShowConsole
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ---------------------------------------------------------------------------
# Keep the PowerShell window off the taskbar
#
# On Windows 11 the console is usually handed over to Windows Terminal, and a
# Windows Terminal window cannot be hidden from inside the script. So when the
# toolkit was launched directly (Start menu shortcut, Run-Toolkit.bat,
# powershell -File ...) and has landed in Windows Terminal, it restarts itself
# once inside a hidden classic console (conhost.exe) and this copy exits. The
# classic console CAN be hidden, so only the toolkit window shows on the
# taskbar.
#
# Skipped when -ShowConsole is used, or when Toolkit.ps1 is run from a
# PowerShell prompt you already have open.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyConsoleHost' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class EndpointguyConsoleHost
{
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);
}
"@
}

$launchedDirectly = [bool](@([Environment]::GetCommandLineArgs()) -like "*$(Split-Path -Leaf $PSCommandPath)*")
$hostHwnd = [EndpointguyConsoleHost]::GetConsoleWindow()

if ($launchedDirectly -and -not $ShowConsole -and -not $env:ENDPOINTGUY_TOOLKIT_RELAUNCHED -and
    $hostHwnd -ne [IntPtr]::Zero) {

    $consoleClass = New-Object System.Text.StringBuilder 256
    [EndpointguyConsoleHost]::GetClassName($hostHwnd, $consoleClass, $consoleClass.Capacity) | Out-Null
    $conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'

    # 'ConsoleWindowClass' is the classic console, which can be hidden later.
    # Anything else (Windows Terminal reports 'PseudoConsoleWindow') cannot.
    if ($consoleClass.ToString() -ne 'ConsoleWindowClass' -and (Test-Path $conhost)) {
        $relaunchArgs = @(
            "`"$PSHOME\powershell.exe`"", '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
            '-WindowStyle', 'Hidden', '-File', "`"$PSCommandPath`""
        )
        if ($TenantId) { $relaunchArgs += @('-TenantId', "`"$TenantId`"") }

        # Guard so the restarted copy can never restart itself again.
        $env:ENDPOINTGUY_TOOLKIT_RELAUNCHED = '1'
        Start-Process -FilePath $conhost -ArgumentList $relaunchArgs -WindowStyle Hidden
        exit 0
    }
}

# ---------------------------------------------------------------------------
# Shared look and feel (used by the toolkit and by every module)
#
#   Show-ToolkitMessageBox   Toolkit-styled replacement for the standard
#                            Windows message box. Same buttons, icons and
#                            return values as [System.Windows.MessageBox].
#   Show-ToolkitProgress     Loading overlay with the moving bar, drawn over
#   Update-ToolkitProgress   a module window while it works. It runs on its
#   Close-ToolkitProgress    own thread, so the bar keeps moving while the
#                            module waits on Microsoft Graph.
# ---------------------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

if (-not ('EndpointguyDialogWin32' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class EndpointguyDialogWin32 {
    [DllImport("user32.dll")] public static extern IntPtr GetActiveWindow();
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
}
"@
}

# The module window the loading overlay was last shown over. Used as a
# fallback owner for message boxes when Windows reports no usable active
# window (for example straight after the overlay closes).
$Script:TkLastOwner = $null

$Script:TkDialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" SizeToContent="WidthAndHeight" ShowInTaskbar="False"
        FontFamily="Segoe UI" UseLayoutRounding="True">
    <Window.Resources>
        <Style x:Key="TkButton" TargetType="Button">
            <Setter Property="MinWidth" Value="96"/>
            <Setter Property="Height" Value="34"/>
            <Setter Property="Margin" Value="8,0,0,0"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Foreground" Value="#F0F0F0"/>
            <Setter Property="Background" Value="#3A3A3A"/>
            <Setter Property="BorderBrush" Value="#4D4D4D"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.88"/>
                            </Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#6BB3E8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border Margin="16" Background="#2B2B2B" BorderBrush="#3D3D3D" BorderThickness="1" CornerRadius="10">
        <Border.Effect>
            <DropShadowEffect BlurRadius="20" ShadowDepth="0" Opacity="0.55" Color="Black"/>
        </Border.Effect>
        <Grid MinWidth="440" MaxWidth="640">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Border x:Name="AccentStrip" Grid.Row="0" Height="4" CornerRadius="10,10,0,0" Background="#2B7CD3"/>
            <Grid x:Name="HeaderBar" Grid.Row="1" Margin="24,18,24,0" Background="Transparent">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid x:Name="IconHost" Grid.Column="0" Width="36" Height="36" Margin="0,0,14,0" VerticalAlignment="Center">
                    <Ellipse x:Name="IconCircle" Fill="#2B7CD3"/>
                    <TextBlock x:Name="IconGlyph" Text="i" FontSize="19" FontWeight="Bold" Foreground="#FFFFFF"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Grid>
                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                    <TextBlock x:Name="TitleText" Text="" FontSize="17" FontWeight="SemiBold" Foreground="#FFFFFF" TextWrapping="Wrap"/>
                    <TextBlock Text="EndpointGuy Intune Toolkit" FontSize="11.5" Foreground="#A6A6A6" Margin="0,2,0,0"/>
                </StackPanel>
            </Grid>
            <TextBox x:Name="MessageText" Grid.Row="2" Margin="24,16,24,20" IsReadOnly="True"
                     Background="Transparent" BorderThickness="0" Foreground="#E8E8E8" FontSize="13.5"
                     TextWrapping="Wrap" MaxHeight="420" VerticalScrollBarVisibility="Auto"
                     SelectionBrush="#2B7CD3" Cursor="Arrow" Padding="0"/>
            <Border Grid.Row="3" Background="#252525" BorderBrush="#3D3D3D" BorderThickness="0,1,0,0"
                    CornerRadius="0,0,10,10" Padding="24,14">
                <StackPanel x:Name="ButtonPanel" Orientation="Horizontal" HorizontalAlignment="Right"/>
            </Border>
        </Grid>
    </Border>
</Window>
"@

function Get-ToolkitDialogOwner {
    <#
        Picks the window a message box should sit over. The handle has to be
        a live, visible window on THIS thread; anything else (for example the
        loading overlay that has just closed, which lives on its own thread)
        makes the styled box fail to open. Returns IntPtr.Zero if nothing
        suitable is found, in which case the box is centred on the screen.
    #>
    $candidates = New-Object System.Collections.Generic.List[IntPtr]
    try { $candidates.Add([EndpointguyDialogWin32]::GetActiveWindow()) } catch { }
    if ($null -ne $Script:TkLastOwner) {
        try { $candidates.Add((New-Object System.Windows.Interop.WindowInteropHelper $Script:TkLastOwner).Handle) } catch { }
    }

    $thisThread = [EndpointguyDialogWin32]::GetCurrentThreadId()
    foreach ($h in $candidates) {
        if ($h -eq [IntPtr]::Zero) { continue }
        if (-not [EndpointguyDialogWin32]::IsWindow($h)) { continue }
        if (-not [EndpointguyDialogWin32]::IsWindowVisible($h)) { continue }
        [uint32]$procId = 0
        $windowThread = [EndpointguyDialogWin32]::GetWindowThreadProcessId($h, [ref]$procId)
        if ($windowThread -ne $thisThread) { continue }
        return $h
    }
    return [IntPtr]::Zero
}

function Show-ToolkitMessageBox {
    <#
        Drop-in replacement for [System.Windows.MessageBox]::Show(...).
        -Button  OK | OKCancel | YesNo | YesNoCancel
        -Icon    None | Information | Warning | Error | Question
                 (Asterisk, Exclamation, Hand and Stop also work)
        -Default the button that Enter presses (Yes, No, OK or Cancel)
        Returns a [System.Windows.MessageBoxResult], just like the original.
    #>
    param(
        [AllowEmptyString()][string]$Message = '',
        [AllowEmptyString()][string]$Title = '',
        [string]$Button = 'OK',
        [string]$Icon = 'None',
        [string]$Default = 'None'
    )

    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = 'EndpointGuy Intune Toolkit' }
    # @( ) keeps this an array even when there is only one button. Without it,
    # 'OK' on its own comes back as a plain string, and under StrictMode its
    # .Count fails - which made every OK-only box fall back to the Windows box.
    $buttonNames = @(switch ($Button) {
        'OKCancel'    { 'OK', 'Cancel' }
        'YesNo'       { 'Yes', 'No' }
        'YesNoCancel' { 'Yes', 'No', 'Cancel' }
        default       { 'OK' }
    })
    # What closing the window (Esc or Alt+F4) returns, same as the standard box.
    $closeResult = if ($buttonNames -contains 'Cancel') { 'Cancel' }
                   elseif ($buttonNames -contains 'No') { 'No' }
                   else { 'OK' }

    # Builds and shows the styled box over $OwnerHandle (or centred on the
    # screen when it is IntPtr.Zero). Returns the button that was pressed.
    $showStyled = {
        param([IntPtr]$OwnerHandle)

        [xml]$doc = $Script:TkDialogXaml
        $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $doc))
        $dlg.Title = $Title
        $dlg.Tag = $closeResult
        $dlg.FindName('TitleText').Text = $Title
        $dlg.FindName('MessageText').Text = $Message

        $iconKind = switch ($Icon) {
            { $_ -in 'Information', 'Asterisk' }  { 'Info' }
            { $_ -in 'Warning', 'Exclamation' }   { 'Warning' }
            { $_ -in 'Error', 'Hand', 'Stop' }    { 'Error' }
            'Question'                            { 'Question' }
            default                               { 'None' }
        }
        $iconColor = switch ($iconKind) {
            'Warning' { '#D29922' }
            'Error'   { '#D13438' }
            default   { '#2B7CD3' }
        }
        $iconText = switch ($iconKind) {
            'Warning'  { '!' }
            'Error'    { 'X' }
            'Question' { '?' }
            default    { 'i' }
        }
        $brush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($iconColor))
        $dlg.FindName('AccentStrip').Background = $brush
        $dlg.FindName('IconCircle').Fill = $brush
        $dlg.FindName('IconGlyph').Text = $iconText
        if ($iconKind -eq 'None') { $dlg.FindName('IconHost').Visibility = [System.Windows.Visibility]::Collapsed }

        # Buttons. The first button is the main action: blue, or red when it
        # confirms something risky (a Warning or Error question).
        $panel = $dlg.FindName('ButtonPanel')
        $style = $dlg.Resources['TkButton']
        $defaultName = if ($buttonNames -contains $Default) { $Default } else { $buttonNames[0] }
        $isRisky = ($buttonNames.Count -gt 1) -and ($iconKind -in 'Warning', 'Error')
        $first = $true
        $focusButton = $null
        foreach ($name in $buttonNames) {
            $b = New-Object System.Windows.Controls.Button
            $b.Style = $style
            $b.Content = $name
            $b.Tag = $name
            if ($first) {
                $main = if ($isRisky) { '#C42B1C' } else { '#2B7CD3' }
                $b.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($main))
                $b.BorderBrush = $b.Background
                $b.Foreground = [System.Windows.Media.Brushes]::White
                $first = $false
            }
            if ($name -eq $defaultName) { $b.IsDefault = $true; $focusButton = $b }
            if ($name -eq $closeResult) { $b.IsCancel = $true }
            $b.Add_Click({
                param($clickSource, $clickArgs)
                $owning = [System.Windows.Window]::GetWindow($clickSource)
                $owning.Tag = [string]$clickSource.Tag
                $owning.Close()
            })
            [void]$panel.Children.Add($b)
        }
        if ($null -ne $focusButton) {
            $dlg.Add_ContentRendered({
                param($renderSource, $renderArgs)
                foreach ($c in $renderSource.FindName('ButtonPanel').Children) {
                    if ($c.IsDefault) { [void]$c.Focus() }
                }
            })
        }

        # Drag the box by its header.
        $dlg.FindName('HeaderBar').Add_MouseLeftButtonDown({
            param($dragSource, $dragArgs)
            try { [System.Windows.Window]::GetWindow($dragSource).DragMove() } catch { }
        })

        # Centre over the toolkit window being used; if there isn't a usable
        # one, centre on the screen and stay on top so it can't get lost.
        if ($OwnerHandle -ne [IntPtr]::Zero) {
            (New-Object System.Windows.Interop.WindowInteropHelper $dlg).Owner = $OwnerHandle
            $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
        }
        else {
            $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
            $dlg.Topmost = $true
        }

        [void]$dlg.ShowDialog()
        return [System.Windows.MessageBoxResult]([string]$dlg.Tag)
    }

    # Dim the module's loading card while the message is on screen.
    $progress = $Script:TkProgress
    if ($null -ne $progress) { $progress.Sync.Hidden = $true }

    try {
        $ownerHandle = [IntPtr]::Zero
        try { $ownerHandle = Get-ToolkitDialogOwner } catch { $ownerHandle = [IntPtr]::Zero }

        # 1. Styled box over the toolkit window.
        try { return (& $showStyled $ownerHandle) }
        catch {
            if (Get-Command -Name Write-ToolkitErrorLog -ErrorAction SilentlyContinue) {
                Write-ToolkitErrorLog -Context "Styled message box '$Title' (with owner) - retrying centred on screen" -ErrorRecord $_ | Out-Null
            }
        }

        # 2. Styled box again, centred on the screen with no owner.
        try { return (& $showStyled ([IntPtr]::Zero)) }
        catch {
            if (Get-Command -Name Write-ToolkitErrorLog -ErrorAction SilentlyContinue) {
                Write-ToolkitErrorLog -Context "Styled message box '$Title' (no owner) - using the standard Windows box" -ErrorRecord $_ | Out-Null
            }
        }

        # 3. Last resort: the standard Windows box, so the message is never lost.
        $mbButton = [System.Windows.MessageBoxButton]$Button
        $mbIcon = [System.Windows.MessageBoxImage]$Icon
        $mbDefault = [System.Windows.MessageBoxResult]$Default
        return [System.Windows.MessageBox]::Show($Message, $Title, $mbButton, $mbIcon, $mbDefault)
    }
    finally {
        if ($null -ne $progress) { $progress.Sync.Hidden = $false }
    }
}

# --- Loading overlay for module windows -------------------------------------
$Script:TkProgress = $null

$Script:TkProgressXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" ShowInTaskbar="False" ShowActivated="False"
        WindowStartupLocation="Manual" FontFamily="Segoe UI" Opacity="0">
    <Grid Background="#B8000000">
        <Border x:Name="LoadingCard" Width="500" Padding="30,26"
                HorizontalAlignment="Center" VerticalAlignment="Center"
                Background="#2B2B2B" BorderBrush="#2B7CD3" BorderThickness="1" CornerRadius="10">
            <Border.Effect>
                <DropShadowEffect BlurRadius="24" ShadowDepth="0" Opacity="0.5" Color="Black"/>
            </Border.Effect>
            <StackPanel>
                <TextBlock x:Name="HeadingText" Text="" FontSize="13" FontWeight="SemiBold"
                           Foreground="#6BB3E8" HorizontalAlignment="Center"/>
                <TextBlock x:Name="LoadingText" Text="Working..." FontSize="21" FontWeight="SemiBold"
                           Foreground="#FFFFFF" HorizontalAlignment="Center" TextAlignment="Center"
                           TextWrapping="Wrap" Margin="0,10,0,18"/>
                <Border x:Name="LoadingTrack" Height="14" CornerRadius="7" Background="#2E2E2E"
                        BorderBrush="#4D4D4D" BorderThickness="1" ClipToBounds="True">
                    <Grid>
                        <Border x:Name="LoadingIndicator" Width="110" Height="12" HorizontalAlignment="Left"
                                CornerRadius="6" Background="#2B7CD3">
                            <Border.RenderTransform>
                                <TranslateTransform X="-110"/>
                            </Border.RenderTransform>
                        </Border>
                    </Grid>
                </Border>
                <TextBlock x:Name="LoadingDetail" Text="" FontSize="11.5" Foreground="#A6A6A6"
                           TextWrapping="Wrap" TextAlignment="Center" HorizontalAlignment="Center" Margin="0,14,0,0"/>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

# Runs on the overlay's own thread.
$Script:TkProgressScript = {
    param($Sync, $Xaml, $Bounds, $OwnerHandle)
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    [xml]$doc = $Xaml
    $w = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $doc))
    $w.Left = [double]$Bounds.Left
    $w.Top = [double]$Bounds.Top
    $w.Width = [double]$Bounds.Width
    $w.Height = [double]$Bounds.Height
    if ($OwnerHandle -ne [IntPtr]::Zero) {
        (New-Object System.Windows.Interop.WindowInteropHelper $w).Owner = $OwnerHandle
    }

    $card = $w.FindName('LoadingCard')
    $heading = $w.FindName('HeadingText')
    $text = $w.FindName('LoadingText')
    $detail = $w.FindName('LoadingDetail')
    $track = $w.FindName('LoadingTrack')
    $slide = $w.FindName('LoadingIndicator').RenderTransform

    $apply = {
        $heading.Text = [string]$Sync.Heading
        $text.Text = [string]$Sync.Text
        $detail.Text = [string]$Sync.Detail
        $card.Visibility = if ($Sync.Hidden) { [System.Windows.Visibility]::Hidden } else { [System.Windows.Visibility]::Visible }
    }
    & $apply

    $w.Add_Loaded({
        # Same bar as the main window: left to right, over and over.
        $trackWidth = [double]$track.ActualWidth
        if ($trackWidth -lt 50) { $trackWidth = 438 }
        $move = New-Object System.Windows.Media.Animation.DoubleAnimation
        $move.From = -110
        $move.To = $trackWidth
        $move.Duration = [System.Windows.Duration]::new([TimeSpan]::FromSeconds(1.2))
        $move.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
        $slide.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $move)

        # Fade in after a short delay, so very quick jobs don't flash.
        $fade = New-Object System.Windows.Media.Animation.DoubleAnimation
        $fade.From = 0
        $fade.To = 1
        $fade.BeginTime = [TimeSpan]::FromMilliseconds(250)
        $fade.Duration = [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(150))
        $w.BeginAnimation([System.Windows.Window]::OpacityProperty, $fade)
    })

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $timer.Add_Tick({
        if ($Sync.Close) {
            $timer.Stop()
            $w.Close()
            return
        }
        & $apply
    })
    $w.Add_Closed({
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvokeShutdown([System.Windows.Threading.DispatcherPriority]::Background)
    })

    if ($Sync.Close) { return }
    $w.Show()
    $timer.Start()
    [System.Windows.Threading.Dispatcher]::Run()
}

function Show-ToolkitProgress {
    <#
        Shows the loading overlay over -Owner (a module window). Calling it
        again while it is showing just changes the text.
    #>
    param(
        $Owner = $null,
        [string]$Heading = '',
        [string]$Text = 'Working...',
        [string]$Detail = 'Please keep this window open while the operation completes.'
    )

    if ($null -ne $Script:TkProgress) {
        $Script:TkProgress.Sync.Heading = $Heading
        $Script:TkProgress.Sync.Text = $Text
        $Script:TkProgress.Sync.Detail = $Detail
        return
    }

    if ($null -ne $Owner) { $Script:TkLastOwner = $Owner }
    try {
        $ownerHandle = [IntPtr]::Zero
        $bounds = $null
        if ($null -ne $Owner) {
            $ownerHandle = (New-Object System.Windows.Interop.WindowInteropHelper $Owner).Handle
            $source = [System.Windows.PresentationSource]::FromVisual($Owner)
            $root = $null
            if ([System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Owner) -gt 0) {
                $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Owner, 0)
            }
            if ($null -ne $source -and $root -is [System.Windows.FrameworkElement] -and $root.ActualWidth -gt 0) {
                $topLeft = $source.CompositionTarget.TransformFromDevice.Transform(
                    $root.PointToScreen((New-Object System.Windows.Point 0, 0)))
                $bounds = @{ Left = $topLeft.X; Top = $topLeft.Y; Width = $root.ActualWidth; Height = $root.ActualHeight }
            }
        }
        if ($null -eq $bounds) { return }

        $sync = [hashtable]::Synchronized(@{
            Heading = $Heading; Text = $Text; Detail = $Detail; Hidden = $false; Close = $false
        })
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = [System.Threading.ApartmentState]::STA
        $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.Open()
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript($Script:TkProgressScript.ToString()).AddArgument($sync).AddArgument($Script:TkProgressXaml).AddArgument($bounds).AddArgument($ownerHandle)
        $handle = $ps.BeginInvoke()
        $Script:TkProgress = @{ PS = $ps; Runspace = $runspace; Handle = $handle; Sync = $sync; Owner = $Owner }
    }
    catch {
        # The overlay is only cosmetic; never let it stop the real work.
        $Script:TkProgress = $null
    }
}

function Update-ToolkitProgress {
    param([string]$Text, [string]$Detail)
    $progress = $Script:TkProgress
    if ($null -eq $progress) { return }
    if ($PSBoundParameters.ContainsKey('Text')) { $progress.Sync.Text = $Text }
    if ($PSBoundParameters.ContainsKey('Detail')) { $progress.Sync.Detail = $Detail }
}

function Close-ToolkitProgress {
    $progress = $Script:TkProgress
    if ($null -eq $progress) { return }
    $Script:TkProgress = $null
    $progress.Sync.Close = $true
    try {
        [void]$progress.Handle.AsyncWaitHandle.WaitOne(3000)
        if ($progress.Handle.IsCompleted) {
            try { [void]$progress.PS.EndInvoke($progress.Handle) } catch { }
            $progress.PS.Dispose()
            $progress.Runspace.Dispose()
        }
    }
    catch { }

    # Hand focus back to the module window, so the next message box opens
    # over it rather than looking for the overlay that has just closed.
    $owner = $progress.Owner
    if ($null -ne $owner) {
        try { if ($owner.IsVisible) { [void]$owner.Activate() } } catch { }
    }
}

# ---------------------------------------------------------------------------
# Graph authentication module
#
# The toolkit installs Microsoft.Graph.Authentication for the current user on
# first launch. No administrator elevation is required. Internet access to the
# PowerShell Gallery is required only when the module is not already installed.
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    try {
        Show-ToolkitMessageBox -Title 'First-time setup' -Button 'OK' -Icon 'Information' -Message (
            "Microsoft.Graph.Authentication is not installed. The toolkit will install it for the current user now.`n`nThis first-time setup can take a few minutes.") | Out-Null

        # Windows PowerShell 5.1 can otherwise negotiate an obsolete protocol
        # when contacting the PowerShell Gallery.
        [Net.ServicePointManager]::SecurityProtocol =
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

        # Install the NuGet provider non-interactively when an older PowerShellGet
        # environment does not already have it.
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }

        Install-Module Microsoft.Graph.Authentication `
            -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    catch {
        Show-ToolkitMessageBox -Title 'First-time setup failed' -Button 'OK' -Icon 'Error' -Message (
            "Microsoft.Graph.Authentication could not be installed automatically.`n`n$($_.Exception.Message)`n`nConfirm that the PowerShell Gallery is reachable, then restart the toolkit.") | Out-Null
        return
    }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ---------------------------------------------------------------------------
# MODULE REGISTRY  <<< TURN MODULES ON / OFF HERE >>>
#
# This is the ONLY place you need to edit to tailor the build to an
# environment. Set Enabled = $false (or delete the whole entry) and the
# module is not dot-sourced, its button disappears from the UI, and its
# Graph scopes are not requested at sign-in.
#
# Key      - internal name, also used by Test-ModuleEnabled.
# Enabled  - $true / $false. The master switch.
# Path     - path to the module file, relative to the toolkit root.
# Button   - x:Name of the button in the XAML that fires this module.
# Scopes   - Graph permissions this module needs. Only merged in when the
#            module is enabled, so a trimmed build asks for less consent.
#
# A missing file is not fatal - the matching button just reports it.
# The legacy flat layout (module sitting next to Toolkit.ps1) is still
# probed as a fallback so a partially migrated folder keeps working.
# ---------------------------------------------------------------------------
$Script:ModuleRegistry = @(
    @{
        Key     = 'CopyDeviceGroups'
        Enabled = $true
        Path    = 'Modules\CopyDeviceGroups\CopyDeviceGroups.ps1'
        Button  = 'BtnCopyGroups'
        Scopes  = @('Device.Read.All','Group.Read.All',
                    'Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Key     = 'RemoveDeviceGroups'
        Enabled = $true
        Path    = 'Modules\RemoveDeviceGroups\RemoveDeviceGroups.ps1'
        Button  = 'BtnRemoveGroups'
        Scopes  = @('Device.Read.All','Group.Read.All',
                    'Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Key     = 'BulkAddToGroup'
        Enabled = $true
        Path    = 'Modules\BulkAddToGroup\BulkAddToGroup.ps1'
        Button  = 'BtnBulkAddGroup'
        Scopes  = @('DeviceManagementManagedDevices.Read.All','Device.Read.All',
                    'Group.Read.All','Group.ReadWrite.All','GroupMember.ReadWrite.All')
    }
    @{
        Key     = 'AppDependencyCheck'
        Enabled = $true
        Path    = 'Modules\AppDependencyCheck\AppDependencyCheck.ps1'
        Button  = 'BtnAppDependency'
        Scopes  = @('DeviceManagementApps.Read.All')
    }
    @{
        Key     = 'AssetStatus'
        Enabled = $true
        Path    = 'Modules\AssetStatus\AssetStatus.ps1'
        Button  = 'BtnAssetStatus'
        Scopes  = @('DeviceManagementManagedDevices.ReadWrite.All')
    }
    @{
        Key     = 'BulkAssetStatus'
        Enabled = $true
        Path    = 'Modules\BulkAssetStatus\BulkAssetStatus.ps1'
        Button  = 'BtnBulkAssetStatus'
        Scopes  = @('DeviceManagementManagedDevices.ReadWrite.All')
    }
    @{
        Key     = 'DevicePrepCheck'
        Enabled = $true
        Path    = 'Modules\DevicePrepCheck\DevicePrepCheck.ps1'
        Button  = 'BtnDevicePrepCheck'
        Scopes  = @('DeviceManagementServiceConfig.Read.All','DeviceManagementManagedDevices.ReadWrite.All',
                    'Device.Read.All','Directory.AccessAsUser.All','GroupMember.Read.All')
    }
)

# ---------------------------------------------------------------------------
# Optional per-machine override file.
#
# Drop a ModuleConfig.psd1 next to Toolkit.ps1 to change settings without
# editing this script - handy when the same build is shared across sites:
#
#     @{
#         Modules = @{ AppDependencyCheck = $false; BulkAddToGroup = $false }
#         AssetStatusValues = @('In-Stock','Retired','Loaner')
#     }
#
# Keys not named in the file keep the values set above.
#
# The older flat layout - module keys sitting at the top level, with no
# Modules sub-table - is still honoured so existing files keep working.
# ---------------------------------------------------------------------------
$Script:ModuleRoot = $PSScriptRoot

# Read by Modules\AssetStatus\AssetStatus.ps1 when it is dot-sourced below.
# $null means 'not configured', which leaves the module's built-in list alone.
# An empty or all-blank list is treated the same way.
$Script:AssetStatusValues = $null

# Branding defaults. ModuleConfig.psd1 can override these values.
$Script:ToolkitName = 'EndpointGuy Intune Toolkit'
$Script:WebsiteText = 'endpointguy.com'
$Script:WebsiteUrl  = 'https://endpointguy.com'
$Script:LogoFile    = 'EGLogoNew.png'

$moduleConfigPath = Join-Path $Script:ModuleRoot 'ModuleConfig.psd1'
if (Test-Path $moduleConfigPath) {
    try {
        $overrides = Import-PowerShellDataFile -Path $moduleConfigPath

        if ($overrides.ContainsKey('Branding') -and $overrides['Branding'] -is [hashtable]) {
            $branding = $overrides['Branding']
            if ($branding.ContainsKey('ToolkitName') -and -not [string]::IsNullOrWhiteSpace([string]$branding['ToolkitName'])) { $Script:ToolkitName = ([string]$branding['ToolkitName']).Trim() }
            if ($branding.ContainsKey('WebsiteText') -and -not [string]::IsNullOrWhiteSpace([string]$branding['WebsiteText'])) { $Script:WebsiteText = ([string]$branding['WebsiteText']).Trim() }
            if ($branding.ContainsKey('WebsiteUrl') -and -not [string]::IsNullOrWhiteSpace([string]$branding['WebsiteUrl'])) { $Script:WebsiteUrl = ([string]$branding['WebsiteUrl']).Trim() }
            if ($branding.ContainsKey('LogoFile') -and -not [string]::IsNullOrWhiteSpace([string]$branding['LogoFile'])) { $Script:LogoFile = ([string]$branding['LogoFile']).Trim() }
        }

        # Prefer the Modules sub-table; fall back to the flat top-level form.
        $moduleSwitches = $overrides
        if ($overrides.ContainsKey('Modules') -and $overrides['Modules'] -is [hashtable]) {
            $moduleSwitches = $overrides['Modules']
        }

        foreach ($entry in $Script:ModuleRegistry) {
            if ($moduleSwitches.ContainsKey($entry.Key)) {
                $entry.Enabled = [bool]$moduleSwitches[$entry.Key]
            }
        }

        # Asset Status lifecycle labels. Trim blanks and drop duplicates so a
        # stray comma or a repeated line cannot produce an empty or doubled
        # drop-down entry.
        if ($overrides.ContainsKey('AssetStatusValues')) {
            $wanted = @($overrides['AssetStatusValues']) |
                      ForEach-Object { [string]$_ } |
                      Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
                      ForEach-Object { $_.Trim() }

            $wanted = $wanted | Select-Object -Unique

            if ($wanted.Count -gt 0) {
                $Script:AssetStatusValues = @($wanted)
            }
            else {
                Show-ToolkitMessageBox -Title 'Module configuration' -Button 'OK' -Icon 'Warning' -Message "AssetStatusValues in ModuleConfig.psd1 is empty, so the built-in statuses are being used." | Out-Null
            }
        }
    }
    catch {
        Show-ToolkitMessageBox -Title 'Module configuration' -Button 'OK' -Icon 'Warning' -Message (
            "ModuleConfig.psd1 could not be read, so the built-in settings are being used.`n`n$($_.Exception.Message)") | Out-Null
    }
}

# ---------------------------------------------------------------------------
# Load every enabled module.
# ---------------------------------------------------------------------------
$Script:ModulesLoaded  = @{}
$Script:ModulesEnabled = @{}

foreach ($entry in $Script:ModuleRegistry) {
    $Script:ModulesEnabled[$entry.Key] = [bool]$entry.Enabled

    # Switched off for this build - do not load it, do not touch the button yet.
    if (-not $entry.Enabled) {
        $Script:ModulesLoaded[$entry.Key] = $false
        continue
    }

    $moduleName = Split-Path -Leaf $entry.Path

    # Preferred subfolder location, then the old flat location.
    $candidates = @(
        (Join-Path $Script:ModuleRoot $entry.Path)
        (Join-Path $Script:ModuleRoot $moduleName)
    )
    $modulePath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($modulePath) {
        try {
            . $modulePath
            $Script:ModulesLoaded[$entry.Key] = $true
        }
        catch {
            $Script:ModulesLoaded[$entry.Key] = $false
            Show-ToolkitMessageBox -Title 'Module load error' -Button 'OK' -Icon 'Warning' -Message (
                "Module '$moduleName' could not be loaded:`n`n$($_.Exception.Message)") | Out-Null
        }
    }
    else {
        $Script:ModulesLoaded[$entry.Key] = $false
    }
}

# Returns $true only when the module is switched on AND its file loaded.
function Test-ModuleEnabled {
    param([string]$Key)
    return ($Script:ModulesEnabled.ContainsKey($Key) -and $Script:ModulesEnabled[$Key])
}

# ---------------------------------------------------------------------------
# Graph scopes
#
# Base scopes cover the shell itself (search, cache, device details). Each
# enabled module then contributes its own scopes, so a build with modules
# switched off asks the tenant for less.
# ---------------------------------------------------------------------------
$Script:GraphScopes = @(
    'DeviceManagementManagedDevices.ReadWrite.All'
    'DeviceManagementConfiguration.Read.All'
    'Device.Read.All'
    'User.Read.All'
    'DeviceManagementServiceConfig.Read.All'
)

foreach ($entry in $Script:ModuleRegistry) {
    if (-not $entry.Enabled) { continue }
    foreach ($scope in $entry.Scopes) {
        if ($Script:GraphScopes -notcontains $scope) { $Script:GraphScopes += $scope }
    }
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$XamlString = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="EndpointGuy Intune Toolkit"
        Height="1000" Width="1600"
        WindowStartupLocation="CenterScreen"
        Background="{DynamicResource WindowBrush}"
        FontFamily="Segoe UI">

    <Window.Resources>
        <!-- Theme brushes (dark). -->
        <!-- Theme brushes. Set-Theme swaps these Color values at runtime. -->
        <SolidColorBrush x:Key="WindowBrush"           Color="#1F1F1F"/>
        <SolidColorBrush x:Key="CardBrush"             Color="#2B2B2B"/>
        <SolidColorBrush x:Key="CardBorderBrush"       Color="#3D3D3D"/>
        <SolidColorBrush x:Key="TextBrush"             Color="#E8E8E8"/>
        <SolidColorBrush x:Key="TitleTextBrush"        Color="#FFFFFF"/>
        <SolidColorBrush x:Key="SubtleTextBrush"       Color="#A6A6A6"/>
        <SolidColorBrush x:Key="BulletBrush"           Color="#6E6E6E"/>
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
        <SolidColorBrush x:Key="InputHoverBrush"       Color="#383838"/>
        <SolidColorBrush x:Key="InputBorderBrush"      Color="#4D4D4D"/>
        <SolidColorBrush x:Key="DropDownBrush"         Color="#2B2B2B"/>
        <SolidColorBrush x:Key="DropDownBorderBrush"   Color="#4D4D4D"/>
        <SolidColorBrush x:Key="ComboArrowBrush"       Color="#C8C8C8"/>
        <SolidColorBrush x:Key="ItemHoverBrush"        Color="#3A3A3A"/>
        <SolidColorBrush x:Key="ItemSelectedBrush"     Color="#0F4C7A"/>
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
        <SolidColorBrush x:Key="SeparatorBrush"        Color="#3D3D3D"/>
        <SolidColorBrush x:Key="ConnGoodBrush"         Color="#4ADE80"/>
        <SolidColorBrush x:Key="ConnBadBrush"          Color="#FF6B6B"/>
        <SolidColorBrush x:Key="OnAccentTextBrush"     Color="#FFFFFF"/>

        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource CardBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
        </Style>

        <Style x:Key="CardHeader" TargetType="TextBlock">
            <Setter Property="FontSize" Value="19"/>
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
                        <Border x:Name="bd" CornerRadius="6"
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

        <Style x:Key="ActionButton" TargetType="Button" BasedOn="{StaticResource NeutralButton}">
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="14,11"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="FontSize" Value="13.5"/>
        </Style>

        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource NeutralButton}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource OnAccentTextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="6"
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

        <Style x:Key="GreenButton" TargetType="Button" BasedOn="{StaticResource PrimaryButton}">
            <Setter Property="Background" Value="{StaticResource GreenBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource GreenBrush}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" CornerRadius="6"
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

        <!-- Fully retemplated ComboBox: the stock control keeps a system-drawn -->
        <!-- chrome that ignores Background, so it renders unreadable dark text -->
        <!-- on a dark surface. Templating it is the only reliable fix.         -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Padding" Value="9,6"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="bd" CornerRadius="3" Background="Transparent"
                                Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                            <ContentPresenter TextBlock.Foreground="{TemplateBinding Foreground}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource ItemHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="{StaticResource ItemSelectedBrush}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="{StaticResource InputBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource InputBorderBrush}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="MainToggle" Focusable="False" ClickMode="Press"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                                <ToggleButton.Style>
                                    <Style TargetType="ToggleButton">
                                        <Setter Property="Template">
                                            <Setter.Value>
                                                <ControlTemplate TargetType="ToggleButton">
                                                    <Border x:Name="tbd" CornerRadius="4"
                                                            Background="{DynamicResource InputBrush}"
                                                            BorderBrush="{DynamicResource InputBorderBrush}"
                                                            BorderThickness="1">
                                                        <Path x:Name="arrow" HorizontalAlignment="Right"
                                                              VerticalAlignment="Center" Margin="0,0,11,0"
                                                              Data="M0,0 L8,0 L4,5 Z"
                                                              Fill="{DynamicResource ComboArrowBrush}"/>
                                                    </Border>
                                                    <ControlTemplate.Triggers>
                                                        <Trigger Property="IsMouseOver" Value="True">
                                                            <Setter TargetName="tbd" Property="Background" Value="{DynamicResource InputHoverBrush}"/>
                                                        </Trigger>
                                                    </ControlTemplate.Triggers>
                                                </ControlTemplate>
                                            </Setter.Value>
                                        </Setter>
                                    </Style>
                                </ToggleButton.Style>
                            </ToggleButton>

                            <ContentPresenter x:Name="SelectionSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              TextBlock.Foreground="{DynamicResource TextBrush}"
                                              Margin="11,0,30,0"
                                              VerticalAlignment="Center" HorizontalAlignment="Left"/>

                            <Popup x:Name="DropDownPopup" Placement="Bottom" Focusable="False"
                                   AllowsTransparency="True" PopupAnimation="Slide"
                                   IsOpen="{TemplateBinding IsDropDownOpen}">
                                <Grid MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}"
                                      SnapsToDevicePixels="True" Margin="0,3,0,0">
                                    <Border Background="{DynamicResource DropDownBrush}"
                                            BorderBrush="{DynamicResource DropDownBorderBrush}"
                                            BorderThickness="1" CornerRadius="4"/>
                                    <ScrollViewer Margin="4" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True"
                                                    KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
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
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Style="{StaticResource Card}" Padding="16,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Image x:Name="LogoImage" Height="30" Width="30" Margin="0,0,11,0"
                           VerticalAlignment="Center" Stretch="Uniform"
                           RenderOptions.BitmapScalingMode="HighQuality"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="HeaderTitle" Text="EndpointGuy Intune Toolkit"
                                   FontSize="19" FontWeight="SemiBold"
                                   Foreground="{StaticResource TitleTextBrush}"/>
                        <TextBlock FontSize="11.5" Margin="0,1,0,0">
                            <Hyperlink x:Name="LnkSite"
                                       NavigateUri="https://endpointguy.com"
                                       Foreground="{StaticResource AccentTextBrush}"
                                       TextDecorations="None"
                                       ToolTip="https://endpointguy.com">endpointguy.com</Hyperlink>
                        </TextBlock>
                    </StackPanel>
                </StackPanel>
                <Border Grid.Column="1" CornerRadius="14" Padding="10,5"
                        Background="{StaticResource InputBrush}"
                        BorderBrush="{StaticResource InputBorderBrush}" BorderThickness="1"
                        VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock x:Name="ConnIcon" Text="&#xE711;"
                                   FontFamily="Segoe MDL2 Assets" FontSize="14"
                                   Foreground="{StaticResource ConnBadBrush}"
                                   VerticalAlignment="Center" Margin="0,0,7,0"/>
                        <StackPanel VerticalAlignment="Center">
                            <TextBlock x:Name="ConnStatus" Text="Not connected"
                                       FontWeight="SemiBold" FontSize="11.5"
                                       Foreground="{StaticResource ConnBadBrush}"/>
                            <TextBlock x:Name="ConnAccount" Text="Sign in to begin"
                                       FontSize="10.5" MaxWidth="210" TextTrimming="CharacterEllipsis"
                                       Foreground="{StaticResource SubtleTextBrush}"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="0,14,0,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="14"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="14"/>
                    <ColumnDefinition Width="500"/>
                </Grid.ColumnDefinitions>

                <Border Grid.Column="2" Style="{StaticResource Card}" Padding="16">
                    <StackPanel>
                        <Grid Margin="0,0,0,12">
                            <TextBlock Text="Tools" Style="{StaticResource CardHeader}" VerticalAlignment="Center"/>
                            <Button x:Name="BtnToolsInfo" Style="{StaticResource NeutralButton}"
                                    Content="&#xE946;" FontFamily="Segoe MDL2 Assets" FontSize="15"
                                    Width="30" Height="30" Padding="0" HorizontalAlignment="Right"
                                    ToolTip="What do these tools do?"/>
                        </Grid>
                        <!-- All four tools sit in an even 2 x 2 grid so every button is the same size. -->
                        <UniformGrid Columns="2" Margin="-4">
                            <Grid x:Name="DeviceToolsSection" Margin="4">
                                <Button x:Name="BtnDeviceMenu" Style="{StaticResource NeutralButton}"
                                        Content="Device Actions  &#x25BE;" Height="34" Padding="14,0" FocusVisualStyle="{x:Null}" IsEnabled="False"/>
                                <Popup x:Name="DeviceMenuPopup" Placement="Bottom"
                                       PlacementTarget="{Binding ElementName=BtnDeviceMenu}"
                                       StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade">
                                    <Border Background="{DynamicResource DropDownBrush}"
                                            BorderBrush="{DynamicResource DropDownBorderBrush}"
                                            BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,4,0,0">
                                        <StackPanel Width="235">
                                            <TextBlock x:Name="SelectedDeviceText" Text="Select a device first"
                                                       FontSize="11.5" TextWrapping="Wrap"
                                                       Foreground="{DynamicResource SubtleTextBrush}" Margin="8,4,8,10"/>
                                            <Button x:Name="BtnCopyGroups" Style="{StaticResource ActionButton}"
                                                    Content="Copy Device Groups"/>
                                            <Button x:Name="BtnRemoveGroups" Style="{StaticResource ActionButton}"
                                                    Content="Remove Device Groups"/>
                                            <Button x:Name="BtnAssetStatus" Style="{StaticResource ActionButton}"
                                                    Content="Update Asset Status" Margin="0"/>
                                        </StackPanel>
                                    </Border>
                                </Popup>
                            </Grid>
                            <Grid x:Name="BulkToolsSection" Margin="4">
                                <Button x:Name="BtnBulkMenu" Style="{StaticResource NeutralButton}"
                                        Content="Bulk Actions  &#x25BE;" Height="34" Padding="14,0" FocusVisualStyle="{x:Null}"/>
                                <Popup x:Name="BulkMenuPopup" Placement="Bottom"
                                       PlacementTarget="{Binding ElementName=BtnBulkMenu}"
                                       StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade">
                                    <Border Background="{DynamicResource DropDownBrush}"
                                            BorderBrush="{DynamicResource DropDownBorderBrush}"
                                            BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,4,0,0">
                                        <StackPanel Width="235">
                                            <Button x:Name="BtnBulkAddGroup" Style="{StaticResource ActionButton}"
                                                    Content="Bulk Add to Group"/>
                                            <Button x:Name="BtnBulkAssetStatus" Style="{StaticResource ActionButton}"
                                                    Content="Bulk Update Asset Status" Margin="0"/>
                                        </StackPanel>
                                    </Border>
                                </Popup>
                            </Grid>
                            <Grid x:Name="ReportToolsSection" Margin="4">
                                <Button x:Name="BtnReportMenu" Style="{StaticResource NeutralButton}"
                                        Content="Reporting  &#x25BE;" Height="34" Padding="14,0" FocusVisualStyle="{x:Null}"/>
                                <Popup x:Name="ReportMenuPopup" Placement="Bottom"
                                       PlacementTarget="{Binding ElementName=BtnReportMenu}"
                                       StaysOpen="False" AllowsTransparency="True" PopupAnimation="Fade">
                                    <Border Background="{DynamicResource DropDownBrush}"
                                            BorderBrush="{DynamicResource DropDownBorderBrush}"
                                            BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,4,0,0">
                                        <StackPanel Width="235">
                                            <Button x:Name="BtnAppDependency" Style="{StaticResource ActionButton}"
                                                    Content="App Dependency Check" Margin="0"/>
                                        </StackPanel>
                                    </Border>
                                </Popup>
                            </Grid>
                            <!-- Standalone tool, outside the drop-down menus. Works with or without a selected device. -->
                            <Button x:Name="BtnDevicePrepCheck" Style="{StaticResource NeutralButton}"
                                    Height="34" Padding="14,0" Margin="4" FocusVisualStyle="{x:Null}"
                                    ToolTip="Check a serial number is ready for Autopilot provisioning">
                                <StackPanel Orientation="Horizontal">
                                    <TextBlock Text="Device Prep Check" VerticalAlignment="Center"/>
                                    <Border Background="#15803D" CornerRadius="8" Padding="7,1"
                                            Margin="10,0,0,0" VerticalAlignment="Center">
                                        <TextBlock Text="NEW" FontSize="10" FontWeight="Bold" Foreground="#FFFFFF"/>
                                    </Border>
                                </StackPanel>
                            </Button>
                        </UniformGrid>
                    </StackPanel>
                </Border>

                <Border Grid.Column="0" Style="{StaticResource Card}" Padding="16">
                    <StackPanel>
                        <TextBlock Text="Device Search" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="150"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="90"/>
                            </Grid.ColumnDefinitions>
                            <ComboBox x:Name="CmbSearchField" Grid.Column="0" SelectedIndex="0">
                                <ComboBoxItem Content="Device Name"/>
                                <ComboBoxItem Content="Serial Number"/>
                                <ComboBoxItem Content="Primary Username"/>
                                <ComboBoxItem Content="Management Name"/>
                            </ComboBox>
                            <ComboBox x:Name="CmbAssetStatus" Grid.Column="1" Margin="10,0,10,0" Visibility="Collapsed"
                                      ToolTip="Asset Status lifecycle labels, from the Asset Status module."/>
                            <TextBox x:Name="TxtSearch" Grid.Column="1" Height="32" Margin="10,0,10,0"/>
                            <Button x:Name="BtnSearch" Grid.Column="2" Style="{StaticResource PrimaryButton}"
                                    Content="Search" Height="32"/>
                        </Grid>
                        <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                            <CheckBox x:Name="ChkExact" Content="Exact match" VerticalAlignment="Center"
                                      FontSize="12.5" Margin="0,0,14,0"/>
                            <Button x:Name="BtnRefreshCache" Style="{StaticResource NeutralButton}"
                                    Content="Refresh Cache" Height="30" Padding="11,4" Margin="0,0,8,0"/>
                            <Button x:Name="BtnAdvanced" Style="{StaticResource NeutralButton}"
                                    Content="Advanced Filter" Height="30" Padding="11,4"
                                    ToolTip="Build a multi-column filter. All rules must match."/>
                        </StackPanel>
                        <Border x:Name="AdvancedPanel" Visibility="Collapsed" Margin="0,12,0,0"
                                Background="{DynamicResource InputBrush}" BorderBrush="{DynamicResource InputBorderBrush}"
                                BorderThickness="1" CornerRadius="6" Padding="12">
                            <StackPanel>
                                <TextBlock Text="Match ALL rules:" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,8"/>
                                <StackPanel x:Name="RulesPanel"/>
                                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                                    <Button x:Name="BtnAddRule" Style="{StaticResource NeutralButton}" Content="+ Add rule" MinHeight="30" Padding="12,5"/>
                                    <Button x:Name="BtnApplyFilter" Style="{StaticResource PrimaryButton}" Content="Apply Filter" MinWidth="110" MinHeight="30" Padding="12,5" Margin="8,0,0,0"/>
                                    <Button x:Name="BtnClearFilter" Style="{StaticResource NeutralButton}" Content="Clear" MinHeight="30" Padding="12,5" Margin="8,0,0,0"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </Border>
            </Grid>

            <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Grid.Row="0" Margin="0,0,0,12">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Text="Search Results" Style="{StaticResource CardHeader}"/>
                            <TextBlock x:Name="ResultCount" Text="" Margin="12,0,0,0"
                                       VerticalAlignment="Bottom" FontSize="13"
                                       Foreground="{StaticResource SubtleTextBrush}"/>
                        </StackPanel>
                        <Button x:Name="BtnExportCsv" Style="{StaticResource NeutralButton}"
                                Content="Export to CSV" HorizontalAlignment="Right"/>
                    </Grid>
                    <DataGrid x:Name="GridDevices" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Single" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                              HorizontalGridLinesBrush="{StaticResource GridLineBrush}" Background="{StaticResource GridRowBrush}"
                              Foreground="{StaticResource TextBrush}" RowBackground="{StaticResource GridRowBrush}"
                              AlternatingRowBackground="{StaticResource GridAltRowBrush}" BorderBrush="{StaticResource GridBorderBrush}"
                              BorderThickness="1" FontSize="12.5" RowHeight="26" CanUserSortColumns="True">
                        <DataGrid.Columns>
                            <!-- Source: 'Intune' is a link that opens the device in the Intune
                                 admin center; 'Autopilot' (not enrolled yet) is plain text. -->
                            <DataGridTemplateColumn Header="Source" Width="95" SortMemberPath="Source">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid VerticalAlignment="Center" Margin="2,0">
                                            <TextBlock x:Name="SourcePlain" Text="{Binding Source, Mode=OneWay}"/>
                                            <TextBlock x:Name="SourceLink" Visibility="Collapsed"
                                                       ToolTip="Open this device in the Intune admin center">
                                                <Hyperlink Foreground="{StaticResource AccentTextBrush}" Cursor="Hand">
                                                    <Run Text="{Binding Source, Mode=OneWay}"/>
                                                </Hyperlink>
                                            </TextBlock>
                                        </Grid>
                                        <DataTemplate.Triggers>
                                            <DataTrigger Binding="{Binding Source}" Value="Intune">
                                                <Setter TargetName="SourcePlain" Property="Visibility" Value="Collapsed"/>
                                                <Setter TargetName="SourceLink" Property="Visibility" Value="Visible"/>
                                            </DataTrigger>
                                        </DataTemplate.Triggers>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <!-- Notes: a button on Intune rows that shows the device's notes from Intune. -->
                            <DataGridTemplateColumn Header="Notes" Width="72" CanUserSort="False">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Grid VerticalAlignment="Center">
                                            <Button x:Name="NotesButton" Tag="DeviceNotes" Visibility="Collapsed"
                                                    Style="{StaticResource NeutralButton}" Padding="0" Width="30" Height="22"
                                                    HorizontalAlignment="Left" Margin="4,0"
                                                    ToolTip="View this device's notes from Intune">
                                                <!-- Note icon: a page with a folded corner and three lines of text. -->
                                                <Path Width="12" Height="14" Stretch="None" StrokeThickness="1.2"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                                                      Stroke="{StaticResource AccentTextBrush}" IsHitTestVisible="False"
                                                      Data="M1,0.6 L7.6,0.6 L11,4 L11,13.4 L1,13.4 Z M7.6,0.6 L7.6,4 L11,4 M3.2,6.8 L8.8,6.8 M3.2,9 L8.8,9 M3.2,11.2 L6.6,11.2"/>
                                            </Button>
                                            <TextBlock x:Name="NotesNone" Text="-" Margin="8,0"
                                                       Foreground="{StaticResource SubtleTextBrush}"
                                                       ToolTip="Not enrolled in Intune yet - no notes"/>
                                        </Grid>
                                        <DataTemplate.Triggers>
                                            <DataTrigger Binding="{Binding Source}" Value="Intune">
                                                <Setter TargetName="NotesButton" Property="Visibility" Value="Visible"/>
                                                <Setter TargetName="NotesNone" Property="Visibility" Value="Collapsed"/>
                                            </DataTrigger>
                                        </DataTemplate.Triggers>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="200"/>
                            <DataGridTextColumn Header="Serial Number" Binding="{Binding SerialNumber}" Width="140"/>
                            <DataGridTextColumn Header="Management Name" Binding="{Binding ManagementName}" Width="200"/>
                            <DataGridTextColumn Header="User" Binding="{Binding User}" Width="220"/>
                            <DataGridTextColumn Header="OS" Binding="{Binding OS}" Width="100"/>
                            <DataGridTextColumn Header="OS Version" Binding="{Binding OSVersion}" Width="130"/>
                            <DataGridTextColumn Header="Compliance" Binding="{Binding Compliance}" Width="110"/>
                            <DataGridTextColumn Header="Model" Binding="{Binding Model}" Width="180"/>
                            <DataGridTextColumn Header="Last Sync" Binding="{Binding LastSync}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>

        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="14,9" Margin="0,14,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusText" Grid.Column="0"
                           Text="Ready. Connect to Microsoft Graph to begin."
                           FontSize="12.5" Foreground="{StaticResource TextBrush}"
                           VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                <TextBlock x:Name="ClockText" Grid.Column="2" Text=""
                           FontSize="12" Foreground="{StaticResource SubtleTextBrush}"
                           VerticalAlignment="Center" HorizontalAlignment="Right"/>
            </Grid>
        </Border>
        <Grid x:Name="LoadingOverlay" Grid.RowSpan="3" Panel.ZIndex="100"
              Visibility="Collapsed" Background="#B8000000">
            <Border Width="500" Padding="30,26" HorizontalAlignment="Center" VerticalAlignment="Center"
                    Background="{DynamicResource CardBrush}" BorderBrush="{DynamicResource AccentBrush}"
                    BorderThickness="1" CornerRadius="10">
                <StackPanel>
                    <TextBlock Text="" FontSize="13" FontWeight="SemiBold"
                               Foreground="{DynamicResource AccentTextBrush}" HorizontalAlignment="Center"/>
                    <TextBlock x:Name="LoadingText" Text="Working..." FontSize="21" FontWeight="SemiBold"
                               Foreground="{DynamicResource TitleTextBrush}" HorizontalAlignment="Center"
                               TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,18"/>
                    <Border x:Name="LoadingTrack" Height="14" CornerRadius="7"
                            Background="{DynamicResource InputBrush}"
                            BorderBrush="{DynamicResource InputBorderBrush}" BorderThickness="1"
                            ClipToBounds="True">
                        <Grid>
                            <Border x:Name="LoadingIndicator" Width="110" Height="12"
                                    HorizontalAlignment="Left" CornerRadius="6"
                                    Background="{DynamicResource AccentBrush}">
                                <Border.RenderTransform>
                                    <TranslateTransform X="-110"/>
                                </Border.RenderTransform>
                            </Border>
                        </Grid>
                    </Border>
                    <TextBlock x:Name="LoadingDetail" Text="Please keep this window open while the operation completes."
                               FontSize="11.5" Foreground="{DynamicResource SubtleTextBrush}" TextWrapping="Wrap" TextAlignment="Center"
                               HorizontalAlignment="Center" Margin="0,14,0,0"/>
                </StackPanel>
            </Border>
        </Grid>

    </Grid>
</Window>
"@

try {
    [xml]$xamlDoc = $XamlString
    $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
    $Window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    Show-ToolkitMessageBox -Title 'XAML load error' -Button 'OK' -Icon 'Error' -Message (
        "The XAML could not be loaded:`n`n$($_.Exception.Message)") | Out-Null
    return
}

# ---------------------------------------------------------------------------
# Map every x:Name into a script-scope variable of the same name.
# NOTE: x:Name lives in the XAML namespace, so GetAttribute('Name') returns
# an empty string; the attribute has to be matched by local-name() instead.
# ---------------------------------------------------------------------------
$namesFound = 0
foreach ($node in $xamlDoc.SelectNodes("//*[@*[local-name()='Name']]")) {
    $attr = $node.Attributes | Where-Object { $_.LocalName -eq 'Name' } | Select-Object -First 1
    if (-not $attr) { continue }

    $ctl = $Window.FindName($attr.Value)

    # Controls declared inside a ControlTemplate (the "bd" borders) are not
    # reachable from window scope and return null - skip them.
    if ($null -ne $ctl) {
        Set-Variable -Name $attr.Value -Value $ctl -Scope Script
        $namesFound++
    }
}

if ($namesFound -eq 0) {
    Show-ToolkitMessageBox -Title 'XAML load error' -Button 'OK' -Icon 'Error' -Message 'No named controls could be resolved from the XAML.' | Out-Null
    return
}

foreach ($required in @('BtnSearch','BtnRefreshCache','GridDevices',
                        'TxtSearch','CmbSearchField','CmbAssetStatus','ChkExact','StatusText',
                        'LoadingOverlay','LoadingText','LoadingDetail','LoadingTrack','LoadingIndicator',
                        'ConnIcon','ConnStatus','ConnAccount','ResultCount',
                        'BtnAdvanced','AdvancedPanel','RulesPanel','BtnAddRule','BtnApplyFilter','BtnClearFilter',
                        'SelectedDeviceText','ClockText','BtnExportCsv',
                        'BtnToolsInfo','BtnDeviceMenu','DeviceMenuPopup','DeviceToolsSection',
                        'BtnBulkMenu','BulkMenuPopup','BulkToolsSection',
                        'BtnReportMenu','ReportMenuPopup','ReportToolsSection',
                        'BtnCopyGroups','BtnRemoveGroups','BtnAssetStatus',
                        'BtnBulkAddGroup','BtnBulkAssetStatus','BtnAppDependency','BtnDevicePrepCheck')) {
    if (-not (Get-Variable -Name $required -Scope Script -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'XAML load error' -Button 'OK' -Icon 'Error' -Message (
            "Control '$required' was not found in the XAML.") | Out-Null
        return
    }
}

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
$Script:DeviceCache    = @()
$Script:CacheLoadedAt  = $null
$Script:SelectedDevice = $null
$Script:Results        = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$GridDevices.ItemsSource = $Script:Results

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Error log and safety net
#
# Any error that escapes a button or timer is written to
#   %LOCALAPPDATA%\EndpointGuy\IntuneToolkit\Toolkit-errors.log
# and shown in a message box, instead of silently closing the toolkit.
# ---------------------------------------------------------------------------
$Script:ErrorLogPath = Join-Path $env:LOCALAPPDATA 'EndpointGuy\IntuneToolkit\Toolkit-errors.log'

function Write-ToolkitErrorLog {
    param([string]$Context, $ErrorRecord)
    try {
        $dir = Split-Path -Parent $Script:ErrorLogPath
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

        $ex = $ErrorRecord
        $isRecord = $ErrorRecord -is [System.Management.Automation.ErrorRecord]
        if ($isRecord) { $ex = $ErrorRecord.Exception }

        $lines = @(
            ('=' * 70)
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Context"
            "Error: $($ex.Message)"
            "Type : $($ex.GetType().FullName)"
        )
        if ($isRecord) {
            if ($ErrorRecord.InvocationInfo) { $lines += "Where: $($ErrorRecord.InvocationInfo.PositionMessage)" }
            if ($ErrorRecord.ScriptStackTrace) { $lines += "Stack:`n$($ErrorRecord.ScriptStackTrace)" }
        }
        if ($ex.InnerException) { $lines += "Inner: $($ex.InnerException.Message)" }
        $lines += "Detail:`n$($ex.ToString())"
        Add-Content -Path $Script:ErrorLogPath -Value $lines -Encoding UTF8
    }
    catch { }
    return $Script:ErrorLogPath
}

# Safety net: an error inside any button, menu or timer is logged and shown,
# and the toolkit stays open instead of closing.
$Window.Dispatcher.Add_UnhandledException({
    param($dispatcherSource, $e)
    $e.Handled = $true
    $logFile = Write-ToolkitErrorLog -Context 'Unhandled window error' -ErrorRecord $e.Exception
    try { $LoadingOverlay.Visibility = [System.Windows.Visibility]::Collapsed } catch { }
    try { $Window.Cursor = $null } catch { }
    Show-ToolkitMessageBox -Title 'EndpointGuy Intune Toolkit' -Button 'OK' -Icon 'Error' -Message (
        "Something went wrong, but the toolkit is still open:`n`n$($e.Exception.Message)`n`nDetails were saved to:`n$logFile") | Out-Null
})

function Set-Status {
    param([string]$Message)
    $StatusText.Text = $Message
    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
}

function Show-ComingSoon {
    param([string]$Feature)
    Show-ToolkitMessageBox -Title 'EndpointGuy Intune Toolkit' -Button 'OK' -Icon 'Information' -Message ("$Feature`n`nComing soon.") | Out-Null
    Set-Status "$Feature - coming soon."
}

function Set-Busy {
    param([bool]$Busy, [string]$Activity = 'Working...')
    $Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    foreach ($b in @($BtnSearch, $BtnRefreshCache)) { $b.IsEnabled = -not $Busy }
    $LoadingText.Text = $Activity
    $LoadingDetail.Text = 'Please keep this window open while the operation completes.'

    $slide = $LoadingIndicator.RenderTransform
    if ($Busy) {
        $LoadingOverlay.Visibility = [System.Windows.Visibility]::Visible

        # Slide the blue segment from left to right, then start again from the
        # left. It starts just off the left edge and finishes just past the
        # right edge of the track, so each pass enters and leaves smoothly.
        $trackWidth = [double]$LoadingTrack.ActualWidth
        if ($trackWidth -lt 50) { $trackWidth = 328 }
        $animation = New-Object System.Windows.Media.Animation.DoubleAnimation
        $animation.From = -110
        $animation.To = $trackWidth
        $animation.Duration = [System.Windows.Duration]::new([TimeSpan]::FromSeconds(1.2))
        $animation.AutoReverse = $false
        $animation.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
        $slide.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $animation)
    }
    else {
        $slide.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $null)
        $slide.X = -110
        $LoadingOverlay.Visibility = [System.Windows.Visibility]::Collapsed
    }

    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
}

function Get-GraphContextSafe {
    try { Get-MgContext } catch { $null }
}

function Test-GraphProperty {
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Update-ConnectionUi {
    $ctx = Get-GraphContextSafe
    if ($ctx -and $ctx.Account) {
        $ConnIcon.Text         = [char]0xE73E
        $ConnIcon.Foreground   = $Window.TryFindResource('ConnGoodBrush')
        $ConnStatus.Text       = 'Connected'
        $ConnStatus.Foreground = $Window.TryFindResource('ConnGoodBrush')
        $ConnAccount.Text      = $ctx.Account
    }
    else {
        $ConnIcon.Text         = [char]0xE711
        $ConnIcon.Foreground   = $Window.TryFindResource('ConnBadBrush')
        $ConnStatus.Text       = 'Not connected'
        $ConnStatus.Foreground = $Window.TryFindResource('ConnBadBrush')
        $ConnAccount.Text      = 'Sign in to begin'
    }
}

function Test-Connected {
    $ctx = Get-GraphContextSafe
    if (-not ($ctx -and $ctx.Account)) {
        Show-ToolkitMessageBox -Title 'Not connected' -Button 'OK' -Icon 'Warning' -Message 'Connect to Microsoft Graph first.' | Out-Null
        return $false
    }
    return $true
}

# ---------------------------------------------------------------------------
# Device cache
# ---------------------------------------------------------------------------
function Get-ManagedDeviceCache {
    $select = 'id,deviceName,managedDeviceName,serialNumber,userPrincipalName,operatingSystem,' +
              'osVersion,complianceState,model,manufacturer,' +
              'lastSyncDateTime,enrolledDateTime,azureADDeviceId,deviceCategoryDisplayName'

    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$select&`$top=1000"
    $all  = New-Object System.Collections.Generic.List[object]
    $page = 0

    while ($uri) {
        $page++
        Set-Status "Loading device cache... page $page ($($all.Count) devices so far)"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

        if (Test-GraphProperty $resp 'value') { $all.AddRange(@($resp.value)) }

        $uri = if (Test-GraphProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
    }

    foreach ($d in $all) {
        [pscustomobject]@{
            DeviceName    = $d.deviceName
            SerialNumber  = $d.serialNumber
            ManagementName = $d.managedDeviceName
            User          = $d.userPrincipalName
            OS            = $d.operatingSystem
            OSVersion     = $d.osVersion
            Compliance    = $d.complianceState
            Model         = $d.model
            Manufacturer  = $d.manufacturer
            Category      = $d.deviceCategoryDisplayName
            LastSync      = if ($d.lastSyncDateTime) { ([datetime]$d.lastSyncDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Enrolled      = if ($d.enrolledDateTime) { ([datetime]$d.enrolledDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            IntuneId      = $d.id
            EntraDeviceId = $d.azureADDeviceId
            Source          = 'Intune'
            EnrollmentState = 'enrolled'
            GroupTag        = ''
        }
    }
}

function Get-AutopilotProperty {
    param($Object, [string]$Name, $Default = '')
    if ((Test-GraphProperty $Object $Name) -and ($null -ne $Object.$Name)) { $Object.$Name }
    else { $Default }
}

function Format-AutopilotDate {
    param($Value)
    if (-not $Value) { return '' }
    try   { ([datetime]$Value).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
    catch { '' }
}

# ---------------------------------------------------------------------------
# Windows Autopilot device identities.
#
# This is the second half of the device cache. A device that has been imported
# into Autopilot has an Entra ID device record from the moment it is imported,
# but it has NO deviceManagement/managedDevices entry until it actually enrols
# - which is why searching for a brand new machine used to return nothing.
#
# No $select is used: the resource returns a small fixed field set, and asking
# for a projection it does not support fails the whole call.
# ---------------------------------------------------------------------------
function Get-AutopilotDeviceCache {
    $uri  = 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=500'
    $all  = New-Object System.Collections.Generic.List[object]
    $page = 0

    while ($uri) {
        $page++
        Set-Status "Loading Autopilot devices... page $page ($($all.Count) so far)"
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject

        if (Test-GraphProperty $resp 'value') { $all.AddRange(@($resp.value)) }

        $uri = if (Test-GraphProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
    }

    foreach ($d in $all) {
        $serial = [string](Get-AutopilotProperty $d 'serialNumber')

        # displayName is usually empty until the device enrols, so fall back to
        # the serial rather than putting a nameless row in the grid.
        $name = [string](Get-AutopilotProperty $d 'displayName')
        if ([string]::IsNullOrWhiteSpace($name)) { $name = $serial }

        # The Entra deviceId is what the group modules bind to. The property
        # was renamed, so accept either spelling.
        $entraId = [string](Get-AutopilotProperty $d 'azureAdDeviceId')
        if (-not $entraId) { $entraId = [string](Get-AutopilotProperty $d 'azureActiveDirectoryDeviceId') }

        $user = [string](Get-AutopilotProperty $d 'userPrincipalName')
        if (-not $user) { $user = [string](Get-AutopilotProperty $d 'addressableUserName') }

        [pscustomobject]@{
            DeviceName      = $name
            SerialNumber    = $serial
            ManagementName  = ''
            User            = $user
            OS              = 'Windows'
            OSVersion       = ''
            Compliance      = ''
            Model           = [string](Get-AutopilotProperty $d 'model')
            Manufacturer    = [string](Get-AutopilotProperty $d 'manufacturer')
            Category        = ''
            LastSync        = Format-AutopilotDate (Get-AutopilotProperty $d 'lastContactedDateTime' $null)
            Enrolled        = ''
            IntuneId        = [string](Get-AutopilotProperty $d 'managedDeviceId')
            EntraDeviceId   = $entraId
            Source          = 'Autopilot'
            EnrollmentState = [string](Get-AutopilotProperty $d 'enrollmentState')
            GroupTag        = [string](Get-AutopilotProperty $d 'groupTag')
        }
    }
}

# Original loader - runs on the window thread. Used only as a fallback when
# the Graph access token cannot be read for the background loader below.
function Update-DeviceCacheSync {
    if (-not (Test-Connected)) { return }
    Set-Busy $true 'Loading device data...'
    try {
        $intune = @(Get-ManagedDeviceCache)

        # A missing DeviceManagementServiceConfig.Read.All consent must not
        # cost us the Intune half of the cache, so this half fails soft.
        $autopilot = @()
        $apNote    = ''
        try { $autopilot = @(Get-AutopilotDeviceCache) }
        catch { $apNote = " Autopilot lookup unavailable: $($_.Exception.Message)" }

        # Serial number is the only identifier both sources always share, so it
        # is the de-duplication key: once a device enrols it appears in both
        # lists and the richer Intune row wins.
        $known = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($d in $intune) {
            if (-not [string]::IsNullOrWhiteSpace([string]$d.SerialNumber)) {
                [void]$known.Add([string]$d.SerialNumber)
            }
        }

        $extra = @($autopilot | Where-Object {
            (-not [string]::IsNullOrWhiteSpace([string]$_.SerialNumber)) -and
            (-not $known.Contains([string]$_.SerialNumber))
        })

        $Script:DeviceCache   = @($intune) + @($extra)
        $Script:CacheLoadedAt = Get-Date
        Set-Status ("Device cache refreshed - $($Script:DeviceCache.Count) device(s): " +
                    "$($intune.Count) from Intune, $($extra.Count) Autopilot-only " +
                    "(not yet enrolled), at $($Script:CacheLoadedAt.ToString('HH:mm:ss')).$apNote")
    }
    catch {
        Set-Status "Cache refresh failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Cache refresh failed' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
    finally { Set-Busy $false }
}

# ---------------------------------------------------------------------------
# Background device cache loader
#
# Loading thousands of devices from Graph takes a while. When that ran on the
# window's own thread, the window could not redraw, so it looked frozen and
# the loading bar stopped moving. The Graph paging now runs in a separate
# PowerShell runspace instead:
#   - the access token from the signed-in Graph session is handed to the
#     background runspace, which pages through Graph with Invoke-RestMethod
#   - progress is written to a synchronized hashtable
#   - a DispatcherTimer on the window thread shows that progress and finishes
#     the job when it is done, so the window stays responsive throughout
#
# If no token can be read from the Graph session, Update-DeviceCacheSync (the
# original in-window loader) is used instead so the cache still loads.
# ---------------------------------------------------------------------------
$Script:CacheJob = $null

$Script:CacheLoaderScript = {
    param([string]$Token, $State)

    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol =
        [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $headers = @{ Authorization = "Bearer $Token" }

    function Get-BgProperty {
        param($Object, [string]$Name, $Default = '')
        if (($null -ne $Object) -and ($Object.PSObject.Properties.Name -contains $Name) -and
            ($null -ne $Object.$Name)) { $Object.$Name } else { $Default }
    }

    function Format-BgDate {
        param($Value)
        if (-not $Value) { return '' }
        try   { ([datetime]$Value).ToLocalTime().ToString('yyyy-MM-dd HH:mm') }
        catch { '' }
    }

    # Pages through a Graph collection. Throttling (429) and transient server
    # errors are retried a few times before giving up.
    function Get-BgPages {
        param([string]$Uri, [string]$Label)
        $all  = New-Object System.Collections.Generic.List[object]
        $page = 0
        while ($Uri) {
            $page++
            $State.Status = "$Label - page $page ($($all.Count) so far)"
            $attempt = 0
            while ($true) {
                try {
                    $resp = Invoke-RestMethod -Method Get -Uri $Uri -Headers $headers -UseBasicParsing
                    break
                }
                catch {
                    $code    = 0
                    $webResp = $_.Exception.Response
                    if ($webResp) { try { $code = [int]$webResp.StatusCode } catch { $code = 0 } }
                    $attempt++
                    if ((($code -eq 429) -or ($code -ge 500)) -and $attempt -lt 5) {
                        $wait = 5
                        try { $ra = $webResp.Headers['Retry-After']; if ($ra) { $wait = [int]$ra } } catch { }
                        $State.Status = "$Label - Graph is busy, retrying in $wait second(s)"
                        Start-Sleep -Seconds $wait
                        continue
                    }
                    throw
                }
            }
            foreach ($item in @(Get-BgProperty $resp 'value' @())) { $all.Add($item) }
            $Uri = [string](Get-BgProperty $resp '@odata.nextLink' '')
        }
        return ,$all
    }

    try {
        # --- Intune managed devices ------------------------------------------
        $select = 'id,deviceName,managedDeviceName,serialNumber,userPrincipalName,operatingSystem,' +
                  'osVersion,complianceState,model,manufacturer,' +
                  'lastSyncDateTime,enrolledDateTime,azureADDeviceId,deviceCategoryDisplayName'
        $uri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=' + $select + '&$top=1000'
        $managed = Get-BgPages -Uri $uri -Label 'Loading Intune devices'

        $State.Status = "Preparing $($managed.Count) Intune device(s)..."
        $intune = New-Object System.Collections.Generic.List[object]
        $known  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($d in $managed) {
            if (-not [string]::IsNullOrWhiteSpace([string]$d.serialNumber)) { [void]$known.Add([string]$d.serialNumber) }
            $intune.Add([pscustomobject]@{
                DeviceName      = $d.deviceName
                SerialNumber    = $d.serialNumber
                ManagementName  = $d.managedDeviceName
                User            = $d.userPrincipalName
                OS              = $d.operatingSystem
                OSVersion       = $d.osVersion
                Compliance      = $d.complianceState
                Model           = $d.model
                Manufacturer    = $d.manufacturer
                Category        = $d.deviceCategoryDisplayName
                LastSync        = Format-BgDate $d.lastSyncDateTime
                Enrolled        = Format-BgDate $d.enrolledDateTime
                IntuneId        = $d.id
                EntraDeviceId   = $d.azureADDeviceId
                Source          = 'Intune'
                EnrollmentState = 'enrolled'
                GroupTag        = ''
            })
        }

        # --- Autopilot devices not yet enrolled -------------------------------
        # A missing DeviceManagementServiceConfig.Read.All consent must not
        # cost us the Intune half of the cache, so this half fails soft.
        $extra = New-Object System.Collections.Generic.List[object]
        try {
            $apUri = 'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=500'
            $autopilot = Get-BgPages -Uri $apUri -Label 'Loading Autopilot devices'

            foreach ($d in $autopilot) {
                $serial = [string](Get-BgProperty $d 'serialNumber')

                # Serial number is the de-duplication key: once a device enrols
                # it appears in both lists and the richer Intune row wins.
                if ([string]::IsNullOrWhiteSpace($serial) -or $known.Contains($serial)) { continue }

                $name = [string](Get-BgProperty $d 'displayName')
                if ([string]::IsNullOrWhiteSpace($name)) { $name = $serial }

                $entraId = [string](Get-BgProperty $d 'azureAdDeviceId')
                if (-not $entraId) { $entraId = [string](Get-BgProperty $d 'azureActiveDirectoryDeviceId') }

                $user = [string](Get-BgProperty $d 'userPrincipalName')
                if (-not $user) { $user = [string](Get-BgProperty $d 'addressableUserName') }

                $extra.Add([pscustomobject]@{
                    DeviceName      = $name
                    SerialNumber    = $serial
                    ManagementName  = ''
                    User            = $user
                    OS              = 'Windows'
                    OSVersion       = ''
                    Compliance      = ''
                    Model           = [string](Get-BgProperty $d 'model')
                    Manufacturer    = [string](Get-BgProperty $d 'manufacturer')
                    Category        = ''
                    LastSync        = Format-BgDate (Get-BgProperty $d 'lastContactedDateTime' $null)
                    Enrolled        = ''
                    IntuneId        = [string](Get-BgProperty $d 'managedDeviceId')
                    EntraDeviceId   = $entraId
                    Source          = 'Autopilot'
                    EnrollmentState = [string](Get-BgProperty $d 'enrollmentState')
                    GroupTag        = [string](Get-BgProperty $d 'groupTag')
                })
            }
        }
        catch { $State.ApNote = " Autopilot lookup unavailable: $($_.Exception.Message)" }

        # Hand the rows back as plain .NET data (ordered dictionaries of
        # strings in an object[]). PowerShell objects created in this runspace
        # can fail with "Argument types do not match" when the main window
        # enumerates them, so only pure .NET types cross back.
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($o in $intune) {
            $h = New-Object System.Collections.Specialized.OrderedDictionary
            foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = [string]$p.Value }
            $rows.Add($h)
        }
        $State.IntuneCount = $rows.Count
        foreach ($o in $extra) {
            $h = New-Object System.Collections.Specialized.OrderedDictionary
            foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = [string]$p.Value }
            $rows.Add($h)
        }
        $State.Rows   = $rows.ToArray()
        $State.Status = 'Device data loaded.'
    }
    catch {
        $State.Error = $_.Exception.Message
    }
}

# Reads the bearer token from the signed-in Graph session by making one tiny
# request and taking the Authorization header the Graph module attached.
function Get-GraphAccessToken {
    $uri  = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=1&$select=id'
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType HttpResponseMessage
    try {
        $auth = $resp.RequestMessage.Headers.Authorization
        if ($auth -and $auth.Parameter) { return [string]$auth.Parameter }
        return $null
    }
    finally { $resp.Dispose() }
}

# Starts a background device load. -OnComplete runs on the window thread once
# the cache has loaded successfully (used by Search and Apply Filter).
function Update-DeviceCache {
    param([scriptblock]$OnComplete)

    if ($Script:CacheJob) {
        Set-Status 'Device data is already loading - please wait...'
        return
    }
    if (-not (Test-Connected)) { return }

    Set-Busy $true 'Loading device data from Intune and Autopilot...'
    Set-Status 'Loading device data from Intune and Autopilot...'

    $token = $null
    try { $token = Get-GraphAccessToken } catch { $token = $null }

    if (-not $token) {
        # Fallback: load on the window thread (the bar pauses while it runs).
        Update-DeviceCacheSync
        if ($OnComplete -and $Script:DeviceCache.Count -gt 0) { & $OnComplete }
        return
    }

    $state = [hashtable]::Synchronized(@{
        Status = 'Connecting to Microsoft Graph...'
        Intune = $null
        Extra  = $null
        ApNote = ''
        Error  = $null
    })

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    [void]$ps.AddScript($Script:CacheLoaderScript.ToString()).AddArgument($token).AddArgument($state)

    $Script:CacheJob = @{
        PS         = $ps
        Runspace   = $runspace
        Handle     = $ps.BeginInvoke()
        State      = $state
        OnComplete = $OnComplete
    }
    $Script:CacheTimer.Start()
}

# Runs on the window thread every 200 ms while a background load is active.
function Complete-DeviceCacheJob {
    # The timer keeps ticking while a message box or a status redraw is
    # running, so ignore any tick that arrives while one is already active.
    if ($Script:CacheCompleting) { return }
    $Script:CacheCompleting = $true
    try {
        $job = $Script:CacheJob
        if ($null -eq $job) { $Script:CacheTimer.Stop(); return }

        $msg = [string]$job.State['Status']
        if ($msg) {
            $StatusText.Text    = $msg
            $LoadingDetail.Text = $msg
        }
        if (-not $job.Handle.IsCompleted) { return }

        $Script:CacheTimer.Stop()
        $Script:CacheJob = $null

        $err = $job.State['Error']
        try { [void]$job.PS.EndInvoke($job.Handle) }
        catch { if (-not $err) { $err = $_.Exception.Message } }

        # Copy the rows into this runspace as plain objects BEFORE the
        # background runspace is closed, so nothing on screen depends on it.
        $cache = New-Object System.Collections.Generic.List[object]
        $intuneCount = 0
        $extraCount  = 0
        if (-not $err) {
            # Walk the rows with a plain .NET enumerator (no @() or +), then
            # rebuild each one as an object owned by this window.
            $rowsData = $job.State['Rows']
            if ($null -ne $rowsData) {
                $enum = $rowsData.GetEnumerator()
                while ($enum.MoveNext()) {
                    $h = $enum.Current
                    if ($null -eq $h) { continue }
                    $keys = New-Object string[] $h.Count
                    $h.Keys.CopyTo($keys, 0)
                    $row = [ordered]@{}
                    for ($i = 0; $i -lt $keys.Length; $i++) { $row[$keys[$i]] = [string]$h[$keys[$i]] }
                    $cache.Add([pscustomobject]$row)
                    if ($row['Source'] -eq 'Intune') { $intuneCount++ } else { $extraCount++ }
                }
            }
        }

        try { $job.PS.Dispose(); $job.Runspace.Dispose() } catch { }

        Set-Busy $false

        if ($err) {
            Set-Status "Cache refresh failed: $err"
            Show-ToolkitMessageBox -Title 'Cache refresh failed' -Button 'OK' -Icon 'Error' -Message ([string]$err) | Out-Null
            return
        }

        $Script:DeviceCache   = $cache.ToArray()
        $Script:CacheLoadedAt = Get-Date
        Set-Status ("Device cache refreshed - $($Script:DeviceCache.Count) device(s): " +
                    "$intuneCount from Intune, $extraCount Autopilot-only " +
                    "(not yet enrolled), at $($Script:CacheLoadedAt.ToString('HH:mm:ss')).$([string]$job.State['ApNote'])")

        if ($job.OnComplete -and $Script:DeviceCache.Count -gt 0) { & $job.OnComplete }
    }
    catch {
        # Never let a failure here close the toolkit - log it and report it.
        $Script:CacheTimer.Stop()
        $Script:CacheJob = $null
        try { Set-Busy $false } catch { }
        $logFile = Write-ToolkitErrorLog -Context 'Finishing device load' -ErrorRecord $_
        $StatusText.Text = "Device load failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Device load error' -Button 'OK' -Icon 'Error' -Message (
            "Device data could not finish loading:`n`n$($_.Exception.Message)`n`nDetails were saved to:`n$logFile") | Out-Null
    }
    finally { $Script:CacheCompleting = $false }
}

$Script:CacheCompleting = $false
$Script:CacheTimer = New-Object System.Windows.Threading.DispatcherTimer
$Script:CacheTimer.Interval = [TimeSpan]::FromMilliseconds(200)
$Script:CacheTimer.Add_Tick({ Complete-DeviceCacheJob })

# ---------------------------------------------------------------------------
# Group actions bind to the Entra ID device record, so they work for an
# Autopilot-only device - but only once that record exists. Anything that
# needs an Intune managedDevice record is blocked for a non-enrolled device.
# ---------------------------------------------------------------------------
function Test-DeviceHasEntraRecord {
    param($Device, [string]$Action)

    $id = [string](Get-AutopilotProperty $Device 'EntraDeviceId')
    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq '00000000-0000-0000-0000-000000000000') {
        Show-ToolkitMessageBox -Title $Action -Button 'OK' -Icon 'Warning' -Message (
            "$($Device.DeviceName) has no Entra ID device record yet, so its group membership cannot be read or changed.`n`nAn Autopilot device gets its Entra record when it is imported. If this device was just imported, refresh the device cache and try again.") | Out-Null
        Set-Status "$Action - $($Device.DeviceName) has no Entra ID device record."
        return $false
    }
    return $true
}

# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------
function Invoke-DeviceSearch {
    if (-not (Test-Connected)) { return }

    $field = switch ($CmbSearchField.SelectedIndex) {
        0 { 'DeviceName' }
        1 { 'SerialNumber' }
        2 { 'User' }
        3 { 'ManagementName' }
        default { 'DeviceName' }
    }

    # The status is compared EXACTLY (case-insensitively) against the whole
    # Management name, because the Asset Status module writes the status
    # verbatim and nothing else.
    $byStatus = ($field -eq 'ManagementName')
    $status   = ''

    if ($byStatus) {
        $status = [string]$CmbAssetStatus.SelectedItem
        if ([string]::IsNullOrWhiteSpace($status)) {
            Set-Status 'Pick an asset status to filter on.'
            return
        }
    }
    elseif ([string]::IsNullOrWhiteSpace($TxtSearch.Text)) {
        Set-Status 'Enter a search term.'
        return
    }

    $term = $TxtSearch.Text.Trim()

    if ($Script:DeviceCache.Count -eq 0) {
        # Load the cache in the background, then run this search when it is ready.
        Update-DeviceCache -OnComplete { Invoke-DeviceSearch }
        return
    }

    $exact = [bool]$ChkExact.IsChecked

    Set-Busy $true 'Searching devices...'
    Set-Status 'Searching devices...'

    try {
        $hits = $Script:DeviceCache | Where-Object {
            if ($byStatus) {
                # Autopilot-only rows carry no Management name, so they drop out.
                $v = [string]$_.ManagementName
                return ((-not [string]::IsNullOrWhiteSpace($v)) -and ($v.Trim() -eq $status))
            }

            $v = $_.$field
            if ([string]::IsNullOrEmpty($v)) { $false }
            elseif ($exact)                  { $v -eq $term }
            else                             { $v -like "*$term*" }
        } | Sort-Object DeviceName

        $Script:Results.Clear()
        foreach ($h in $hits) { $Script:Results.Add($h) }

        $ResultCount.Text = "$($Script:Results.Count) device(s) found"
        $what = if ($byStatus) { $status } else { "'$term'" }
        Set-Status "Search complete - $($Script:Results.Count) result(s) for $what in $field."
    }
    catch {
        Set-Status "Search failed: $($_.Exception.Message)"
    }
    finally { Set-Busy $false }
}

# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# Sign-in prompt placement
#
# Connect-MgGraph shows its sign-in box (Windows Web Account Manager) over
# the PowerShell CONSOLE window that launched the toolkit - not over the
# toolkit window itself. So if the console is on one display and the toolkit
# is on another, the sign-in box opens on the console's display.
#
# Just before sign-in, the console is moved to the centre of WHICHEVER
# display the toolkit window is currently on, so the sign-in box opens next
# to the toolkit. The console is put back where it was afterwards. If the
# console is already on the same display as the toolkit it is left alone.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyConsoleWindow' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class EndpointguyConsoleWindow
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
}

# SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE - move only, never resize,
# re-stack or steal focus.
$Script:SwpMoveOnly = [uint32]0x15

function Get-ToolkitScreen {
    # The display that holds most of the toolkit window. Falls back to the
    # primary display if the toolkit window handle is not available.
    try {
        $toolkitHwnd = (New-Object System.Windows.Interop.WindowInteropHelper $Window).Handle
        if ($toolkitHwnd -ne [IntPtr]::Zero) {
            return [System.Windows.Forms.Screen]::FromHandle($toolkitHwnd)
        }
    }
    catch { }
    return [System.Windows.Forms.Screen]::PrimaryScreen
}

function Move-ConsoleToToolkitScreen {
    # Returns the console's original position so it can be put back, or
    # $null when there is nothing to move. Never throws - a failure here
    # must not stop the sign-in itself.
    try {
        $hWnd = [EndpointguyConsoleWindow]::GetConsoleWindow()
        if ($hWnd -eq [IntPtr]::Zero) { return $null }

        $rect = New-Object 'EndpointguyConsoleWindow+RECT'
        if (-not [EndpointguyConsoleWindow]::GetWindowRect($hWnd, [ref]$rect)) { return $null }

        # A minimised window reports about -32000; leave it alone.
        if ($rect.Left -le -32000) { return $null }

        $target = Get-ToolkitScreen

        # Console already on the toolkit's display - nothing to do.
        $consoleScreen = [System.Windows.Forms.Screen]::FromHandle($hWnd)
        if ($consoleScreen.DeviceName -eq $target.DeviceName) { return $null }

        $width  = $rect.Right  - $rect.Left
        $height = $rect.Bottom - $rect.Top
        $area   = $target.WorkingArea
        $x = $area.Left + [Math]::Max(0, [int](($area.Width  - $width)  / 2))
        $y = $area.Top  + [Math]::Max(0, [int](($area.Height - $height) / 2))

        [EndpointguyConsoleWindow]::SetWindowPos($hWnd, [IntPtr]::Zero, $x, $y, 0, 0, $Script:SwpMoveOnly) | Out-Null
        return [pscustomobject]@{ Handle = $hWnd; Left = $rect.Left; Top = $rect.Top }
    }
    catch { return $null }
}

function Restore-ConsolePosition {
    param($Saved)
    if ($null -eq $Saved) { return }
    try {
        [EndpointguyConsoleWindow]::SetWindowPos($Saved.Handle, [IntPtr]::Zero,
            $Saved.Left, $Saved.Top, 0, 0, $Script:SwpMoveOnly) | Out-Null
    }
    catch { }
}

function Connect-ToolkitGraph {
    param([switch]$LoadDeviceData)

    Set-Busy $true 'Waiting for Microsoft Graph sign-in...'
    $consoleSpot = $null
    $connected = $false
    try {
        Set-Status 'Opening Microsoft Graph sign-in prompt...'

        # Bring the sign-in box up on the same display as the toolkit.
        $consoleSpot = Move-ConsoleToToolkitScreen

        $params = @{ Scopes = $Script:GraphScopes; NoWelcome = $true }
        if ($TenantId) { $params['TenantId'] = $TenantId }
        Connect-MgGraph @params
        Update-ConnectionUi
        $ctx = Get-GraphContextSafe
        $connected = [bool]($ctx -and $ctx.Account)

        if ($connected) {
            Set-Status "Connected to tenant $($ctx.TenantId) as $($ctx.Account)."
        }
    }
    catch {
        Update-ConnectionUi
        Set-Status "Connection failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Connection failed' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
    finally {
        Restore-ConsolePosition $consoleSpot
        # Keep the overlay up when the device load follows straight on.
        if (-not ($connected -and $LoadDeviceData)) { Set-Busy $false }
    }

    if ($connected -and $LoadDeviceData) {
        Set-Status 'Signed in. Loading device data from Intune and Autopilot...'
        Update-DeviceCache
    }
    return $connected
}


$BtnSearch.Add_Click({ Invoke-DeviceSearch })

# ---------------------------------------------------------------------------
# Management Name search mode.
#
# The status list is owned by the Asset Status module: when that module is
# enabled it has already been dot-sourced and has set $Script:AstStatuses
# from ModuleConfig.psd1, so the drop-down here and the drop-down that writes
# the label cannot drift apart. When the module is switched off we fall back
# to the same seven built-in labels, because devices in the tenant may still be
# carrying them from an earlier build.
# ---------------------------------------------------------------------------
$Script:SearchStatuses = @('Assigned','In-Stock','Retired','Recycled','Stolen','Legalhold','Lost')

if ((Get-Variable -Name AstStatuses -Scope Script -ErrorAction SilentlyContinue) -and
    $null -ne $Script:AstStatuses) {

    $fromModule = @($Script:AstStatuses) |
                  ForEach-Object { [string]$_ } |
                  Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
                  ForEach-Object { $_.Trim() }

    if ($fromModule.Count -gt 0) { $Script:SearchStatuses = @($fromModule) }
}

# ---------------------------------------------------------------------------
# ADVANCED FILTER
#
# Builds a list of rules (column + operator + value) that are ANDed together.
# Rule rows are generated at runtime rather than declared in the XAML, so the
# number of criteria is open-ended.
#
# Columns marked with a fixed value list render a drop-down instead of a text
# box, so Management Name can only ever be filtered on a real Asset Status
# label - no free typing, no typos, no silently empty result sets.
# ---------------------------------------------------------------------------
$Script:FilterColumns = [ordered]@{
    "Device Name"     = "DeviceName"
    "Serial Number"   = "SerialNumber"
    "Management Name" = "ManagementName"
    "User"            = "User"
    "OS"              = "OS"
    "OS Version"      = "OSVersion"
    "Compliance"      = "Compliance"
    "Model"           = "Model"
    "Manufacturer"    = "Manufacturer"
    "Category"        = "Category"
    "Last Sync"       = "LastSync"
    "Source"          = "Source"
}

# Operators available to every column.
$Script:FilterOperators = @(
    "contains","does not contain","starts with","ends with",
    "equals","does not equal","is blank","is not blank"
)

# Columns whose values come from a fixed list. Anything not named here gets a
# free-text box. Compliance is a Graph enumeration, so it is safe to
# constrain too.
function Get-FilterValueList {
    param([string]$Property)

    switch ($Property) {
        "ManagementName" { return @($Script:SearchStatuses) }
        "Compliance"     { return @("compliant","noncompliant","conflict","error","inGracePeriod","unknown") }
        "Source"         { return @("Intune","Autopilot") }
        default          { return $null }
    }
}

# Operators that take no value - the value control is hidden for these.
function Test-FilterOperatorNeedsValue {
    param([string]$Operator)
    return ($Operator -ne "is blank" -and $Operator -ne "is not blank")
}

# Creates one rule row and appends it to RulesPanel.
function Add-FilterRule {
    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = "Horizontal"
    $row.Margin      = "0,0,0,8"

    $cmbCol = New-Object System.Windows.Controls.ComboBox
    $cmbCol.Width = 170
    foreach ($k in $Script:FilterColumns.Keys) { [void]$cmbCol.Items.Add($k) }
    # SelectedIndex is set after the handlers are wired, below.

    $cmbOp = New-Object System.Windows.Controls.ComboBox
    $cmbOp.Width  = 150
    $cmbOp.Margin = "8,0,0,0"
    foreach ($o in $Script:FilterOperators) { [void]$cmbOp.Items.Add($o) }
    # SelectedIndex is set after the handlers are wired, below.
    # Both a text box and a drop-down are created; exactly one is shown,
    # decided by the column. Swapping visibility is far simpler than
    # rebuilding the row every time the column changes.
    $txtVal = New-Object System.Windows.Controls.TextBox
    $txtVal.Width  = 240
    $txtVal.Height = 32
    $txtVal.Margin = "8,0,0,0"

    $cmbVal = New-Object System.Windows.Controls.ComboBox
    $cmbVal.Width      = 240
    $cmbVal.Margin     = "8,0,0,0"
    $cmbVal.Visibility = [System.Windows.Visibility]::Collapsed

    $btnDel = New-Object System.Windows.Controls.Button
    $btnDel.Content = [char]0x2212   # minus sign - clearer than X for "remove rule"
    $btnDel.Width   = 30
    $btnDel.Height  = 30
    $btnDel.Margin  = "8,0,0,0"
    $btnDel.ToolTip = "Remove this rule"
    $style = $Window.TryFindResource("NeutralButton")
    if ($null -ne $style) { $btnDel.Style = $style }
    # Overrides go after the style: NeutralButton sets Padding/FontSize
    # that would otherwise crush the glyph on a 30x30 button.
    $btnDel.Padding    = "0"
    $btnDel.FontSize   = 18
    $btnDel.FontWeight = "Bold"
    $red = $Window.TryFindResource("ConnBadBrush")
    if ($null -eq $red) {
        $red = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,107,107))
    }
    $btnDel.Foreground = $red

    # Show the drop-down for fixed-list columns, the text box otherwise, and
    # hide both when the operator does not take a value.
    # Local copy: GetNewClosure() captures locals, but $Script: inside the
    # closure resolves to the closure module, where the script vars are null.
    $colMap = $Script:FilterColumns

    $syncRow = {
        $colName = [string]$cmbCol.SelectedItem
        $prop    = [string]$colMap[$colName]
        $list    = Get-FilterValueList -Property $prop
        $needs   = Test-FilterOperatorNeedsValue ([string]$cmbOp.SelectedItem)

        if (-not $needs) {
            $txtVal.Visibility = [System.Windows.Visibility]::Collapsed
            $cmbVal.Visibility = [System.Windows.Visibility]::Collapsed
            return
        }

        if ($null -ne $list) {
            # Repopulate only when the list actually changed, so the current
            # selection survives an operator change.
            $existing = @($cmbVal.Items | ForEach-Object { [string]$_ })
            if (($existing -join "|") -ne (@($list) -join "|")) {
                $cmbVal.Items.Clear()
                foreach ($v in $list) { [void]$cmbVal.Items.Add($v) }
                $cmbVal.SelectedIndex = 0
            }
            $cmbVal.Visibility = [System.Windows.Visibility]::Visible
            $txtVal.Visibility = [System.Windows.Visibility]::Collapsed
        }
        else {
            $cmbVal.Visibility = [System.Windows.Visibility]::Collapsed
            $txtVal.Visibility = [System.Windows.Visibility]::Visible
        }
    }.GetNewClosure()

    # Select defaults BEFORE wiring SelectionChanged, otherwise the handler
    # fires against a row that is not fully built yet.
    $cmbCol.SelectedIndex = 0
    $cmbOp.SelectedIndex  = 0

    # GetNewClosure() pins each handler to the controls of the row it was
    # built for; without it every row would end up driving the last row added.
    $cmbCol.Add_SelectionChanged($syncRow)
    $cmbOp.Add_SelectionChanged($syncRow)
    # Remove via the row's own parent - no captured panel reference needed.
    $btnDel.Add_Click({ $p = $row.Parent; if ($null -ne $p) { [void]$p.Children.Remove($row) } }.GetNewClosure())

    [void]$row.Children.Add($cmbCol)
    [void]$row.Children.Add($cmbOp)
    [void]$row.Children.Add($txtVal)
    [void]$row.Children.Add($cmbVal)
    [void]$row.Children.Add($btnDel)

    # Tag carries the controls, so Get-FilterRules can read a rule back
    # without relying on child positions.
    $row.Tag = [pscustomobject]@{
        Column   = $cmbCol
        Operator = $cmbOp
        TextBox  = $txtVal
        ComboBox = $cmbVal
    }

    [void]$RulesPanel.Children.Add($row)
    & $syncRow
}

# Reads every rule row back into plain objects.
function Get-FilterRules {
    $rules = New-Object System.Collections.Generic.List[object]

    foreach ($row in $RulesPanel.Children) {
        $t = $row.Tag
        if ($null -eq $t) { continue }

        $colName = [string]$t.Column.SelectedItem
        if ([string]::IsNullOrWhiteSpace($colName)) { continue }

        $op    = [string]$t.Operator.SelectedItem
        $needs = Test-FilterOperatorNeedsValue $op

        $val = ""
        if ($needs) {
            $val = if ($t.ComboBox.Visibility -eq [System.Windows.Visibility]::Visible) {
                       [string]$t.ComboBox.SelectedItem
                   } else {
                       [string]$t.TextBox.Text
                   }
            # A rule with no value yet is skipped rather than matching nothing.
            if ([string]::IsNullOrWhiteSpace($val)) { continue }
        }

        $rules.Add([pscustomobject]@{
            ColumnName = $colName
            Property   = [string]$Script:FilterColumns[$colName]
            Operator   = $op
            Value      = $val.Trim()
        })
    }

    return ,$rules
}

# Evaluates one rule against one device. Comparisons are case-insensitive,
# matching the rest of the toolkit.
function Test-FilterRule {
    param($Device,$Rule)

    $v = [string]$Device.($Rule.Property)
    if ($null -eq $v) { $v = "" }
    $v = $v.Trim()
    $t = $Rule.Value

    # Escape *, ? and [ so a value typed by the user is matched literally -
    # -like would otherwise treat them as wildcards. Only the operators that
    # use -like need this; -eq is unaffected.
    $tLike = [System.Management.Automation.WildcardPattern]::Escape($t)

    switch ($Rule.Operator) {
        "is blank"         { return [string]::IsNullOrWhiteSpace($v) }
        "is not blank"     { return (-not [string]::IsNullOrWhiteSpace($v)) }
        "contains"         { return ($v -like "*$tLike*") }
        "does not contain" { return (-not ($v -like "*$tLike*")) }
        "starts with"      { return ($v -like "$tLike*") }
        "ends with"        { return ($v -like "*$tLike") }
        "equals"           { return ($v -eq $t) }
        "does not equal"   { return ($v -ne $t) }
        default            { return $true }
    }
}

# Runs every rule against the cache. All rules must match (AND).
function Invoke-AdvancedFilter {
    if (-not (Test-Connected)) { return }

    $rules = Get-FilterRules
    if ($rules.Count -eq 0) {
        Set-Status "Add at least one rule with a value, then select Apply Filter."
        return
    }

    if ($Script:DeviceCache.Count -eq 0) {
        # Load the cache in the background, then run the filter when it is ready.
        Update-DeviceCache -OnComplete { Invoke-AdvancedFilter }
        return
    }

    Set-Busy $true 'Applying advanced filter...'
    try {
        $hits = $Script:DeviceCache | Where-Object {
            $device = $_
            $keep   = $true
            foreach ($rule in $rules) {
                if (-not (Test-FilterRule -Device $device -Rule $rule)) { $keep = $false; break }
            }
            $keep
        } | Sort-Object DeviceName

        $Script:Results.Clear()
        foreach ($h in $hits) { $Script:Results.Add($h) }

        $ResultCount.Text = "$($Script:Results.Count) device(s) found"

        $summary = ($rules | ForEach-Object {
            if (Test-FilterOperatorNeedsValue $_.Operator) { "$($_.ColumnName) $($_.Operator) ''$($_.Value)''" }
            else { "$($_.ColumnName) $($_.Operator)" }
        }) -join " AND "

        Set-Status "Filter complete - $($Script:Results.Count) result(s) for: $summary"
    }
    catch {
        Set-Status "Filter failed: $($_.Exception.Message)"
    }
    finally { Set-Busy $false }
}

# Simple reference window for the Tools area.
function Show-ToolsInfo {
    $help = New-Object System.Windows.Window
    $help.Title = 'Toolkit Help'
    $help.Width = 540
    $help.Height = 790
    $help.ResizeMode = [System.Windows.ResizeMode]::NoResize
    $help.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $help.Owner = $Window
    $help.Background = $Window.TryFindResource('WindowBrush')

    $panel = New-Object System.Windows.Controls.StackPanel
    $panel.Margin = '24'

    $heading = New-Object System.Windows.Controls.TextBlock
    $heading.Text = 'Toolkit Tools'
    $heading.FontSize = 22
    $heading.FontWeight = [System.Windows.FontWeights]::SemiBold
    $heading.Foreground = $Window.TryFindResource('TitleTextBrush')
    $heading.Margin = '0,0,0,4'
    [void]$panel.Children.Add($heading)

    $intro = New-Object System.Windows.Controls.TextBlock
    $intro.Text = 'A quick guide to the available actions.'
    $intro.FontSize = 12.5
    $intro.Foreground = $Window.TryFindResource('SubtleTextBrush')
    $intro.Margin = '0,0,0,18'
    [void]$panel.Children.Add($intro)

    $sections = @(
        @{ Name = 'DEVICE ACTIONS'; Items = @(
            @('Copy Device Groups', 'Copies assigned security-group memberships from the selected device to another device.'),
            @('Remove Device Groups', 'Removes the selected device from assigned security groups after confirmation.'),
            @('Update Asset Status', 'Changes the selected device Management name to an approved lifecycle status.')
        ) },
        @{ Name = 'BULK ACTIONS'; Items = @(
            @('Bulk Add to Group', 'Adds devices listed in a CSV file to one assigned security group.'),
            @('Bulk Update Asset Status', 'Applies one approved lifecycle status to devices listed in a CSV file.')
        ) },
        @{ Name = 'REPORTING'; Items = @(
            @('App Dependency Check', 'Shows which Win32 apps depend on a selected app. This report is read-only.'),
            @('Device Prep Check  (NEW)', 'Checks a serial number is ready for Autopilot: hash imported, enabled in Entra ID, and no old Intune record.')
        ) }


    )

    foreach ($section in $sections) {
        $sectionName = New-Object System.Windows.Controls.TextBlock
        $sectionName.Text = $section.Name
        $sectionName.FontSize = 10.5
        $sectionName.FontWeight = [System.Windows.FontWeights]::Bold
        $sectionName.Foreground = $Window.TryFindResource('InfoLabelBrush')
        $sectionName.Margin = '0,4,0,9'
        [void]$panel.Children.Add($sectionName)

        foreach ($item in $section.Items) {
            $name = New-Object System.Windows.Controls.TextBlock
            $name.Text = $item[0]
            $name.FontSize = 13.5
            $name.FontWeight = [System.Windows.FontWeights]::SemiBold
            $name.Foreground = $Window.TryFindResource('AccentTextBrush')
            $name.Margin = '0,0,0,2'
            [void]$panel.Children.Add($name)

            $description = New-Object System.Windows.Controls.TextBlock
            $description.Text = $item[1]
            $description.FontSize = 12.5
            $description.TextWrapping = [System.Windows.TextWrapping]::Wrap
            $description.Foreground = $Window.TryFindResource('TextBrush')
            $description.Margin = '0,0,0,12'
            [void]$panel.Children.Add($description)
        }
    }

    $close = New-Object System.Windows.Controls.Button
    $close.Content = 'Close'
    $close.Width = 90
    $close.Height = 32
    $close.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $close.Background = $Window.TryFindResource('AccentBrush')
    $close.Foreground = $Window.TryFindResource('OnAccentTextBrush')
    $close.BorderBrush = $Window.TryFindResource('AccentBrush')
    $close.Add_Click({ $help.Close() })
    [void]$panel.Children.Add($close)

    $border = New-Object System.Windows.Controls.Border
    $border.Background = $Window.TryFindResource('CardBrush')
    $border.BorderBrush = $Window.TryFindResource('CardBorderBrush')
    $border.BorderThickness = '1'
    $border.Margin = '12'
    $border.CornerRadius = '8'
    $border.Child = $panel
    $help.Content = $border
    [void]$help.ShowDialog()
}

$BtnToolsInfo.Add_Click({ Show-ToolsInfo })

# Open one Tools menu at a time.
$BtnDeviceMenu.Add_Click({
    $BulkMenuPopup.IsOpen = $false
    $ReportMenuPopup.IsOpen = $false
    $DeviceMenuPopup.IsOpen = -not $DeviceMenuPopup.IsOpen
})
$BtnBulkMenu.Add_Click({
    $DeviceMenuPopup.IsOpen = $false
    $ReportMenuPopup.IsOpen = $false
    $BulkMenuPopup.IsOpen = -not $BulkMenuPopup.IsOpen
})
$BtnReportMenu.Add_Click({
    $DeviceMenuPopup.IsOpen = $false
    $BulkMenuPopup.IsOpen = $false
    $ReportMenuPopup.IsOpen = -not $ReportMenuPopup.IsOpen
})
$BtnCopyGroups.Add_Click({ $DeviceMenuPopup.IsOpen = $false })
$BtnRemoveGroups.Add_Click({ $DeviceMenuPopup.IsOpen = $false })
$BtnAssetStatus.Add_Click({ $DeviceMenuPopup.IsOpen = $false })
$BtnBulkAddGroup.Add_Click({ $BulkMenuPopup.IsOpen = $false })
$BtnBulkAssetStatus.Add_Click({ $BulkMenuPopup.IsOpen = $false })
$BtnAppDependency.Add_Click({ $ReportMenuPopup.IsOpen = $false })

# Show/hide the panel. The simple search bar stays usable either way, so
# nothing is taken away by opening this.
$BtnAdvanced.Add_Click({
  try {
    $open = ($AdvancedPanel.Visibility -eq [System.Windows.Visibility]::Visible)
    $AdvancedPanel.Visibility = if ($open) { [System.Windows.Visibility]::Collapsed }
                               else        { [System.Windows.Visibility]::Visible }
    # Start with one blank rule so the panel is never empty on first open.
    if (-not $open -and $RulesPanel.Children.Count -eq 0) { Add-FilterRule }
  }
  catch { Show-ToolkitMessageBox -Title 'Advanced filter' -Button 'OK' -Icon 'Error' -Message ("Advanced filter failed to open:`n`n$($_.Exception.Message)") | Out-Null }
})

$BtnAddRule.Add_Click({ Add-FilterRule })
$BtnApplyFilter.Add_Click({ Invoke-AdvancedFilter })
$BtnClearFilter.Add_Click({
    $RulesPanel.Children.Clear()
    Add-FilterRule
    Set-Status "Filter rules cleared."
})

foreach ($s in $Script:SearchStatuses) { [void]$CmbAssetStatus.Items.Add($s) }
if ($CmbAssetStatus.Items.Count -gt 0) { $CmbAssetStatus.SelectedIndex = 0 }

# Swap the free-text box for the status drop-down when Management Name is
# picked. The status is matched exactly, so a free-text term would have
# nothing left to narrow - the text box and Exact match only are both
# hidden/disabled in this mode rather than left on screen doing nothing.
function Update-SearchFieldUi {
    $byStatus = ($CmbSearchField.SelectedIndex -eq 3)

    $CmbAssetStatus.Visibility = if ($byStatus) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $TxtSearch.Visibility      = if ($byStatus) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    $ChkExact.IsEnabled        = -not $byStatus
}

$CmbSearchField.Add_SelectionChanged({ Update-SearchFieldUi })
Update-SearchFieldUi

$CmbAssetStatus.Add_SelectionChanged({ if ((Get-GraphContextSafe) -and $Script:DeviceCache.Count -gt 0) { Invoke-DeviceSearch } })

# WPF event scriptblocks populate $args[0] (sender) and $args[1] (event args).
# $_ is NOT set here, which would trip Set-StrictMode.
$TxtSearch.Add_KeyDown({
    if ($args[1].Key -eq 'Return') { Invoke-DeviceSearch }
})

$BtnRefreshCache.Add_Click({ Update-DeviceCache })

$GridDevices.Add_SelectionChanged({
    $d = $GridDevices.SelectedItem
    $hasDevice = ($null -ne $d)
    $Script:SelectedDevice = if ($hasDevice) { $d } else { $null }

    $deviceActionReady = $Script:ModulesLoaded['CopyDeviceGroups'] -or
                         $Script:ModulesLoaded['RemoveDeviceGroups'] -or
                         $Script:ModulesLoaded['AssetStatus']
    $BtnDeviceMenu.IsEnabled = $hasDevice -and $deviceActionReady

    if ($hasDevice) {
        $sourceNote = if ($d.Source -eq 'Autopilot') { "`nAutopilot: $($d.EnrollmentState)" } else { '' }
        $SelectedDeviceText.Text = "$($d.DeviceName)`nSerial: $($d.SerialNumber)$sourceNote"
        $SelectedDeviceText.Foreground = $Window.TryFindResource('TextBrush')
        Set-Status "Selected $($d.DeviceName)  |  Serial: $($d.SerialNumber)  |  Source: $($d.Source)"
    }
    else {
        $SelectedDeviceText.Text = 'Select a device first'
        $SelectedDeviceText.Foreground = $Window.TryFindResource('SubtleTextBrush')
        $DeviceMenuPopup.IsOpen = $false
    }
})

$BtnExportCsv.Add_Click({
    if ($Script:Results.Count -eq 0) { Set-Status 'Nothing to export.'; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "Devices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:Results | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Set-Status "Exported $($Script:Results.Count) row(s) to $($dlg.FileName)"
    }
})

# --- Stubbed actions -------------------------------------------------------
$BtnCopyGroups.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'CopyDeviceGroups')) {
        Set-Status 'Copy Device Groups is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        Show-ToolkitMessageBox -Title 'Copy Device Groups' -Button 'OK' -Icon 'Warning' -Message 'Select a device in the search results first.' | Out-Null
        return
    }

    if (-not (Test-DeviceHasEntraRecord -Device $Script:SelectedDevice -Action 'Copy Device Groups')) { return }

    if (-not (Get-Command -Name Show-CopyDeviceGroupsWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module not available' -Button 'OK' -Icon 'Error' -Message (
            "CopyDeviceGroups.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\CopyDeviceGroups')") | Out-Null
        Set-Status 'Copy Device Groups module is not available.'
        return
    }

    Set-Status "Opening Copy Device Groups for $($Script:SelectedDevice.DeviceName)..."
    try {
        Show-CopyDeviceGroupsWindow -SourceDevice $Script:SelectedDevice `
                                    -DeviceCache  $Script:DeviceCache `
                                    -Owner        $Window
        Set-Status "Copy Device Groups closed - source device was $($Script:SelectedDevice.DeviceName)."
    }
    catch {
        Set-Status "Copy Device Groups failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Copy Device Groups' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})
$BtnRemoveGroups.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'RemoveDeviceGroups')) {
        Set-Status 'Remove Device Groups is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        Show-ToolkitMessageBox -Title 'Remove Device Groups' -Button 'OK' -Icon 'Warning' -Message 'Select a device in the search results first.' | Out-Null
        return
    }

    if (-not (Test-DeviceHasEntraRecord -Device $Script:SelectedDevice -Action 'Remove Device Groups')) { return }

    if (-not (Get-Command -Name Show-RemoveDeviceGroupsWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module not available' -Button 'OK' -Icon 'Error' -Message (
            "RemoveDeviceGroups.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\RemoveDeviceGroups')") | Out-Null
        Set-Status 'Remove Device Groups module is not available.'
        return
    }

    Set-Status "Opening Remove Device Groups for $($Script:SelectedDevice.DeviceName)..."
    try {
        Show-RemoveDeviceGroupsWindow -Device $Script:SelectedDevice `
                                      -Owner  $Window
        Set-Status "Remove Device Groups closed - device was $($Script:SelectedDevice.DeviceName)."
    }
    catch {
        Set-Status "Remove Device Groups failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Remove Device Groups' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})
$BtnBulkAddGroup.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'BulkAddToGroup')) {
        Set-Status 'Bulk Add to Group is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if (-not (Get-Command -Name Show-BulkAddToGroupWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module missing' -Button 'OK' -Icon 'Warning' -Message (
            "BulkAddToGroup.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\BulkAddToGroup')") | Out-Null
        return
    }

    Set-Status 'Opening Bulk Add to Group...'
    try {
        Show-BulkAddToGroupWindow -DeviceCache $Script:DeviceCache `
                                  -Owner       $Window
        Set-Status 'Bulk Add to Group closed.'
    }
    catch {
        Set-Status "Bulk Add to Group failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Bulk Add to Group' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})
$BtnAppDependency.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'AppDependencyCheck')) {
        Set-Status 'App Dependency Check is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if (-not (Get-Command -Name Show-AppDependencyCheckWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module missing' -Button 'OK' -Icon 'Warning' -Message (
            "AppDependencyCheck.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\AppDependencyCheck')") | Out-Null
        return
    }

    Set-Status 'Opening App Dependency Check...'
    try {
        Show-AppDependencyCheckWindow -Owner $Window
        Set-Status 'App Dependency Check closed.'
    }
    catch {
        Set-Status "App Dependency Check failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'App Dependency Check' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})
$BtnDevicePrepCheck.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'DevicePrepCheck')) {
        Set-Status 'Device Prep Check is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if (-not (Get-Command -Name Show-DevicePrepCheckWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module missing' -Button 'OK' -Icon 'Warning' -Message (
            "DevicePrepCheck.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\DevicePrepCheck')") | Out-Null
        return
    }

    # Works with or without a selected device: the serial number is only a
    # starting point and can be changed in the window.
    $serial = ''
    if ($null -ne $Script:SelectedDevice -and $Script:SelectedDevice.PSObject.Properties['SerialNumber']) {
        $serial = [string]$Script:SelectedDevice.SerialNumber
    }

    Set-Status 'Opening Device Prep Check...'
    try {
        Show-DevicePrepCheckWindow -SerialNumber $serial -Owner $Window
        Set-Status 'Device Prep Check closed.'
    }
    catch {
        Set-Status "Device Prep Check failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Device Prep Check' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})
$BtnAssetStatus.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'AssetStatus')) {
        Set-Status 'Asset Status is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        Show-ToolkitMessageBox -Title 'Asset Status' -Button 'OK' -Icon 'Warning' -Message 'Select a device in the search results first.' | Out-Null
        return
    }

    # The Management name lives on the Intune managedDevice record, so this
    # action needs an IntuneId - not the Entra record the group actions use.
    # An Autopilot device that has not enrolled yet has no managedDevice.
    if ([string]::IsNullOrWhiteSpace([string]$Script:SelectedDevice.IntuneId)) {
        Show-ToolkitMessageBox -Title 'Asset Status' -Button 'OK' -Icon 'Warning' -Message (
            "$($Script:SelectedDevice.DeviceName) is not enrolled in Intune yet, so it has no Management name to set.`n`nA device imported into Autopilot only gets a Management name once it enrols.") | Out-Null
        Set-Status 'Asset Status needs a device that has enrolled in Intune.'
        return
    }

    if (-not (Get-Command -Name Show-AssetStatusWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module not available' -Button 'OK' -Icon 'Error' -Message (
            "AssetStatus.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\AssetStatus')") | Out-Null
        Set-Status 'Asset Status module is not available.'
        return
    }

    Set-Status "Opening Asset Status for $($Script:SelectedDevice.DeviceName)..."
    try {
        Show-AssetStatusWindow -Device $Script:SelectedDevice `
                               -Owner  $Window
        Set-Status "Asset Status closed - device was $($Script:SelectedDevice.DeviceName)."
    }
    catch {
        Set-Status "Asset Status failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Asset Status' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})

$BtnBulkAssetStatus.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'BulkAssetStatus')) {
        Set-Status 'Bulk Update Asset Status is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    # Unlike the single-device Asset Status action, this one works from a
    # CSV, so it does NOT need a device selected in the search results.
    if (-not (Get-Command -Name Show-BulkAssetStatusWindow -ErrorAction SilentlyContinue)) {
        Show-ToolkitMessageBox -Title 'Module not available' -Button 'OK' -Icon 'Error' -Message (
            "BulkAssetStatus.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\BulkAssetStatus')") | Out-Null
        Set-Status 'Bulk Update Asset Status module is not available.'
        return
    }

    Set-Status 'Opening Bulk Update Asset Status...'
    try {
        Show-BulkAssetStatusWindow -DeviceCache $Script:DeviceCache `
                                   -Owner       $Window
        Set-Status 'Bulk Update Asset Status closed.'
    }
    catch {
        Set-Status "Bulk Update Asset Status failed: $($_.Exception.Message)"
        Show-ToolkitMessageBox -Title 'Bulk Update Asset Status' -Button 'OK' -Icon 'Error' -Message ($_.Exception.Message) | Out-Null
    }
})

# ---------------------------------------------------------------------------
# Header web link - open endpointguy.com in the default browser.
# WPF will not navigate on its own; the RequestNavigate event must be handled.
# ---------------------------------------------------------------------------
$HeaderTitle = $Window.FindName('HeaderTitle')
$LnkSite = $Window.FindName('LnkSite')
if ($HeaderTitle) { $HeaderTitle.Text = $Script:ToolkitName }
if ($Window) { $Window.Title = $Script:ToolkitName }
if ($LnkSite) {
    try {
        $LnkSite.Inlines.Clear()
        [void]$LnkSite.Inlines.Add($Script:WebsiteText)
        $LnkSite.NavigateUri = New-Object System.Uri($Script:WebsiteUrl)
        $LnkSite.ToolTip = $Script:WebsiteUrl
    }
    catch { $LnkSite.Visibility = [System.Windows.Visibility]::Collapsed }
    $LnkSite.Add_RequestNavigate({
        param($eventSender, $e)
        try { Start-Process $e.Uri.AbsoluteUri }
        catch { Set-Status "Could not open $($e.Uri.AbsoluteUri)" }
        $e.Handled = $true
    })
}

# ---------------------------------------------------------------------------
# Source column link - clicking 'Intune' in the results grid opens that
# device in the Intune admin center in the default browser. Autopilot-only
# rows are not enrolled yet, so they have no Intune page and show plain text.
# ---------------------------------------------------------------------------
$Script:IntuneDeviceUrl = 'https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/{0}'

$GridDevices.AddHandler([System.Windows.Documents.Hyperlink]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param($linkSource, $linkArgs)
    $link = $linkArgs.OriginalSource
    if ($link -isnot [System.Windows.Documents.Hyperlink]) { return }
    $linkArgs.Handled = $true

    $device = $link.DataContext
    $deviceName = [string]$(if ($null -ne $device -and $device.PSObject.Properties['DeviceName']) { $device.DeviceName })
    $intuneId = [string]$(if ($null -ne $device -and $device.PSObject.Properties['IntuneId']) { $device.IntuneId })
    $parsedId = [guid]::Empty
    if (-not [guid]::TryParse($intuneId, [ref]$parsedId)) {
        Set-Status "No Intune device ID is available for $deviceName."
        return
    }

    $url = $Script:IntuneDeviceUrl -f $parsedId.ToString()
    try {
        Start-Process $url
        Set-Status "Opened $deviceName in the Intune admin center."
    }
    catch {
        Set-Status "Could not open the browser for $deviceName."
        Show-ToolkitMessageBox -Title 'Open in Intune' -Button 'OK' -Icon 'Error' -Message (
            "The browser could not be opened.`n`n$($_.Exception.Message)`n`nLink:`n$url") | Out-Null
    }
})

# ---------------------------------------------------------------------------
# Notes column - the Notes button on each Intune row reads that device's
# notes from Graph and opens them in an editable toolkit dialog. Save writes
# the edited notes back to the device in Intune. Notes are fetched when the
# button is clicked (Graph only returns them for a single device), so the
# device cache stays fast.
# ---------------------------------------------------------------------------
function Show-DeviceNotesEditor {
    <#
        Opens one device's notes in an editable toolkit-styled dialog.
        Returns the edited text when Save is pressed, or $null on Cancel.
    #>
    param(
        [string]$DeviceName,
        [AllowEmptyString()][string]$Notes = ''
    )

    # Reuse the toolkit message box layout, then make the text area editable.
    [xml]$doc = $Script:TkDialogXaml
    $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $doc))
    $dlg.Title = "Notes - $DeviceName"
    $dlg.Tag = 'Cancel'
    $dlg.FindName('TitleText').Text = "Notes - $DeviceName"

    $box = $dlg.FindName('MessageText')
    $box.IsReadOnly = $false
    $box.AcceptsReturn = $true
    $box.Cursor = [System.Windows.Input.Cursors]::IBeam
    $box.Width = 520
    $box.MinHeight = 200
    $box.MaxHeight = 420
    $box.Padding = New-Object System.Windows.Thickness 8
    $box.BorderThickness = New-Object System.Windows.Thickness 1
    $box.BorderBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#4D4D4D'))
    $box.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#1F1F1F'))
    $box.CaretBrush = [System.Windows.Media.Brushes]::White
    $box.Text = $Notes

    # Save (blue) and Cancel. Esc also cancels.
    $panel = $dlg.FindName('ButtonPanel')
    $style = $dlg.Resources['TkButton']
    foreach ($name in @('Save', 'Cancel')) {
        $b = New-Object System.Windows.Controls.Button
        $b.Style = $style
        $b.Content = $name
        $b.Tag = $name
        if ($name -eq 'Save') {
            $b.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#2B7CD3'))
            $b.BorderBrush = $b.Background
            $b.Foreground = [System.Windows.Media.Brushes]::White
        }
        else {
            $b.IsCancel = $true
        }
        $b.Add_Click({
            param($clickSource, $clickArgs)
            $owning = [System.Windows.Window]::GetWindow($clickSource)
            $owning.Tag = [string]$clickSource.Tag
            $owning.Close()
        })
        [void]$panel.Children.Add($b)
    }

    # Put the cursor at the end of the notes, ready to type.
    $dlg.Add_ContentRendered({
        param($renderSource, $renderArgs)
        $t = $renderSource.FindName('MessageText')
        [void]$t.Focus()
        $t.CaretIndex = $t.Text.Length
    })

    # Drag the box by its header.
    $dlg.FindName('HeaderBar').Add_MouseLeftButtonDown({
        param($dragSource, $dragArgs)
        try { [System.Windows.Window]::GetWindow($dragSource).DragMove() } catch { }
    })

    $ownerHandle = [IntPtr]::Zero
    try { $ownerHandle = Get-ToolkitDialogOwner } catch { $ownerHandle = [IntPtr]::Zero }
    if ($ownerHandle -ne [IntPtr]::Zero) {
        (New-Object System.Windows.Interop.WindowInteropHelper $dlg).Owner = $ownerHandle
        $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    }
    else {
        $dlg.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
        $dlg.Topmost = $true
    }

    [void]$dlg.ShowDialog()
    if ([string]$dlg.Tag -ne 'Save') { return $null }
    return [string]$box.Text
}

$GridDevices.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]{
    param($notesSource, $notesArgs)
    $btn = $notesArgs.OriginalSource
    if ($btn -isnot [System.Windows.Controls.Button] -or [string]$btn.Tag -ne 'DeviceNotes') { return }
    $notesArgs.Handled = $true

    $device = $btn.DataContext
    $deviceName = [string]$(if ($null -ne $device -and $device.PSObject.Properties['DeviceName']) { $device.DeviceName })
    $intuneId = [string]$(if ($null -ne $device -and $device.PSObject.Properties['IntuneId']) { $device.IntuneId })
    $parsedId = [guid]::Empty
    if (-not [guid]::TryParse($intuneId, [ref]$parsedId)) {
        Show-ToolkitMessageBox -Title "Notes - $deviceName" -Button 'OK' -Icon 'Information' -Message (
            'This device is not enrolled in Intune yet, so it has no notes.') | Out-Null
        return
    }
    $deviceUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($parsedId.ToString())"

    $notes = ''
    $failure = $null
    $btn.IsEnabled = $false
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    Set-Status "Loading notes for $deviceName..."
    try {
        # Notes come from the beta endpoint; v1.0 often returns them empty.
        $resp = Invoke-MgGraphRequest -Method GET -Uri "$deviceUri`?`$select=id,notes" -OutputType PSObject
        if ((Test-GraphProperty $resp 'notes') -and $null -ne $resp.notes) { $notes = [string]$resp.notes }
        Set-Status "Loaded notes for $deviceName."
    }
    catch {
        $failure = $_.Exception.Message
        Set-Status "Could not load notes for $deviceName."
    }
    finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        $btn.IsEnabled = $true
    }

    if ($failure) {
        Show-ToolkitMessageBox -Title "Notes - $deviceName" -Button 'OK' -Icon 'Error' -Message (
            "The notes for this device could not be loaded from Graph.`n`n$failure") | Out-Null
        return
    }

    # Edit, then save back to Intune. If the save fails the editor reopens
    # with the typed text, so nothing is lost.
    $draft = $notes
    while ($true) {
        $edited = Show-DeviceNotesEditor -DeviceName $deviceName -Notes $draft
        if ($null -eq $edited) {
            Set-Status "Notes for $deviceName were not changed."
            return
        }
        if ($edited.TrimEnd() -ceq $notes.TrimEnd()) {
            Set-Status "No changes to the notes for $deviceName."
            return
        }

        $saveFailure = $null
        $btn.IsEnabled = $false
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        Set-Status "Saving notes for $deviceName..."
        try {
            $body = @{ notes = $edited } | ConvertTo-Json -Compress
            Invoke-MgGraphRequest -Method PATCH -Uri $deviceUri -Body $body -ContentType 'application/json' | Out-Null
            Set-Status "Saved notes for $deviceName."
        }
        catch {
            $saveFailure = $_.Exception.Message
            Write-ToolkitErrorLog -Context "Save notes for $deviceName" -ErrorRecord $_ | Out-Null
            Set-Status "Could not save notes for $deviceName."
        }
        finally {
            [System.Windows.Input.Mouse]::OverrideCursor = $null
            $btn.IsEnabled = $true
        }

        if (-not $saveFailure) {
            return
        }

        $retry = Show-ToolkitMessageBox -Title "Notes - $deviceName" -Button 'YesNo' -Icon 'Error' -Default 'Yes' -Message (
            "The notes could not be saved to Intune.`n`n$saveFailure`n`nGo back to the notes and try again?")
        if ($retry -ne [System.Windows.MessageBoxResult]::Yes) { return }
        $draft = $edited
    }
})

$LogoPath = Join-Path $Script:ModuleRoot $Script:LogoFile
if ($LogoImage -and (Test-Path $LogoPath)) {
    try {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.UriSource   = New-Object System.Uri($LogoPath)
        # Load the bytes now so the file is not left locked on disk.
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $LogoImage.Source = $bmp
    }
    # A missing or corrupt logo must never stop the toolkit loading.
    catch { $LogoImage.Visibility = [System.Windows.Visibility]::Collapsed }
}
else { if ($LogoImage) { $LogoImage.Visibility = [System.Windows.Visibility]::Collapsed } }


# ---------------------------------------------------------------------------
# Clock
# ---------------------------------------------------------------------------
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({ $ClockText.Text = (Get-Date).ToString('HH:mm:ss') })
$timer.Start()

$Window.Add_Closed({
    $timer.Stop()
    $Script:CacheTimer.Stop()

    # Stop a device load that is still running when the window is closed.
    if ($Script:CacheJob) {
        try { $Script:CacheJob.PS.Stop() } catch { }
        try { $Script:CacheJob.PS.Dispose(); $Script:CacheJob.Runspace.Dispose() } catch { }
        $Script:CacheJob = $null
    }
})

# ---------------------------------------------------------------------------
# Apply the module registry to the UI.
#
# A module that is switched off has its button removed from the layout
# (Collapsed, not just greyed out) so the panel closes up and the operator
# is never shown an action this build cannot perform. A module that is
# switched ON but whose file is missing stays visible but disabled, because
# that is a deployment fault worth seeing rather than hiding.
# ---------------------------------------------------------------------------
function Update-ModuleUi {
    foreach ($entry in $Script:ModuleRegistry) {
        $btn = $Window.FindName($entry.Button)
        if ($null -eq $btn) { continue }

        if (-not $entry.Enabled) {
            $btn.Visibility = [System.Windows.Visibility]::Collapsed
            continue
        }

        $btn.Visibility = [System.Windows.Visibility]::Visible

        if (-not $Script:ModulesLoaded[$entry.Key]) {
            $btn.IsEnabled = $false
            $btn.ToolTip   = "$($entry.Key) is enabled but its file was not found at '$($entry.Path)'."
        }
    }

    # Hide a menu when every module in that category is switched off.
    $deviceAvailable = (Test-ModuleEnabled -Key 'CopyDeviceGroups') -or
                       (Test-ModuleEnabled -Key 'RemoveDeviceGroups') -or
                       (Test-ModuleEnabled -Key 'AssetStatus')
    $bulkAvailable = (Test-ModuleEnabled -Key 'BulkAddToGroup') -or
                     (Test-ModuleEnabled -Key 'BulkAssetStatus')
    $reportAvailable = Test-ModuleEnabled -Key 'AppDependencyCheck'
    $DeviceToolsSection.Visibility = if ($deviceAvailable) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $BulkToolsSection.Visibility = if ($bulkAvailable) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $ReportToolsSection.Visibility = if ($reportAvailable) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $offCount = @($Script:ModuleRegistry | Where-Object { -not $_.Enabled }).Count
    if ($offCount -gt 0) {
        $names = ($Script:ModuleRegistry | Where-Object { -not $_.Enabled } |
                  ForEach-Object { $_.Key }) -join ', '
        Set-Status "Ready. $offCount module(s) switched off for this build: $names"
    }
}

Update-ModuleUi

# ---------------------------------------------------------------------------
# Recent changes - shown once per toolkit version for each Windows user.
# Bump ToolkitRelease whenever the list below changes for a new release.
# ---------------------------------------------------------------------------
$Script:ToolkitRelease = '2026.10.22.1'

function Show-RecentChangesIfNeeded {
    $stateRoot = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'EndpointGuy\IntuneToolkit'
    $stateFile = Join-Path $stateRoot 'recent-changes-version.txt'
    $seenVersion = ''
    try {
        if (Test-Path $stateFile) { $seenVersion = ([string](Get-Content -LiteralPath $stateFile -Raw -ErrorAction Stop)).Trim() }
    }
    catch { Write-Warning "Recent changes state could not be read: $($_.Exception.Message)" }
    if ($seenVersion -eq $Script:ToolkitRelease) { return }

    $changes = @(
        @{ Title = 'Updated Interface'; Text = 'A refreshed interface introduces the new Tools section, prominent loading indicators for immediate feedback, and several other visual improvements throughout the toolkit.' },
        @{ Title = 'Automatic Graph Sign-In and Device Loading'; Text = 'The toolkit now signs in to Microsoft Graph when it opens and immediately begins loading Intune and Autopilot device data, so it is ready to use sooner.' },
        @{ Title = 'Faster Access to Intune Details'; Text = 'Open a device directly in the Intune admin center by selecting Intune in the Source column. To view or update its notes, select the Notes button in the same result row.'; ShowNotes = $true },
        @{ Title = 'Device Prep Check'; Text = 'Check Autopilot readiness from a serial number, including the hardware hash, Entra ID status, existing Intune records, and group memberships.' }
    )

    $dialog = New-Object System.Windows.Window
    $dialog.Title = 'Recent Changes'
    $dialog.Width = 650
    $dialog.Height = 690
    $dialog.MinWidth = 560
    $dialog.MinHeight = 560
    $dialog.ResizeMode = [System.Windows.ResizeMode]::CanResize
    $dialog.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $dialog.Owner = $Window
    $dialog.ShowInTaskbar = $false
    $dialog.Background = $Window.TryFindResource('WindowBrush')

    $root = New-Object System.Windows.Controls.Grid
    $root.Margin = '18'
    [void]$root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = 'Auto' }))
    [void]$root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = '*' }))
    [void]$root.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{ Height = 'Auto' }))

    $header = New-Object System.Windows.Controls.Border
    $header.Background = $Window.TryFindResource('CardBrush')
    $header.BorderBrush = $Window.TryFindResource('CardBorderBrush')
    $header.BorderThickness = '1'
    $header.CornerRadius = '7'
    $header.Padding = '20,16'
    $headerPanel = New-Object System.Windows.Controls.StackPanel
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = 'Recent Changes'
    $title.FontSize = 24
    $title.FontWeight = [System.Windows.FontWeights]::Bold
    $title.Foreground = $Window.TryFindResource('TitleTextBrush')
    [void]$headerPanel.Children.Add($title)
    $release = New-Object System.Windows.Controls.TextBlock
    $release.Text = "Toolkit update $($Script:ToolkitRelease)"
    $release.FontSize = 12.5
    $release.Foreground = $Window.TryFindResource('AccentTextBrush')
    $release.Margin = '0,3,0,0'
    [void]$headerPanel.Children.Add($release)
    $intro = New-Object System.Windows.Controls.TextBlock
    $intro.Text = 'Here is a quick look at what changed.'
    $intro.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $intro.FontSize = 12.5
    $intro.Foreground = $Window.TryFindResource('SubtleTextBrush')
    $intro.Margin = '0,8,0,0'
    [void]$headerPanel.Children.Add($intro)
    $header.Child = $headerPanel
    [System.Windows.Controls.Grid]::SetRow($header, 0)
    [void]$root.Children.Add($header)

    $scroll = New-Object System.Windows.Controls.ScrollViewer
    $scroll.VerticalScrollBarVisibility = [System.Windows.Controls.ScrollBarVisibility]::Auto
    $scroll.Margin = '0,14,0,14'
    $list = New-Object System.Windows.Controls.StackPanel
    foreach ($change in $changes) {
        $card = New-Object System.Windows.Controls.Border
        $card.Background = $Window.TryFindResource('CardBrush')
        $card.BorderBrush = $Window.TryFindResource('CardBorderBrush')
        $card.BorderThickness = '1'
        $card.CornerRadius = '6'
        $card.Padding = '16,12'
        $card.Margin = '0,0,0,10'
        $content = New-Object System.Windows.Controls.StackPanel
        $changeTitle = New-Object System.Windows.Controls.TextBlock
        $changeTitle.Text = $change.Title
        $changeTitle.FontSize = 15
        $changeTitle.FontWeight = [System.Windows.FontWeights]::SemiBold
        $changeTitle.Foreground = $Window.TryFindResource('AccentTextBrush')
        [void]$content.Children.Add($changeTitle)
        $changeText = New-Object System.Windows.Controls.TextBlock
        $changeText.Text = $change.Text
        $changeText.FontSize = 12.5
        $changeText.LineHeight = 18
        $changeText.TextWrapping = [System.Windows.TextWrapping]::Wrap
        $changeText.Foreground = $Window.TryFindResource('TextBrush')
        $changeText.Margin = '0,4,0,0'
        [void]$content.Children.Add($changeText)

        if ($change.ContainsKey('ShowNotes') -and [bool]$change.ShowNotes) {
            $sampleRow = New-Object System.Windows.Controls.StackPanel
            $sampleRow.Orientation = [System.Windows.Controls.Orientation]::Horizontal
            $sampleRow.Margin = '0,10,0,0'
            $sampleLabel = New-Object System.Windows.Controls.TextBlock
            $sampleLabel.Text = 'Notes button:'
            $sampleLabel.FontSize = 12
            $sampleLabel.Foreground = $Window.TryFindResource('SubtleTextBrush')
            $sampleLabel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            $sampleLabel.Margin = '0,0,8,0'
            [void]$sampleRow.Children.Add($sampleLabel)
            $notesSample = New-Object System.Windows.Controls.Button
            $notesSample.Style = $Window.TryFindResource('NeutralButton')
            $notesSample.Padding = '0'
            $notesSample.Width = 30
            $notesSample.Height = 22
            $notesSample.IsHitTestVisible = $false
            $noteIcon = New-Object System.Windows.Shapes.Path
            $noteIcon.Width = 12
            $noteIcon.Height = 14
            $noteIcon.Stretch = [System.Windows.Media.Stretch]::None
            $noteIcon.StrokeThickness = 1.2
            $noteIcon.StrokeLineJoin = [System.Windows.Media.PenLineJoin]::Round
            $noteIcon.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
            $noteIcon.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
            $noteIcon.Stroke = $Window.TryFindResource('AccentTextBrush')
            $noteIcon.Data = [System.Windows.Media.Geometry]::Parse('M1,0.6 L7.6,0.6 L11,4 L11,13.4 L1,13.4 Z M7.6,0.6 L7.6,4 L11,4 M3.2,6.8 L8.8,6.8 M3.2,9 L8.8,9 M3.2,11.2 L6.6,11.2')
            $notesSample.Content = $noteIcon
            [void]$sampleRow.Children.Add($notesSample)
            [void]$content.Children.Add($sampleRow)
        }

        $card.Child = $content
        [void]$list.Children.Add($card)
    }
    $scroll.Content = $list
    [System.Windows.Controls.Grid]::SetRow($scroll, 1)
    [void]$root.Children.Add($scroll)

    $close = New-Object System.Windows.Controls.Button
    $close.Content = 'Continue'
    $close.Width = 110
    $close.Height = 36
    $close.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $close.Background = $Window.TryFindResource('AccentBrush')
    $close.Foreground = $Window.TryFindResource('OnAccentTextBrush')
    $close.BorderBrush = $Window.TryFindResource('AccentBrush')
    $close.FontWeight = [System.Windows.FontWeights]::SemiBold
    $close.Add_Click({ $dialog.Close() })
    [System.Windows.Controls.Grid]::SetRow($close, 2)
    [void]$root.Children.Add($close)

    $dialog.Content = $root
    [void]$dialog.ShowDialog()

    try {
        if (-not (Test-Path $stateRoot)) { New-Item -ItemType Directory -Path $stateRoot -Force | Out-Null }
        Set-Content -LiteralPath $stateFile -Value $Script:ToolkitRelease -Encoding ASCII -Force
    }
    catch { Write-Warning "Recent changes state could not be saved: $($_.Exception.Message)" }
}

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------
Update-ConnectionUi

# Show release notes first, then prompt for Graph sign-in and preload the
# Intune and Autopilot cache. InvokeAsync lets the first frame paint first.
$Script:StartupInitialized = $false
$Window.Add_ContentRendered({
    if ($Script:StartupInitialized) { return }
    $Script:StartupInitialized = $true
    $Window.Dispatcher.InvokeAsync([action]{
        Show-RecentChangesIfNeeded
        [void](Connect-ToolkitGraph -LoadDeviceData)
    }, [Windows.Threading.DispatcherPriority]::Background) | Out-Null
})

# Hide the PowerShell console window so only the toolkit appears on the
# taskbar. This only happens when the process was started just to run the
# toolkit (Start menu shortcut, Run-Toolkit.bat, powershell -File ...).
# Running .\Toolkit.ps1 from your own PowerShell prompt leaves that prompt
# alone. Use -ShowConsole to keep the console visible for troubleshooting.
$Script:ConsoleHwnd = [EndpointguyConsoleWindow]::GetConsoleWindow()
$ownLaunch = [bool](@([Environment]::GetCommandLineArgs()) -like "*$(Split-Path -Leaf $PSCommandPath)*")
if ($ownLaunch -and -not $ShowConsole -and $Script:ConsoleHwnd -ne [IntPtr]::Zero) {
    [EndpointguyConsoleWindow]::ShowWindow($Script:ConsoleHwnd, 0) | Out-Null   # SW_HIDE
}

try {
    $Window.ShowDialog() | Out-Null
}
catch {
    $logFile = Write-ToolkitErrorLog -Context 'Toolkit window' -ErrorRecord $_
    # Bring the console back so the error (and the .bat pause) can be seen.
    if ($Script:ConsoleHwnd -ne [IntPtr]::Zero) {
        [EndpointguyConsoleWindow]::ShowWindow($Script:ConsoleHwnd, 5) | Out-Null   # SW_SHOW
    }
    Show-ToolkitMessageBox -Title 'EndpointGuy Intune Toolkit' -Button 'OK' -Icon 'Error' -Message (
        "The toolkit closed because of an error:`n`n$($_.Exception.Message)`n`nDetails were saved to:`n$logFile") | Out-Null
    throw
}
