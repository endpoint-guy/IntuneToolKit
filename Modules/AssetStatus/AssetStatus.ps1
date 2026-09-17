<#
.SYNOPSIS
    Asset Status - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that sets the Intune Management name (managedDeviceName)
    of the selected managed device to one of six lifecycle statuses:

        In-Stock, Retired, Recycled, Stolen, Legalhold, Lost

      - The device is passed in from the Toolkit (the device selected in the
        search results) and is pre-filled in the window.
      - The current Management name is read back from Graph when the window
        opens, so the operator can see what is about to be replaced.
      - Picking a status shows the resulting Management name before anything
        is written.
      - A confirmation prompt naming the device, the old name and the new
        name is shown before the write, and it defaults to No.
      - The grid is re-read from Graph after a successful write, so what is
        on screen is what Intune actually holds.

    The Management name is a label only. Changing it does NOT retire, wipe,
    unenrol or otherwise act on the device, and it leaves the device name,
    serial number, group memberships and assignments untouched.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so AssetStatus.xaml
    is not required at runtime; embedded XAML is used by default
    when launched from Toolkit.ps1. Pass -XamlPath to load an external
    AssetStatus.xaml instead while iterating on the UI.

    Every function here is Ast-prefixed on purpose. This module,
    CopyDeviceGroups.ps1, RemoveDeviceGroups.ps1, BulkAddToGroup.ps1 and
    AppDependencyCheck.ps1 are dot-sourced into the SAME session by
    Toolkit.ps1, so a shared name would mean the last file loaded silently
    wins and could change the behaviour of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\AssetStatus\AssetStatus.ps1"
        Show-AssetStatusWindow -Device $Script:SelectedDevice `
                               -Owner  $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: DeviceManagementManagedDevices.ReadWrite.All
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external AssetStatus.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# The statuses this module is allowed to write.
#
# Kept in one place so the drop-down, the validation and the standalone test
# path cannot drift apart. The drop-down is built from this list at runtime,
# and anything passed that is not on the list is refused.
#
# Toolkit.ps1 sets $Script:AssetStatusValues from ModuleConfig.psd1 before
# dot-sourcing this file. When that is absent - or when this module is run
# standalone - the six built-in statuses below are used.
# ---------------------------------------------------------------------------
$AstDefaultStatuses = @('In-Stock','Retired','Recycled','Stolen','Legalhold','Lost')

$Script:AstStatuses = $AstDefaultStatuses

if ((Get-Variable -Name AssetStatusValues -Scope Script -ErrorAction SilentlyContinue) -and
    $null -ne $Script:AssetStatusValues) {

    $astConfigured = @($Script:AssetStatusValues) |
                     ForEach-Object { [string]$_ } |
                     Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
                     ForEach-Object { $_.Trim() }

    if ($astConfigured.Count -gt 0) {
        $Script:AstStatuses = @($astConfigured)
    }
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$AstXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Asset Status"
        Height="760" Width="1100"
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
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Padding="22,14">
            <StackPanel>
                <TextBlock Text="Asset Status" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Sets the Intune Management name of the selected device to a lifecycle status. The Management name is a label only - it does not retire, wipe or unenrol the device."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Device -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <StackPanel>
                <TextBlock Text="Device" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                <Border CornerRadius="4"
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
            </StackPanel>
        </Border>

        <!-- Status -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <StackPanel>
                <TextBlock Text="Management name" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>

                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="18"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0">
                        <TextBlock Text="CURRENT" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <TextBox x:Name="TxtCurrentName" IsReadOnly="True" Text=""/>
                    </StackPanel>

                    <StackPanel Grid.Column="2">
                        <TextBlock Text="NEW STATUS" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <!-- Items are added at runtime from $Script:AstStatuses. -->
                        <ComboBox x:Name="CmbStatus" Height="32" FontSize="13"/>
                    </StackPanel>
                </Grid>

                <Border CornerRadius="4" Margin="0,14,0,0"
                        Background="{StaticResource InfoPanelBrush}"
                        BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                        Padding="12,10">
                    <StackPanel>
                        <TextBlock Text="RESULTING MANAGEMENT NAME" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <TextBlock x:Name="TxtPreview" Text="Pick a status to see the result."
                                   FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap"/>
                        <TextBlock Text="The current Management name is replaced entirely. Intune keeps the device name, serial number and all assignments unchanged."
                                   FontSize="11.5" Foreground="{StaticResource SubtleTextBrush}"
                                   TextWrapping="Wrap" Margin="0,6,0,0"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>

        <!-- Diagnostics console -->
        <Border x:Name="ConsolePanel" Grid.Row="3" Visibility="Collapsed"
                Style="{StaticResource Card}" Padding="18,14" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,10">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Diagnostics console" Style="{StaticResource CardHeader}"/>
                        <TextBlock Text="Live Graph calls and errors" Margin="12,0,0,0"
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

                <TextBox x:Name="TxtLog" Grid.Row="1" MinHeight="160"
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
                <TextBlock x:Name="StatusText" Text="Ready."
                           FontSize="12.5" VerticalAlignment="Center" TextWrapping="NoWrap"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnReload" Style="{StaticResource NeutralButton}"
                            Content="Reload" MinWidth="110" Height="38" Margin="0,0,12,0"/>
                    <Button x:Name="BtnApply" Style="{StaticResource PrimaryButton}"
                            Content="Update Management name" MinWidth="220" Height="38" Margin="0,0,12,0"/>
                    <Button x:Name="BtnClose" Style="{StaticResource NeutralButton}"
                            Content="Close" MinWidth="110" Height="38"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# ---------------------------------------------------------------------------
# Diagnostics log
#
# The window installs a sink here so the helper functions - which live
# outside its scope - can write to the on-screen console. When no window is
# open the sink is null and the lines go to Write-Verbose instead, so the
# module still behaves when dot-sourced or run headless.
# ---------------------------------------------------------------------------
$Script:AstLogSink    = $null
$Script:AstLogVerbose = $false

function Write-AstLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:AstLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:AstLogSink) {
        try { & $Script:AstLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-AstLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-AstGraphError $ErrorRecord
    Write-AstLog "$Context failed: $detail" 'ERROR'
    if ($Script:AstLogVerbose -and $ErrorRecord) {
        $ex = Get-AstProperty $ErrorRecord 'Exception'
        if ($ex) { Write-AstLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-AstProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-AstLog "At: $((Get-AstProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-AstProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-AstProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-AstProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-AstBodyError {
    # Pulls code/message out of a Graph error body, whether it arrives as a
    # parsed object, a hashtable, or a raw JSON string.
    param($Body)

    if ($null -eq $Body) { return $null }

    $obj = $Body
    if ($Body -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Body)) { return $null }
        try { $obj = $Body | ConvertFrom-Json } catch { return $Body }
    }
    if ($obj -is [System.Collections.IDictionary]) {
        $obj = [pscustomobject]$obj
    }

    if (Test-AstProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-AstProperty $err 'code' ''
        $msg  = Get-AstProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-AstGraphError {
    <#
        Invoke-MgGraphRequest throws a generic exception whose Message is only
        'Response status code does not indicate success: BadRequest (Bad
        Request).' The useful text is in the response body, which lands in a
        different place depending on the module version - hence every probe
        below. Whatever is found is returned as a readable string.
    #>
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) { return 'Unknown error.' }

    # 1. ErrorDetails.Message - usually the raw JSON body.
    if (Test-AstProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-AstProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-AstBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-AstProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-AstProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-AstProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-AstBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-AstProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-AstProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-AstProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

# ---------------------------------------------------------------------------
# Graph
# ---------------------------------------------------------------------------
function Get-AstManagedDeviceId {
    <#
        The Management name lives on the Intune managedDevice object, so this
        module needs a managedDevice id - not the Entra ID device object the
        group modules bind to.

        The Toolkit cache already carries it as IntuneId for an enrolled
        device. An Autopilot-only device has no managedDevice record at all,
        which is why that case is refused with a clear message rather than a
        Graph 404.
    #>
    param([Parameter(Mandatory)]$Device)

    $name = Get-AstProperty $Device 'DeviceName' '(unknown device)'

    $id = [string](Get-AstProperty $Device 'IntuneId' '')
    if ([string]::IsNullOrWhiteSpace($id)) { $id = [string](Get-AstProperty $Device 'Id' '') }

    if ([string]::IsNullOrWhiteSpace($id)) {
        $src = [string](Get-AstProperty $Device 'Source' '')
        if ($src -eq 'Autopilot') {
            throw "$name has been imported into Autopilot but has not enrolled in Intune yet, so it has no Management name to set. Once the device enrols, refresh the device cache and try again."
        }
        throw "$name has no Intune managed device id, so its Management name cannot be read or written."
    }

    return $id
}

function Get-AstManagedDevice {
    # Reads the live managedDevice record so the window shows what Intune
    # actually holds rather than whatever the cache was populated with.
    param([Parameter(Mandatory)][string]$ManagedDeviceId)

    $select = 'id,deviceName,managedDeviceName,serialNumber,userPrincipalName,operatingSystem,managementState'
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$ManagedDeviceId`?`$select=$select"

    Write-AstLog "GET $uri" 'GRAPH'
    $sw   = [System.Diagnostics.Stopwatch]::StartNew()
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    $sw.Stop()
    Write-AstLog "  -> responded in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

    if ($null -eq $resp) {
        throw "No Intune managed device was found for id $ManagedDeviceId."
    }

    [pscustomobject]@{
        Id                = Get-AstProperty $resp 'id' $ManagedDeviceId
        DeviceName        = Get-AstProperty $resp 'deviceName' ''
        ManagedDeviceName = Get-AstProperty $resp 'managedDeviceName' ''
        SerialNumber      = Get-AstProperty $resp 'serialNumber' ''
        User              = Get-AstProperty $resp 'userPrincipalName' ''
        OS                = Get-AstProperty $resp 'operatingSystem' ''
        ManagementState   = Get-AstProperty $resp 'managementState' ''
    }
}

function Set-AstManagementName {
    <#
        Writes managedDeviceName on the managedDevice.

        PATCH /deviceManagement/managedDevices/<id> with a managedDeviceName
        body is the supported way to change the Management name; it returns
        200 with the updated object, or 204 on some tenants. Nothing else
        about the device is sent, so no other property can be disturbed.
    #>
    param(
        [Parameter(Mandatory)][string]$ManagedDeviceId,
        [Parameter(Mandatory)][string]$NewName
    )

    if ([string]::IsNullOrWhiteSpace($ManagedDeviceId)) {
        throw 'The device has no Intune managed device id, so its Management name cannot be set.'
    }
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        throw 'The new Management name is empty.'
    }

    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$ManagedDeviceId"
    $body = @{ managedDeviceName = $NewName } | ConvertTo-Json -Compress

    Write-AstLog "PATCH $uri" 'GRAPH'
    Write-AstLog "  -> body $body" 'DEBUG'

    # Preferred path: ask Graph not to throw, so the error body is returned
    # intact instead of being flattened into a useless exception message.
    $canSkipThrow = $false
    try {
        $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction Stop
        $canSkipThrow = $cmd.Parameters.ContainsKey('SkipHttpErrorCheck') -and
                        $cmd.Parameters.ContainsKey('StatusCodeVariable')
    }
    catch { $canSkipThrow = $false }

    Write-AstLog "  -> using $(if ($canSkipThrow) { '-SkipHttpErrorCheck (full error body)' } else { 'try/catch (legacy Graph module)' })" 'DEBUG'

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body `
                    -ContentType 'application/json' -OutputType PSObject `
                    -SkipHttpErrorCheck -StatusCodeVariable 'code'

        Write-AstLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) {
            Write-AstLog '  -> Management name updated' 'OK'
            return
        }

        $detail = Get-AstBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        Write-AstLog "  -> Graph said: $detail" 'DEBUG'
        throw "HTTP $code - $detail"
    }

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType 'application/json' | Out-Null
        Write-AstLog '  -> Management name updated' 'OK'
    }
    catch {
        throw (Get-AstGraphError $_)
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-AssetStatusWindow {
    <#
    .SYNOPSIS
        Opens the Asset Status window for the supplied device.
    .PARAMETER Device
        The device object selected in the Toolkit results grid. Needs at least
        DeviceName and IntuneId.
    .PARAMETER Owner
        The Toolkit window, so this one centres on it and stays on top.
    .PARAMETER ShowConsole
        Opens with the diagnostics console visible. Leave it off for normal
        use; turn it on when troubleshooting a Graph error. F12 toggles it at
        any time.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Device,
        $Owner = $null,
        [string]$XamlPath,

        # Reveals the diagnostics console. Off by default - the log is still
        # captured in memory, so ticking it open mid-session shows the
        # history rather than starting blank.
        [switch]$ShowConsole
    )

    # --- load XAML ---------------------------------------------------------
    $xamlText = $AstXamlString
    if ($XamlPath -and (Test-Path $XamlPath)) {
        $xamlText = Get-Content -Path $XamlPath -Raw
    }

    try {
        [xml]$xamlDoc = $xamlText
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        $win    = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "The Asset Status XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtDeviceName','TxtDeviceDetail','TxtCurrentName','CmbStatus',
                     'TxtPreview','StatusText','BtnApply','BtnReload','BtnClose',
                     'TxtLog','ChkVerbose','ChkAutoScroll','BtnCopyLog','BtnSaveLog','BtnClearLog',
                     'ConsolePanel')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the Asset Status XAML.",
                'XAML load error','OK','Error') | Out-Null
            return
        }
        $ui[$n] = $ctl
    }

    if ($Owner) { $win.Owner = $Owner }

    # --- state -------------------------------------------------------------
    # A hashtable, not plain variables: assigning inside an event scriptblock
    # would otherwise create a local copy and the value would be lost.
    $state = @{
        ManagedId = $null
        Current   = $null
        Busy      = $false
    }

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # --- diagnostics console ----------------------------------------------
    # The sink is what Write-AstLog calls. Appending to a TextBox is cheap
    # and keeps the whole session scrollable; the box is trimmed so a long
    # run cannot grow it without bound.
    $appendLog = {
        param([string]$Line)
        $ui.TxtLog.AppendText($Line + [Environment]::NewLine)

        if ($ui.TxtLog.LineCount -gt 2000) {
            $ui.TxtLog.Text = $ui.TxtLog.Text.Substring($ui.TxtLog.GetCharacterIndexFromLineIndex(500))
        }

        if ([bool]$ui.ChkAutoScroll.IsChecked) { $ui.TxtLog.ScrollToEnd() }
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # Point the module-scope sink at this window.
    $Script:AstLogSink    = $appendLog
    $Script:AstLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
        $win.Height = 940
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnApply, $ui.BtnReload)) { $b.IsEnabled = -not $Busy }
        $ui.CmbStatus.IsEnabled = -not $Busy
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # --- selected status ---------------------------------------------------
    $getStatus = {
        # The ComboBox holds ComboBoxItem objects, so the Content is what
        # carries the status text.
        $item = $ui.CmbStatus.SelectedItem
        if ($null -eq $item) { return '' }
        if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
        return [string]$item
    }

    $refreshPreview = {
        $status = & $getStatus
        if ([string]::IsNullOrWhiteSpace($status)) {
            $ui.TxtPreview.Text = 'Pick a status to see the result.'
            return
        }
        $ui.TxtPreview.Text = $status
    }

    # --- load --------------------------------------------------------------
    $loadDevice = {
        & $setBusy $true
        try {
            & $setStatus "Reading the Intune record for $(Get-AstProperty $Device 'DeviceName' 'the device')..."
            $state.ManagedId = Get-AstManagedDeviceId -Device $Device
            $state.Current   = Get-AstManagedDevice -ManagedDeviceId $state.ManagedId

            $ui.TxtDeviceName.Text = if ($state.Current.DeviceName) { $state.Current.DeviceName }
                                     else { Get-AstProperty $Device 'DeviceName' '(unknown device)' }

            $ui.TxtDeviceDetail.Text = "Serial: $(if ($state.Current.SerialNumber) { $state.Current.SerialNumber } else { 'n/a' })`n" +
                                       "User: $(if ($state.Current.User) { $state.Current.User } else { 'n/a' })`n" +
                                       "Intune device id: $($state.Current.Id)"

            $ui.TxtCurrentName.Text = if ([string]::IsNullOrWhiteSpace($state.Current.ManagedDeviceName))
                                      { '(not set)' } else { $state.Current.ManagedDeviceName }

            & $refreshPreview
            & $setStatus "Loaded. Current Management name: $($ui.TxtCurrentName.Text)"
            Write-AstLog "Current Management name: $($ui.TxtCurrentName.Text)" 'INFO'
        }
        catch {
            $detail = Write-AstLogException $_ 'Loading the device'
            & $setStatus "Could not load the device: $detail"
            $ui.TxtCurrentName.Text = ''
            [System.Windows.MessageBox]::Show($detail,'Asset Status','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- apply -------------------------------------------------------------
    $doApply = {
        $status = & $getStatus
        if ([string]::IsNullOrWhiteSpace($status)) {
            [System.Windows.MessageBox]::Show('Pick a status first.',
                'Asset Status','OK','Warning') | Out-Null
            return
        }

        # Belt and braces: the drop-down is fixed, but a status that is not on
        # the approved list must never reach Graph.
        if ($Script:AstStatuses -notcontains $status) {
            [System.Windows.MessageBox]::Show("'$status' is not one of the allowed statuses.",
                'Asset Status','OK','Error') | Out-Null
            return
        }

        if ([string]::IsNullOrWhiteSpace([string]$state.ManagedId)) {
            [System.Windows.MessageBox]::Show('The device has not loaded yet, so nothing can be written. Select Reload and try again.',
                'Asset Status','OK','Warning') | Out-Null
            return
        }

        $deviceName = $ui.TxtDeviceName.Text
        $oldName    = $ui.TxtCurrentName.Text
        $newName    = $status

        if ($oldName -ceq $newName) {
            [System.Windows.MessageBox]::Show("The Management name of $deviceName is already '$newName'. Nothing to do.",
                'Asset Status','OK','Information') | Out-Null
            & $setStatus "Management name is already $newName - nothing was written."
            return
        }

        # The old name is overwritten and Intune keeps no history of it, so
        # the prompt shows both names and defaults to No.
        $confirm = [System.Windows.MessageBox]::Show(
            "Set the Management name of $deviceName to '$newName'?`n`n" +
            "  From: $oldName`n" +
            "  To:   $newName`n`n" +
            "This replaces the Management name in Intune. The device name, serial number, group memberships and assignments are NOT changed, and the device is not retired, wiped or unenrolled.`n`n" +
            "Intune keeps no history of the previous Management name, so it cannot be restored from here.",
            'Confirm Management name change','YesNo','Warning',[System.Windows.MessageBoxResult]::No)

        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-AstLog 'Update cancelled at the confirmation prompt.' 'INFO'
            & $setStatus 'Cancelled - nothing was written.'
            return
        }

        & $setBusy $true
        try {
            & $setStatus "Setting the Management name to $newName..."
            Write-AstLog "Setting Management name: '$oldName' -> '$newName'" 'INFO'

            Set-AstManagementName -ManagedDeviceId $state.ManagedId -NewName $newName

            $summary = "Management name of $deviceName set to $newName."
            & $setStatus $summary
            Write-AstLog $summary 'OK'

            [System.Windows.MessageBox]::Show(
                "$summary`n`nFrom: $oldName`nTo:   $newName",
                'Asset Status','OK','Information') | Out-Null
        }
        catch {
            $detail = Write-AstLogException $_ 'Updating the Management name'
            & $setStatus "Update failed: $detail"
            [System.Windows.MessageBox]::Show($detail,'Asset Status','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }

        # Re-read from Graph so the window shows what Intune actually holds.
        & $loadDevice
    }

    # --- events ------------------------------------------------------------
    $ui.CmbStatus.Add_SelectionChanged({ & $refreshPreview })

    $ui.ChkVerbose.Add_Checked({
        $Script:AstLogVerbose = $true
        Write-AstLog 'Verbose logging on - Graph URIs, bodies and timings will be shown.' 'INFO'
    })
    $ui.ChkVerbose.Add_Unchecked({
        $Script:AstLogVerbose = $false
        Write-AstLog 'Verbose logging off.' 'INFO'
    })

    $ui.BtnClearLog.Add_Click({
        $ui.TxtLog.Clear()
        Write-AstLog 'Log cleared.' 'INFO'
    })

    $ui.BtnCopyLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
        & $setStatus 'Diagnostics log copied to the clipboard.'
    })

    $ui.BtnSaveLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $file  = Join-Path ([Environment]::GetFolderPath('Desktop')) "AssetStatus-$stamp.log"
        try {
            $ui.TxtLog.Text | Set-Content -Path $file -Encoding UTF8
            & $setStatus "Log saved to $file"
            Write-AstLog "Log saved to $file" 'OK'
        }
        catch { Write-AstLog "Could not save the log: $($_.Exception.Message)" 'ERROR' }
    })

    $ui.BtnReload.Add_Click({ & $loadDevice })
    $ui.BtnApply.Add_Click({  & $doApply })
    $ui.BtnClose.Add_Click({  $win.Close() })

    # Drop the sink so a closed window is never written to.
    $win.Add_Closed({ $Script:AstLogSink = $null })

    # F12 toggles the console without reopening the window.
    $win.Add_KeyDown({
        if ($args[1].Key -ne 'F12') { return }
        if ($ui.ConsolePanel.Visibility -eq [System.Windows.Visibility]::Visible) {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Collapsed
            $win.Height = 760
        }
        else {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
            $win.Height = 940
        }
    })

    # --- go ----------------------------------------------------------------
    # Build the drop-down from the approved list, so a status added in
    # ModuleConfig.psd1 appears here without touching the XAML.
    $ui.CmbStatus.Items.Clear()
    foreach ($s in $Script:AstStatuses) {
        $ui.CmbStatus.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{ Content = $s })) | Out-Null
    }
    Write-AstLog "Statuses offered: $($Script:AstStatuses -join ', ')" 'DEBUG'

    $ui.TxtDeviceName.Text = Get-AstProperty $Device 'DeviceName' '(unknown device)'
    Write-AstLog 'Asset Status opened.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-AstLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-AstLog "Device: $(Get-AstProperty $Device 'DeviceName' '(unknown)')" 'INFO'
    $win.Add_ContentRendered({ & $loadDevice })
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# prompt for a device by name so the module can be tested on its own.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.ReadWrite.All' -NoWelcome

    $name = Read-Host 'Device name'
    $esc  = $name.Replace("'", "''")
    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$esc'&`$select=id,deviceName,managedDeviceName,serialNumber,userPrincipalName"

    Write-AstLog "GET $uri" 'GRAPH'
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    $hits = @()
    if (Test-AstProperty $resp 'value') { $hits = @($resp.value) }

    if ($hits.Count -eq 0) { Write-Warning "No managed device named '$name' was found."; return }

    $dev = [pscustomobject]@{
        DeviceName   = Get-AstProperty $hits[0] 'deviceName' ''
        SerialNumber = Get-AstProperty $hits[0] 'serialNumber' ''
        User         = Get-AstProperty $hits[0] 'userPrincipalName' ''
        IntuneId     = Get-AstProperty $hits[0] 'id' ''
        Source       = 'Intune'
    }

    Show-AssetStatusWindow -Device $dev -XamlPath $XamlPath -ShowConsole
}
