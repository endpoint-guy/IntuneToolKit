<#
.SYNOPSIS
    Remove Device Groups - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that removes the selected managed device from the Entra ID
    security groups it currently belongs to.

      - The device is passed in from the Toolkit (the device selected in the
        search results) and is pre-filled in the window.
      - Every group the device belongs to is listed, with each group marked
        eligible or ineligible. Only assigned (static) security groups are
        eligible; dynamic groups are shown but cannot be selected because
        their membership is evaluated by rule and cannot be written to.
      - Every eligible group is ticked by default, so the common case -
        strip this device out of all of its groups - is one click.
      - A confirmation prompt naming the device and the group count is shown
        before anything is removed, and it defaults to No.
      - Each group is removed one at a time and its outcome is written back
        into the Result column, with a summary at the end.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so RemoveDeviceGroups.xaml
    is not required at runtime and this file can be compiled with PS2EXE
    alongside Toolkit.ps1. Pass -XamlPath to load an external
    RemoveDeviceGroups.xaml instead while iterating on the UI.

    Every function here is Rdg-prefixed on purpose. This module and
    CopyDeviceGroups.ps1 are dot-sourced into the SAME session by Toolkit.ps1,
    so a shared name would mean the last file loaded silently wins and could
    change the behaviour of the other module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\RemoveDeviceGroups\RemoveDeviceGroups.ps1"
        Show-RemoveDeviceGroupsWindow -Device $Script:SelectedDevice `
                                      -Owner  $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: Device.Read.All, Group.Read.All,
                  Group.ReadWrite.All, GroupMember.ReadWrite.All
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external RemoveDeviceGroups.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# Row type for the groups grid.
#
# This is deliberately a separate type from EndpointguyGroupRow (used by the
# Copy module): it adds a Result column so each removal outcome can be written
# back into the grid as it happens. Add-Type cannot redefine a type that is
# already loaded in the session, so reusing the name would leave whichever
# module loaded second without its Result property.
#
# Selected and Result both raise PropertyChanged so the grid repaints when
# Select all / Select none is used and while the removal loop runs.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyRemoveGroupRow' -as [type])) {
    Add-Type -ReferencedAssemblies System.ComponentModel.TypeConverter -TypeDefinition @'
using System.ComponentModel;

public class EndpointguyRemoveGroupRow : INotifyPropertyChanged
{
    private bool   _selected;
    private string _result = "";

    public bool Selected
    {
        get { return _selected; }
        set
        {
            if (_selected != value)
            {
                _selected = value;
                OnPropertyChanged("Selected");
            }
        }
    }

    public string Result
    {
        get { return _result; }
        set
        {
            if (_result != value)
            {
                _result = value;
                OnPropertyChanged("Result");
            }
        }
    }

    public bool   IsEligible  { get; set; }
    public string DisplayName { get; set; }
    public string GroupType   { get; set; }
    public string Status      { get; set; }
    public string GroupId     { get; set; }

    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged(string name)
    {
        PropertyChangedEventHandler handler = PropertyChanged;
        if (handler != null) { handler(this, new PropertyChangedEventArgs(name)); }
    }
}
'@
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$RdgXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Remove Device Groups"
        Height="820" Width="1150"
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
                <TextBlock Text="Remove Device Groups" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Removes the selected device from its assigned (static) security groups. Dynamic groups are listed but cannot be removed because their membership is evaluated by rule."
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

        <!-- Groups -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Group memberships" Style="{StaticResource CardHeader}"/>
                        <TextBlock x:Name="GroupCount" Text="" Margin="12,0,0,0"
                                   VerticalAlignment="Bottom" FontSize="13"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <CheckBox x:Name="ChkShowIneligible" Content="Show ineligible groups"
                                  VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"/>
                        <Button x:Name="BtnSelectAll"    Style="{StaticResource NeutralButton}" Content="Select all"   Margin="0,0,10,0"/>
                        <Button x:Name="BtnSelectNone"   Style="{StaticResource NeutralButton}" Content="Select none"  Margin="0,0,10,0"/>
                        <Button x:Name="BtnReloadGroups" Style="{StaticResource NeutralButton}" Content="Reload groups"/>
                    </StackPanel>
                </Grid>

                <DataGrid x:Name="GridGroups" Grid.Row="1"
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
                        <DataGridCheckBoxColumn Header="Remove" Width="70"
                                                Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                            <DataGridCheckBoxColumn.ElementStyle>
                                <Style TargetType="CheckBox">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                    <Setter Property="VerticalAlignment" Value="Center"/>
                                    <Setter Property="IsEnabled" Value="{Binding IsEligible}"/>
                                </Style>
                            </DataGridCheckBoxColumn.ElementStyle>
                            <DataGridCheckBoxColumn.EditingElementStyle>
                                <Style TargetType="CheckBox">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                    <Setter Property="VerticalAlignment" Value="Center"/>
                                    <Setter Property="IsEnabled" Value="{Binding IsEligible}"/>
                                </Style>
                            </DataGridCheckBoxColumn.EditingElementStyle>
                        </DataGridCheckBoxColumn>
                        <DataGridTextColumn Header="Group Name"  Binding="{Binding DisplayName}" Width="320" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Type"        Binding="{Binding GroupType}"   Width="170" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Eligibility" Binding="{Binding Status}"      Width="280" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Result"      Binding="{Binding Result}"      Width="160" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Group Id"    Binding="{Binding GroupId}"     Width="*"   IsReadOnly="True"/>
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
                        <TextBlock Text="Live Graph calls, eligibility decisions and errors" Margin="12,0,0,0"
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

                <TextBox x:Name="TxtLog" Grid.Row="1" Height="200"
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
                    <Button x:Name="BtnRemove" Style="{StaticResource DangerButton}"
                            Content="Remove selected groups" MinWidth="220" Height="38" Margin="0,0,12,0"/>
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
$Script:RdgLogSink    = $null
$Script:RdgLogVerbose = $false

function Write-RdgLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:RdgLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:RdgLogSink) {
        try { & $Script:RdgLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-RdgLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-RdgGraphError $ErrorRecord
    Write-RdgLog "$Context failed: $detail" 'ERROR'
    if ($Script:RdgLogVerbose -and $ErrorRecord) {
        $ex = Get-RdgProperty $ErrorRecord 'Exception'
        if ($ex) { Write-RdgLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-RdgProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-RdgLog "At: $((Get-RdgProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-RdgProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-RdgProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-RdgProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-RdgGraphPaged {
    # Walks @odata.nextLink and returns every item in the 'value' array.
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($Uri) {
        $page++
        Write-RdgLog "GET $Uri" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        $sw.Stop()

        $count = 0
        if (Test-RdgProperty $resp 'value') { $count = @($resp.value).Count; $all.AddRange(@($resp.value)) }
        Write-RdgLog "  -> page $page returned $count item(s) in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        $Uri = if (Test-RdgProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        if ($Uri) { Write-RdgLog '  -> following @odata.nextLink for the next page' 'DEBUG' }
    }
    return $all
}

function Get-RdgEntraDevice {
    <#
        Group membership lives on the Entra ID device object, not on the Intune
        managed device. managedDevice.azureADDeviceId is the Entra deviceId
        (not the directory object id), so it has to be resolved first.
    #>
    param([Parameter(Mandatory)]$Device)

    $deviceId = Get-RdgProperty $Device 'EntraDeviceId'
    if (-not $deviceId) { $deviceId = Get-RdgProperty $Device 'azureADDeviceId' }

    $name = Get-RdgProperty $Device 'DeviceName' '(unknown device)'

    if ([string]::IsNullOrWhiteSpace($deviceId) -or
        $deviceId -eq '00000000-0000-0000-0000-000000000000') {
        throw "$name has no Entra ID device record (it may be Intune-only or co-managed without Entra join), so its groups cannot be read or written."
    }

    $uri  = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$deviceId'&`$select=id,displayName,deviceId,accountEnabled,onPremisesSyncEnabled"
    $hits = @(Get-RdgGraphPaged -Uri $uri)

    if ($hits.Count -eq 0) {
        throw "No Entra ID device object was found for $name (deviceId $deviceId)."
    }

    [pscustomobject]@{
        ObjectId    = $hits[0].id
        DisplayName = Get-RdgProperty $hits[0] 'displayName' $name
        DeviceId    = $deviceId
        DeviceName  = $name
    }
}

function Get-RdgGroupDetail {
    <#
        Reads a group's real properties straight from /groups/<id>.

        The memberOf endpoint returns a directoryObject collection, and Graph
        does not reliably honour a $select of group-only properties on it -
        groupTypes in particular is often missing. A group with a missing
        groupTypes then looks static, gets marked eligible, and the removal
        fails with a bare 400. Reading each group directly is the only way to
        know whether it is dynamic.
    #>
    param([Parameter(Mandatory)][string]$GroupId)

    $select = 'id,displayName,groupTypes,securityEnabled,mailEnabled,' +
              'membershipRule,membershipRuleProcessingState,onPremisesSyncEnabled'

    return Invoke-MgGraphRequest -Method GET -OutputType PSObject `
        -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=$select"
}

function Get-RdgGroupRows {
    <#
        Returns one EndpointguyRemoveGroupRow per group the device belongs to.

        Eligible   = assigned (static) security group that Graph will accept a
                     device member being removed from.
        Ineligible = dynamic, on-premises synced, mail-enabled, Microsoft 365,
                     or a directory role / administrative unit.
    #>
    param([Parameter(Mandatory)][string]$EntraObjectId)

    $uri  = "https://graph.microsoft.com/v1.0/devices/$EntraObjectId/memberOf`?`$top=999"
    $rows = New-Object System.Collections.Generic.List[object]

    Write-RdgLog "Reading group memberships for device object $EntraObjectId" 'INFO'

    foreach ($o in (Get-RdgGraphPaged -Uri $uri)) {
        $odataType = Get-RdgProperty $o '@odata.type' ''
        $name      = Get-RdgProperty $o 'displayName' '(unnamed)'
        $groupId   = Get-RdgProperty $o 'id' ''

        # memberOf also returns directory roles and administrative units.
        if ($odataType -and $odataType -ne '#microsoft.graph.group') {
            $kind = ($odataType -replace '#microsoft\.graph\.', '')
            $row  = New-Object EndpointguyRemoveGroupRow
            $row.DisplayName = $name
            $row.GroupType   = $kind
            $row.Status      = 'Not a group - cannot be removed here'
            $row.GroupId     = $groupId
            $row.IsEligible  = $false
            $row.Selected    = $false
            $rows.Add($row) | Out-Null
            continue
        }

        # Re-read the group itself so groupTypes is guaranteed to be present.
        $g = $o
        try   { if ($groupId) { $g = Get-RdgGroupDetail -GroupId $groupId } }
        catch { $g = $o; Write-RdgLog "Could not re-read group '$name' ($groupId); falling back to the memberOf projection - eligibility may be wrong. $(Get-RdgGraphError $_)" 'WARN' }

        if (Test-RdgProperty $g 'displayName') { $name = $g.displayName }

        $groupTypes  = @(Get-RdgProperty $g 'groupTypes' @())
        $isDynamic   = ($groupTypes -contains 'DynamicMembership')
        $isUnified   = ($groupTypes -contains 'Unified')
        $secEnabled  = [bool](Get-RdgProperty $g 'securityEnabled' $false)
        $mailEnabled = [bool](Get-RdgProperty $g 'mailEnabled' $false)
        $onPrem      = [bool](Get-RdgProperty $g 'onPremisesSyncEnabled' $false)

        # A membership rule is the definitive tell: some tenants return an
        # empty groupTypes on a group that is still rule-driven.
        $rule = Get-RdgProperty $g 'membershipRule' ''
        if (-not [string]::IsNullOrWhiteSpace($rule)) { $isDynamic = $true }

        $type =
            if     ($isUnified)                    { 'Microsoft 365' }
            elseif ($isDynamic)                    { 'Dynamic security' }
            elseif ($secEnabled -and $mailEnabled) { 'Mail-enabled security' }
            elseif ($secEnabled)                   { 'Assigned security' }
            elseif ($mailEnabled)                  { 'Distribution' }
            else                                   { 'Other' }

        $eligible = $false
        $status   = ''

        if ($isDynamic) {
            $status = 'Dynamic - membership is set by rule, cannot be removed'
        }
        elseif ($onPrem) {
            $status = 'Synced from on-premises AD - read only in Entra ID'
        }
        elseif ($isUnified) {
            $status = 'Microsoft 365 group - devices cannot be members'
        }
        elseif ($mailEnabled) {
            $status = 'Mail-enabled group - devices cannot be members'
        }
        elseif (-not $secEnabled) {
            $status = 'Not a security group - devices cannot be members'
        }
        else {
            $eligible = $true
            $status   = 'Eligible - assigned security group'
        }

        $row = New-Object EndpointguyRemoveGroupRow
        $row.DisplayName = $name
        $row.GroupType   = $type
        $row.Status      = $status
        $row.GroupId     = $groupId
        $row.IsEligible  = $eligible
        $row.Selected    = $eligible   # pre-tick everything that can be removed
        $row.Result      = ''

        $verdict = if ($eligible) { 'ELIGIBLE' } else { 'skipped ' }
        Write-RdgLog "  $verdict | $name | $type | $status" 'DEBUG'
        $rows.Add($row) | Out-Null
    }

    $elig = @($rows | Where-Object { $_.IsEligible }).Count
    Write-RdgLog "Evaluated $($rows.Count) membership(s): $elig eligible, $($rows.Count - $elig) ineligible" 'OK'

    return ($rows | Sort-Object -Property @{Expression={ -not $_.IsEligible }}, DisplayName)
}

function Get-RdgBodyError {
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

    if (Test-RdgProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-RdgProperty $err 'code' ''
        $msg  = Get-RdgProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-RdgGraphError {
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
    if (Test-RdgProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-RdgProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-RdgBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-RdgProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-RdgProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-RdgProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-RdgBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-RdgProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-RdgProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-RdgProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

function Remove-RdgDeviceFromGroup {
    <#
        Removes the device directory object from a group. Returns 'Removed' or
        'Not a member', or throws with the real Graph message.

        DELETE /groups/<id>/members/<deviceObjectId>/$ref is the only supported
        way to drop a single member; it returns 204 on success and 404 when the
        device was not in the group to begin with, which is not an error worth
        showing the operator as a failure.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$DeviceObjectId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw 'The group has no id, so it cannot be updated.'
    }
    if ([string]::IsNullOrWhiteSpace($DeviceObjectId)) {
        throw 'The device has no Entra ID object id, so it cannot be removed from a group.'
    }

    Write-RdgLog "DELETE group $GroupId <- device object $DeviceObjectId" 'GRAPH'
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$DeviceObjectId/`$ref"

    # Preferred path: ask Graph not to throw, so the error body is returned
    # intact instead of being flattened into a useless exception message.
    $canSkipThrow = $false
    try {
        $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction Stop
        $canSkipThrow = $cmd.Parameters.ContainsKey('SkipHttpErrorCheck') -and
                        $cmd.Parameters.ContainsKey('StatusCodeVariable')
    }
    catch { $canSkipThrow = $false }

    Write-RdgLog "  -> using $(if ($canSkipThrow) { '-SkipHttpErrorCheck (full error body)' } else { 'try/catch (legacy Graph module)' })" 'DEBUG'

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest -Method DELETE -Uri $uri -OutputType PSObject `
                    -SkipHttpErrorCheck -StatusCodeVariable 'code'

        Write-RdgLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) { Write-RdgLog '  -> removed' 'OK'; return 'Removed' }
        if ($code -eq 404) { Write-RdgLog '  -> device was not a member' 'INFO'; return 'Not a member' }

        $detail = Get-RdgBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        Write-RdgLog "  -> Graph said: $detail" 'DEBUG'

        if ($detail -match 'does not exist|Resource .* does not exist') { return 'Not a member' }
        throw "HTTP $code - $detail"
    }

    try {
        Invoke-MgGraphRequest -Method DELETE -Uri $uri | Out-Null
        return 'Removed'
    }
    catch {
        $detail = Get-RdgGraphError $_
        if ($detail -match 'does not exist|not found|ResourceNotFound') { return 'Not a member' }
        throw $detail
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-RemoveDeviceGroupsWindow {
    <#
    .SYNOPSIS
        Opens the Remove Device Groups window for the supplied device.
    .PARAMETER Device
        The device object selected in the Toolkit results grid. Needs at least
        DeviceName and EntraDeviceId.
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
    $xamlText = $RdgXamlString
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
            "The Remove Device Groups XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtDeviceName','TxtDeviceDetail','GridGroups','GroupCount',
                     'ChkShowIneligible','BtnSelectAll','BtnSelectNone','BtnReloadGroups',
                     'StatusText','BtnRemove','BtnClose',
                     'TxtLog','ChkVerbose','ChkAutoScroll','BtnCopyLog','BtnSaveLog','BtnClearLog',
                     'ConsolePanel')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the Remove Device Groups XAML.",
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
        Entra = $null
        Rows  = @()
        Busy  = $false
    }

    $groupRows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $ui.GridGroups.ItemsSource = $groupRows

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # --- diagnostics console ----------------------------------------------
    # The sink is what Write-RdgLog calls. Appending to a TextBox is cheap
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
    $Script:RdgLogSink    = $appendLog
    $Script:RdgLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
        $win.Height = 1000
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnRemove, $ui.BtnReloadGroups,
                         $ui.BtnSelectAll, $ui.BtnSelectNone)) { $b.IsEnabled = -not $Busy }
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    $refreshGrid = {
        # Rebuilds the visible rows, honouring the 'Show ineligible' toggle.
        $showAll = [bool]$ui.ChkShowIneligible.IsChecked
        $groupRows.Clear()
        foreach ($r in $state.Rows) {
            if ($showAll -or $r.IsEligible) { $groupRows.Add($r) }
        }
        $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
        $total    = @($state.Rows).Count
        $ui.GroupCount.Text = "$eligible of $total group(s) eligible to remove"
    }

    $loadGroups = {
        & $setBusy $true
        try {
            & $setStatus "Reading Entra ID device record for $($Device.DeviceName)..."
            $state.Entra = Get-RdgEntraDevice -Device $Device

            $ui.TxtDeviceName.Text   = $state.Entra.DeviceName
            $ui.TxtDeviceDetail.Text = "Serial: $(Get-RdgProperty $Device 'SerialNumber' 'n/a')`nUser: $(Get-RdgProperty $Device 'User' 'n/a')`nEntra object: $($state.Entra.ObjectId)"

            & $setStatus 'Reading group memberships...'
            $state.Rows = @(Get-RdgGroupRows -EntraObjectId $state.Entra.ObjectId)
            & $refreshGrid

            $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
            & $setStatus "Loaded $(@($state.Rows).Count) group(s) for $($state.Entra.DeviceName) - $eligible can be removed."
        }
        catch {
            $detail = Write-RdgLogException $_ 'Loading groups'
            & $setStatus "Could not load groups: $detail"
            [System.Windows.MessageBox]::Show($detail,'Remove Device Groups','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- remove ------------------------------------------------------------
    $doRemove = {
        $picked = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected })
        if ($picked.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Tick at least one eligible group to remove.',
                'Remove Device Groups','OK','Warning') | Out-Null
            return
        }

        $deviceName = $state.Entra.DisplayName

        # Destructive and not undoable, so the prompt names the device, states
        # the count, and defaults to No.
        $names = ($picked | Select-Object -First 15 | ForEach-Object { "  - $($_.DisplayName)" }) -join "`n"
        if ($picked.Count -gt 15) { $names = "$names`n  ... and $($picked.Count - 15) more" }

        $confirm = [System.Windows.MessageBox]::Show(
            "Remove $deviceName from $($picked.Count) group(s)?`n`n$names`n`nThis takes the device out of those groups in Entra ID. Any policies, apps or configuration targeted at them will stop applying. This CANNOT be undone from here - the memberships would have to be added back by hand.",
            'Confirm removal','YesNo','Warning',[System.Windows.MessageBoxResult]::No)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-RdgLog 'Removal cancelled at the confirmation prompt.' 'INFO'
            & $setStatus 'Cancelled - nothing was removed.'
            return
        }

        & $setBusy $true
        $removed = 0; $notMember = 0; $failed = 0
        $lines = New-Object System.Collections.Generic.List[string]
        $i = 0

        foreach ($row in $picked) {
            $i++
            & $setStatus "Removing $i of $($picked.Count): $($row.DisplayName)..."
            $row.Result = 'Working...'
            try {
                $result = Remove-RdgDeviceFromGroup -GroupId $row.GroupId -DeviceObjectId $state.Entra.ObjectId
                $row.Result = $result
                Write-RdgLog "$($row.DisplayName) - $result" 'OK'
                if ($result -eq 'Removed') { $removed++ } else { $notMember++ }
                $lines.Add("$($row.DisplayName) - $result") | Out-Null
            }
            catch {
                $failed++
                $detail = Write-RdgLogException $_ "Remove '$($row.DisplayName)'"
                $row.Result = "FAILED: $detail"
                $lines.Add("$($row.DisplayName) - FAILED: $detail") | Out-Null
            }
        }

        & $setBusy $false
        $summary = "Removal complete - $removed removed, $notMember not a member, $failed failed."
        & $setStatus $summary
        Write-RdgLog $summary $(if ($failed -gt 0) { 'WARN' } else { 'OK' })

        $icon = if ($failed -gt 0) { 'Warning' } else { 'Information' }
        [System.Windows.MessageBox]::Show(
            "$summary`n`n$($lines -join "`n")",
            'Remove Device Groups','OK',$icon) | Out-Null

        # Re-read from Graph so the grid reflects what is actually left.
        if ($removed -gt 0) { & $loadGroups }
    }

    # --- events ------------------------------------------------------------
    $ui.ChkShowIneligible.Add_Checked({   & $refreshGrid })
    $ui.ChkShowIneligible.Add_Unchecked({ & $refreshGrid })

    $ui.BtnSelectAll.Add_Click({
        foreach ($r in $state.Rows) { if ($r.IsEligible) { $r.Selected = $true } }
    })
    $ui.BtnSelectNone.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $false }
    })

    $ui.ChkVerbose.Add_Checked({
        $Script:RdgLogVerbose = $true
        Write-RdgLog 'Verbose logging on - per-group decisions and Graph timings will be shown.' 'INFO'
    })
    $ui.ChkVerbose.Add_Unchecked({
        $Script:RdgLogVerbose = $false
        Write-RdgLog 'Verbose logging off.' 'INFO'
    })

    $ui.BtnClearLog.Add_Click({
        $ui.TxtLog.Clear()
        Write-RdgLog 'Log cleared.' 'INFO'
    })

    $ui.BtnCopyLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
        & $setStatus 'Diagnostics log copied to the clipboard.'
    })

    $ui.BtnSaveLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $file  = Join-Path ([Environment]::GetFolderPath('Desktop')) "RemoveDeviceGroups-$stamp.log"
        try {
            $ui.TxtLog.Text | Set-Content -Path $file -Encoding UTF8
            & $setStatus "Log saved to $file"
            Write-RdgLog "Log saved to $file" 'OK'
        }
        catch { Write-RdgLog "Could not save the log: $($_.Exception.Message)" 'ERROR' }
    })

    $ui.BtnReloadGroups.Add_Click({ & $loadGroups })
    $ui.BtnRemove.Add_Click({      & $doRemove })
    $ui.BtnClose.Add_Click({       $win.Close() })

    # Drop the sink so a closed window is never written to.
    $win.Add_Closed({ $Script:RdgLogSink = $null })

    # F12 toggles the console without reopening the window.
    $win.Add_KeyDown({
        if ($args[1].Key -ne 'F12') { return }
        if ($ui.ConsolePanel.Visibility -eq [System.Windows.Visibility]::Visible) {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Collapsed
            $win.Height = 820
        }
        else {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
            $win.Height = 1000
        }
    })

    # --- go ----------------------------------------------------------------
    $ui.TxtDeviceName.Text = Get-RdgProperty $Device 'DeviceName' '(unknown device)'
    Write-RdgLog 'Remove Device Groups opened.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-RdgLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-RdgLog "Device: $(Get-RdgProperty $Device 'DeviceName' '(unknown)')" 'INFO'
    $win.Add_ContentRendered({ & $loadGroups })
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# prompt for a device by name so the module can be tested on its own.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'Device.Read.All','Group.Read.All','Group.ReadWrite.All',
                            'GroupMember.ReadWrite.All','DeviceManagementManagedDevices.Read.All' -NoWelcome

    $name = Read-Host 'Device name'
    $esc  = $name.Replace("'", "''")
    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$esc'&`$select=id,deviceName,serialNumber,userPrincipalName,azureADDeviceId"
    $dev  = @(Get-RdgGraphPaged -Uri $uri | ForEach-Object {
        [pscustomobject]@{
            DeviceName    = Get-RdgProperty $_ 'deviceName' ''
            SerialNumber  = Get-RdgProperty $_ 'serialNumber' ''
            User          = Get-RdgProperty $_ 'userPrincipalName' ''
            EntraDeviceId = Get-RdgProperty $_ 'azureADDeviceId' ''
        }
    })

    if ($dev.Count -eq 0) { Write-Warning "No managed device named '$name' was found."; return }
    Show-RemoveDeviceGroupsWindow -Device $dev[0] -XamlPath $XamlPath -ShowConsole
}
