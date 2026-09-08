<#
.SYNOPSIS
    Copy Device Groups - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that copies Entra ID security group memberships from a
    source managed device to a target managed device.

      - The source device is passed in from the Toolkit (the device selected
        in the search results) and is pre-filled in the window.
      - The operator searches for and picks the target device.
      - Every group the source device belongs to is listed, with each group
        marked eligible or ineligible. Only assigned (static) security groups
        are eligible; dynamic groups are shown but cannot be selected because
        their membership is evaluated by rule and cannot be written to.
      - Selected groups are added to the target device one at a time, with a
        per-group result summary at the end.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so CopyDeviceGroups.xaml
    is not required at runtime and this file can be compiled with PS2EXE
    alongside Toolkit.ps1. Pass -XamlPath to load an external
    CopyDeviceGroups.xaml instead while iterating on the UI.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\CopyDeviceGroups\CopyDeviceGroups.ps1"
        Show-CopyDeviceGroupsWindow -SourceDevice $Script:SelectedDevice `
                                    -DeviceCache  $Script:DeviceCache `
                                    -Owner        $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: Device.Read.All, Group.Read.All,
                  Group.ReadWrite.All, GroupMember.ReadWrite.All
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external CopyDeviceGroups.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# Row type for the groups grid.
#
# A [pscustomobject] cannot raise change notifications, so the Select all /
# Select none buttons would not repaint the check boxes. A small class that
# implements INotifyPropertyChanged keeps the grid and the data in step, and
# IsEligible drives the per-row IsEnabled binding on the check box column.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyGroupRow' -as [type])) {
    Add-Type -ReferencedAssemblies System.ComponentModel.TypeConverter -TypeDefinition @'
using System.ComponentModel;

public class EndpointguyGroupRow : INotifyPropertyChanged
{
    private bool _selected;

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
$CdgXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Copy Device Groups"
        Height="860" Width="1200"
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
                <TextBlock Text="Copy Device Groups" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Copies assigned (static) security group memberships from the source device to a target device. Dynamic groups cannot be copied because membership is evaluated by rule."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Source / target devices -->
        <Grid Grid.Row="1" Margin="0,14,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="1.35*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Style="{StaticResource Card}" Padding="18">
                <StackPanel>
                    <TextBlock Text="Source device (copy FROM)" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                    <Border CornerRadius="4"
                            Background="{StaticResource InfoPanelBrush}"
                            BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                            Padding="12,10">
                        <StackPanel>
                            <TextBlock Text="SELECTED IN TOOLKIT" FontSize="10.5" FontWeight="Bold"
                                       Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                            <TextBlock x:Name="TxtSourceDevice" Text="No device selected"
                                       FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
                            <TextBlock x:Name="TxtSourceDetail" Text=""
                                       FontSize="12" Foreground="{StaticResource SubtleTextBrush}"
                                       TextWrapping="Wrap" LineHeight="17" Margin="0,6,0,0"/>
                        </StackPanel>
                    </Border>
                </StackPanel>
            </Border>

            <Border Grid.Column="2" Style="{StaticResource Card}" Padding="18">
                <StackPanel>
                    <TextBlock Text="Target device (copy TO)" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                    <StackPanel Orientation="Horizontal">
                        <TextBox x:Name="TxtTargetSearch" Width="320" Height="32"/>
                        <Button x:Name="BtnFindTarget" Style="{StaticResource PrimaryButton}"
                                Content="Find device" Width="130" Height="32" Margin="12,0,0,0"/>
                        <TextBlock Text="Name, serial or user" VerticalAlignment="Center" FontSize="12"
                                   Foreground="{StaticResource SubtleTextBrush}" Margin="12,0,0,0"/>
                    </StackPanel>
                    <DataGrid x:Name="GridTargets" Height="118" Margin="0,12,0,0"
                              AutoGenerateColumns="False" IsReadOnly="True"
                              SelectionMode="Single" HeadersVisibility="Column"
                              GridLinesVisibility="Horizontal"
                              HorizontalGridLinesBrush="{StaticResource GridLineBrush}"
                              Background="{StaticResource GridRowBrush}"
                              RowBackground="{StaticResource GridRowBrush}"
                              AlternatingRowBackground="{StaticResource GridAltRowBrush}"
                              BorderBrush="{StaticResource GridBorderBrush}"
                              BorderThickness="1" FontSize="12.5" RowHeight="24">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Device Name"   Binding="{Binding DeviceName}"   Width="200"/>
                            <DataGridTextColumn Header="Serial Number" Binding="{Binding SerialNumber}" Width="130"/>
                            <DataGridTextColumn Header="User"          Binding="{Binding User}"         Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock x:Name="TxtTargetSelected" Text="No target device selected."
                               FontSize="12.5" Margin="2,10,0,0" TextWrapping="Wrap"
                               Foreground="{StaticResource WarnTextBrush}"/>
                </StackPanel>
            </Border>
        </Grid>

        <!-- Groups on the source device -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Groups on source device" Style="{StaticResource CardHeader}"/>
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
                        <DataGridCheckBoxColumn Header="Copy" Width="60"
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
                        <DataGridTextColumn Header="Group Name"  Binding="{Binding DisplayName}" Width="340" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Type"        Binding="{Binding GroupType}"   Width="180" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Eligibility" Binding="{Binding Status}"      Width="290" IsReadOnly="True"/>
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
                    <Button x:Name="BtnCopy"  Style="{StaticResource GreenButton}"
                            Content="Copy selected groups" MinWidth="210" Height="38" Margin="0,0,12,0"/>
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
$Script:CdgLogSink    = $null
$Script:CdgLogVerbose = $false

function Write-CdgLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:CdgLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:CdgLogSink) {
        try { & $Script:CdgLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-CdgLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-CdgGraphError $ErrorRecord
    Write-CdgLog "$Context failed: $detail" 'ERROR'
    if ($Script:CdgLogVerbose -and $ErrorRecord) {
        $ex = Get-CdgProperty $ErrorRecord 'Exception'
        if ($ex) { Write-CdgLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-CdgProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-CdgLog "At: $((Get-CdgProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-CdgProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-CdgProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-CdgProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-CdgGraphPaged {
    # Walks @odata.nextLink and returns every item in the 'value' array.
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($Uri) {
        $page++
        Write-CdgLog "GET $Uri" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        $sw.Stop()

        $count = 0
        if (Test-CdgProperty $resp 'value') { $count = @($resp.value).Count; $all.AddRange(@($resp.value)) }
        Write-CdgLog "  -> page $page returned $count item(s) in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        $Uri = if (Test-CdgProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        if ($Uri) { Write-CdgLog '  -> following @odata.nextLink for the next page' 'DEBUG' }
    }
    return $all
}

function Get-CdgEntraDevice {
    <#
        Group membership lives on the Entra ID device object, not on the Intune
        managed device. managedDevice.azureADDeviceId is the Entra deviceId
        (not the directory object id), so it has to be resolved first.
    #>
    param([Parameter(Mandatory)]$Device)

    $deviceId = Get-CdgProperty $Device 'EntraDeviceId'
    if (-not $deviceId) { $deviceId = Get-CdgProperty $Device 'azureADDeviceId' }

    $name = Get-CdgProperty $Device 'DeviceName' '(unknown device)'

    if ([string]::IsNullOrWhiteSpace($deviceId) -or
        $deviceId -eq '00000000-0000-0000-0000-000000000000') {
        throw "$name has no Entra ID device record (it may be Intune-only or co-managed without Entra join), so its groups cannot be read or written."
    }

    $uri  = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$deviceId'&`$select=id,displayName,deviceId,accountEnabled,onPremisesSyncEnabled"
    $hits = @(Get-CdgGraphPaged -Uri $uri)

    if ($hits.Count -eq 0) {
        throw "No Entra ID device object was found for $name (deviceId $deviceId)."
    }

    [pscustomobject]@{
        ObjectId    = $hits[0].id
        DisplayName = Get-CdgProperty $hits[0] 'displayName' $name
        DeviceId    = $deviceId
        DeviceName  = $name
    }
}

function Get-CdgGroupDetail {
    <#
        Reads a group's real properties straight from /groups/<id>.

        The memberOf endpoint returns a directoryObject collection, and Graph
        does not reliably honour a $select of group-only properties on it -
        groupTypes in particular is often missing. A group with a missing
        groupTypes then looks static, gets marked eligible, and the add fails
        with a bare 400 at copy time. Reading each group directly is the only
        way to know whether it is dynamic.
    #>
    param([Parameter(Mandatory)][string]$GroupId)

    $select = 'id,displayName,groupTypes,securityEnabled,mailEnabled,' +
              'membershipRule,membershipRuleProcessingState,onPremisesSyncEnabled'

    return Invoke-MgGraphRequest -Method GET -OutputType PSObject `
        -Uri "https://graph.microsoft.com/v1.0/groups/$GroupId`?`$select=$select"
}

function Get-CdgGroupRows {
    <#
        Returns one EndpointguyGroupRow per group the device belongs to.

        Eligible   = assigned (static) security group that Graph will accept a
                     device member on.
        Ineligible = dynamic, on-premises synced, mail-enabled, Microsoft 365,
                     or a directory role / administrative unit.
    #>
    param([Parameter(Mandatory)][string]$EntraObjectId)

    $uri  = "https://graph.microsoft.com/v1.0/devices/$EntraObjectId/memberOf`?`$top=999"
    $rows = New-Object System.Collections.Generic.List[object]

    Write-CdgLog "Reading group memberships for device object $EntraObjectId" 'INFO'

    foreach ($o in (Get-CdgGraphPaged -Uri $uri)) {
        $odataType = Get-CdgProperty $o '@odata.type' ''
        $name      = Get-CdgProperty $o 'displayName' '(unnamed)'
        $groupId   = Get-CdgProperty $o 'id' ''

        # memberOf also returns directory roles and administrative units.
        if ($odataType -and $odataType -ne '#microsoft.graph.group') {
            $kind = ($odataType -replace '#microsoft\.graph\.', '')
            $row  = New-Object EndpointguyGroupRow
            $row.DisplayName = $name
            $row.GroupType   = $kind
            $row.Status      = 'Not a group - cannot be copied'
            $row.GroupId     = $groupId
            $row.IsEligible  = $false
            $row.Selected    = $false
            $rows.Add($row) | Out-Null
            continue
        }

        # Re-read the group itself so groupTypes is guaranteed to be present.
        $g = $o
        try   { if ($groupId) { $g = Get-CdgGroupDetail -GroupId $groupId } }
        catch { $g = $o; Write-CdgLog "Could not re-read group '$name' ($groupId); falling back to the memberOf projection - eligibility may be wrong. $(Get-CdgGraphError $_)" 'WARN' }

        if (Test-CdgProperty $g 'displayName') { $name = $g.displayName }

        $groupTypes  = @(Get-CdgProperty $g 'groupTypes' @())
        $isDynamic   = ($groupTypes -contains 'DynamicMembership')
        $isUnified   = ($groupTypes -contains 'Unified')
        $secEnabled  = [bool](Get-CdgProperty $g 'securityEnabled' $false)
        $mailEnabled = [bool](Get-CdgProperty $g 'mailEnabled' $false)
        $onPrem      = [bool](Get-CdgProperty $g 'onPremisesSyncEnabled' $false)

        # A membership rule is the definitive tell: some tenants return an
        # empty groupTypes on a group that is still rule-driven.
        $rule = Get-CdgProperty $g 'membershipRule' ''
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
            $status = 'Dynamic - membership is set by rule, cannot be copied'
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

        $row = New-Object EndpointguyGroupRow
        $row.DisplayName = $name
        $row.GroupType   = $type
        $row.Status      = $status
        $row.GroupId     = $groupId
        $row.IsEligible  = $eligible
        $row.Selected    = $eligible   # pre-tick everything that can be copied

        $verdict = if ($eligible) { 'ELIGIBLE' } else { 'skipped ' }
        Write-CdgLog "  $verdict | $name | $type | $status" 'DEBUG'
        $rows.Add($row) | Out-Null
    }

    $elig = @($rows | Where-Object { $_.IsEligible }).Count
    Write-CdgLog "Evaluated $($rows.Count) membership(s): $elig eligible, $($rows.Count - $elig) ineligible" 'OK'

    return ($rows | Sort-Object -Property @{Expression={ -not $_.IsEligible }}, DisplayName)
}

function Get-CdgBodyError {
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

    if (Test-CdgProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-CdgProperty $err 'code' ''
        $msg  = Get-CdgProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-CdgGraphError {
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
    if (Test-CdgProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-CdgProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-CdgBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-CdgProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-CdgProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-CdgProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-CdgBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-CdgProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-CdgProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-CdgProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

function Add-CdgDeviceToGroup {
    <#
        Adds the device directory object to a group. Returns 'Added',
        'Already a member', or throws with the real Graph message.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$DeviceObjectId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw 'The group has no id, so it cannot be updated.'
    }
    if ([string]::IsNullOrWhiteSpace($DeviceObjectId)) {
        throw 'The target device has no Entra ID object id, so it cannot be added to a group.'
    }

    Write-CdgLog "POST group $GroupId <- device object $DeviceObjectId" 'GRAPH'
    $uri  = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"
    $body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$DeviceObjectId" } |
            ConvertTo-Json -Compress

    # Preferred path: ask Graph not to throw, so the error body is returned
    # intact instead of being flattened into a useless exception message.
    $canSkipThrow = $false
    try {
        $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction Stop
        $canSkipThrow = $cmd.Parameters.ContainsKey('SkipHttpErrorCheck') -and
                        $cmd.Parameters.ContainsKey('StatusCodeVariable')
    }
    catch { $canSkipThrow = $false }

    Write-CdgLog "  -> using $(if ($canSkipThrow) { '-SkipHttpErrorCheck (full error body)' } else { 'try/catch (legacy Graph module)' })" 'DEBUG'

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body `
                    -ContentType 'application/json' -OutputType PSObject `
                    -SkipHttpErrorCheck -StatusCodeVariable 'code'

        Write-CdgLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) { Write-CdgLog '  -> added' 'OK'; return 'Added' }

        $detail = Get-CdgBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        Write-CdgLog "  -> Graph said: $detail" 'DEBUG'

        if ($detail -match 'already exist') { return 'Already a member' }
        throw "HTTP $code - $detail"
    }

    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body `
            -ContentType 'application/json' | Out-Null
        return 'Added'
    }
    catch {
        $detail = Get-CdgGraphError $_
        if ($detail -match 'already exist') { return 'Already a member' }
        throw $detail
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-CopyDeviceGroupsWindow {
    <#
    .SYNOPSIS
        Opens the Copy Device Groups window for the supplied source device.
    .PARAMETER SourceDevice
        The device object selected in the Toolkit results grid. Needs at least
        DeviceName and EntraDeviceId.
    .PARAMETER DeviceCache
        The Toolkit device cache, used for the target device lookup so the
        search is instant. Falls back to a live Graph query when empty.
    .PARAMETER Owner
        The Toolkit window, so this one centres on it and stays on top.
    .PARAMETER ShowConsole
        Opens with the diagnostics console visible. Leave it off for normal
        use; turn it on when troubleshooting a new module or a Graph error.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$SourceDevice,
        $DeviceCache = @(),
        $Owner       = $null,
        [string]$XamlPath,

        # Reveals the diagnostics console. Off by default - the log is still
        # captured in memory, so ticking it open mid-session shows the
        # history rather than starting blank.
        [switch]$ShowConsole
    )

    # --- load XAML ---------------------------------------------------------
    $xamlText = $CdgXamlString
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
            "The Copy Device Groups XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtSourceDevice','TxtSourceDetail','TxtTargetSearch','BtnFindTarget',
                     'GridTargets','TxtTargetSelected','GridGroups','GroupCount',
                     'ChkShowIneligible','BtnSelectAll','BtnSelectNone','BtnReloadGroups',
                     'StatusText','BtnCopy','BtnClose',
                     'TxtLog','ChkVerbose','ChkAutoScroll','BtnCopyLog','BtnSaveLog','BtnClearLog',
                     'ConsolePanel')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the Copy Device Groups XAML.",
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
        Source       = $null
        Target       = $null
        TargetDevice = $null
        Rows         = @()
        Busy         = $false
    }

    $targets    = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $groupRows  = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $ui.GridTargets.ItemsSource = $targets
    $ui.GridGroups.ItemsSource  = $groupRows

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # --- diagnostics console ----------------------------------------------
    # The sink is what Write-CdgLog calls. Appending to a TextBox is cheap
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
    $Script:CdgLogSink    = $appendLog
    $Script:CdgLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
        $win.Height = 1040
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy   = $Busy
        $win.Cursor   = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnCopy, $ui.BtnFindTarget, $ui.BtnReloadGroups,
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
        $ui.GroupCount.Text = "$eligible of $total group(s) eligible to copy"
    }

    $loadGroups = {
        & $setBusy $true
        try {
            & $setStatus "Reading Entra ID device record for $($SourceDevice.DeviceName)..."
            $state.Source = Get-CdgEntraDevice -Device $SourceDevice

            $ui.TxtSourceDevice.Text = $state.Source.DeviceName
            $ui.TxtSourceDetail.Text = "Serial: $(Get-CdgProperty $SourceDevice 'SerialNumber' 'n/a')`nUser: $(Get-CdgProperty $SourceDevice 'User' 'n/a')`nEntra object: $($state.Source.ObjectId)"

            & $setStatus 'Reading group memberships...'
            $state.Rows = @(Get-CdgGroupRows -EntraObjectId $state.Source.ObjectId)
            & $refreshGrid

            $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
            & $setStatus "Loaded $(@($state.Rows).Count) group(s) for $($state.Source.DeviceName) - $eligible can be copied."
        }
        catch {
            $detail = Write-CdgLogException $_ 'Loading groups'
            & $setStatus "Could not load groups: $detail"
            [System.Windows.MessageBox]::Show($detail,'Copy Device Groups','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- target device lookup ----------------------------------------------
    $findTargets = {
        $term = $ui.TxtTargetSearch.Text
        if ([string]::IsNullOrWhiteSpace($term)) {
            & $setStatus 'Enter a device name, serial number or user to find the target device.'
            return
        }
        $term = $term.Trim()

        & $setBusy $true
        try {
            $hits = @()
            if (@($DeviceCache).Count -gt 0) {
                $hits = @($DeviceCache | Where-Object {
                    ((Get-CdgProperty $_ 'DeviceName'   '') -like "*$term*") -or
                    ((Get-CdgProperty $_ 'SerialNumber' '') -like "*$term*") -or
                    ((Get-CdgProperty $_ 'User'         '') -like "*$term*")
                } | Sort-Object DeviceName | Select-Object -First 200)
            }
            else {
                # No cache handed in - exact-match query straight against Graph.
                & $setStatus 'No device cache available, querying Graph...'
                $esc = $term.Replace("'", "''")
                $filter = "deviceName eq '$esc' or serialNumber eq '$esc' or userPrincipalName eq '$esc'"
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=$filter&`$select=id,deviceName,serialNumber,userPrincipalName,azureADDeviceId"
                $hits = @(Get-CdgGraphPaged -Uri $uri | ForEach-Object {
                    [pscustomobject]@{
                        DeviceName    = Get-CdgProperty $_ 'deviceName' ''
                        SerialNumber  = Get-CdgProperty $_ 'serialNumber' ''
                        User          = Get-CdgProperty $_ 'userPrincipalName' ''
                        EntraDeviceId = Get-CdgProperty $_ 'azureADDeviceId' ''
                    }
                })
            }

            # Never let the source device be its own target.
            $srcName = Get-CdgProperty $SourceDevice 'DeviceName' ''
            $hits = @($hits | Where-Object { (Get-CdgProperty $_ 'DeviceName' '') -ne $srcName })

            $targets.Clear()
            foreach ($h in $hits) { $targets.Add($h) }
            & $setStatus "$($targets.Count) matching device(s) - pick the target device in the list."
        }
        catch {
            & $setStatus "Target lookup failed: $(Write-CdgLogException $_ 'Target lookup')"
        }
        finally { & $setBusy $false }
    }

    # --- copy --------------------------------------------------------------
    $doCopy = {
        if ($null -eq $state.Target) {
            [System.Windows.MessageBox]::Show('Select a target device first.',
                'Copy Device Groups','OK','Warning') | Out-Null
            return
        }

        $picked = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected })
        if ($picked.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Tick at least one eligible group to copy.',
                'Copy Device Groups','OK','Warning') | Out-Null
            return
        }

        $confirm = [System.Windows.MessageBox]::Show(
            "Add $($picked.Count) group(s) from`n  $($state.Source.DeviceName)`nto`n  $($state.Target.DisplayName)?`n`nExisting memberships on the target device are left alone - nothing is removed.",
            'Confirm copy','YesNo','Question')
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

        & $setBusy $true
        $added = 0; $existing = 0; $failed = 0
        $lines = New-Object System.Collections.Generic.List[string]
        $i = 0

        foreach ($row in $picked) {
            $i++
            & $setStatus "Copying $i of $($picked.Count): $($row.DisplayName)..."
            try {
                $result = Add-CdgDeviceToGroup -GroupId $row.GroupId -DeviceObjectId $state.Target.ObjectId
                Write-CdgLog "$($row.DisplayName) - $result" 'OK'
                if ($result -eq 'Added') { $added++ } else { $existing++ }
                $lines.Add("$($row.DisplayName) - $result") | Out-Null
            }
            catch {
                $failed++
                $detail = Write-CdgLogException $_ "Add '$($row.DisplayName)'"
                $lines.Add("$($row.DisplayName) - FAILED: $detail") | Out-Null
            }
        }

        & $setBusy $false
        $summary = "Copy complete - $added added, $existing already a member, $failed failed."
        & $setStatus $summary
        Write-CdgLog $summary $(if ($failed -gt 0) { 'WARN' } else { 'OK' })

        $icon = if ($failed -gt 0) { 'Warning' } else { 'Information' }
        [System.Windows.MessageBox]::Show(
            "$summary`n`n$($lines -join "`n")",
            'Copy Device Groups','OK',$icon) | Out-Null
    }

    # --- events ------------------------------------------------------------
    $ui.BtnFindTarget.Add_Click({ & $findTargets })
    $ui.TxtTargetSearch.Add_KeyDown({ if ($args[1].Key -eq 'Return') { & $findTargets } })

    $ui.GridTargets.Add_SelectionChanged({
        $d = $ui.GridTargets.SelectedItem
        if (-not $d) { return }
        try {
            $state.TargetDevice = $d
            $state.Target       = Get-CdgEntraDevice -Device $d
            $ui.TxtTargetSelected.Text = "Target: $($state.Target.DisplayName)  |  Serial: $(Get-CdgProperty $d 'SerialNumber' 'n/a')  |  Entra object: $($state.Target.ObjectId)"
            $ui.TxtTargetSelected.Foreground = $win.TryFindResource('AccentTextBrush')
            & $setStatus "Target set to $($state.Target.DisplayName)."
        }
        catch {
            $state.Target = $null
            $ui.TxtTargetSelected.Text = $_.Exception.Message
            $ui.TxtTargetSelected.Foreground = $win.TryFindResource('ConnBadBrush')
        }
    })

    $ui.ChkShowIneligible.Add_Checked({   & $refreshGrid })
    $ui.ChkShowIneligible.Add_Unchecked({ & $refreshGrid })

    $ui.BtnSelectAll.Add_Click({
        foreach ($r in $state.Rows) { if ($r.IsEligible) { $r.Selected = $true } }
    })
    $ui.BtnSelectNone.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $false }
    })

    $ui.ChkVerbose.Add_Checked({
        $Script:CdgLogVerbose = $true
        Write-CdgLog 'Verbose logging on - per-group decisions and Graph timings will be shown.' 'INFO'
    })
    $ui.ChkVerbose.Add_Unchecked({
        $Script:CdgLogVerbose = $false
        Write-CdgLog 'Verbose logging off.' 'INFO'
    })

    $ui.BtnClearLog.Add_Click({
        $ui.TxtLog.Clear()
        Write-CdgLog 'Log cleared.' 'INFO'
    })

    $ui.BtnCopyLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
        & $setStatus 'Diagnostics log copied to the clipboard.'
    })

    $ui.BtnSaveLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $file  = Join-Path ([Environment]::GetFolderPath('Desktop')) "CopyDeviceGroups-$stamp.log"
        try {
            $ui.TxtLog.Text | Set-Content -Path $file -Encoding UTF8
            & $setStatus "Log saved to $file"
            Write-CdgLog "Log saved to $file" 'OK'
        }
        catch { Write-CdgLog "Could not save the log: $($_.Exception.Message)" 'ERROR' }
    })

    $ui.BtnReloadGroups.Add_Click({ & $loadGroups })
    $ui.BtnCopy.Add_Click({        & $doCopy })
    $ui.BtnClose.Add_Click({       $win.Close() })

    # Drop the sink so a closed window is never written to.
    $win.Add_Closed({ $Script:CdgLogSink = $null })

    # F12 toggles the console without reopening the window.
    $win.Add_KeyDown({
        if ($args[1].Key -ne 'F12') { return }
        if ($ui.ConsolePanel.Visibility -eq [System.Windows.Visibility]::Visible) {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Collapsed
            $win.Height = 860
        }
        else {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
            $win.Height = 1040
        }
    })

    # --- go ----------------------------------------------------------------
    $ui.TxtSourceDevice.Text = Get-CdgProperty $SourceDevice 'DeviceName' '(unknown device)'
    Write-CdgLog 'Copy Device Groups opened.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-CdgLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-CdgLog "Source device: $(Get-CdgProperty $SourceDevice 'DeviceName' '(unknown)')" 'INFO'
    Write-CdgLog "Device cache: $(@($DeviceCache).Count) entry(ies)" 'DEBUG'
    $win.Add_ContentRendered({ & $loadGroups })
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# prompt for a source device by name so the module can be tested on its own.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'Device.Read.All','Group.Read.All','Group.ReadWrite.All',
                            'GroupMember.ReadWrite.All','DeviceManagementManagedDevices.Read.All' -NoWelcome

    $name = Read-Host 'Source device name'
    $esc  = $name.Replace("'", "''")
    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$esc'&`$select=id,deviceName,serialNumber,userPrincipalName,azureADDeviceId"
    $src  = @(Get-CdgGraphPaged -Uri $uri | ForEach-Object {
        [pscustomobject]@{
            DeviceName    = Get-CdgProperty $_ 'deviceName' ''
            SerialNumber  = Get-CdgProperty $_ 'serialNumber' ''
            User          = Get-CdgProperty $_ 'userPrincipalName' ''
            EntraDeviceId = Get-CdgProperty $_ 'azureADDeviceId' ''
        }
    })

    if ($src.Count -eq 0) { Write-Warning "No managed device named '$name' was found."; return }
    Show-CopyDeviceGroupsWindow -SourceDevice $src[0] -XamlPath $XamlPath -ShowConsole
}
