<#
.SYNOPSIS
    Bulk Add to Group - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that adds many devices, listed in a CSV, to one Entra ID
    security group.

      - The operator browses for a CSV. Every line of the file is read -
        nothing is skipped - and each one supplies one device name, taken
        from the first column.
      - Each name is matched against the Intune managed devices (the Toolkit
        device cache is reused when it has been loaded, otherwise the list is
        read from Graph once).
      - Every name is listed with its match state, so a typo or a retired
        machine is visible before anything is written. Only rows that matched
        exactly one device that has an Entra ID device record are ticked.
      - The operator searches for the target group and picks it from a grid.
        Only assigned (static) security groups can be written to; dynamic
        groups are listed but cannot be selected because their membership is
        evaluated by rule.
      - A confirmation prompt naming the group and the device count is shown
        before anything is written, and it defaults to No.
      - Each device is added one at a time and its outcome is written back
        into the Result column, with a summary at the end. The results can be
        exported to CSV.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so BulkAddToGroup.xaml
    is not required at runtime and this file can be compiled with PS2EXE
    alongside Toolkit.ps1. Pass -XamlPath to load an external
    BulkAddToGroup.xaml instead while iterating on the UI.

    Every function here is Bag-prefixed on purpose. This module,
    CopyDeviceGroups.ps1 and RemoveDeviceGroups.ps1 are dot-sourced into the
    SAME session by Toolkit.ps1, so a shared name would mean the last file
    loaded silently wins and could change the behaviour of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\BulkAddToGroup\BulkAddToGroup.ps1"
        Show-BulkAddToGroupWindow -DeviceCache $Script:DeviceCache `
                                  -Owner       $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: DeviceManagementManagedDevices.Read.All, Device.Read.All,
                  Group.Read.All, Group.ReadWrite.All,
                  GroupMember.ReadWrite.All
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external BulkAddToGroup.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# Row types for the two grids.
#
# Separate types from the Copy and Remove modules on purpose: Add-Type cannot
# redefine a type that is already loaded in the session, so reusing a name
# would leave whichever module loaded second without its own properties.
#
# Selected and Result raise PropertyChanged so the grid repaints when
# Select all / Select none is used and while the add loop runs.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyBulkDeviceRow' -as [type])) {
    Add-Type -ReferencedAssemblies System.ComponentModel.TypeConverter -TypeDefinition @'
using System.ComponentModel;

public class EndpointguyBulkDeviceRow : INotifyPropertyChanged
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

    public bool   IsEligible    { get; set; }
    public string CsvName       { get; set; }
    public string DeviceName    { get; set; }
    public string Status        { get; set; }
    public string SerialNumber  { get; set; }
    public string User          { get; set; }
    public string EntraDeviceId { get; set; }
    public string EntraObjectId { get; set; }

    public event PropertyChangedEventHandler PropertyChanged;

    private void OnPropertyChanged(string name)
    {
        PropertyChangedEventHandler handler = PropertyChanged;
        if (handler != null) { handler(this, new PropertyChangedEventArgs(name)); }
    }
}
'@
}

if (-not ('EndpointguyBulkGroupRow' -as [type])) {
    Add-Type -TypeDefinition @'
public class EndpointguyBulkGroupRow
{
    public bool   IsEligible  { get; set; }
    public string DisplayName { get; set; }
    public string GroupType   { get; set; }
    public string Status      { get; set; }
    public string GroupId     { get; set; }
}
'@
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$BagXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bulk Add to Group"
        Height="900" Width="1250"
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
                <TextBlock Text="Bulk Add to Group" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Reads device names from the first column of a CSV - every line of the file is read - matches each one to an Intune managed device, and adds every matched device to a single assigned (static) security group. Dynamic groups are listed but cannot be written to because their membership is evaluated by rule."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- 1. Device list -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Text="1. Device list (CSV)" Style="{StaticResource CardHeader}"/>
                    <TextBlock x:Name="TxtCsvSummary" Text="No file loaded" Margin="12,0,0,0"
                               VerticalAlignment="Bottom" FontSize="13"
                               Foreground="{StaticResource SubtleTextBrush}"/>
                </StackPanel>

                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtCsvPath" Grid.Column="0" IsReadOnly="True" Text=""
                             ToolTip="The CSV currently loaded."/>
                    <Button x:Name="BtnBrowseCsv" Grid.Column="1" Style="{StaticResource PrimaryButton}"
                            Content="Browse for CSV..." MinWidth="180" Margin="10,0,0,0"/>
                </Grid>

                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
                    <TextBlock Text="Device names are read from the first column. Every line of the file is read."
                               VerticalAlignment="Center" FontSize="12.5"
                               Foreground="{StaticResource SubtleTextBrush}" Margin="0,0,16,0"/>
                    <Button x:Name="BtnReloadCsv" Style="{StaticResource NeutralButton}" Content="Reload file"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- 2. Target group -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Text="2. Target group" Style="{StaticResource CardHeader}"/>
                    <TextBlock x:Name="TxtSelectedGroup" Text="No group selected" Margin="12,0,0,0"
                               VerticalAlignment="Bottom" FontSize="13"
                               Foreground="{StaticResource SubtleTextBrush}"/>
                </StackPanel>

                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtGroupSearch" Grid.Column="0"
                             ToolTip="Type part of the group name and press Enter."/>
                    <Button x:Name="BtnSearchGroups" Grid.Column="1" Style="{StaticResource PrimaryButton}"
                            Content="Search groups" MinWidth="180" Margin="10,0,0,0"/>
                </Grid>

                <DataGrid x:Name="GridGroups" Grid.Row="2" Height="150" Margin="0,12,0,0"
                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          IsReadOnly="True"
                          SelectionMode="Single"
                          SelectionUnit="FullRow"
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
                        <DataGridTextColumn Header="Group Name"  Binding="{Binding DisplayName}" Width="330"/>
                        <DataGridTextColumn Header="Type"        Binding="{Binding GroupType}"   Width="170"/>
                        <DataGridTextColumn Header="Eligibility" Binding="{Binding Status}"      Width="330"/>
                        <DataGridTextColumn Header="Group Id"    Binding="{Binding GroupId}"     Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>

        <!-- 3. Devices -->
        <Border Grid.Row="3" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="3. Devices from the file" Style="{StaticResource CardHeader}"/>
                        <TextBlock x:Name="DeviceCount" Text="" Margin="12,0,0,0"
                                   VerticalAlignment="Bottom" FontSize="13"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <CheckBox x:Name="ChkShowUnmatched" Content="Show unmatched devices" IsChecked="True"
                                  VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"/>
                        <Button x:Name="BtnSelectAll"  Style="{StaticResource NeutralButton}" Content="Select all"  Margin="0,0,10,0"/>
                        <Button x:Name="BtnSelectNone" Style="{StaticResource NeutralButton}" Content="Select none" Margin="0,0,10,0"/>
                        <Button x:Name="BtnExportResults" Style="{StaticResource NeutralButton}" Content="Export results"/>
                    </StackPanel>
                </Grid>

                <DataGrid x:Name="GridDevices" Grid.Row="1"
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
                        <DataGridCheckBoxColumn Header="Add" Width="60"
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
                        <DataGridTextColumn Header="Name in file" Binding="{Binding CsvName}"      Width="190" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}"    Width="190" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Match"       Binding="{Binding Status}"        Width="270" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Serial"      Binding="{Binding SerialNumber}"  Width="150" IsReadOnly="True"/>
                        <DataGridTextColumn Header="User"        Binding="{Binding User}"          Width="190" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Result"      Binding="{Binding Result}"        Width="180" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Entra Object Id" Binding="{Binding EntraObjectId}" Width="*" IsReadOnly="True"/>
                    </DataGrid.Columns>
                </DataGrid>
            </Grid>
        </Border>

        <!-- Footer -->
        <Border Grid.Row="4" Style="{StaticResource Card}" Padding="14,10" Margin="0,14,0,0">
            <Grid>
                <TextBlock x:Name="StatusText" Text="Ready."
                           FontSize="12.5" VerticalAlignment="Center" TextWrapping="NoWrap"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnAdd" Style="{StaticResource GreenButton}"
                            Content="Add devices to group" MinWidth="220" Height="38" Margin="0,0,12,0"/>
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
# There is no on-screen console, so the log lines go to Write-Verbose. The
# sink is left in place - and stays null unless something sets it - so the
# module still behaves when dot-sourced or run headless.
# ---------------------------------------------------------------------------
$Script:BagLogSink    = $null
$Script:BagLogVerbose = $false

function Write-BagLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG','CMD')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:BagLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:BagLogSink) {
        # Out-Null is essential: the sink is a WPF scriptblock and
        # Dispatcher.Invoke returns a value. Without this, that value becomes
        # the OUTPUT of Write-BagLog, and every function that logs then
        # returns extra junk on its pipeline. A helper returning a hashtable
        # becomes an Object[] in the caller, and the first string index into
        # it fails with "Argument types do not match".
        try { & $Script:BagLogSink $line | Out-Null } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-BagLogException {
    <#
        Turns an ErrorRecord into a readable console entry.

        Everything that identifies a type-binding failure - exception type,
        the inner exception chain, the fully qualified error id and the
        script stack trace - is written EVERY time, not only when Verbose is
        ticked. A message like "argument types do not match" says nothing on
        its own; the stack trace is what names the call that actually failed.
    #>
    param($ErrorRecord, [string]$Context = 'Operation')

    $detail = Get-BagGraphError $ErrorRecord
    Write-BagLog "$Context failed: $detail" 'ERROR'

    if (-not $ErrorRecord) { return $detail }

    $ex = Get-BagProperty $ErrorRecord 'Exception'
    if ($ex) {
        Write-BagLog "  exception  : $($ex.GetType().FullName)" 'ERROR'

        # Walk the whole inner chain - a reflection binding failure keeps the
        # useful message two or three levels down.
        $inner = Get-BagProperty $ex 'InnerException'
        $depth = 1
        while ($inner -and $depth -le 5) {
            Write-BagLog "  inner[$depth]  : $($inner.GetType().FullName): $($inner.Message)" 'ERROR'
            $inner = Get-BagProperty $inner 'InnerException'
            $depth++
        }
    }

    $fqid = [string](Get-BagProperty $ErrorRecord 'FullyQualifiedErrorId' '')
    if ($fqid) { Write-BagLog "  errorId    : $fqid" 'ERROR' }

    $cat = Get-BagProperty $ErrorRecord 'CategoryInfo'
    if ($cat) { Write-BagLog "  category   : $cat" 'ERROR' }

    $pos = Get-BagProperty $ErrorRecord 'InvocationInfo'
    if ($pos) {
        $line = ([string](Get-BagProperty $pos 'Line' '')).Trim()
        $no   = Get-BagProperty $pos 'ScriptLineNumber' 0
        if ($line) { Write-BagLog "  at line $no : $line" 'ERROR' }
    }

    $trace = [string](Get-BagProperty $ErrorRecord 'ScriptStackTrace' '')
    if ($trace) {
        foreach ($l in ($trace -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($l)) { Write-BagLog "  stack      : $l" 'ERROR' }
        }
    }

    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-BagProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-BagProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-BagProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-BagGraphPaged {
    # Walks @odata.nextLink and returns every item in the 'value' array.
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($Uri) {
        $page++
        Write-BagLog "GET $Uri" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        $sw.Stop()

        $count = 0
        if (Test-BagProperty $resp 'value') { $count = @($resp.value).Count; $all.AddRange(@($resp.value)) }
        Write-BagLog "  -> page $page returned $count item(s) in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        $Uri = if (Test-BagProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        if ($Uri) { Write-BagLog '  -> following @odata.nextLink for the next page' 'DEBUG' }
    }
    return $all
}

function Get-BagBodyError {
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

    if (Test-BagProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-BagProperty $err 'code' ''
        $msg  = Get-BagProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-BagGraphError {
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
    if (Test-BagProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-BagProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-BagBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-BagProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-BagProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-BagProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-BagBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-BagProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-BagProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-BagProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
function Import-BagDeviceNameFile {
    <#
        Reads the device list: one device name per line, taken from the first
        column. EVERY line of the file is read - no line is ever skipped.

        There is no spacer row and no header row any more: line 1 is treated
        exactly like every other line, so a device name sitting on it is
        imported. If a file does arrive with a heading such as 'DeviceName'
        on top, that heading is read as a name and reported as not matched,
        which is visible in the grid rather than silently applied.

        Blank lines are ignored wherever they appear.

        Parsing goes through ConvertFrom-Csv with a single column header, so a
        quoted name containing a comma survives and any extra columns to the
        right are ignored.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "The file '$Path' could not be found."
    }

    $raw = @(Get-Content -LiteralPath $Path)
    if ($raw.Count -eq 0) { throw "'$Path' is empty." }

    $dataLines = @($raw | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $blank     = $raw.Count - $dataLines.Count

    # A HashSet keyed on the lower-cased name.
    #
    # Deliberately the DEFAULT constructor: passing [StringComparer] to
    # New-Object makes PowerShell pick a constructor overload by reflection,
    # which is one of the ways this module produced "Argument types do not
    # match". Lower-casing the key gives the same case-insensitive behaviour
    # with nothing for the binder to resolve.
    $seen  = New-Object System.Collections.Generic.HashSet[string]
    $names = New-Object System.Collections.Generic.List[string]
    $dupes = 0

    foreach ($row in @($dataLines | ConvertFrom-Csv -Header 'DeviceName')) {
        $value = [string](Get-BagProperty $row 'DeviceName' '')
        $value = $value.Trim()
        if ([string]::IsNullOrWhiteSpace($value)) { $blank++; continue }
        # HashSet.Add returns false when the value was already present, so this
        # is the de-duplication test and the insert in one typed call.
        if (-not $seen.Add([string]$value.ToLowerInvariant())) { $dupes++; continue }
        $names.Add($value) | Out-Null
    }

    if ($blank -gt 0) { Write-BagLog "Ignored $blank blank line(s)." 'INFO' }
    if ($dupes -gt 0) { Write-BagLog "Ignored $dupes duplicate name(s)." 'INFO' }

    if ($names.Count -eq 0) {
        throw "No device names were found in the first column of '$Path'."
    }

    Write-BagLog "Read $($names.Count) device name(s) from $Path." 'OK'

    [pscustomobject]@{
        Names          = $names
        BlankCount     = $blank
        DuplicateCount = $dupes
    }
}

# ---------------------------------------------------------------------------
# Groups
# ---------------------------------------------------------------------------
function Get-BagGroupSearchRows {
    <#
        Searches groups by name and returns one EndpointguyBulkGroupRow each.

        Eligible   = assigned (static) security group that Graph will accept a
                     device member being added to.
        Ineligible = dynamic, on-premises synced, mail-enabled, Microsoft 365.

        A pasted object id is looked up directly, so a group found elsewhere
        in the portal can be used without knowing its exact name.
    #>
    param([Parameter(Mandatory)][string]$Term)

    $Term = $Term.Trim()
    if ([string]::IsNullOrWhiteSpace($Term)) { throw 'Type part of a group name (or a group object id) to search for.' }

    $select = 'id,displayName,groupTypes,securityEnabled,mailEnabled,' +
              'membershipRule,membershipRuleProcessingState,onPremisesSyncEnabled'

    $found = @()
    $guid  = [ref][guid]::Empty
    if ([guid]::TryParse($Term, $guid)) {
        Write-BagLog "'$Term' looks like an object id - reading the group directly." 'INFO'
        $found = @(Invoke-MgGraphRequest -Method GET -OutputType PSObject `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$Term`?`$select=$select")
    }
    else {
        $esc = $Term.Replace("'", "''")
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=startswith(displayName,'$esc')&`$select=$select&`$top=999"
        $found = @(Get-BagGraphPaged -Uri $uri)
    }

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($g in $found) {
        $name        = Get-BagProperty $g 'displayName' '(unnamed)'
        $groupTypes  = @(Get-BagProperty $g 'groupTypes' @())
        $isDynamic   = ($groupTypes -contains 'DynamicMembership')
        $isUnified   = ($groupTypes -contains 'Unified')
        $secEnabled  = [bool](Get-BagProperty $g 'securityEnabled' $false)
        $mailEnabled = [bool](Get-BagProperty $g 'mailEnabled' $false)
        $onPrem      = [bool](Get-BagProperty $g 'onPremisesSyncEnabled' $false)

        # A membership rule is the definitive tell: some tenants return an
        # empty groupTypes on a group that is still rule-driven.
        $rule = Get-BagProperty $g 'membershipRule' ''
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
            $status = 'Dynamic - membership is set by rule, cannot be written to'
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

        $row = New-Object EndpointguyBulkGroupRow
        $row.DisplayName = $name
        $row.GroupType   = $type
        $row.Status      = $status
        $row.GroupId     = Get-BagProperty $g 'id' ''
        $row.IsEligible  = $eligible

        $verdict = if ($eligible) { 'ELIGIBLE' } else { 'skipped ' }
        Write-BagLog "  $verdict | $name | $type | $status" 'DEBUG'
        $rows.Add($row) | Out-Null
    }

    $elig = @($rows | Where-Object { $_.IsEligible }).Count
    Write-BagLog "Group search for '$Term' returned $($rows.Count) group(s): $elig eligible." 'OK'

    return ($rows | Sort-Object -Property @{Expression={ -not $_.IsEligible }}, DisplayName)
}

# ---------------------------------------------------------------------------
# Devices
# ---------------------------------------------------------------------------
function Get-BagManagedDeviceLookup {
    <#
        Builds a device-name -> managed device(s) lookup. The Toolkit cache is
        reused when it has been loaded so a bulk run costs no extra Graph
        calls; otherwise the managed device list is read once and paged.

        The value is a list, not a single device, so a duplicated device name
        can be reported rather than silently resolving to the wrong machine.
    #>
    param($DeviceCache)

    $devices = @()
    if ($DeviceCache -and @($DeviceCache).Count -gt 0) {
        $devices = @($DeviceCache)
        Write-BagLog "Using the Toolkit device cache - $($devices.Count) managed device(s)." 'INFO'
    }
    else {
        Write-BagLog 'No device cache was supplied - reading managed devices from Graph.' 'INFO'
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,serialNumber,userPrincipalName,azureADDeviceId&`$top=999"
        $devices = @(Get-BagGraphPaged -Uri $uri | ForEach-Object {
            [pscustomobject]@{
                DeviceName    = Get-BagProperty $_ 'deviceName' ''
                SerialNumber  = Get-BagProperty $_ 'serialNumber' ''
                User          = Get-BagProperty $_ 'userPrincipalName' ''
                EntraDeviceId = Get-BagProperty $_ 'azureADDeviceId' ''
            }
        })
        Write-BagLog "Read $($devices.Count) managed device(s) from Graph." 'OK'
    }

    # A generic Dictionary, NOT a hashtable.
    #
    # Reading a hashtable as $h[$k] does not call Hashtable.get_Item directly:
    # PowerShell routes it through its parameterized-property binder, which
    # picks an indexer overload by reflection. On a large hashtable reached
    # from a typed variable that resolution can fail outright with the
    # unhelpful "Argument types do not match" - which is exactly the error
    # this module kept dying on.
    #
    # Dictionary.TryGetValue and .Add are plain strongly typed calls with one
    # overload each, so no binder and no reflection is involved. Keys are
    # already lower-cased here, so the default ordinal comparer is correct.
    $lookup = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($d in $devices) {
        $name = [string](Get-BagProperty $d 'DeviceName' '')
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = [string]$name.Trim().ToLowerInvariant()

        $bucket = $null
        if (-not $lookup.TryGetValue($key, [ref]$bucket)) {
            $bucket = New-Object System.Collections.Generic.List[object]
            $lookup.Add($key, $bucket)
        }
        $bucket.Add($d) | Out-Null
    }
    # Comma keeps this a single hashtable on the pipeline rather than letting
    # PowerShell enumerate it into loose DictionaryEntry objects.
    return ,$lookup
}

function Get-BagEntraDeviceMap {
    <#
        Resolves Entra deviceId -> directory object in as few calls as
        possible. Group membership is written against the directory object id,
        not the Intune device id, so this hop cannot be skipped - but the ids
        are batched into OR filters rather than one request per device.
    #>
    param([string[]]$DeviceIds)

    $map = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    $ids = @($DeviceIds |
             Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and
                            $_ -ne '00000000-0000-0000-0000-000000000000' } |
             Select-Object -Unique)
    if ($ids.Count -eq 0) { return ,$map }

    $chunkSize = 15
    Write-BagLog "Resolving $($ids.Count) Entra ID device object(s) in batches of $chunkSize." 'INFO'

    for ($i = 0; $i -lt $ids.Count; $i += $chunkSize) {
        $last  = [Math]::Min($i + $chunkSize - 1, $ids.Count - 1)
        $slice = @($ids[$i..$last])
        $filter = ($slice | ForEach-Object { "deviceId eq '$_'" }) -join ' or '
        $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=$filter&`$select=id,displayName,deviceId&`$top=999"

        foreach ($d in (Get-BagGraphPaged -Uri $uri)) {
            $key = [string](Get-BagProperty $d 'deviceId' '')
            if ($key -and -not $map.ContainsKey($key)) { $map.Add($key, $d) }
        }
    }

    Write-BagLog "Resolved $($map.Count) of $($ids.Count) Entra ID device object(s)." 'OK'
    # Comma keeps this a single hashtable on the pipeline rather than letting
    # PowerShell enumerate it into loose DictionaryEntry objects.
    return ,$map
}

function Resolve-BagDeviceRows {
    <#
        Turns the names read from the CSV into grid rows.

        Only a row that matched exactly one managed device AND has an Entra ID
        device object is eligible, because those are the only ones Graph will
        accept as a group member.
    #>
    param(
        [Parameter(Mandatory)][string[]]$Names,
        $DeviceCache
    )

    [System.Collections.Generic.Dictionary[string,object]]$lookup = Get-BagManagedDeviceLookup -DeviceCache $DeviceCache
    $rows   = New-Object System.Collections.Generic.List[object]

    foreach ($n in $Names) {
        $row = New-Object EndpointguyBulkDeviceRow
        $row.CsvName    = $n
        $row.Result     = ''
        $row.IsEligible = $false
        $row.Selected   = $false

        $key = [string]$n.Trim().ToLowerInvariant()

        # NOTE the shape here. Writing
        #     $hits = if (...) { @($bucket) } else { @() }
        # looks equivalent but is not: an `if` that yields @() emits NOTHING to
        # the pipeline, so $hits becomes $null rather than an empty array.
        # Toolkit.ps1 sets Set-StrictMode -Version Latest, which turns the very
        # next $hits.Count into a terminating PropertyNotFoundException.
        #
        # Starting from a real empty array and only overwriting it on a hit
        # keeps $hits an array on every path.
        $bucket = $null
        $hits   = @()
        if ($lookup.TryGetValue($key, [ref]$bucket)) { $hits = @($bucket) }

        if ($hits.Count -eq 0) {
            $row.DeviceName = ''
            $row.Status     = 'Not found in Intune'
        }
        elseif ($hits.Count -gt 1) {
            $row.DeviceName = Get-BagProperty $hits[0] 'DeviceName' ''
            $row.Status     = "Ambiguous - $($hits.Count) managed devices share this name"
        }
        else {
            $d = $hits[0]
            $row.DeviceName    = Get-BagProperty $d 'DeviceName' ''
            $row.SerialNumber  = [string](Get-BagProperty $d 'SerialNumber' '')
            $row.User          = [string](Get-BagProperty $d 'User' '')
            $row.EntraDeviceId = [string](Get-BagProperty $d 'EntraDeviceId' '')
            $row.Status        = 'Matched'
            $row.IsEligible    = $true
        }

        Write-BagLog "  $($row.CsvName) -> $($row.Status)" 'DEBUG'
        $rows.Add($row) | Out-Null
    }

    # Resolve the directory object for everything that matched.
    $matched = @($rows | Where-Object { $_.IsEligible })
    if ($matched.Count -gt 0) {
        # Generic Dictionary, not a hashtable - see Get-BagManagedDeviceLookup
        # for why the index operator is avoided on this path.
        [System.Collections.Generic.Dictionary[string,object]]$map = Get-BagEntraDeviceMap -DeviceIds @($matched | ForEach-Object { [string]$_.EntraDeviceId })

        foreach ($row in $matched) {
            if ([string]::IsNullOrWhiteSpace($row.EntraDeviceId) -or
                $row.EntraDeviceId -eq '00000000-0000-0000-0000-000000000000') {
                $row.IsEligible = $false
                $row.Status     = 'No Entra ID device record - Intune only'
                continue
            }
            # TryGetValue, not ContainsKey + [] - a strongly typed method call
            # with exactly one overload, so PowerShell never has to resolve an
            # indexer by reflection.
            $entra = $null
            if (-not $map.TryGetValue([string]$row.EntraDeviceId, [ref]$entra)) {
                $row.IsEligible = $false
                $row.Status     = 'No Entra ID device object found'
                continue
            }
            $row.EntraObjectId = [string](Get-BagProperty $entra 'id' '')
            $row.Status        = 'Matched - ready to add'
            $row.Selected      = $true   # pre-tick everything that can be added
        }
    }

    $elig = @($rows | Where-Object { $_.IsEligible }).Count
    Write-BagLog "Matched $elig of $($rows.Count) name(s) from the file." 'OK'

    return ($rows | Sort-Object -Property @{Expression={ -not $_.IsEligible }}, CsvName)
}

function Add-BagDeviceToGroup {
    <#
        Adds the device directory object to a group. Returns 'Added' or
        'Already a member', or throws with the real Graph message.

        POST /groups/<id>/members/$ref with an @odata.id body is the supported
        way to add a single member; it returns 204 on success and 400 with
        'already exist' when the device is in the group already, which is not
        worth showing the operator as a failure.
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][string]$DeviceObjectId
    )

    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw 'The group has no id, so it cannot be updated.'
    }
    if ([string]::IsNullOrWhiteSpace($DeviceObjectId)) {
        throw 'The device has no Entra ID object id, so it cannot be added to a group.'
    }

    Write-BagLog "POST group $GroupId <- device object $DeviceObjectId" 'GRAPH'
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

    Write-BagLog "  -> using $(if ($canSkipThrow) { '-SkipHttpErrorCheck (full error body)' } else { 'try/catch (legacy Graph module)' })" 'DEBUG'

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body `
                    -ContentType 'application/json' -OutputType PSObject `
                    -SkipHttpErrorCheck -StatusCodeVariable 'code'

        Write-BagLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) { Write-BagLog '  -> added' 'OK'; return 'Added' }

        $detail = Get-BagBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        Write-BagLog "  -> Graph said: $detail" 'DEBUG'

        if ($detail -match 'already exist') { return 'Already a member' }
        throw "HTTP $code - $detail"
    }

    try {
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body `
            -ContentType 'application/json' | Out-Null
        return 'Added'
    }
    catch {
        $detail = Get-BagGraphError $_
        if ($detail -match 'already exist') { return 'Already a member' }
        throw $detail
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-BulkAddToGroupWindow {
    <#
    .SYNOPSIS
        Opens the Bulk Add to Group window.
    .PARAMETER DeviceCache
        The Toolkit device cache. Passing it means the CSV names are matched
        without any extra Graph calls; omit it and the list is read from Graph.
    .PARAMETER Owner
        The Toolkit window, so this one centres on it and stays on top.
    .PARAMETER CsvPath
        Optional CSV to load as soon as the window opens.
    #>
    [CmdletBinding()]
    param(
        $DeviceCache = @(),
        $Owner = $null,
        [string]$CsvPath,
        [string]$XamlPath
    )

    # --- load XAML ---------------------------------------------------------
    $xamlText = $BagXamlString
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
            "The Bulk Add to Group XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtCsvPath','BtnBrowseCsv','TxtCsvSummary','BtnReloadCsv',
                     'TxtGroupSearch','BtnSearchGroups','GridGroups','TxtSelectedGroup',
                     'GridDevices','DeviceCount','ChkShowUnmatched','BtnSelectAll','BtnSelectNone','BtnExportResults',
                     'StatusText','BtnAdd','BtnClose')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the Bulk Add to Group XAML.",
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
        Rows     = @()
        Group    = $null
        Busy     = $false
        Cache    = @($DeviceCache)
    }

    $deviceRows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $groupGrid  = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $ui.GridDevices.ItemsSource = $deviceRows
    $ui.GridGroups.ItemsSource  = $groupGrid

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render) | Out-Null
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnAdd, $ui.BtnBrowseCsv, $ui.BtnReloadCsv, $ui.BtnSearchGroups,
                         $ui.BtnSelectAll, $ui.BtnSelectNone)) { $b.IsEnabled = -not $Busy }
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render) | Out-Null
    }

    $refreshGrid = {
        # Rebuilds the visible rows, honouring the 'Show unmatched' toggle.
        $showAll = [bool]$ui.ChkShowUnmatched.IsChecked
        $deviceRows.Clear()
        foreach ($r in $state.Rows) {
            if ($showAll -or $r.IsEligible) { $deviceRows.Add($r) }
        }
        $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
        $total    = @($state.Rows).Count
        $ui.DeviceCount.Text = "$eligible of $total name(s) matched a device"
    }

    # --- csv ---------------------------------------------------------------
    $loadCsv = {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }

        & $setBusy $true
        try {
            & $setStatus "Reading $Path..."
            $file = Import-BagDeviceNameFile -Path $Path

            $ui.TxtCsvPath.Text = $Path

            $summary = "$(@($file.Names).Count) device name(s)"
            if ($file.DuplicateCount -gt 0) { $summary = "$summary, $($file.DuplicateCount) duplicate(s) ignored" }
            if ($file.BlankCount -gt 0)     { $summary = "$summary, $($file.BlankCount) blank line(s) ignored" }
            $ui.TxtCsvSummary.Text = $summary

            & $setStatus "Matching $(@($file.Names).Count) name(s) against Intune..."
            $state.Rows = @(Resolve-BagDeviceRows -Names @($file.Names) -DeviceCache $state.Cache)
            & $refreshGrid

            $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
            & $setStatus "$eligible of $(@($file.Names).Count) name(s) matched a device that can be added."
        }
        catch {
            $state.Rows = @()
            & $refreshGrid
            $ui.TxtCsvSummary.Text = 'No file loaded'
            $detail = Write-BagLogException $_ 'Reading the device list'
            & $setStatus "Could not read the file: $detail"
            [System.Windows.MessageBox]::Show($detail,'Bulk Add to Group','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- groups ------------------------------------------------------------
    $searchGroups = {
        $term = $ui.TxtGroupSearch.Text
        if ([string]::IsNullOrWhiteSpace($term)) {
            [System.Windows.MessageBox]::Show('Type part of a group name to search for.',
                'Bulk Add to Group','OK','Warning') | Out-Null
            return
        }

        & $setBusy $true
        try {
            & $setStatus "Searching for groups starting with '$term'..."
            $groupGrid.Clear()
            foreach ($g in @(Get-BagGroupSearchRows -Term $term)) { $groupGrid.Add($g) }

            $eligible = @($groupGrid | Where-Object { $_.IsEligible }).Count
            if ($groupGrid.Count -eq 0) {
                & $setStatus "No group name starts with '$term'."
            }
            else {
                & $setStatus "$($groupGrid.Count) group(s) found - $eligible can be written to. Pick one."
            }
        }
        catch {
            $detail = Write-BagLogException $_ 'Group search'
            & $setStatus "Group search failed: $detail"
            [System.Windows.MessageBox]::Show($detail,'Bulk Add to Group','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- add ---------------------------------------------------------------
    $doAdd = {
        if ($null -eq $state.Group) {
            [System.Windows.MessageBox]::Show('Search for a group and pick an eligible one first.',
                'Bulk Add to Group','OK','Warning') | Out-Null
            return
        }

        $picked = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected })
        if ($picked.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Tick at least one matched device to add.',
                'Bulk Add to Group','OK','Warning') | Out-Null
            return
        }

        $groupName = $state.Group.DisplayName

        # Writing to a group changes what policies and apps target these
        # machines, so the prompt names the group, states the count, and
        # defaults to No.
        $names = ($picked | Select-Object -First 15 | ForEach-Object { "  - $($_.DeviceName)" }) -join "`n"
        if ($picked.Count -gt 15) { $names = "$names`n  ... and $($picked.Count - 15) more" }

        $confirm = [System.Windows.MessageBox]::Show(
            "Add $($picked.Count) device(s) to '$groupName'?`n`n$names`n`nAny policies, apps or configuration targeted at that group will start applying to these devices.",
            'Confirm bulk add','YesNo','Warning',[System.Windows.MessageBoxResult]::No)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-BagLog 'Bulk add cancelled at the confirmation prompt.' 'INFO'
            & $setStatus 'Cancelled - nothing was added.'
            return
        }

        & $setBusy $true
        $added = 0; $already = 0; $failed = 0
        $lines = New-Object System.Collections.Generic.List[string]
        $i = 0

        foreach ($row in $picked) {
            $i++
            & $setStatus "Adding $i of $($picked.Count): $($row.DeviceName)..."
            $row.Result = 'Working...'
            try {
                $result = Add-BagDeviceToGroup -GroupId $state.Group.GroupId -DeviceObjectId $row.EntraObjectId
                $row.Result = $result
                Write-BagLog "$($row.DeviceName) - $result" 'OK'
                if ($result -eq 'Added') { $added++ } else { $already++ }
                $lines.Add("$($row.DeviceName) - $result") | Out-Null
            }
            catch {
                $failed++
                $detail = Write-BagLogException $_ "Add '$($row.DeviceName)'"
                $row.Result = "FAILED: $detail"
                $lines.Add("$($row.DeviceName) - FAILED: $detail") | Out-Null
            }
        }

        & $setBusy $false
        $summary = "Bulk add complete - $added added, $already already a member, $failed failed."
        & $setStatus $summary
        Write-BagLog $summary $(if ($failed -gt 0) { 'WARN' } else { 'OK' })

        $shown = @($lines | Select-Object -First 25) -join "`n"
        if ($lines.Count -gt 25) { $shown = "$shown`n... and $($lines.Count - 25) more - use Export results for the full list." }

        $icon = if ($failed -gt 0) { 'Warning' } else { 'Information' }
        [System.Windows.MessageBox]::Show(
            "$summary`n`n$shown",
            'Bulk Add to Group','OK',$icon) | Out-Null
    }

    # --- events ------------------------------------------------------------
    $ui.BtnBrowseCsv.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title  = 'Pick the CSV with the device names'
        $dlg.Filter = 'CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*'
        if ($dlg.ShowDialog()) { & $loadCsv $dlg.FileName }
    })

    $ui.BtnReloadCsv.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtCsvPath.Text)) {
            & $setStatus 'Browse for a CSV first.'
            return
        }
        & $loadCsv $ui.TxtCsvPath.Text
    })

    $ui.BtnSearchGroups.Add_Click({ & $searchGroups })
    $ui.TxtGroupSearch.Add_KeyDown({
        if ($args[1].Key -eq 'Return') { & $searchGroups }
    })

    $ui.GridGroups.Add_SelectionChanged({
        $g = $ui.GridGroups.SelectedItem
        if ($null -eq $g) { return }
        if (-not $g.IsEligible) {
            $state.Group = $null
            $ui.TxtSelectedGroup.Text = "'$($g.DisplayName)' cannot be written to - $($g.Status)"
            return
        }
        $state.Group = $g
        $ui.TxtSelectedGroup.Text = "Target: $($g.DisplayName)"
        & $setStatus "Target group set to '$($g.DisplayName)'."
    })

    $ui.ChkShowUnmatched.Add_Checked({   & $refreshGrid })
    $ui.ChkShowUnmatched.Add_Unchecked({ & $refreshGrid })

    $ui.BtnSelectAll.Add_Click({
        foreach ($r in $state.Rows) { if ($r.IsEligible) { $r.Selected = $true } }
    })
    $ui.BtnSelectNone.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $false }
    })

    $ui.BtnExportResults.Add_Click({
        if (@($state.Rows).Count -eq 0) { & $setStatus 'Nothing to export.'; return }
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Title      = 'Export the bulk add results'
        $dlg.Filter     = 'CSV files (*.csv)|*.csv'
        $dlg.FileName   = "BulkAddToGroup-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if (-not $dlg.ShowDialog()) { return }
        try {
            $state.Rows |
                Select-Object CsvName, DeviceName, Status, SerialNumber, User, Result, EntraObjectId |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            & $setStatus "Exported $(@($state.Rows).Count) row(s) to $($dlg.FileName)"
            Write-BagLog "Results exported to $($dlg.FileName)" 'OK'
        }
        catch { Write-BagLog "Could not export the results: $($_.Exception.Message)" 'ERROR' }
    })

    $ui.BtnAdd.Add_Click({   & $doAdd })
    $ui.BtnClose.Add_Click({ $win.Close() })

    # --- go ----------------------------------------------------------------
    Write-BagLog 'Bulk Add to Group opened.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-BagLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-BagLog "Device cache supplied: $(@($state.Cache).Count) device(s)" 'INFO'

    if ($CsvPath) { $win.Add_ContentRendered({ & $loadCsv $CsvPath }) }
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# connect and open the window on its own so the module can be tested.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'Device.Read.All','Group.Read.All','Group.ReadWrite.All',
                            'GroupMember.ReadWrite.All','DeviceManagementManagedDevices.Read.All' -NoWelcome

    Show-BulkAddToGroupWindow -XamlPath $XamlPath
}































































































