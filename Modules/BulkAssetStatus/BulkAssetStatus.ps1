<#
.SYNOPSIS
    Bulk Update Asset Status - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that sets the Intune Management name (managedDeviceName) of
    MANY managed devices at once, from a list of device names in a CSV.

      - The operator browses for a CSV. Every line of the file is read -
        nothing is skipped - and each one supplies one device name, taken
        from the first column.
      - Each name is matched against the Intune managed devices. Only a name
        that matches exactly ONE managed device can be written to; names that
        match nothing, or match several devices, are listed with the reason
        and cannot be ticked.
      - The current Management name of every matched device is read from
        Graph and shown in the grid, so the operator can see exactly what is
        about to be replaced before anything is written.
      - ONE status is picked for the whole batch. Every ticked device is set
        to that same status.
      - A confirmation prompt naming the status and the device count is shown
        before the first write, and it defaults to No.
      - Each row reports its own result as the run proceeds, so a partial
        failure is visible per device rather than as one opaque error.

    The Management name is a label only. Changing it does NOT retire, wipe,
    unenrol or otherwise act on the device, and it leaves the device name,
    serial number, group memberships and assignments untouched.

    This is the bulk counterpart of the Asset Status module. Both read the
    same approved status list, so the two can never drift apart.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so there is no separate
    .xaml file to deploy or keep in sync.

    Every function here is Bas-prefixed on purpose. This module,
    AssetStatus.ps1, CopyDeviceGroups.ps1, RemoveDeviceGroups.ps1,
    BulkAddToGroup.ps1 and AppDependencyCheck.ps1 are dot-sourced into the
    SAME session by Toolkit.ps1, so a shared name would mean the last file
    loaded silently wins and could change the behaviour of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\BulkAssetStatus\BulkAssetStatus.ps1"
        Show-BulkAssetStatusWindow -DeviceCache $Script:DeviceCache `
                                   -Owner       $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: DeviceManagementManagedDevices.ReadWrite.All
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# The statuses this module is allowed to write.
#
# Deliberately the SAME source as the single-device Asset Status module:
# Toolkit.ps1 sets $Script:AssetStatusValues from ModuleConfig.psd1 before
# dot-sourcing either file, so one edit in ModuleConfig.psd1 moves both
# drop-downs and both validation lists together. When that is absent - or
# when this module is run standalone - the seven built-in statuses are used.
#
# This is also the approved list: a value that is not on it is refused
# before anything is written to Graph.
# ---------------------------------------------------------------------------
$BasDefaultStatuses = @('Assigned','In-Stock','Retired','Recycled','Stolen','Legalhold','Lost')

$Script:BasStatuses = $BasDefaultStatuses

if ((Get-Variable -Name AssetStatusValues -Scope Script -ErrorAction SilentlyContinue) -and
    $null -ne $Script:AssetStatusValues) {

    $basConfigured = @($Script:AssetStatusValues) |
                     ForEach-Object { [string]$_ } |
                     Where-Object   { -not [string]::IsNullOrWhiteSpace($_) } |
                     ForEach-Object { $_.Trim() }

    if ($basConfigured.Count -gt 0) {
        $Script:BasStatuses = @($basConfigured)
    }
}

# ---------------------------------------------------------------------------
# Row type for the device grid.
#
# A separate type from the other bulk module on purpose: Add-Type cannot
# redefine a type that is already loaded in the session, so reusing
# EndpointguyBulkDeviceRow would leave whichever module loaded second
# without the properties it needs.
#
# Selected, Result and NewName raise PropertyChanged so the grid repaints
# when Select all / Select none is used, when the status is changed, and
# while the write loop runs.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyBulkAssetStatusRow' -as [type])) {
    Add-Type -ReferencedAssemblies System.ComponentModel.TypeConverter -TypeDefinition @'
using System.ComponentModel;

public class EndpointguyBulkAssetStatusRow : INotifyPropertyChanged
{
    private bool   _selected;
    private string _result  = "";
    private string _newName = "";
    private string _currentName = "";

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

    public string NewName
    {
        get { return _newName; }
        set
        {
            if (_newName != value)
            {
                _newName = value;
                OnPropertyChanged("NewName");
            }
        }
    }

    public string CurrentName
    {
        get { return _currentName; }
        set
        {
            if (_currentName != value)
            {
                _currentName = value;
                OnPropertyChanged("CurrentName");
            }
        }
    }

    public bool   IsEligible   { get; set; }
    public string CsvName      { get; set; }
    public string DeviceName   { get; set; }
    public string Status       { get; set; }
    public string SerialNumber { get; set; }
    public string User         { get; set; }
    public string IntuneId     { get; set; }

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
$BasXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Bulk Update Asset Status"
        Height="900" Width="1250"
        WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False"
        Background="{DynamicResource WindowBrush}"
        FontFamily="Segoe UI">

    <Window.Resources>

        <!-- Theme brushes. Kept in sync with Toolkit.ps1 so the child     -->
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
                <TextBlock Text="Bulk Update Asset Status" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Reads device names from the first column of a CSV - every line of the file is read - matches each one to an Intune managed device, shows the Management name each device carries today, and sets every ticked device to the SAME status. A name that matches no device, or matches more than one, is listed with the reason and cannot be ticked."
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

        <!-- 2. Status for the batch -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,12">
                    <TextBlock Text="2. Status to apply" Style="{StaticResource CardHeader}"/>
                    <TextBlock Text="One status is applied to every ticked device." Margin="12,0,0,0"
                               VerticalAlignment="Bottom" FontSize="13"
                               Foreground="{StaticResource SubtleTextBrush}"/>
                </StackPanel>

                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="320"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <StackPanel Grid.Column="0">
                        <TextBlock Text="NEW STATUS" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <!-- Items are added at runtime from $Script:BasStatuses. -->
                        <ComboBox x:Name="CmbStatus" Height="32" FontSize="13"/>
                    </StackPanel>
                </Grid>

                <Border Grid.Row="2" CornerRadius="4" Margin="0,14,0,0"
                        Background="{StaticResource InfoPanelBrush}"
                        BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                        Padding="12,10">
                    <StackPanel>
                        <TextBlock Text="RESULT" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,4"/>
                        <TextBlock x:Name="TxtPreview" Text="Pick a status to see what will be written."
                                   FontSize="13" TextWrapping="Wrap"/>
                        <TextBlock Text="The Management name is a label. Setting it does not retire, wipe or unenrol the device, and the device name, serial number and group memberships are left untouched."
                                   FontSize="11.5" TextWrapping="Wrap" Margin="0,8,0,0"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                </Border>
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
                        <Button x:Name="BtnSelectAll"    Style="{StaticResource NeutralButton}" Content="Select all"      Margin="0,0,10,0"/>
                        <Button x:Name="BtnSelectNone"   Style="{StaticResource NeutralButton}" Content="Select none"     Margin="0,0,10,0"/>
                        <Button x:Name="BtnSelectHighlighted" Style="{StaticResource NeutralButton}" Content="Tick highlighted rows"/>
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
                        <DataGridCheckBoxColumn Header="Set" Width="60"
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
                        <DataGridTextColumn Header="Name in file"    Binding="{Binding CsvName}"      Width="180" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Device Name"     Binding="{Binding DeviceName}"   Width="180" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Match"           Binding="{Binding Status}"       Width="250" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Management name now" Binding="{Binding CurrentName}" Width="200" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Will become"     Binding="{Binding NewName}"      Width="160" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Serial"          Binding="{Binding SerialNumber}" Width="140" IsReadOnly="True"/>
                        <DataGridTextColumn Header="User"            Binding="{Binding User}"         Width="180" IsReadOnly="True"/>
                        <DataGridTextColumn Header="Result"          Binding="{Binding Result}"       Width="*"   IsReadOnly="True"/>
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
                    <Button x:Name="BtnApply" Style="{StaticResource GreenButton}"
                            Content="Set Management name" MinWidth="220" Height="38" Margin="0,0,12,0"/>
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
$Script:BasLogSink    = $null
$Script:BasLogVerbose = $false

function Write-BasLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:BasLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:BasLogSink) {
        try { & $Script:BasLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-BasLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-BasGraphError $ErrorRecord
    Write-BasLog "$Context failed: $detail" 'ERROR'
    if ($Script:BasLogVerbose -and $ErrorRecord) {
        $ex = Get-BasProperty $ErrorRecord 'Exception'
        if ($ex) { Write-BasLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-BasProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-BasLog "At: $((Get-BasProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-BasProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-BasProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-BasProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-BasGraphPaged {
    # Walks @odata.nextLink and returns every item in the 'value' array.
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($Uri) {
        $page++
        Write-BasLog "GET $Uri" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        $sw.Stop()

        $count = 0
        if (Test-BasProperty $resp 'value') { $count = @($resp.value).Count; $all.AddRange(@($resp.value)) }
        Write-BasLog "  -> page $page returned $count item(s) in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        $Uri = if (Test-BasProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        if ($Uri) { Write-BasLog '  -> following @odata.nextLink for the next page' 'DEBUG' }
    }
    return $all
}

function Get-BasBodyError {
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

    if (Test-BasProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-BasProperty $err 'code' ''
        $msg  = Get-BasProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-BasGraphError {
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
    if (Test-BasProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-BasProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-BasBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-BasProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-BasProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-BasProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-BasBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-BasProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-BasProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-BasProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------
function Import-BasDeviceNameFile {
    <#
        Reads the device list: one device name per line, taken from the first
        column. EVERY line of the file is read - no line is ever skipped.

        There is no spacer row and no header row: line 1 is treated exactly
        like every other line, so a device name sitting on it is imported. If
        a file does arrive with a heading such as 'DeviceName' on top, that
        heading is read as a name and reported as not matched, which is
        visible in the grid rather than silently applied.

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
    # which is one of the ways the bulk modules produced "Argument types do
    # not match". Lower-casing the key gives the same case-insensitive
    # behaviour with nothing for the binder to resolve.
    $seen  = New-Object System.Collections.Generic.HashSet[string]
    $names = New-Object System.Collections.Generic.List[string]
    $dupes = 0

    foreach ($row in @($dataLines | ConvertFrom-Csv -Header 'DeviceName')) {
        $value = [string](Get-BasProperty $row 'DeviceName' '')
        $value = $value.Trim()
        if ([string]::IsNullOrWhiteSpace($value)) { $blank++; continue }
        # HashSet.Add returns false when the value was already present, so this
        # is the de-duplication test and the insert in one typed call.
        if (-not $seen.Add([string]$value.ToLowerInvariant())) { $dupes++; continue }
        $names.Add($value) | Out-Null
    }

    if ($blank -gt 0) { Write-BasLog "Ignored $blank blank line(s)." 'INFO' }
    if ($dupes -gt 0) { Write-BasLog "Ignored $dupes duplicate name(s)." 'INFO' }

    if ($names.Count -eq 0) {
        throw "No device names were found in the first column of '$Path'."
    }

    Write-BasLog "Read $($names.Count) device name(s) from $Path." 'OK'

    [pscustomobject]@{
        Names          = $names
        BlankCount     = $blank
        DuplicateCount = $dupes
    }
}

# ---------------------------------------------------------------------------
# Devices
# ---------------------------------------------------------------------------
function Get-BasManagedDeviceLookup {
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
        Write-BasLog "Using the Toolkit device cache - $($devices.Count) managed device(s)." 'INFO'
    }
    else {
        Write-BasLog 'No device cache was supplied - reading managed devices from Graph.' 'INFO'
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=id,deviceName,managedDeviceName,serialNumber,userPrincipalName&`$top=999"
        $devices = @(Get-BasGraphPaged -Uri $uri | ForEach-Object {
            [pscustomobject]@{
                DeviceName        = Get-BasProperty $_ 'deviceName' ''
                ManagedDeviceName = Get-BasProperty $_ 'managedDeviceName' ''
                SerialNumber      = Get-BasProperty $_ 'serialNumber' ''
                User              = Get-BasProperty $_ 'userPrincipalName' ''
                IntuneId          = Get-BasProperty $_ 'id' ''
            }
        })
        Write-BasLog "Read $($devices.Count) managed device(s) from Graph." 'OK'
    }

    # A generic Dictionary, NOT a hashtable.
    #
    # Reading a hashtable as $h[$k] does not call Hashtable.get_Item directly:
    # PowerShell routes it through its parameterized-property binder, which
    # picks an indexer overload by reflection. On a large hashtable reached
    # from a typed variable that resolution can fail outright with the
    # unhelpful "Argument types do not match".
    #
    # Dictionary.TryGetValue and .Add are plain strongly typed calls with one
    # overload each, so no binder and no reflection is involved. Keys are
    # already lower-cased here, so the default ordinal comparer is correct.
    $lookup = New-Object 'System.Collections.Generic.Dictionary[string,object]'
    foreach ($d in $devices) {
        $name = [string](Get-BasProperty $d 'DeviceName' '')
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $key = [string]$name.Trim().ToLowerInvariant()

        $bucket = $null
        if (-not $lookup.TryGetValue($key, [ref]$bucket)) {
            $bucket = New-Object System.Collections.Generic.List[object]
            $lookup.Add($key, $bucket)
        }
        $bucket.Add($d) | Out-Null
    }
    # Comma keeps this a single dictionary on the pipeline rather than letting
    # PowerShell enumerate it into loose KeyValuePair objects.
    return ,$lookup
}

function Get-BasCurrentManagementName {
    <#
        Reads the live managedDeviceName for one device.

        The cache is populated by a device search that does not always carry
        managedDeviceName, and even when it does the cache can be stale. The
        operator is about to overwrite this value, so it is read back from
        Graph rather than trusted from the cache.

        Returns the name, or throws with the real Graph message.
    #>
    param([Parameter(Mandatory)][string]$ManagedDeviceId)

    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$ManagedDeviceId`?`$select=id,managedDeviceName"
    Write-BasLog "GET $uri" 'GRAPH'
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    if ($null -eq $resp) {
        throw "No Intune managed device was found for id $ManagedDeviceId."
    }
    return [string](Get-BasProperty $resp 'managedDeviceName' '')
}

function Resolve-BasDeviceRows {
    <#
        Turns the names read from the CSV into grid rows.

        Only a row that matched exactly one managed device AND carries an
        Intune managed device id is eligible, because the Management name
        lives on the Intune managedDevice object - an Autopilot-only device
        that has never enrolled has no such record to write to.
    #>
    param(
        [Parameter(Mandatory)][string[]]$Names,
        $DeviceCache
    )

    [System.Collections.Generic.Dictionary[string,object]]$lookup = Get-BasManagedDeviceLookup -DeviceCache $DeviceCache
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($n in $Names) {
        $row = New-Object EndpointguyBulkAssetStatusRow
        $row.CsvName     = $n
        $row.Result      = ''
        $row.NewName     = ''
        $row.CurrentName = ''
        $row.IsEligible  = $false
        $row.Selected    = $false

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
            $row.DeviceName = Get-BasProperty $hits[0] 'DeviceName' ''
            $row.Status     = "Ambiguous - $($hits.Count) managed devices share this name"
        }
        else {
            $d = $hits[0]
            $row.DeviceName   = Get-BasProperty $d 'DeviceName' ''
            $row.SerialNumber = [string](Get-BasProperty $d 'SerialNumber' '')
            $row.User         = [string](Get-BasProperty $d 'User' '')
            $row.IntuneId     = [string](Get-BasProperty $d 'IntuneId' '')
            # The cache calls it IntuneId; a device read straight from Graph
            # calls it Id. Fall back so both shapes resolve.
            if ([string]::IsNullOrWhiteSpace($row.IntuneId)) {
                $row.IntuneId = [string](Get-BasProperty $d 'Id' '')
            }
            # The Toolkit cache calls this ManagementName; a device read
            # straight from Graph by this module calls it ManagedDeviceName.
            # Probe both so either shape populates the column. This is only
            # the provisional value - it is replaced by a live read before
            # anything is written.
            $row.CurrentName = [string](Get-BasProperty $d 'ManagementName' '')
            if ([string]::IsNullOrWhiteSpace($row.CurrentName)) {
                $row.CurrentName = [string](Get-BasProperty $d 'ManagedDeviceName' '')
            }

            if ([string]::IsNullOrWhiteSpace($row.IntuneId)) {
                $src = [string](Get-BasProperty $d 'Source' '')
                if ($src -eq 'Autopilot') {
                    $row.Status = 'Imported into Autopilot but not enrolled - no Management name to set'
                }
                else {
                    $row.Status = 'No Intune managed device id - Management name cannot be set'
                }
            }
            else {
                $row.Status     = 'Matched - ready to set'
                $row.IsEligible = $true
                $row.Selected   = $true   # pre-tick everything that can be written
            }
        }

        Write-BasLog "  $($row.CsvName) -> $($row.Status)" 'DEBUG'
        $rows.Add($row) | Out-Null
    }

    $elig = @($rows | Where-Object { $_.IsEligible }).Count
    Write-BasLog "Matched $elig of $($rows.Count) name(s) from the file." 'OK'

    return ($rows | Sort-Object -Property @{Expression={ -not $_.IsEligible }}, CsvName)
}

function Set-BasManagementName {
    <#
        Writes managedDeviceName on one managedDevice.

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

    Write-BasLog "PATCH $uri" 'GRAPH'
    Write-BasLog "  -> body $body" 'DEBUG'

    # Preferred path: ask Graph not to throw, so the error body is returned
    # intact instead of being flattened into a useless exception message.
    $canSkipThrow = $false
    try {
        $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction Stop
        $canSkipThrow = $cmd.Parameters.ContainsKey('SkipHttpErrorCheck') -and
                        $cmd.Parameters.ContainsKey('StatusCodeVariable')
    }
    catch { $canSkipThrow = $false }

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body `
                    -ContentType 'application/json' -OutputType PSObject `
                    -SkipHttpErrorCheck -StatusCodeVariable 'code'

        Write-BasLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) { return }

        $detail = Get-BasBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        Write-BasLog "  -> Graph said: $detail" 'DEBUG'
        throw "HTTP $code - $detail"
    }

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $body -ContentType 'application/json' | Out-Null
    }
    catch {
        throw (Get-BasGraphError $_)
    }
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-BulkAssetStatusWindow {
    <#
    .SYNOPSIS
        Opens the Bulk Update Asset Status window.
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
        [string]$CsvPath
    )

    # --- load XAML ---------------------------------------------------------
    $xamlText = $BasXamlString
    try {
        [xml]$xamlDoc = $xamlText
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        $win    = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "The Bulk Update Asset Status XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtCsvPath','BtnBrowseCsv','TxtCsvSummary','BtnReloadCsv',
                     'CmbStatus','TxtPreview',
                     'GridDevices','DeviceCount','ChkShowUnmatched','BtnSelectAll','BtnSelectNone',
                     'BtnSelectHighlighted','StatusText','BtnApply','BtnClose')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the Bulk Update Asset Status XAML.",
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
        Rows  = @()
        Busy  = $false
        Cache = @($DeviceCache)
    }

    $deviceRows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $ui.GridDevices.ItemsSource = $deviceRows

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render) | Out-Null
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnApply, $ui.BtnBrowseCsv, $ui.BtnReloadCsv,
                         $ui.BtnSelectAll, $ui.BtnSelectNone, $ui.BtnSelectHighlighted)) {
            $b.IsEnabled = -not $Busy
        }
        $ui.CmbStatus.IsEnabled = -not $Busy
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render) | Out-Null
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

    $refreshGrid = {
        # Rebuilds the visible rows, honouring the 'Show unmatched' toggle.
        $showAll = [bool]$ui.ChkShowUnmatched.IsChecked
        $deviceRows.Clear()
        foreach ($r in $state.Rows) {
            if ($showAll -or $r.IsEligible) { $deviceRows.Add($r) }
        }
        $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
        $total    = @($state.Rows).Count
        $ticked   = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected }).Count
        $ui.DeviceCount.Text = "$eligible of $total name(s) matched a device - $ticked ticked"
    }

    # Shows the chosen status on every eligible row, so the grid states what
    # each device will end up with before anything is written.
    $refreshPreview = {
        $status = & $getStatus
        foreach ($r in $state.Rows) {
            if ($r.IsEligible) { $r.NewName = $status }
        }

        $ticked = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected }).Count
        if ([string]::IsNullOrWhiteSpace($status)) {
            $ui.TxtPreview.Text = 'Pick a status to see what will be written.'
        }
        elseif ($ticked -eq 0) {
            $ui.TxtPreview.Text = "Management name will be set to '$status'. No device is ticked yet."
        }
        else {
            $ui.TxtPreview.Text = "Management name will be set to '$status' on $ticked device(s)."
        }
        & $refreshGrid
    }

    # --- csv ---------------------------------------------------------------
    $loadCsv = {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }

        & $setBusy $true
        try {
            & $setStatus "Reading $Path..."
            $file = Import-BasDeviceNameFile -Path $Path

            $ui.TxtCsvPath.Text = $Path

            $summary = "$(@($file.Names).Count) device name(s)"
            if ($file.DuplicateCount -gt 0) { $summary = "$summary, $($file.DuplicateCount) duplicate(s) ignored" }
            if ($file.BlankCount -gt 0)     { $summary = "$summary, $($file.BlankCount) blank line(s) ignored" }
            $ui.TxtCsvSummary.Text = $summary

            & $setStatus "Matching $(@($file.Names).Count) name(s) against Intune..."
            $state.Rows = @(Resolve-BasDeviceRows -Names @($file.Names) -DeviceCache $state.Cache)

            # Read back the live Management name for every matched device, so
            # the operator sees what is about to be replaced rather than
            # whatever the cache happened to hold.
            $matched = @($state.Rows | Where-Object { $_.IsEligible })
            $i = 0
            foreach ($row in $matched) {
                $i++
                & $setStatus "Reading current Management name $i of $($matched.Count): $($row.DeviceName)..."
                try {
                    $row.CurrentName = Get-BasCurrentManagementName -ManagedDeviceId $row.IntuneId
                }
                catch {
                    # A device that cannot be read cannot be safely written
                    # to either, so it is taken out of the run rather than
                    # left ticked with an unknown current value.
                    $detail = Write-BasLogException $_ "Reading '$($row.DeviceName)'"
                    $row.CurrentName = '(could not be read)'
                    $row.IsEligible  = $false
                    $row.Selected    = $false
                    $row.Status      = "Could not be read: $detail"
                }
            }

            & $refreshPreview

            $eligible = @($state.Rows | Where-Object { $_.IsEligible }).Count
            & $setStatus "$eligible of $(@($file.Names).Count) name(s) matched a device that can be set."
        }
        catch {
            $state.Rows = @()
            & $refreshGrid
            $ui.TxtCsvSummary.Text = 'No file loaded'
            $detail = Write-BasLogException $_ 'Reading the device list'
            & $setStatus "Could not read the file: $detail"
            [System.Windows.MessageBox]::Show($detail,'Bulk Update Asset Status','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- apply -------------------------------------------------------------
    $doApply = {
        $status = & $getStatus
        if ([string]::IsNullOrWhiteSpace($status)) {
            [System.Windows.MessageBox]::Show('Pick the status to apply first.',
                'Bulk Update Asset Status','OK','Warning') | Out-Null
            return
        }

        # The approved list is checked again here, not just when the drop-down
        # was built, so nothing outside the list can reach Graph.
        if ($Script:BasStatuses -notcontains $status) {
            [System.Windows.MessageBox]::Show(
                "'$status' is not one of the approved statuses, so it will not be written.",
                'Bulk Update Asset Status','OK','Error') | Out-Null
            return
        }

        $picked = @($state.Rows | Where-Object { $_.IsEligible -and $_.Selected })
        if ($picked.Count -eq 0) {
            [System.Windows.MessageBox]::Show('Tick at least one matched device to set.',
                'Bulk Update Asset Status','OK','Warning') | Out-Null
            return
        }

        # The Management name is what reporting keys off, so the prompt names
        # the status, states the count, and defaults to No.
        $names = ($picked | Select-Object -First 15 |
                  ForEach-Object { "  - $($_.DeviceName)  [$($_.CurrentName)]" }) -join "`n"
        if ($picked.Count -gt 15) { $names = "$names`n  ... and $($picked.Count - 15) more" }

        $confirm = [System.Windows.MessageBox]::Show(
            "Set the Intune Management name of $($picked.Count) device(s) to '$status'?`n`n$names`n`nThe name each device carries today is shown in brackets and will be replaced. This is a label only - no device is retired, wiped or unenrolled.",
            'Confirm bulk status update','YesNo','Warning',[System.Windows.MessageBoxResult]::No)
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-BasLog 'Bulk status update cancelled at the confirmation prompt.' 'INFO'
            & $setStatus 'Cancelled - nothing was changed.'
            return
        }

        & $setBusy $true
        $done = 0; $failed = 0
        $lines = New-Object System.Collections.Generic.List[string]
        $i = 0

        foreach ($row in $picked) {
            $i++
            & $setStatus "Setting $i of $($picked.Count): $($row.DeviceName)..."
            $row.Result = 'Working...'
            try {
                Set-BasManagementName -ManagedDeviceId $row.IntuneId -NewName $status
                $done++
                $row.CurrentName = $status
                $row.Result      = 'Set'
                Write-BasLog "$($row.DeviceName) - set to $status" 'OK'
                $lines.Add("$($row.DeviceName) - set to $status") | Out-Null
            }
            catch {
                $failed++
                $detail = Write-BasLogException $_ "Set '$($row.DeviceName)'"
                $row.Result = "FAILED: $detail"
                $lines.Add("$($row.DeviceName) - FAILED: $detail") | Out-Null
            }
        }

        & $setBusy $false
        $summary = "Bulk status update complete - $done set, $failed failed."
        & $setStatus $summary
        Write-BasLog $summary $(if ($failed -gt 0) { 'WARN' } else { 'OK' })

        $shown = @($lines | Select-Object -First 25) -join "`n"
        if ($lines.Count -gt 25) { $shown = "$shown`n... and $($lines.Count - 25) more - see the grid for the full list." }

        $icon = if ($failed -gt 0) { 'Warning' } else { 'Information' }
        [System.Windows.MessageBox]::Show(
            "$summary`n`n$shown",
            'Bulk Update Asset Status','OK',$icon) | Out-Null
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

    $ui.ChkShowUnmatched.Add_Checked({   & $refreshGrid })
    $ui.ChkShowUnmatched.Add_Unchecked({ & $refreshGrid })

    $ui.BtnSelectAll.Add_Click({
        foreach ($r in $state.Rows) { if ($r.IsEligible) { $r.Selected = $true } }
        & $refreshPreview
    })
    $ui.BtnSelectNone.Add_Click({
        foreach ($r in $state.Rows) { $r.Selected = $false }
        & $refreshPreview
    })

    # Ticks whatever is highlighted in the grid. Highlighting a block with
    # Shift or Ctrl and pressing this is the quick way to pick out a subset
    # without clicking every box.
    $ui.BtnSelectHighlighted.Add_Click({
        $sel = @($ui.GridDevices.SelectedItems)
        if ($sel.Count -eq 0) {
            & $setStatus 'Highlight one or more rows in the grid first, then press Tick highlighted rows.'
            return
        }
        $n = 0
        foreach ($r in $sel) {
            if ($r.IsEligible) { $r.Selected = $true; $n++ }
        }
        & $refreshPreview
        & $setStatus "Ticked $n highlighted device(s)."
    })

    $ui.BtnApply.Add_Click({ & $doApply })
    $ui.BtnClose.Add_Click({ $win.Close() })

    # --- go ----------------------------------------------------------------
    # Build the drop-down from the approved list, so a status added in
    # ModuleConfig.psd1 appears here without touching the XAML.
    #
    # The SelectionChanged handler is attached AFTER the items are added:
    # setting the selection raises the event synchronously, and a handler
    # that ran while the window was still half-built would touch controls
    # that are not there yet.
    $ui.CmbStatus.Items.Clear()
    foreach ($s in $Script:BasStatuses) {
        $ui.CmbStatus.Items.Add((New-Object System.Windows.Controls.ComboBoxItem -Property @{ Content = $s })) | Out-Null
    }
    $ui.CmbStatus.Add_SelectionChanged({ & $refreshPreview })

    Write-BasLog 'Bulk Update Asset Status opened.' 'INFO'
    Write-BasLog "Statuses offered: $($Script:BasStatuses -join ', ')" 'DEBUG'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-BasLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-BasLog "Device cache supplied: $(@($state.Cache).Count) device(s)" 'INFO'

    if ($CsvPath) { $win.Add_ContentRendered({ & $loadCsv $CsvPath }) }
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# connect and open the window on its own so the module can be tested.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'DeviceManagementManagedDevices.ReadWrite.All' -NoWelcome

    Show-BulkAssetStatusWindow
}
