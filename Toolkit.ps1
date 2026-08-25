<#
.SYNOPSIS
    Endpointguy Intune Toolkit - WPF front end for Microsoft Graph / Intune administration.

.DESCRIPTION
    Phase 1 build:
      - Connect to Microsoft Graph (interactive) with live connection status indicator
      - Device search by Device Name / Serial Number / Primary Username (contains or exact)
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
      - Remaining Device Actions are stubbed with "Coming soon"

.NOTES
    SELF-CONTAINED UI - the WPF XAML is embedded in this file, so there is no
    MainWindow.xaml dependency at runtime. Use -XamlPath for UI development.

    Requires: Windows PowerShell 5.1; Graph authentication module installs automatically
        First launch requires access to the PowerShell Gallery

    Run: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1
    Or double-click Run-Toolkit.bat
#>

[CmdletBinding()]
param(
    [string]$TenantId,

    # Optional dev override: point at an external MainWindow.xaml to tweak the UI
    # without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ---------------------------------------------------------------------------
# Graph authentication module
#
# The toolkit installs Microsoft.Graph.Authentication for the current user on
# first launch. No administrator elevation is required. Internet access to the
# PowerShell Gallery is required only when the module is not already installed.
# ---------------------------------------------------------------------------
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    try {
        [System.Windows.MessageBox]::Show(
            "Microsoft.Graph.Authentication is not installed. The toolkit will install it for the current user now.`n`nThis first-time setup can take a few minutes.",
            'First-time setup','OK','Information') | Out-Null

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
        [System.Windows.MessageBox]::Show(
            "Microsoft.Graph.Authentication could not be installed automatically.`n`n$($_.Exception.Message)`n`nConfirm that the PowerShell Gallery is reachable, then restart the toolkit.",
            'First-time setup failed','OK','Error') | Out-Null
        return
    }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# ---------------------------------------------------------------------------
# Modules
#
# Each action lives in its own folder under Modules\.
# Modules are resolved relative to the Toolkit.ps1 directory.
# A missing module is not fatal - the matching button just reports it.
#
# Each entry is a path relative to the toolkit root. The legacy flat layout
# (module sitting next to Toolkit.ps1) is still probed as a fallback so a
# partially migrated folder keeps working.
# ---------------------------------------------------------------------------
$Script:ModuleRoot = $PSScriptRoot

$Script:ModulesLoaded = @{}

foreach ($module in @('Modules\CopyDeviceGroups\CopyDeviceGroups.ps1',
                      'Modules\RemoveDeviceGroups\RemoveDeviceGroups.ps1',
                      'Modules\BulkAddToGroup\BulkAddToGroup.ps1',
                      'Modules\AppDependencyCheck\AppDependencyCheck.ps1')) {
    $moduleName = Split-Path -Leaf $module

    # Preferred subfolder location, then the old flat location.
    $candidates = @(
        (Join-Path $Script:ModuleRoot $module)
        (Join-Path $Script:ModuleRoot $moduleName)
    )
    $modulePath = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($modulePath) {
        try {
            . $modulePath
            $Script:ModulesLoaded[$moduleName] = $true
        }
        catch {
            $Script:ModulesLoaded[$moduleName] = $false
            [System.Windows.MessageBox]::Show(
                "Module '$moduleName' could not be loaded:`n`n$($_.Exception.Message)",
                'Module load error','OK','Warning') | Out-Null
        }
    }
    else {
        $Script:ModulesLoaded[$moduleName] = $false
    }
}

# Scopes needed now, plus headroom for the group actions coming next
$Script:GraphScopes = @(
    'DeviceManagementManagedDevices.ReadWrite.All'
    'DeviceManagementConfiguration.Read.All'
    'Device.Read.All'
    'Group.Read.All'
    'Group.ReadWrite.All'
    'GroupMember.ReadWrite.All'
    'User.Read.All'
    'DeviceManagementApps.Read.All'
)

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$XamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Endpointguy Intune Toolkit"
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
            <Setter Property="CornerRadius" Value="6"/>
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

        <Border Grid.Row="0" Style="{StaticResource Card}" Padding="22,16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <TextBlock Text="Endpointguy Intune Toolkit"
                               FontSize="24" FontWeight="Bold"
                               Foreground="{StaticResource TitleTextBrush}"/>
                    <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                        <TextBlock Text="Device lookup and Intune administration tools"
                                   Foreground="{StaticResource SubtleTextBrush}" FontSize="13"/>
                        <TextBlock Text="  &#8226;  " Foreground="{StaticResource BulletBrush}" FontSize="13"/>
                        <TextBlock x:Name="LinkSite" Text="endpointguy.com"
                                   Foreground="{StaticResource AccentTextBrush}" FontSize="13"
                                   Cursor="Hand"/>
                    </StackPanel>
                </StackPanel>

                <Button x:Name="BtnConnect" Grid.Column="1"
                        Style="{StaticResource GreenButton}"
                        Content="Connect to Graph"
                        MinWidth="170" Height="42" Margin="0,0,18,0"/>

                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock x:Name="ConnIcon" Text="&#xE711;"
                               FontFamily="Segoe MDL2 Assets" FontSize="20"
                               Foreground="{StaticResource ConnBadBrush}"
                               VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="ConnStatus" Text="Not connected"
                                   FontWeight="SemiBold" FontSize="13.5"
                                   Foreground="{StaticResource ConnBadBrush}"/>
                        <TextBlock x:Name="ConnAccount" Text="Sign in to begin"
                                   FontSize="12" Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="0,14,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="290"/>
                <ColumnDefinition Width="14"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0" Style="{StaticResource Card}" Padding="18">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Device Actions" Style="{StaticResource CardHeader}"/>

                    <Border Grid.Row="1" Margin="0,16,0,18" CornerRadius="4"
                            Background="{StaticResource InfoPanelBrush}"
                            BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                            Padding="12,10">
                        <StackPanel>
                            <TextBlock Text="SELECTED DEVICE" FontSize="10.5" FontWeight="Bold"
                                       Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,5"/>
                            <TextBlock x:Name="SelectedDeviceText"
                                       Text="No device selected"
                                       FontSize="12.5" Foreground="{StaticResource TextBrush}"
                                       TextWrapping="Wrap" LineHeight="18"/>
                        </StackPanel>
                    </Border>

                    <StackPanel Grid.Row="2">
                        <Button x:Name="BtnCopyGroups"   Style="{StaticResource ActionButton}"
                                Content="Copy Device Groups"
                                ToolTip="Copy assigned security groups from the selected device to another device."/>
                        <Button x:Name="BtnRemoveGroups" Style="{StaticResource ActionButton}"
                                Content="Remove Device Groups"
                                ToolTip="Remove the selected device from its assigned security groups. Asks for confirmation first."/>
                    </StackPanel>

                    <StackPanel Grid.Row="3" Margin="0,8,0,0">
                        <Separator Background="{StaticResource SeparatorBrush}" Margin="0,0,0,12"/>
                        <TextBlock Text="Bulk Actions" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                        <Button x:Name="BtnBulkAddGroup" Style="{StaticResource ActionButton}" Content="Bulk Add to Group"/>
                        <Separator Background="{StaticResource SeparatorBrush}" Margin="0,14,0,12"/>
                        <TextBlock Text="Reporting" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                        <Button x:Name="BtnAppDependency" Style="{StaticResource ActionButton}"
                                Content="App Dependency Check"
                                ToolTip="Read-only. Lists every app that depends on a selected Win32 app, so you can see what breaks before changing it."/>
                    </StackPanel>

                    <TextBlock Grid.Row="5" TextWrapping="Wrap" FontSize="11.5"
                               Foreground="{StaticResource SubtleTextBrush}"
                               Text="Actions apply to the device selected in the search results."/>
                </Grid>
            </Border>

            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="14"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Style="{StaticResource Card}" Padding="18">
                    <StackPanel>
                        <TextBlock Text="Device Search" Style="{StaticResource CardHeader}" Margin="0,0,0,14"/>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Search by:" VerticalAlignment="Center"
                                       FontSize="13" Margin="0,0,10,0"/>
                            <ComboBox x:Name="CmbSearchField" Width="180" SelectedIndex="0">
                                <ComboBoxItem Content="Device Name"/>
                                <ComboBoxItem Content="Serial Number"/>
                                <ComboBoxItem Content="Primary Username"/>
                            </ComboBox>
                            <TextBox x:Name="TxtSearch" Width="300" Height="32" Margin="14,0,0,0"/>
                            <Button x:Name="BtnSearch" Style="{StaticResource PrimaryButton}"
                                    Content="Search" Width="110" Height="32" Margin="14,0,0,0"/>
                            <CheckBox x:Name="ChkExact" Content="Exact match only"
                                      VerticalAlignment="Center" FontSize="13" Margin="18,0,0,0"/>
                            <Button x:Name="BtnRefreshCache" Style="{StaticResource NeutralButton}"
                                    Content="Refresh Device Cache" Height="32" Margin="18,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                </Border>

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

                        <DataGrid x:Name="GridDevices" Grid.Row="1"
                                  AutoGenerateColumns="False"
                                  IsReadOnly="True"
                                  SelectionMode="Single"
                                  HeadersVisibility="Column"
                                  GridLinesVisibility="Horizontal"
                                  HorizontalGridLinesBrush="{StaticResource GridLineBrush}"
                                  Background="{StaticResource GridRowBrush}"
                                  Foreground="{StaticResource TextBrush}"
                                  RowBackground="{StaticResource GridRowBrush}"
                                  AlternatingRowBackground="{StaticResource GridAltRowBrush}"
                                  BorderBrush="{StaticResource GridBorderBrush}"
                                  BorderThickness="1"
                                  FontSize="12.5"
                                  RowHeight="26"
                                  CanUserSortColumns="True">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Device Name"   Binding="{Binding DeviceName}"   Width="200"/>
                                <DataGridTextColumn Header="Serial Number" Binding="{Binding SerialNumber}" Width="140"/>
                                <DataGridTextColumn Header="User"          Binding="{Binding User}"         Width="220"/>
                                <DataGridTextColumn Header="OS"            Binding="{Binding OS}"           Width="100"/>
                                <DataGridTextColumn Header="OS Version"    Binding="{Binding OSVersion}"    Width="130"/>
                                <DataGridTextColumn Header="Compliance"    Binding="{Binding Compliance}"   Width="110"/>
                                <DataGridTextColumn Header="Ownership"     Binding="{Binding Ownership}"    Width="100"/>
                                <DataGridTextColumn Header="Model"         Binding="{Binding Model}"        Width="180"/>
                                <DataGridTextColumn Header="Last Sync"     Binding="{Binding LastSync}"     Width="150"/>
                            </DataGrid.Columns>
                        </DataGrid>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="14,9" Margin="0,14,0,0">
            <Grid>
                <TextBlock x:Name="StatusText" Text="Ready. Connect to Microsoft Graph to begin."
                           FontSize="12.5" Foreground="{StaticResource TextBrush}" VerticalAlignment="Center"/>
                <TextBlock x:Name="ClockText" Text="" HorizontalAlignment="Right"
                           FontSize="12" Foreground="{StaticResource SubtleTextBrush}"
                           VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

if ($XamlPath -and (Test-Path $XamlPath)) {
    $XamlString = Get-Content -Path $XamlPath -Raw
}

try {
    [xml]$xamlDoc = $XamlString
    $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
    $Window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.MessageBox]::Show(
        "The XAML could not be loaded:`n`n$($_.Exception.Message)",
        'XAML load error','OK','Error') | Out-Null
    return
}

# ---------------------------------------------------------------------------
# Map every x:Name into a script-scope variable of the same name.
# NOTE: x:Name lives in the XAML namespace, so GetAttribute('Name') returns
# nothing. The attribute must be matched by LocalName instead.
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
    [System.Windows.MessageBox]::Show(
        'No named controls could be resolved from the XAML.',
        'XAML load error','OK','Error') | Out-Null
    return
}

foreach ($required in @('BtnConnect','BtnSearch','BtnRefreshCache','GridDevices',
                        'TxtSearch','CmbSearchField','ChkExact','StatusText',
                        'ConnIcon','ConnStatus','ConnAccount','ResultCount',
                        'SelectedDeviceText','ClockText','BtnExportCsv','LinkSite',
                        'BtnCopyGroups','BtnRemoveGroups',
                        'BtnBulkAddGroup','BtnAppDependency')) {
    if (-not (Get-Variable -Name $required -Scope Script -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "Control '$required' was not found in the XAML.",
            'XAML load error','OK','Error') | Out-Null
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
function Set-Status {
    param([string]$Message)
    $StatusText.Text = $Message
    $Window.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
}

function Show-ComingSoon {
    param([string]$Feature)
    [System.Windows.MessageBox]::Show("$Feature`n`nComing soon.",
        'Endpointguy Intune Toolkit','OK','Information') | Out-Null
    Set-Status "$Feature - coming soon."
}

function Set-Busy {
    param([bool]$Busy)
    $Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    foreach ($b in @($BtnSearch, $BtnRefreshCache, $BtnConnect)) { $b.IsEnabled = -not $Busy }
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
        $BtnConnect.Content    = 'Reconnect'
    }
    else {
        $ConnIcon.Text         = [char]0xE711
        $ConnIcon.Foreground   = $Window.TryFindResource('ConnBadBrush')
        $ConnStatus.Text       = 'Not connected'
        $ConnStatus.Foreground = $Window.TryFindResource('ConnBadBrush')
        $ConnAccount.Text      = 'Sign in to begin'
        $BtnConnect.Content    = 'Connect to Graph'
    }
}

function Test-Connected {
    $ctx = Get-GraphContextSafe
    if (-not ($ctx -and $ctx.Account)) {
        [System.Windows.MessageBox]::Show('Connect to Microsoft Graph first.',
            'Not connected','OK','Warning') | Out-Null
        return $false
    }
    return $true
}

# ---------------------------------------------------------------------------
# Device cache
# ---------------------------------------------------------------------------
function Get-ManagedDeviceCache {
    $select = 'id,deviceName,serialNumber,userPrincipalName,operatingSystem,' +
              'osVersion,complianceState,managedDeviceOwnerType,model,manufacturer,' +
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
            User          = $d.userPrincipalName
            OS            = $d.operatingSystem
            OSVersion     = $d.osVersion
            Compliance    = $d.complianceState
            Ownership     = $d.managedDeviceOwnerType
            Model         = $d.model
            Manufacturer  = $d.manufacturer
            Category      = $d.deviceCategoryDisplayName
            LastSync      = if ($d.lastSyncDateTime) { ([datetime]$d.lastSyncDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            Enrolled      = if ($d.enrolledDateTime) { ([datetime]$d.enrolledDateTime).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' }
            IntuneId      = $d.id
            EntraDeviceId = $d.azureADDeviceId
        }
    }
}

function Update-DeviceCache {
    if (-not (Test-Connected)) { return }
    Set-Busy $true
    try {
        $Script:DeviceCache   = @(Get-ManagedDeviceCache)
        $Script:CacheLoadedAt = Get-Date
        Set-Status "Device cache refreshed - $($Script:DeviceCache.Count) devices loaded at $($Script:CacheLoadedAt.ToString('HH:mm:ss'))."
    }
    catch {
        Set-Status "Cache refresh failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Cache refresh failed','OK','Error') | Out-Null
    }
    finally { Set-Busy $false }
}

# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------
function Invoke-DeviceSearch {
    if (-not (Test-Connected)) { return }

    if ([string]::IsNullOrWhiteSpace($TxtSearch.Text)) {
        Set-Status 'Enter a search term.'
        return
    }
    $term = $TxtSearch.Text.Trim()

    if ($Script:DeviceCache.Count -eq 0) {
        Update-DeviceCache
        if ($Script:DeviceCache.Count -eq 0) { return }
    }

    $field = switch ($CmbSearchField.SelectedIndex) {
        0 { 'DeviceName' }
        1 { 'SerialNumber' }
        2 { 'User' }
        default { 'DeviceName' }
    }
    $exact = [bool]$ChkExact.IsChecked

    Set-Busy $true
    try {
        $hits = $Script:DeviceCache | Where-Object {
            $v = $_.$field
            if ([string]::IsNullOrEmpty($v)) { $false }
            elseif ($exact)                  { $v -eq $term }
            else                             { $v -like "*$term*" }
        } | Sort-Object DeviceName

        $Script:Results.Clear()
        foreach ($h in $hits) { $Script:Results.Add($h) }

        $ResultCount.Text = "$($Script:Results.Count) device(s) found"
        Set-Status "Search complete - $($Script:Results.Count) result(s) for '$term' in $field."
    }
    catch {
        Set-Status "Search failed: $($_.Exception.Message)"
    }
    finally { Set-Busy $false }
}

# ---------------------------------------------------------------------------
# Event handlers
# ---------------------------------------------------------------------------
$BtnConnect.Add_Click({
    Set-Busy $true
    try {
        Set-Status 'Opening sign-in prompt...'
        $params = @{ Scopes = $Script:GraphScopes; NoWelcome = $true }
        if ($TenantId) { $params['TenantId'] = $TenantId }
        Connect-MgGraph @params
        Update-ConnectionUi
        $ctx = Get-GraphContextSafe
        Set-Status "Connected to tenant $($ctx.TenantId) as $($ctx.Account)."
    }
    catch {
        Update-ConnectionUi
        Set-Status "Connection failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Connection failed','OK','Error') | Out-Null
    }
    finally { Set-Busy $false }
})

$BtnSearch.Add_Click({ Invoke-DeviceSearch })

# WPF event scriptblocks populate $args[0] (sender) and $args[1] (event args).
# $_ is NOT set here, which would trip Set-StrictMode.
$TxtSearch.Add_KeyDown({
    if ($args[1].Key -eq 'Return') { Invoke-DeviceSearch }
})

$BtnRefreshCache.Add_Click({ Update-DeviceCache })

$GridDevices.Add_SelectionChanged({
    $d = $GridDevices.SelectedItem
    if ($d) {
        $Script:SelectedDevice   = $d
        $SelectedDeviceText.Text = "$($d.DeviceName)`nSerial: $($d.SerialNumber)`n$($d.User)"
        Set-Status "Selected $($d.DeviceName)  |  Serial: $($d.SerialNumber)"
    }
})

$BtnExportCsv.Add_Click({
    if ($Script:Results.Count -eq 0) { Set-Status 'Nothing to export.'; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV file (*.csv)|*.csv'
    $dlg.FileName = "IntuneDevices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:Results | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Set-Status "Exported $($Script:Results.Count) row(s) to $($dlg.FileName)"
    }
})

# --- Stubbed actions -------------------------------------------------------
$BtnCopyGroups.Add_Click({
    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        [System.Windows.MessageBox]::Show(
            'Select a device in the search results first.',
            'Copy Device Groups','OK','Warning') | Out-Null
        return
    }

    if (-not (Get-Command -Name Show-CopyDeviceGroupsWindow -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "CopyDeviceGroups.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\CopyDeviceGroups')",
            'Module not available','OK','Error') | Out-Null
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Copy Device Groups','OK','Error') | Out-Null
    }
})
$BtnRemoveGroups.Add_Click({
    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        [System.Windows.MessageBox]::Show(
            'Select a device in the search results first.',
            'Remove Device Groups','OK','Warning') | Out-Null
        return
    }

    if (-not (Get-Command -Name Show-RemoveDeviceGroupsWindow -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "RemoveDeviceGroups.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\RemoveDeviceGroups')",
            'Module not available','OK','Error') | Out-Null
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Remove Device Groups','OK','Error') | Out-Null
    }
})
$BtnBulkAddGroup.Add_Click({
    if (-not (Test-Connected)) { return }

    if (-not (Get-Command -Name Show-BulkAddToGroupWindow -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "BulkAddToGroup.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\BulkAddToGroup')",
            'Module missing','OK','Warning') | Out-Null
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Bulk Add to Group','OK','Error') | Out-Null
    }
})
$BtnAppDependency.Add_Click({
    if (-not (Test-Connected)) { return }

    if (-not (Get-Command -Name Show-AppDependencyCheckWindow -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "AppDependencyCheck.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\AppDependencyCheck')",
            'Module missing','OK','Warning') | Out-Null
        return
    }

    Set-Status 'Opening App Dependency Check...'
    try {
        Show-AppDependencyCheckWindow -Owner $Window
        Set-Status 'App Dependency Check closed.'
    }
    catch {
        Set-Status "App Dependency Check failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($_.Exception.Message,'App Dependency Check','OK','Error') | Out-Null
    }
})

$LinkSite.Add_MouseLeftButtonUp({ Start-Process 'https://endpointguy.com' })

# ---------------------------------------------------------------------------
# Clock
# ---------------------------------------------------------------------------
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({ $ClockText.Text = (Get-Date).ToString('HH:mm:ss') })
$timer.Start()

$Window.Add_Closed({ $timer.Stop() })

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------
Update-ConnectionUi
$Window.ShowDialog() | Out-Null
