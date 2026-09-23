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
    [string]$TenantId
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

$moduleConfigPath = Join-Path $Script:ModuleRoot 'ModuleConfig.psd1'
if (Test-Path $moduleConfigPath) {
    try {
        $overrides = Import-PowerShellDataFile -Path $moduleConfigPath

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
                [System.Windows.MessageBox]::Show(
                    "AssetStatusValues in ModuleConfig.psd1 is empty, so the built-in statuses are being used.",
                    'Module configuration','OK','Warning') | Out-Null
            }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "ModuleConfig.psd1 could not be read, so the built-in settings are being used.`n`n$($_.Exception.Message)",
            'Module configuration','OK','Warning') | Out-Null
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
            [System.Windows.MessageBox]::Show(
                "Module '$moduleName' could not be loaded:`n`n$($_.Exception.Message)",
                'Module load error','OK','Warning') | Out-Null
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
$XamlString = @'
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

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Image x:Name="LogoImage" Height="40" Margin="0,0,14,0"
                           VerticalAlignment="Center" Stretch="Uniform"
                           RenderOptions.BitmapScalingMode="HighQuality"/>
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="EndpointGuy Intune Toolkit"
                                   FontSize="24" FontWeight="Bold"
                                   Foreground="{StaticResource TitleTextBrush}"/>
                        <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
                            <TextBlock Text="Device lookup and Intune administration tools"
                                       Foreground="{StaticResource SubtleTextBrush}" FontSize="13"/>
                            <TextBlock Margin="10,0,0,0" FontSize="13">
                                <Hyperlink x:Name="LnkSite"
                                           NavigateUri="https://endpointguy.com"
                                           Foreground="{StaticResource AccentTextBrush}"
                                           TextDecorations="Underline"
                                           ToolTip="https://endpointguy.com">endpointguy.com</Hyperlink>
                            </TextBlock>
                        </StackPanel>
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
                        <Button x:Name="BtnAssetStatus"  Style="{StaticResource ActionButton}"
                                Content="Update Asset Status"
                                ToolTip="Set the Intune Management name of the selected device to Assigned, In-Stock, Retired, Recycled, Stolen, Legalhold or Lost. Asks for confirmation first."/>
                    </StackPanel>

                    <StackPanel Grid.Row="3" Margin="0,8,0,0">
                        <Separator x:Name="SepBulkActions" Background="{StaticResource SeparatorBrush}" Margin="0,0,0,12"/>
                        <TextBlock x:Name="HdrBulkActions" Text="Bulk Actions" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                        <Button x:Name="BtnBulkAddGroup" Style="{StaticResource ActionButton}" Content="Bulk Add to Group"/>
                        <Button x:Name="BtnBulkAssetStatus" Style="{StaticResource ActionButton}"
                                Content="Bulk Update Asset Status"
                                ToolTip="Set the Intune Management name of many devices at once from a CSV of device names. Asks for confirmation first."/>
                        <Separator x:Name="SepReporting" Background="{StaticResource SeparatorBrush}" Margin="0,14,0,12"/>
                        <TextBlock x:Name="HdrReporting" Text="Reporting" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
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
                                <ComboBoxItem Content="Management Name"/>
                            </ComboBox>
                            <ComboBox x:Name="CmbAssetStatus" Width="170" Margin="14,0,0,0" Visibility="Collapsed" ToolTip="Asset Status lifecycle labels, from the Asset Status module."/>
                            <TextBox x:Name="TxtSearch" Width="300" Height="32" Margin="14,0,0,0"/>
                            <Button x:Name="BtnSearch" Style="{StaticResource PrimaryButton}"
                                    Content="Search" Width="110" Height="32" Margin="14,0,0,0"/>
                            <CheckBox x:Name="ChkExact" Content="Exact match only"
                                      VerticalAlignment="Center" FontSize="13" Margin="18,0,0,0"/>
                            <Button x:Name="BtnRefreshCache" Style="{StaticResource NeutralButton}"
                                    Content="Refresh Device Cache" MinHeight="32" Padding="14,6" VerticalAlignment="Center" Margin="18,0,0,0"/>
                            <Button x:Name="BtnAdvanced" Style="{StaticResource NeutralButton}"
                                    Content="Advanced Filter" MinHeight="32" Padding="14,6" VerticalAlignment="Center" Margin="18,0,0,0"
                                    ToolTip="Build a multi-column filter. All rules must match."/>
                        </StackPanel>

                        <!-- ADVANCED FILTER PANEL -->
                        <!-- Rows are generated at runtime into RulesPanel, so the number -->
                        <!-- of criteria is not fixed by the markup.                      -->
                        <Border x:Name="AdvancedPanel" Visibility="Collapsed" Margin="0,16,0,0"
                                Background="{DynamicResource InputBrush}"
                                BorderBrush="{DynamicResource InputBorderBrush}"
                                BorderThickness="1" CornerRadius="4" Padding="14">
                            <StackPanel>
                                <TextBlock Text="Match ALL of the following rules:"
                                           FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                                <StackPanel x:Name="RulesPanel"/>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnAddRule" Style="{StaticResource NeutralButton}"
                                            Content="+ Add rule" MinHeight="30" Padding="12,5"/>
                                    <Button x:Name="BtnApplyFilter" Style="{StaticResource PrimaryButton}"
                                            Content="Apply Filter" MinWidth="120" MinHeight="30" Padding="12,5" Margin="10,0,0,0"/>
                                    <Button x:Name="BtnClearFilter" Style="{StaticResource NeutralButton}"
                                            Content="Clear" MinHeight="30" Padding="12,5" Margin="10,0,0,0"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
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
                                <DataGridTextColumn Header="Source"        Binding="{Binding Source}"       Width="95"/>
                                <DataGridTextColumn Header="Device Name"   Binding="{Binding DeviceName}"   Width="200"/>
                                <DataGridTextColumn Header="Serial Number" Binding="{Binding SerialNumber}" Width="140"/>
                                <DataGridTextColumn Header="Management Name" Binding="{Binding ManagementName}" Width="200"/>
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
    [System.Windows.MessageBox]::Show(
        'No named controls could be resolved from the XAML.',
        'XAML load error','OK','Error') | Out-Null
    return
}

foreach ($required in @('BtnConnect','BtnSearch','BtnRefreshCache','GridDevices',
                        'TxtSearch','CmbSearchField','CmbAssetStatus','ChkExact','StatusText',
                        'ConnIcon','ConnStatus','ConnAccount','ResultCount',
                        'BtnAdvanced','AdvancedPanel','RulesPanel','BtnAddRule','BtnApplyFilter','BtnClearFilter',
                        'SelectedDeviceText','ClockText','BtnExportCsv',
                        'BtnCopyGroups','BtnRemoveGroups','BtnAssetStatus',
                        'BtnBulkAddGroup','BtnBulkAssetStatus','BtnAppDependency')) {
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
        'EndpointGuy Intune Toolkit','OK','Information') | Out-Null
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
    $select = 'id,deviceName,managedDeviceName,serialNumber,userPrincipalName,operatingSystem,' +
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
            ManagementName = $d.managedDeviceName
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
            Ownership       = ''
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

function Update-DeviceCache {
    if (-not (Test-Connected)) { return }
    Set-Busy $true
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Cache refresh failed','OK','Error') | Out-Null
    }
    finally { Set-Busy $false }
}

# ---------------------------------------------------------------------------
# Group actions bind to the Entra ID device record, so they work for an
# Autopilot-only device - but only once that record exists. Anything that
# needs an Intune managedDevice record is blocked for a non-enrolled device.
# ---------------------------------------------------------------------------
function Test-DeviceHasEntraRecord {
    param($Device, [string]$Action)

    $id = [string](Get-AutopilotProperty $Device 'EntraDeviceId')
    if ([string]::IsNullOrWhiteSpace($id) -or $id -eq '00000000-0000-0000-0000-000000000000') {
        [System.Windows.MessageBox]::Show(
            "$($Device.DeviceName) has no Entra ID device record yet, so its group membership cannot be read or changed.`n`nAn Autopilot device gets its Entra record when it is imported. If this device was just imported, refresh the device cache and try again.",
            $Action,'OK','Warning') | Out-Null
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
        Update-DeviceCache
        if ($Script:DeviceCache.Count -eq 0) { return }
    }

    $exact = [bool]$ChkExact.IsChecked

    Set-Busy $true
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
    "Ownership"       = "Ownership"
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
# free-text box. Compliance and Ownership are Graph enumerations, so they are
# safe to constrain too.
function Get-FilterValueList {
    param([string]$Property)

    switch ($Property) {
        "ManagementName" { return @($Script:SearchStatuses) }
        "Compliance"     { return @("compliant","noncompliant","conflict","error","inGracePeriod","unknown") }
        "Ownership"      { return @("company","personal","unknown") }
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
        Update-DeviceCache
        if ($Script:DeviceCache.Count -eq 0) { return }
    }

    Set-Busy $true
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
  catch { [System.Windows.MessageBox]::Show("Advanced filter failed to open:`n`n$($_.Exception.Message)",'Advanced filter','OK','Error') | Out-Null }
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
    if ($d) {
        $Script:SelectedDevice   = $d
        $extraLine = if ($d.Source -eq 'Autopilot') { "`nAutopilot - not enrolled in Intune ($($d.EnrollmentState))" } else { '' }
        $SelectedDeviceText.Text = "$($d.DeviceName)`nSerial: $($d.SerialNumber)`n$($d.User)$extraLine"
        Set-Status "Selected $($d.DeviceName)  |  Serial: $($d.SerialNumber)  |  Source: $($d.Source)"
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
        [System.Windows.MessageBox]::Show(
            'Select a device in the search results first.',
            'Copy Device Groups','OK','Warning') | Out-Null
        return
    }

    if (-not (Test-DeviceHasEntraRecord -Device $Script:SelectedDevice -Action 'Copy Device Groups')) { return }

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
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'RemoveDeviceGroups')) {
        Set-Status 'Remove Device Groups is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        [System.Windows.MessageBox]::Show(
            'Select a device in the search results first.',
            'Remove Device Groups','OK','Warning') | Out-Null
        return
    }

    if (-not (Test-DeviceHasEntraRecord -Device $Script:SelectedDevice -Action 'Remove Device Groups')) { return }

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
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'BulkAddToGroup')) {
        Set-Status 'Bulk Add to Group is switched off in this build.'
        return
    }

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
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'AppDependencyCheck')) {
        Set-Status 'App Dependency Check is switched off in this build.'
        return
    }

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
$BtnAssetStatus.Add_Click({
    # Switched off in the module registry at the top of this script.
    if (-not (Test-ModuleEnabled -Key 'AssetStatus')) {
        Set-Status 'Asset Status is switched off in this build.'
        return
    }

    if (-not (Test-Connected)) { return }

    if ($null -eq $Script:SelectedDevice) {
        [System.Windows.MessageBox]::Show(
            'Select a device in the search results first.',
            'Asset Status','OK','Warning') | Out-Null
        return
    }

    # The Management name lives on the Intune managedDevice record, so this
    # action needs an IntuneId - not the Entra record the group actions use.
    # An Autopilot device that has not enrolled yet has no managedDevice.
    if ([string]::IsNullOrWhiteSpace([string]$Script:SelectedDevice.IntuneId)) {
        [System.Windows.MessageBox]::Show(
            "$($Script:SelectedDevice.DeviceName) is not enrolled in Intune yet, so it has no Management name to set.`n`nA device imported into Autopilot only gets a Management name once it enrols.",
            'Asset Status','OK','Warning') | Out-Null
        Set-Status 'Asset Status needs a device that has enrolled in Intune.'
        return
    }

    if (-not (Get-Command -Name Show-AssetStatusWindow -ErrorAction SilentlyContinue)) {
        [System.Windows.MessageBox]::Show(
            "AssetStatus.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\AssetStatus')",
            'Module not available','OK','Error') | Out-Null
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Asset Status','OK','Error') | Out-Null
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
        [System.Windows.MessageBox]::Show(
            "BulkAssetStatus.ps1 was not found.`n`nExpected in:`n  $(Join-Path $Script:ModuleRoot 'Modules\BulkAssetStatus')",
            'Module not available','OK','Error') | Out-Null
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
        [System.Windows.MessageBox]::Show($_.Exception.Message,'Bulk Update Asset Status','OK','Error') | Out-Null
    }
})

# ---------------------------------------------------------------------------
# Header web link - open endpointguy.com in the default browser.
# WPF will not navigate on its own; the RequestNavigate event must be handled.
# ---------------------------------------------------------------------------
$LnkSite = $Window.FindName('LnkSite')
if ($LnkSite) {
    $LnkSite.Add_RequestNavigate({
        param($eventSender, $e)
        try { Start-Process $e.Uri.AbsoluteUri }
        catch { Set-Status "Could not open $($e.Uri.AbsoluteUri)" }
        $e.Handled = $true
    })
}

$LogoPath = Join-Path $Script:ModuleRoot 'EGLogoNew.png'
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

$Window.Add_Closed({ $timer.Stop() })

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

    # Hide a section header and its separator when every button under it is gone.
    foreach ($section in @(
        @{ Keys = @('BulkAddToGroup','BulkAssetStatus'); Header = 'HdrBulkActions'; Separator = 'SepBulkActions' }
        @{ Keys = @('AppDependencyCheck'); Header = 'HdrReporting';   Separator = 'SepReporting'   }
    )) {
        $anyVisible = $false
        foreach ($key in $section.Keys) {
            if (Test-ModuleEnabled -Key $key) { $anyVisible = $true }
        }
        if (-not $anyVisible) {
            foreach ($name in @($section.Header, $section.Separator)) {
                $el = $Window.FindName($name)
                if ($el) { $el.Visibility = [System.Windows.Visibility]::Collapsed }
            }
        }
    }

    $offCount = @($Script:ModuleRegistry | Where-Object { -not $_.Enabled }).Count
    if ($offCount -gt 0) {
        $names = ($Script:ModuleRegistry | Where-Object { -not $_.Enabled } |
                  ForEach-Object { $_.Key }) -join ', '
        Set-Status "Ready. $offCount module(s) switched off for this build: $names"
    }
}

Update-ModuleUi

# ---------------------------------------------------------------------------
# Show
# ---------------------------------------------------------------------------
Update-ConnectionUi
$Window.ShowDialog() | Out-Null