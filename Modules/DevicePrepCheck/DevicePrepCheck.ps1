<#
.SYNOPSIS
    Device Prep Check - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a window that checks whether a device is ready for Windows
    Autopilot provisioning, from its serial number. Three checks are run:

      1. Autopilot hash imported - the serial number has a Windows Autopilot
         device identity, so the hardware hash has been uploaded.
      2. Enabled in Entra ID - the Entra ID device object linked to the
         Autopilot record exists and is enabled. A disabled object blocks
         the Autopilot sign-in.
      3. No Intune record - there is no existing Intune managed device with
         this serial number. A stale record left behind by a previous build
         gets in the way of a clean enrolment.

      - The serial number is pre-filled from the device selected in the
        Toolkit, and can be edited to check any other device.
      - Each check shows a tick or a cross, what was found, and what
        needs attention. An overall Ready / Not ready banner sits on top.
      - A disabled Entra ID device can be enabled from the window, behind a
        confirmation that defaults to No.
      - An existing Intune record can be opened in the Intune admin center.
        This module never deletes anything.
      - The Entra ID groups the device is a member of are listed at the
        bottom, for information only.
      - Recheck runs all three checks again.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so there is no separate
    .xaml file to deploy or keep in sync.

    Every function here is Dpc-prefixed on purpose. All modules are
    dot-sourced into the SAME session by Toolkit.ps1, so a shared name would
    mean the last file loaded silently wins and could change the behaviour
    of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\DevicePrepCheck\DevicePrepCheck.ps1"
        Show-DevicePrepCheckWindow -SerialNumber $serial -Owner $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: DeviceManagementServiceConfig.Read.All,
                  DeviceManagementManagedDevices.Read.All,
                  Device.Read.All,
                  Directory.AccessAsUser.All (enable in Entra ID),
                  GroupMember.Read.All (group memberships)

    Enabling an Entra ID device also needs the signed-in admin to hold at
    least the Cloud Device Administrator role.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------------------
# Shared dialog and loading overlay. Toolkit.ps1 provides the styled
# versions; these plain stand-ins are only used if this module is run on its
# own, outside the toolkit.
# ---------------------------------------------------------------------------
if (-not (Get-Command -Name Show-ToolkitMessageBox -ErrorAction SilentlyContinue)) {
    function Show-ToolkitMessageBox {
        param([string]$Message = '', [string]$Title = '', [string]$Button = 'OK',
              [string]$Icon = 'None', [string]$Default = 'None')
        return [System.Windows.MessageBox]::Show($Message, $Title,
            [System.Windows.MessageBoxButton]$Button, [System.Windows.MessageBoxImage]$Icon,
            [System.Windows.MessageBoxResult]$Default)
    }
}
if (-not (Get-Command -Name Show-ToolkitProgress -ErrorAction SilentlyContinue)) {
    function Show-ToolkitProgress { param($Owner, [string]$Heading, [string]$Text, [string]$Detail) }
    function Update-ToolkitProgress { param([string]$Text, [string]$Detail) }
    function Close-ToolkitProgress { }
}

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$DpcXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Device Prep Check"
        Height="860" Width="1000"
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
        <SolidColorBrush x:Key="GridBorderBrush"       Color="#3D3D3D"/>
        <SolidColorBrush x:Key="InfoPanelBrush"        Color="#16324A"/>
        <SolidColorBrush x:Key="InfoPanelBorderBrush"  Color="#2D5B85"/>
        <SolidColorBrush x:Key="InfoLabelBrush"        Color="#6BB3E8"/>
        <SolidColorBrush x:Key="WarnTextBrush"         Color="#F0B429"/>
        <SolidColorBrush x:Key="ConsoleBrush"          Color="#141414"/>
        <SolidColorBrush x:Key="ConsoleTextBrush"      Color="#D4D4D4"/>
        <SolidColorBrush x:Key="OnAccentTextBrush"     Color="#FFFFFF"/>
        <SolidColorBrush x:Key="PendingIconBrush"      Color="#4D4D4D"/>

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

        <!-- One row per check: status circle, text, optional fix button. -->
        <Style x:Key="CheckIcon" TargetType="Border">
            <Setter Property="Width" Value="36"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="CornerRadius" Value="18"/>
            <Setter Property="Background" Value="{StaticResource PendingIconBrush}"/>
            <Setter Property="VerticalAlignment" Value="Top"/>
            <Setter Property="Margin" Value="0,2,16,0"/>
        </Style>

        <Style x:Key="CheckGlyph" TargetType="TextBlock">
            <Setter Property="FontSize" Value="18"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="HorizontalAlignment" Value="Center"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <Style x:Key="CheckDetail" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Foreground" Value="{StaticResource SubtleTextBrush}"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="LineHeight" Value="18"/>
            <Setter Property="Margin" Value="0,4,0,0"/>
        </Style>

        <Style x:Key="CheckAction" TargetType="TextBlock">
            <Setter Property="FontSize" Value="12.5"/>
            <Setter Property="Foreground" Value="{StaticResource WarnTextBrush}"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="LineHeight" Value="18"/>
            <Setter Property="Margin" Value="0,6,0,0"/>
            <Setter Property="Visibility" Value="Collapsed"/>
        </Style>

    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Padding="22,14">
            <StackPanel>
                <TextBlock Text="Device Prep Check" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Checks if a serial number is ready for Windows Autopilot provisioning: the hardware hash is imported, the device is enabled in Entra ID, and there is no old Intune record in the way."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Serial number -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <StackPanel>
                <TextBlock Text="SERIAL NUMBER" FontSize="10.5" FontWeight="Bold"
                           Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox x:Name="TxtSerial" Grid.Column="0" Height="38" FontSize="14"/>
                    <Button x:Name="BtnCheck" Grid.Column="1" Style="{StaticResource PrimaryButton}"
                            Content="Check device" MinWidth="150" Height="38" Margin="12,0,0,0"/>
                </Grid>
                <TextBlock Text="Enter the serial number, not the device name."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="11.5" Margin="0,7,0,0"/>
            </StackPanel>
        </Border>

        <!-- Overall result -->
        <Border x:Name="BannerBorder" Grid.Row="2" CornerRadius="6" BorderThickness="1" Padding="18,14" Margin="0,14,0,0"
                Background="{StaticResource InfoPanelBrush}" BorderBrush="{StaticResource InfoPanelBorderBrush}">
            <StackPanel>
                <TextBlock x:Name="BannerTitle" Text="Enter a serial number and select Check device."
                           FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TitleTextBrush}" TextWrapping="Wrap"/>
                <TextBlock x:Name="BannerDetail" Text="" FontSize="12.5" TextWrapping="Wrap"
                           LineHeight="18" Margin="0,6,0,0" Visibility="Collapsed"/>
            </StackPanel>
        </Border>

        <!-- Checks -->
        <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Margin="0,14,0,0">
            <StackPanel>

                <!-- 1. Autopilot -->
                <Border Style="{StaticResource Card}" Padding="18,16">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name="Icon1" Grid.Column="0" Style="{StaticResource CheckIcon}">
                            <TextBlock x:Name="Glyph1" Text="-" Style="{StaticResource CheckGlyph}"/>
                        </Border>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="1. Autopilot hash imported" Style="{StaticResource CardHeader}"/>
                            <TextBlock x:Name="Detail1" Text="Not checked yet." Style="{StaticResource CheckDetail}"/>
                            <TextBlock x:Name="Action1" Text="" Style="{StaticResource CheckAction}"/>
                        </StackPanel>
                    </Grid>
                </Border>

                <!-- 2. Entra ID -->
                <Border Style="{StaticResource Card}" Padding="18,16" Margin="0,10,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name="Icon2" Grid.Column="0" Style="{StaticResource CheckIcon}">
                            <TextBlock x:Name="Glyph2" Text="-" Style="{StaticResource CheckGlyph}"/>
                        </Border>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="2. Enabled in Entra ID" Style="{StaticResource CardHeader}"/>
                            <TextBlock x:Name="Detail2" Text="Not checked yet." Style="{StaticResource CheckDetail}"/>
                            <TextBlock x:Name="Action2" Text="" Style="{StaticResource CheckAction}"/>
                        </StackPanel>
                        <Button x:Name="BtnEnableEntra" Grid.Column="2" Style="{StaticResource GreenButton}"
                                Content="Enable in Entra ID" MinWidth="170" Height="36" Margin="16,0,0,0"
                                VerticalAlignment="Top" Visibility="Collapsed"/>
                    </Grid>
                </Border>

                <!-- 3. Intune -->
                <Border Style="{StaticResource Card}" Padding="18,16" Margin="0,10,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <Border x:Name="Icon3" Grid.Column="0" Style="{StaticResource CheckIcon}">
                            <TextBlock x:Name="Glyph3" Text="-" Style="{StaticResource CheckGlyph}"/>
                        </Border>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="3. Intune enrollment status" Style="{StaticResource CardHeader}"/>
                            <TextBlock x:Name="Detail3" Text="Not checked yet." Style="{StaticResource CheckDetail}"/>
                            <TextBlock x:Name="Action3" Text="" Style="{StaticResource CheckAction}"/>
                        </StackPanel>
                        <Button x:Name="BtnOpenIntune" Grid.Column="2" Style="{StaticResource NeutralButton}"
                                Content="Open in Intune" MinWidth="170" Height="36" Margin="16,0,0,0"
                                ToolTip="Open the existing record in the Intune admin center"
                                VerticalAlignment="Top" Visibility="Collapsed"/>
                    </Grid>
                </Border>

                <!-- Group memberships (information only) -->
                <Border Style="{StaticResource Card}" Padding="18,16" Margin="0,10,0,0">
                    <StackPanel>
                        <TextBlock Text="Group memberships" Style="{StaticResource CardHeader}"/>
                        <TextBlock x:Name="GroupsSummary" Text="Not checked yet." Style="{StaticResource CheckDetail}"/>
                        <Border x:Name="GroupsPanel" Margin="0,12,0,0" CornerRadius="4" BorderThickness="1"
                                BorderBrush="{StaticResource CardBorderBrush}" Background="{StaticResource ConsoleBrush}"
                                Visibility="Collapsed">
                            <ItemsControl x:Name="GroupsList" AlternationCount="2">
                                <ItemsControl.ItemTemplate>
                                    <DataTemplate>
                                        <Border x:Name="Row" Padding="12,8" Background="Transparent">
                                            <Grid>
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>
                                                <TextBlock Grid.Column="0" Text="{Binding Name}" FontSize="13.5"
                                                           Foreground="#F0F0F0" TextTrimming="CharacterEllipsis"
                                                           ToolTip="{Binding Name}" VerticalAlignment="Center"/>
                                                <Border x:Name="Badge" Grid.Column="1" CornerRadius="10" Padding="10,3"
                                                        Margin="12,0,0,0" MinWidth="78" VerticalAlignment="Center"
                                                        Background="#3A3A3A" BorderBrush="#5A5A5A" BorderThickness="1">
                                                    <TextBlock x:Name="BadgeText" Text="{Binding Type}" FontSize="11.5"
                                                               FontWeight="SemiBold" Foreground="#E8E8E8"
                                                               HorizontalAlignment="Center"/>
                                                </Border>
                                            </Grid>
                                        </Border>
                                        <DataTemplate.Triggers>
                                            <DataTrigger Binding="{Binding Path=(ItemsControl.AlternationIndex), RelativeSource={RelativeSource TemplatedParent}}" Value="1">
                                                <Setter TargetName="Row" Property="Background" Value="#202020"/>
                                            </DataTrigger>
                                            <DataTrigger Binding="{Binding Type}" Value="Dynamic">
                                                <Setter TargetName="Badge" Property="Background" Value="#16324A"/>
                                                <Setter TargetName="Badge" Property="BorderBrush" Value="#2D5B85"/>
                                                <Setter TargetName="BadgeText" Property="Foreground" Value="#8CC8F0"/>
                                            </DataTrigger>
                                        </DataTemplate.Triggers>
                                    </DataTemplate>
                                </ItemsControl.ItemTemplate>
                            </ItemsControl>
                        </Border>
                    </StackPanel>
                </Border>

            </StackPanel>
        </ScrollViewer>

        <!-- Diagnostics console (F12) -->
        <Border x:Name="ConsolePanel" Grid.Row="4" Visibility="Collapsed" Height="220"
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
                        <Button x:Name="BtnCopyLog"  Style="{StaticResource NeutralButton}" Content="Copy log" Margin="0,0,10,0"/>
                        <Button x:Name="BtnClearLog" Style="{StaticResource NeutralButton}" Content="Clear"/>
                    </StackPanel>
                </Grid>

                <TextBox x:Name="TxtLog" Grid.Row="1"
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
        <Border Grid.Row="5" Style="{StaticResource Card}" Padding="14,10" Margin="0,14,0,0">
            <Grid>
                <TextBlock x:Name="StatusText" Text="Ready."
                           FontSize="12.5" VerticalAlignment="Center" TextWrapping="NoWrap"
                           TextTrimming="CharacterEllipsis" Margin="0,0,280,0"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="BtnRecheck" Style="{StaticResource NeutralButton}"
                            Content="Recheck" MinWidth="110" Height="38" Margin="0,0,12,0"/>
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
# open the sink is null and the lines go to Write-Verbose instead.
# ---------------------------------------------------------------------------
$Script:DpcLogSink    = $null
$Script:DpcLogVerbose = $false

function Write-DpcLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:DpcLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:DpcLogSink) {
        try { & $Script:DpcLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-DpcLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-DpcGraphError $ErrorRecord
    Write-DpcLog "$Context failed: $detail" 'ERROR'
    if ($Script:DpcLogVerbose -and $ErrorRecord) {
        $ex = Get-DpcProperty $ErrorRecord 'Exception'
        if ($ex) { Write-DpcLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
        $pos = Get-DpcProperty $ErrorRecord 'InvocationInfo'
        if ($pos) { Write-DpcLog "At: $((Get-DpcProperty $pos 'PositionMessage' '').Trim())" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-DpcProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-DpcProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-DpcProperty $Object $Name) {
        $v = $Object.$Name
        if ($null -eq $v) { return $Default }
        return $v
    }
    return $Default
}

function Get-DpcBodyError {
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

    if (Test-DpcProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-DpcProperty $err 'code' ''
        $msg  = Get-DpcProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-DpcGraphError {
    # Invoke-MgGraphRequest hides the useful error text in the response body,
    # which lands in a different place depending on the module version.
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) { return 'Unknown error.' }
    if ($ErrorRecord -is [string]) { return $ErrorRecord }

    # 1. ErrorDetails.Message - usually the raw JSON body.
    if (Test-DpcProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-DpcProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-DpcBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-DpcProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-DpcProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-DpcProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-DpcBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-DpcProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-DpcProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-DpcProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

function ConvertTo-DpcODataString {
    # Makes a value safe to drop inside '...' in an OData $filter: single
    # quotes are doubled, then the whole thing is URL-encoded.
    param([string]$Value)
    return [uri]::EscapeDataString($Value.Replace("'", "''"))
}

function Format-DpcDate {
    # Graph dates arrive as strings or DateTime depending on the module
    # version. Returns a short local date/time, or 'never'.
    param($Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return 'never' }
    try {
        $d = [datetime]$Value
        if ($d.Year -le 1601) { return 'never' }
        return $d.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
    }
    catch { return [string]$Value }
}

# ---------------------------------------------------------------------------
# Graph
# ---------------------------------------------------------------------------
function Invoke-DpcGet {
    # GET that follows @odata.nextLink and returns every item in 'value'.
    param([Parameter(Mandatory)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next  = $Uri
    while ($next) {
        Write-DpcLog "GET $next" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        $sw.Stop()
        Write-DpcLog "  -> responded in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        foreach ($v in @(Get-DpcProperty $resp 'value' @())) {
            if ($null -ne $v) { $items.Add($v) }
        }
        $next = [string](Get-DpcProperty $resp '@odata.nextLink' '')
    }
    return $items.ToArray()
}

function Invoke-DpcWrite {
    <#
        PATCH with the full Graph error body on failure. Uses
        -SkipHttpErrorCheck when the installed Graph module has it, otherwise
        falls back to try/catch.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('PATCH')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [string]$Body
    )

    Write-DpcLog "$Method $Uri" 'GRAPH'
    if ($Body) { Write-DpcLog "  -> body $Body" 'DEBUG' }

    $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject' }
    if ($Body) { $params.Body = $Body; $params.ContentType = 'application/json' }

    $canSkipThrow = $false
    try {
        $cmd = Get-Command Invoke-MgGraphRequest -ErrorAction Stop
        $canSkipThrow = $cmd.Parameters.ContainsKey('SkipHttpErrorCheck') -and
                        $cmd.Parameters.ContainsKey('StatusCodeVariable')
    }
    catch { $canSkipThrow = $false }

    if ($canSkipThrow) {
        $code = 0
        $resp = Invoke-MgGraphRequest @params -SkipHttpErrorCheck -StatusCodeVariable 'code'
        Write-DpcLog "  -> HTTP $code" 'DEBUG'
        if ($code -ge 200 -and $code -lt 300) { return }

        $detail = Get-DpcBodyError $resp
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "HTTP $code" }
        throw "HTTP $code - $detail"
    }

    try { Invoke-MgGraphRequest @params | Out-Null }
    catch { throw (Get-DpcGraphError $_) }
}

function Get-DpcAutopilotRecord {
    <#
        Returns the Windows Autopilot device identities whose serial number
        matches exactly (case-insensitive, trimmed). Graph only offers
        'contains' on serialNumber, so the exact match is done here.
    #>
    param([Parameter(Mandatory)][string]$SerialNumber)

    $serial = $SerialNumber.Trim()
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$(ConvertTo-DpcODataString $serial)')"

    $hits = @(Invoke-DpcGet -Uri $uri | Where-Object {
        ([string](Get-DpcProperty $_ 'serialNumber' '')).Trim() -ieq $serial
    })
    Write-DpcLog "Autopilot records for '$serial': $($hits.Count)" 'INFO'
    return $hits
}

function Get-DpcEntraDevice {
    # Finds the Entra ID device object by its deviceId (the id the Autopilot
    # record calls azureActiveDirectoryDeviceId).
    param([Parameter(Mandatory)][string]$DeviceId)

    $esc = ConvertTo-DpcODataString $DeviceId
    $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$esc'&`$select=id,deviceId,displayName,accountEnabled"
    $hits = @(Invoke-DpcGet -Uri $uri)
    Write-DpcLog "Entra ID devices with deviceId $($DeviceId): $($hits.Count)" 'INFO'
    return $hits
}

function Get-DpcIntuneRecords {
    # Every Intune managed device carrying this serial number.
    param([Parameter(Mandatory)][string]$SerialNumber)

    $esc = ConvertTo-DpcODataString $SerialNumber.Trim()
    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$esc'&`$select=id,deviceName,userPrincipalName,lastSyncDateTime,enrolledDateTime,operatingSystem,serialNumber"
    $hits = @(Invoke-DpcGet -Uri $uri | Where-Object {
        ([string](Get-DpcProperty $_ 'serialNumber' '')).Trim() -ieq $SerialNumber.Trim()
    })
    Write-DpcLog "Intune records for '$($SerialNumber.Trim())': $($hits.Count)" 'INFO'
    return $hits
}

function Get-DpcIntuneRecordsByDeviceName {
    # Exact device-name lookup used only to catch a common input mistake before
    # a device name is reported as a missing Autopilot serial number.
    param([Parameter(Mandatory)][string]$DeviceName)

    $name = $DeviceName.Trim()
    $esc  = ConvertTo-DpcODataString $name
    $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$esc'&`$select=id,deviceName,serialNumber,userPrincipalName"
    $hits = @(Invoke-DpcGet -Uri $uri | Where-Object {
        ([string](Get-DpcProperty $_ 'deviceName' '')).Trim() -ieq $name
    })
    Write-DpcLog "Intune device-name matches for '$name': $($hits.Count)" 'INFO'
    return $hits
}

function Enable-DpcEntraDevice {
    # Sets accountEnabled = true on the Entra ID device object. Needs the
    # Directory.AccessAsUser.All delegated scope and at least the Cloud
    # Device Administrator role.
    param([Parameter(Mandatory)][string]$ObjectId)
    $uri  = "https://graph.microsoft.com/v1.0/devices/$ObjectId"
    $body = @{ accountEnabled = $true } | ConvertTo-Json -Compress
    Invoke-DpcWrite -Method PATCH -Uri $uri -Body $body
    Write-DpcLog "  -> Entra ID device $ObjectId enabled" 'OK'
}

function Get-DpcDeviceGroups {
    # Groups the Entra ID device object is a direct member of. Needs
    # Device.Read.All plus GroupMember.Read.All (or Group.Read.All).
    param([Parameter(Mandatory)][string]$ObjectId)

    $uri = "https://graph.microsoft.com/v1.0/devices/$ObjectId/memberOf?`$select=id,displayName,groupTypes"
    $groups = @(Invoke-DpcGet -Uri $uri | Where-Object {
        [string](Get-DpcProperty $_ '@odata.type' '') -eq '#microsoft.graph.group'
    })
    Write-DpcLog "Groups for Entra ID device $($ObjectId): $($groups.Count)" 'INFO'
    return $groups
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-DevicePrepCheckWindow {
    <#
    .SYNOPSIS
        Opens the Device Prep Check window.
    .PARAMETER SerialNumber
        Serial number to pre-fill. When supplied, the checks run as soon as
        the window opens. Leave it empty to type one in.
    .PARAMETER Owner
        The Toolkit window, so this one centres on it and stays on top.
    .PARAMETER ShowConsole
        Opens with the diagnostics console visible. F12 toggles it at any time.
    #>
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$SerialNumber = '',
        $Owner = $null,
        [switch]$ShowConsole
    )

    # --- load XAML ---------------------------------------------------------
    try {
        [xml]$xamlDoc = $DpcXamlString
        $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
        $win    = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        Show-ToolkitMessageBox -Title 'XAML load error' -Button 'OK' -Icon 'Error' -Message (
            "The Device Prep Check XAML could not be loaded:`n`n$($_.Exception.Message)") | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('TxtSerial','BtnCheck','BannerBorder','BannerTitle','BannerDetail',
                     'Icon1','Icon2','Icon3','Glyph1','Glyph2','Glyph3',
                     'Detail1','Detail2','Detail3','Action1','Action2','Action3',
                     'BtnEnableEntra','BtnOpenIntune','GroupsSummary','GroupsList','GroupsPanel',
                     'ConsolePanel','ChkVerbose','BtnCopyLog','BtnClearLog','TxtLog',
                     'StatusText','BtnRecheck','BtnClose')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            Show-ToolkitMessageBox -Title 'XAML load error' -Button 'OK' -Icon 'Error' -Message ("Control '$n' was not found in the Device Prep Check XAML.") | Out-Null
            return
        }
        $ui[$n] = $ctl
    }

    if ($Owner) { $win.Owner = $Owner }
    else { $win.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen }

    # --- state -------------------------------------------------------------
    # A hashtable, not plain variables: assigning inside an event scriptblock
    # would otherwise create a local copy and the value would be lost.
    $state = @{
        Busy          = $false
        Serial        = ''
        EntraObjectId = $null
        EntraName     = ''
        IntuneRecords = @()
    }

    $tick   = [string][char]0x2713
    $cross  = [string][char]0x2717
    $dash   = [string][char]0x2013
    $bullet = [string][char]0x2022

    $brush = {
        param([string]$Hex)
        New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($Hex))
    }

    $render = { $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render) }

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        Update-ToolkitProgress -Text $Message
        & $render
    }

    # --- diagnostics console ----------------------------------------------
    $appendLog = {
        param([string]$Line)
        $ui.TxtLog.AppendText($Line + [Environment]::NewLine)
        if ($ui.TxtLog.LineCount -gt 2000) {
            $ui.TxtLog.Text = $ui.TxtLog.Text.Substring($ui.TxtLog.GetCharacterIndexFromLineIndex(500))
        }
        $ui.TxtLog.ScrollToEnd()
    }

    # Point the module-scope sink at this window.
    $Script:DpcLogSink    = $appendLog
    $Script:DpcLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    $normalHeight  = 860
    $consoleHeight = [math]::Min(1100, [math]::Floor([System.Windows.SystemParameters]::WorkArea.Height))
    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
        $win.Height = $consoleHeight
    }

    $setBusy = {
        param([bool]$Busy, [string]$Text = 'Working...')
        $state.Busy = $Busy
        # Loading overlay with the moving bar over this window while it works.
        if ($Busy) { Show-ToolkitProgress -Owner $win -Heading 'Device Prep Check' -Text $Text } else { Close-ToolkitProgress }
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnCheck, $ui.BtnRecheck, $ui.BtnEnableEntra, $ui.BtnOpenIntune)) { $b.IsEnabled = -not $Busy }
        $ui.TxtSerial.IsEnabled = -not $Busy
        & $render
    }

    # --- check rows ----------------------------------------------------------
    # Result: Pass (green tick), Fail / Error (red cross), Skip / Pending (grey dash).
    $setCheck = {
        param([int]$N, [string]$Result, [string]$Detail, [string]$Action = '')
        $icon  = $ui["Icon$N"]
        $glyph = $ui["Glyph$N"]
        if ($Result -eq 'Pass') {
            $icon.Background = & $brush '#15803D'
            $glyph.Text = $tick
        }
        elseif ($Result -eq 'Fail' -or $Result -eq 'Error') {
            $icon.Background = & $brush '#D13438'
            $glyph.Text = $cross
        }
        else {
            $icon.Background = & $brush '#4D4D4D'
            $glyph.Text = $dash
        }

        $ui["Detail$N"].Text = $Detail
        $act = $ui["Action$N"]
        if ([string]::IsNullOrWhiteSpace($Action)) {
            $act.Text = ''
            $act.Visibility = [System.Windows.Visibility]::Collapsed
        }
        else {
            $act.Text = "What to do: $Action"
            $act.Visibility = [System.Windows.Visibility]::Visible
        }
        Write-DpcLog "Check $($N): $Result - $Detail" $(if ($Result -eq 'Pass') { 'OK' } elseif ($Result -in 'Fail','Error') { 'WARN' } else { 'INFO' })
    }

    $setBanner = {
        param([string]$Kind, [string]$Title, [string]$Detail = '')
        switch ($Kind) {
            'Ready'    { $bg = '#12351F'; $bd = '#15803D' }
            'NotReady' { $bg = '#3A1A1A'; $bd = '#D13438' }
            default    { $bg = '#16324A'; $bd = '#2D5B85' }
        }
        $ui.BannerBorder.Background  = & $brush $bg
        $ui.BannerBorder.BorderBrush = & $brush $bd
        $ui.BannerTitle.Text = $Title
        if ([string]::IsNullOrWhiteSpace($Detail)) {
            $ui.BannerDetail.Text = ''
            $ui.BannerDetail.Visibility = [System.Windows.Visibility]::Collapsed
        }
        else {
            $ui.BannerDetail.Text = $Detail
            $ui.BannerDetail.Visibility = [System.Windows.Visibility]::Visible
        }
    }

    $resetChecks = {
        foreach ($i in 1..3) { & $setCheck $i 'Pending' 'Checking...' }
        $ui.BtnEnableEntra.Visibility  = [System.Windows.Visibility]::Collapsed
        $ui.BtnOpenIntune.Visibility   = [System.Windows.Visibility]::Collapsed
        $ui.GroupsSummary.Text         = 'Checking...'
        $ui.GroupsList.ItemsSource     = $null
        $ui.GroupsPanel.Visibility     = [System.Windows.Visibility]::Collapsed
        & $setBanner 'Neutral' "Checking $($state.Serial)..."
    }

    # --- run the three checks ----------------------------------------------
    $runChecks = {
        if ($state.Busy) { return }

        $serial = ([string]$ui.TxtSerial.Text).Trim()
        if ([string]::IsNullOrWhiteSpace($serial)) {
            Show-ToolkitMessageBox -Title 'Device Prep Check' -Button 'OK' -Icon 'Warning' -Message 'Type a serial number first.' | Out-Null
            [void]$ui.TxtSerial.Focus()
            return
        }

        $ui.TxtSerial.Text   = $serial
        $state.Serial        = $serial
        $state.EntraObjectId = $null
        $state.EntraName     = ''
        $state.IntuneRecords = @()

        $actions = New-Object System.Collections.Generic.List[string]
        $failed  = 0

        Write-DpcLog "Checking serial number '$serial'" 'INFO'
        & $resetChecks
        & $setBusy $true "Checking $serial..."
        try {
            # 1. Autopilot hash imported ------------------------------------
            $ap = $null
            try {
                & $setStatus 'Checking Windows Autopilot...'
                $apHits = @(Get-DpcAutopilotRecord -SerialNumber $serial)
                if ($apHits.Count -eq 0) {
                    # A device name can look like a serial number. Before showing
                    # a misleading failure, check Intune for an exact name match.
                    $nameHits = @(Get-DpcIntuneRecordsByDeviceName -DeviceName $serial)
                    if ($nameHits.Count -gt 0) {
                        $serialOptions = @($nameHits | ForEach-Object {
                            ([string](Get-DpcProperty $_ 'serialNumber' '')).Trim()
                        } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

                        if ($serialOptions.Count -eq 1 -and $serialOptions[0] -ine $serial) {
                            $foundSerial = [string]$serialOptions[0]
                            $answer = Show-ToolkitMessageBox -Title 'Device name detected' -Button 'YesNo' -Icon 'Information' -Default 'Yes' -Message (
                                "'$serial' appears to be a device name, not a serial number.`n`n" +
                                "Its serial number is $foundSerial.`n`nUse this serial number and continue?")
                            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
                                foreach ($i in 1..3) { & $setCheck $i 'Skip' 'Not checked - enter a serial number to continue.' }
                                $ui.GroupsSummary.Text = 'Not checked - enter a serial number to continue.'
                                & $setBanner 'Neutral' 'Serial number required' 'The value entered appears to be a device name. Enter its serial number, then select Check device.'
                                & $setStatus 'Check stopped - a serial number is required.'
                                return
                            }

                            Write-DpcLog "Device name '$serial' replaced with serial number '$foundSerial'" 'INFO'
                            $serial = $foundSerial
                            $ui.TxtSerial.Text = $serial
                            $state.Serial = $serial
                            & $setBanner 'Neutral' "Checking $serial..." "The device name was replaced with its serial number."
                            & $setStatus "Using serial number $serial..."
                            $apHits = @(Get-DpcAutopilotRecord -SerialNumber $serial)
                        }
                        elseif ($serialOptions.Count -gt 1) {
                            $choices = @($serialOptions | ForEach-Object { "  $bullet $_" }) -join "`n"
                            foreach ($i in 1..3) { & $setCheck $i 'Skip' 'Not checked - enter a serial number to continue.' }
                            $ui.GroupsSummary.Text = 'Not checked - enter a serial number to continue.'
                            & $setBanner 'Neutral' 'Several devices use that name' 'Choose the correct serial number from the message, enter it above, then select Check device.'
                            & $setStatus 'Check stopped - several serial numbers match that device name.'
                            Show-ToolkitMessageBox -Title 'Device name detected' -Button 'OK' -Icon 'Warning' -Message (
                                "'$serial' is a device name shared by more than one Intune record.`n`n" +
                                "Enter the correct serial number:`n$choices") | Out-Null
                            return
                        }
                        else {
                            foreach ($i in 1..3) { & $setCheck $i 'Skip' 'Not checked - enter a serial number to continue.' }
                            $ui.GroupsSummary.Text = 'Not checked - enter a serial number to continue.'
                            & $setBanner 'Neutral' 'Device name detected' 'The Intune record does not contain a serial number. Find the device serial number and enter it above.'
                            & $setStatus 'Check stopped - the entered value is a device name.'
                            Show-ToolkitMessageBox -Title 'Device name detected' -Button 'OK' -Icon 'Warning' -Message (
                                "'$serial' appears to be a device name, but its Intune record does not contain a serial number.") | Out-Null
                            return
                        }
                    }
                }

                if ($apHits.Count -eq 0) {
                    $failed++
                    $act = 'The hardware hash needs to be imported.'
                    & $setCheck 1 'Fail' "No Windows Autopilot record was found for serial number $serial." $act
                    $actions.Add($act)
                }
                else {
                    $ap = $apHits[0]
                    $make  = [string](Get-DpcProperty $ap 'manufacturer' '')
                    $model = [string](Get-DpcProperty $ap 'model' '')
                    $tag   = [string](Get-DpcProperty $ap 'groupTag' '')
                    $hw    = ("$make $model").Trim()
                    $detail = 'The hardware hash is imported.'
                    if ($hw)  { $detail += " Hardware: $hw." }
                    $detail += " Group tag: $(if ($tag) { $tag } else { '(none)' })."
                    if ($apHits.Count -gt 1) {
                        $detail += " Note: $($apHits.Count) Autopilot records share this serial number; the first one was used."
                    }
                    & $setCheck 1 'Pass' $detail
                }
            }
            catch {
                $failed++
                $d = Write-DpcLogException $_ 'Autopilot check'
                $act = 'Check the Graph connection and the Autopilot and managed-device read permissions, then select Recheck.'
                & $setCheck 1 'Error' "Autopilot could not be checked: $d" $act
                $actions.Add($act)
            }

            # 2. Enabled in Entra ID ----------------------------------------
            try {
                & $setStatus 'Checking Entra ID...'
                if ($null -eq $ap) {
                    & $setCheck 2 'Skip' "Not checked $dash the hardware hash needs to be imported first."
                }
                else {
                    $aadId = [string](Get-DpcProperty $ap 'azureActiveDirectoryDeviceId' '')
                    $waitAct = 'The device needs to be registered in Entra ID.'
                    if ([string]::IsNullOrWhiteSpace($aadId) -or $aadId -eq '00000000-0000-0000-0000-000000000000') {
                        $failed++
                        & $setCheck 2 'Fail' 'The Autopilot record is not linked to an Entra ID device object yet.' $waitAct
                        $actions.Add($waitAct)
                    }
                    else {
                        $entra = @(Get-DpcEntraDevice -DeviceId $aadId)
                        if ($entra.Count -eq 0) {
                            $failed++
                            & $setCheck 2 'Fail' "No Entra ID device object with device id $aadId was found." $waitAct
                            $actions.Add($waitAct)
                        }
                        else {
                            $e = $entra[0]
                            $state.EntraObjectId = [string](Get-DpcProperty $e 'id' '')
                            $state.EntraName     = [string](Get-DpcProperty $e 'displayName' '')
                            $shown = if ($state.EntraName) { $state.EntraName } else { $aadId }
                            $enabled = Get-DpcProperty $e 'accountEnabled' $null
                            if ($enabled -eq $true) {
                                & $setCheck 2 'Pass' "Enabled. Entra ID device: $shown (device id $aadId)."
                            }
                            else {
                                $failed++
                                $act = 'The device needs to be enabled in Entra ID.'
                                & $setCheck 2 'Fail' "The Entra ID device $shown is disabled. A disabled device cannot complete Autopilot." $act
                                $ui.BtnEnableEntra.Visibility = [System.Windows.Visibility]::Visible
                                $actions.Add($act)
                            }
                        }
                    }
                }
            }
            catch {
                $failed++
                $d = Write-DpcLogException $_ 'Entra ID check'
                $act = 'Check the Graph connection and the Device.Read.All permission, then select Recheck.'
                & $setCheck 2 'Error' "Entra ID could not be checked: $d" $act
                $actions.Add($act)
            }

            # 3. No existing Intune record ----------------------------------
            try {
                & $setStatus 'Checking Intune...'
                $recs = @(Get-DpcIntuneRecords -SerialNumber $serial)
                $state.IntuneRecords = $recs
                if ($recs.Count -eq 0) {
                    & $setCheck 3 'Pass' 'No Intune managed device uses this serial number.'
                }
                else {
                    $failed++
                    $recLines = foreach ($r in $recs) {
                        $rn = [string](Get-DpcProperty $r 'deviceName' '(no name)')
                        $ru = [string](Get-DpcProperty $r 'userPrincipalName' '')
                        if (-not $ru) { $ru = 'no user' }
                        "  $bullet $rn $dash $ru $dash last check-in $(Format-DpcDate (Get-DpcProperty $r 'lastSyncDateTime' $null))"
                    }
                    $detail = "$($recs.Count) existing Intune record(s) found:`n" + (@($recLines) -join "`n")
                    $act = 'The device needs to be wiped or deleted from the Intune portal.'
                    & $setCheck 3 'Fail' $detail $act
                    $ui.BtnOpenIntune.Visibility = [System.Windows.Visibility]::Visible
                    $actions.Add($act)
                }
            }
            catch {
                $failed++
                $d = Write-DpcLogException $_ 'Intune check'
                $act = 'Check the Graph connection and the DeviceManagementManagedDevices.Read.All permission, then select Recheck.'
                & $setCheck 3 'Error' "Intune could not be checked: $d" $act
                $actions.Add($act)
            }

            # Group memberships (information only - not part of the result) -
            try {
                & $setStatus 'Reading group memberships...'
                if ([string]::IsNullOrWhiteSpace([string]$state.EntraObjectId)) {
                    $ui.GroupsSummary.Text = "Not checked $dash no Entra ID device was found."
                }
                else {
                    $groups = @(Get-DpcDeviceGroups -ObjectId $state.EntraObjectId)
                    if ($groups.Count -eq 0) {
                        $ui.GroupsSummary.Text = 'The device is not a member of any groups.'
                    }
                    else {
                        $rows = @(foreach ($g in $groups) {
                            $types = @(Get-DpcProperty $g 'groupTypes' @())
                            $kind  = 'Assigned'
                            if ($types -contains 'DynamicMembership') { $kind = 'Dynamic' }
                            [pscustomobject]@{
                                Name = [string](Get-DpcProperty $g 'displayName' '(no name)')
                                Type = $kind
                            }
                        })
                        $rows = @($rows | Sort-Object Name)
                        $ui.GroupsSummary.Text     = "Member of $($rows.Count) group(s)."
                        $ui.GroupsList.ItemsSource = $rows
                        $ui.GroupsPanel.Visibility = [System.Windows.Visibility]::Visible
                    }
                }
            }
            catch {
                $d = Write-DpcLogException $_ 'Group memberships'
                $ui.GroupsSummary.Text = "Group memberships could not be read: $d"
            }

            # Overall result ------------------------------------------------
            if ($failed -eq 0) {
                & $setBanner 'Ready' "$tick Ready for provisioning" "All three checks passed for $serial."
                & $setStatus "$serial is ready for provisioning."
                Write-DpcLog "$serial is ready for provisioning." 'OK'
            }
            else {
                $n = 0
                $numbered = foreach ($a in $actions) { $n++; "$n. $a" }
                & $setBanner 'NotReady' "$cross Not ready $dash $failed item(s) need attention" (@($numbered) -join "`n")
                & $setStatus "$serial is not ready: $failed item(s) need attention."
                Write-DpcLog "$serial is not ready: $failed item(s) need attention." 'WARN'
            }
        }
        finally { & $setBusy $false }
    }

    # --- fix: enable in Entra ID -------------------------------------------
    $doEnable = {
        if ($state.Busy) { return }
        $objId = [string]$state.EntraObjectId
        if ([string]::IsNullOrWhiteSpace($objId)) {
            Show-ToolkitMessageBox -Title 'Device Prep Check' -Button 'OK' -Icon 'Warning' -Message 'Run the checks first so the Entra ID device can be found.' | Out-Null
            return
        }
        $name = if ($state.EntraName) { $state.EntraName } else { $state.Serial }

        $confirm = Show-ToolkitMessageBox -Title 'Enable in Entra ID' -Button 'YesNo' -Icon 'Question' -Default 'No' -Message (
            "Enable the Entra ID device $($name)?`n`n" +
            "  Serial number: $($state.Serial)`n" +
            "  Object id:     $objId`n`n" +
            "This sets the device to Enabled in Entra ID so it can sign in during Autopilot. Nothing else about the device is changed.`n`n" +
            "Your account needs at least the Cloud Device Administrator role.")
        if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-DpcLog 'Enable cancelled at the confirmation prompt.' 'INFO'
            & $setStatus 'Cancelled - nothing was changed.'
            return
        }

        $ok = $false
        & $setBusy $true "Enabling $name in Entra ID..."
        try {
            & $setStatus "Enabling $name in Entra ID..."
            Enable-DpcEntraDevice -ObjectId $objId
            $ok = $true
            & $setStatus "$name enabled in Entra ID."
        }
        catch {
            $d = Write-DpcLogException $_ 'Enabling the Entra ID device'
            & $setStatus "Enable failed: $d"
            $hint = ''
            if ($d -match 'Authorization_RequestDenied|Insufficient privileges|403') {
                $hint = "`n`nThis usually means your account does not hold the Cloud Device Administrator role (or higher), or the Directory.AccessAsUser.All permission has not been consented for the toolkit."
            }
            Show-ToolkitMessageBox -Title 'Enable in Entra ID' -Button 'OK' -Icon 'Error' -Message ("$name could not be enabled:`n`n$d$hint") | Out-Null
        }
        finally { & $setBusy $false }

        if ($ok) { & $runChecks }
    }

    # --- open the existing Intune record(s) in the Intune admin center ------
    # Read-only: this only opens the browser. Nothing is changed in Intune.
    $doOpen = {
        $recs = @($state.IntuneRecords)
        if ($recs.Count -eq 0) {
            Show-ToolkitMessageBox -Title 'Device Prep Check' -Button 'OK' -Icon 'Information' -Message 'There is no Intune record to open. Select Recheck to check again.' | Out-Null
            return
        }
        foreach ($r in $recs) {
            $rid = [string](Get-DpcProperty $r 'id' '')
            if ([string]::IsNullOrWhiteSpace($rid)) { continue }
            $url = 'https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/mdmDeviceId/' + $rid
            Write-DpcLog "Opening Intune record $rid in the browser" 'INFO'
            try { Start-Process $url }
            catch {
                $d = Write-DpcLogException $_ 'Opening the Intune admin center'
                Show-ToolkitMessageBox -Title 'Open in Intune' -Button 'OK' -Icon 'Error' -Message ('The Intune admin center could not be opened: ' + $d) | Out-Null
                return
            }
        }
    }

    # --- events ------------------------------------------------------------
    $ui.BtnCheck.Add_Click({   & $runChecks })
    $ui.BtnRecheck.Add_Click({ & $runChecks })
    $ui.BtnEnableEntra.Add_Click({  & $doEnable })
    $ui.BtnOpenIntune.Add_Click({   & $doOpen })
    $ui.BtnClose.Add_Click({ $win.Close() })

    # Enter in the serial number box runs the checks.
    $ui.TxtSerial.Add_KeyDown({
        if ($args[1].Key -ne [System.Windows.Input.Key]::Enter) { return }
        $args[1].Handled = $true
        & $runChecks
    })

    $ui.ChkVerbose.Add_Checked({
        $Script:DpcLogVerbose = $true
        Write-DpcLog 'Verbose logging on - Graph URIs, bodies and timings will be shown.' 'INFO'
    })
    $ui.ChkVerbose.Add_Unchecked({
        $Script:DpcLogVerbose = $false
        Write-DpcLog 'Verbose logging off.' 'INFO'
    })

    $ui.BtnClearLog.Add_Click({
        $ui.TxtLog.Clear()
        Write-DpcLog 'Log cleared.' 'INFO'
    })

    $ui.BtnCopyLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
        & $setStatus 'Diagnostics log copied to the clipboard.'
    })

    # Drop the sink so a closed window is never written to.
    $win.Add_Closed({ $Script:DpcLogSink = $null })

    # F12 toggles the console without reopening the window.
    $win.Add_KeyDown({
        if ($args[1].Key -ne [System.Windows.Input.Key]::F12) { return }
        if ($ui.ConsolePanel.Visibility -eq [System.Windows.Visibility]::Visible) {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Collapsed
            $win.Height = $normalHeight
        }
        else {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
            $win.Height = $consoleHeight
        }
    })

    # --- go ----------------------------------------------------------------
    Write-DpcLog 'Device Prep Check opened.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-DpcLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'

    $ui.TxtSerial.Text = ([string]$SerialNumber).Trim()
    $win.Add_ContentRendered({
        if ([string]::IsNullOrWhiteSpace($ui.TxtSerial.Text)) {
            & $setStatus 'Type a serial number and press Enter.'
            [void]$ui.TxtSerial.Focus()
        }
        else { & $runChecks }
    })
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# sign in and prompt for a serial number so the module can be tested on its own.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -NoWelcome -Scopes @(
        'DeviceManagementServiceConfig.Read.All',
        'DeviceManagementManagedDevices.ReadWrite.All',
        'Device.Read.All',
        'Directory.AccessAsUser.All'
    )

    $serial = Read-Host 'Serial number (leave blank to type it in the window)'
    Show-DevicePrepCheckWindow -SerialNumber $serial -ShowConsole
}
