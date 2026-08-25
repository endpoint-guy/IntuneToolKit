<#
.SYNOPSIS
    App Dependency Check - module for the Endpointguy Intune Toolkit.

.DESCRIPTION
    Opens a read-only window that answers one question: which apps depend on
    the app I am about to change?

      - Every Win32 app in the tenant is read once when the window opens,
        together with its assignments, and listed A-Z in a drop-down.
      - Picking an app and selecting Find dependents walks the relationship
        graph UPWARDS and lists every app that would be affected if the
        selected app were changed, replaced or removed.
      - Direct dependents (level 1) and indirect dependents reached through
        another app are both listed, with the full chain shown.
      - NOTHING IS WRITTEN BACK. Every Graph call this module makes is a
        GET, so it is safe to run against production at any time.
      - The list can be exported to CSV.

.NOTES
    SELF-CONTAINED - the WPF XAML is embedded below, so AppDependencyCheck.xaml
    is not required at runtime; embedded XAML is used by default
    when launched from Toolkit.ps1. Pass -XamlPath to load an external
    AppDependencyCheck.xaml instead while iterating on the UI.

    Every function here is Adc-prefixed on purpose. This module,
    CopyDeviceGroups.ps1, RemoveDeviceGroups.ps1 and BulkAddToGroup.ps1 are
    dot-sourced into the SAME session by Toolkit.ps1, so a shared name would
    mean the last file loaded silently wins and could change the behaviour
    of another module.

    Dot-source it from Toolkit.ps1, then call the entry point:
        . "$PSScriptRoot\Modules\AppDependencyCheck\AppDependencyCheck.ps1"
        Show-AppDependencyCheckWindow -Owner $Window

    Requires: Windows PowerShell 5.1 (-STA), Microsoft.Graph.Authentication
    Graph scopes: DeviceManagementApps.Read.All
#>

[CmdletBinding()]
param(
    # Optional dev override: point at an external AppDependencyCheck.xaml to
    # tweak the UI without recompiling. When omitted, the embedded XAML is used.
    [string]$XamlPath
)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ---------------------------------------------------------------------------
# Types used by the window.
#
# Separate types from the other modules on purpose: Add-Type cannot redefine
# a type that is already loaded in the session, so reusing a name would leave
# whichever module loaded second without its own properties.
#
# EndpointguyAppListItem overrides ToString() so the app drop-down renders the
# app name. A [pscustomobject] cannot do that - a retemplated ComboBox falls
# back to ToString() for the closed selection box, and PowerShell renders an
# object as "@{Id=...; Label=...}", which is what put the raw GUID on screen.
# ---------------------------------------------------------------------------
if (-not ('EndpointguyAppDependentRow' -as [type]) -or -not ('EndpointguyAppListItem' -as [type])) {
    Add-Type -TypeDefinition @'
public class EndpointguyAppDependentRow
{
    public string DependentApp    { get; set; }
    public string Version         { get; set; }
    public string Publisher       { get; set; }
    public string Relationship    { get; set; }
    public string Level           { get; set; }
    public string DependencyType  { get; set; }
    public string Assigned        { get; set; }
    public string Chain           { get; set; }
    public string AppId           { get; set; }
}

public class EndpointguyAppListItem
{
    public string Id    { get; set; }
    public string Label { get; set; }

    // The drop-down shows this. Overriding ToString() keeps the closed
    // selection box correct even when the templated ContentPresenter has
    // no ContentTemplate to work with.
    public override string ToString() { return Label; }
}
'@
}

# Safety stop for the upward walk. A dependency graph should never be this
# deep; the limit only exists so a malformed loop cannot spin forever.
# Cycles are already handled by the visited set in Get-AdcDependentApps.
$Script:AdcMaxDepth = 25

# ---------------------------------------------------------------------------
# Embedded XAML
# ---------------------------------------------------------------------------
$AdcXamlString = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Dependency Check"
        Height="880" Width="1300"
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
        <SolidColorBrush x:Key="InputHoverBrush"       Color="#383838"/>
        <SolidColorBrush x:Key="DropDownBrush"         Color="#2B2B2B"/>
        <SolidColorBrush x:Key="DropDownBorderBrush"   Color="#4D4D4D"/>
        <SolidColorBrush x:Key="ComboArrowBrush"       Color="#C8C8C8"/>
        <SolidColorBrush x:Key="ItemHoverBrush"        Color="#3A3A3A"/>
        <SolidColorBrush x:Key="ItemSelectedBrush"     Color="#0F4C7A"/>

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

        <!-- Fully retemplated ComboBox: the stock control keeps a system-drawn -->
        <!-- chrome that ignores Background, so it renders unreadable dark text -->
        <!-- on a dark surface. Templating it is the only reliable fix.         -->
        <!-- Kept in sync with the ComboBox in MainWindow.xaml.                 -->
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
                <TextBlock Text="App Dependency Check" FontSize="22" FontWeight="Bold"
                           Foreground="{StaticResource TitleTextBrush}"/>
                <TextBlock Text="Read only. Pick a Win32 app and this lists every app that depends on it - the apps that would be affected if it were changed, replaced or removed. Direct and indirect dependents are both shown. Nothing is written back to Intune."
                           Foreground="{StaticResource SubtleTextBrush}" FontSize="12.5"
                           TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
        </Border>

        <!-- Scope -->
        <Border Grid.Row="1" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <StackPanel>
                <TextBlock Text="Scope" Style="{StaticResource CardHeader}" Margin="0,0,0,12"/>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Show apps that depend on:" VerticalAlignment="Center"
                               FontSize="13" Margin="0,0,10,0"/>
                    <ComboBox x:Name="CmbApp" Width="430" Margin="0,0,14,0"
                              MaxDropDownHeight="420" VerticalAlignment="Center"
                              IsEnabled="False"
                              ToolTip="Every Win32 app in the tenant, listed A-Z. With the list open, type the first few letters to jump to an app."/>
                    <CheckBox x:Name="ChkIncludeSupersedence" Content="Include supersedence"
                              IsChecked="True" VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"
                              ToolTip="Also list apps that supersede the selected app, not just apps that depend on it."/>
                    <Button x:Name="BtnScan" Style="{StaticResource PrimaryButton}"
                            Content="Find dependents" MinWidth="150" Height="34" Margin="0,0,10,0"/>
                    <Button x:Name="BtnReload" Style="{StaticResource NeutralButton}"
                            Content="Reload app list" Height="34"
                            ToolTip="Re-reads every Win32 app from Graph and rebuilds the list."/>
                </StackPanel>

                <Border CornerRadius="4" Margin="0,14,0,0"
                        Background="{StaticResource InfoPanelBrush}"
                        BorderBrush="{StaticResource InfoPanelBorderBrush}" BorderThickness="1"
                        Padding="12,10">
                    <StackPanel>
                        <TextBlock Text="RESULT" FontSize="10.5" FontWeight="Bold"
                                   Foreground="{StaticResource InfoLabelBrush}" Margin="0,0,0,6"/>
                        <TextBlock x:Name="TxtSummary" Text="Loading the app list..."
                                   FontSize="12.5" TextWrapping="Wrap" LineHeight="18"/>
                    </StackPanel>
                </Border>
            </StackPanel>
        </Border>

        <!-- Dependents -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="18" Margin="0,14,0,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <Grid Grid.Row="0" Margin="0,0,0,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <TextBlock Text="Dependent apps" Style="{StaticResource CardHeader}"/>
                        <TextBlock x:Name="FindingCount" Text="" Margin="12,0,0,0"
                                   VerticalAlignment="Bottom" FontSize="13"
                                   Foreground="{StaticResource SubtleTextBrush}"/>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <CheckBox x:Name="ChkDirectOnly" Content="Direct dependents only"
                                  VerticalAlignment="Center" FontSize="13" Margin="0,0,16,0"
                                  ToolTip="Hide apps that only depend on the selected app through another app."/>
                    </StackPanel>
                </Grid>

                <DataGrid x:Name="GridFindings" Grid.Row="1"
                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          IsReadOnly="True"
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
                        <DataGridTextColumn Header="Dependent app"  Binding="{Binding DependentApp}"   Width="300"/>
                        <DataGridTextColumn Header="Version"        Binding="{Binding Version}"        Width="110"/>
                        <DataGridTextColumn Header="Publisher"      Binding="{Binding Publisher}"      Width="150"/>
                        <DataGridTextColumn Header="Relationship"   Binding="{Binding Relationship}"   Width="120"/>
                        <DataGridTextColumn Header="Level"          Binding="{Binding Level}"          Width="110"/>
                        <DataGridTextColumn Header="Install"        Binding="{Binding DependencyType}" Width="100"/>
                        <DataGridTextColumn Header="Assigned"       Binding="{Binding Assigned}"       Width="85"/>
                        <DataGridTextColumn Header="Chain"          Binding="{Binding Chain}"          Width="*"/>
                        <DataGridTextColumn Header="App Id"         Binding="{Binding AppId}"          Width="240"/>
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
                        <TextBlock Text="Live Graph calls, relationship walks and errors" Margin="12,0,0,0"
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
                    <Button x:Name="BtnExportCsv" Style="{StaticResource NeutralButton}"
                            Content="Export list to CSV" MinWidth="200" Height="38" Margin="0,0,12,0"/>
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
$Script:AdcLogSink    = $null
$Script:AdcLogVerbose = $false

function Write-AdcLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','OK','WARN','ERROR','GRAPH','DEBUG')]
        [string]$Level = 'INFO'
    )

    # DEBUG lines only surface when the Verbose box is ticked.
    if ($Level -eq 'DEBUG' -and -not $Script:AdcLogVerbose) { return }

    $line = '[{0}] {1,-5} {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message

    if ($Script:AdcLogSink) {
        try { & $Script:AdcLogSink $line } catch { Write-Verbose $line }
    }
    else { Write-Verbose $line }
}

function Write-AdcLogException {
    # One place that turns an ErrorRecord into a readable console entry.
    param($ErrorRecord, [string]$Context = 'Operation')
    $detail = Get-AdcGraphError $ErrorRecord
    Write-AdcLog "$Context failed: $detail" 'ERROR'
    if ($Script:AdcLogVerbose -and $ErrorRecord) {
        $ex = Get-AdcProperty $ErrorRecord 'Exception'
        if ($ex) { Write-AdcLog "Exception type: $($ex.GetType().FullName)" 'DEBUG' }
    }
    return $detail
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Test-AdcProperty {
    # Set-StrictMode -Version Latest makes a missing property a terminating
    # error, so every Graph response property is probed before it is read.
    param($Object, [string]$Name)
    return ($null -ne $Object) -and
           ($Object.PSObject.Properties.Name -contains $Name)
}

function Get-AdcProperty {
    param($Object, [string]$Name, $Default = $null)
    if (Test-AdcProperty $Object $Name) { return $Object.$Name }
    return $Default
}

function Get-AdcGraphPaged {
    # Walks @odata.nextLink and returns every item in the 'value' array.
    param([Parameter(Mandatory)][string]$Uri)
    $all = New-Object System.Collections.Generic.List[object]
    $page = 0
    while ($Uri) {
        $page++
        Write-AdcLog "GET $Uri" 'GRAPH'
        $sw   = [System.Diagnostics.Stopwatch]::StartNew()
        $resp = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject
        $sw.Stop()

        $count = 0
        if (Test-AdcProperty $resp 'value') { $count = @($resp.value).Count; $all.AddRange(@($resp.value)) }
        Write-AdcLog "  -> page $page returned $count item(s) in $($sw.ElapsedMilliseconds) ms" 'DEBUG'

        $Uri = if (Test-AdcProperty $resp '@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
        if ($Uri) { Write-AdcLog '  -> following @odata.nextLink for the next page' 'DEBUG' }
    }
    return $all
}

function Get-AdcBodyError {
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

    if (Test-AdcProperty $obj 'error') {
        $err = $obj.error
        if ($err -is [System.Collections.IDictionary]) { $err = [pscustomobject]$err }
        $code = Get-AdcProperty $err 'code' ''
        $msg  = Get-AdcProperty $err 'message' ''
        if (-not [string]::IsNullOrWhiteSpace($msg)) {
            if ($code) { return "$code - $msg" }
            return $msg
        }
    }
    return $null
}

function Get-AdcGraphError {
    <#
        Invoke-MgGraphRequest throws a generic exception whose Message is only
        'Response status code does not indicate success: Forbidden
        (Forbidden).' The useful text is in the response body, which lands in
        a different place depending on the module version - hence every probe
        below. Whatever is found is returned as a readable string.
    #>
    param($ErrorRecord)

    if ($null -eq $ErrorRecord) { return 'Unknown error.' }

    # 1. ErrorDetails.Message - usually the raw JSON body.
    if (Test-AdcProperty $ErrorRecord 'ErrorDetails') {
        $raw = Get-AdcProperty $ErrorRecord.ErrorDetails 'Message'
        $parsed = Get-AdcBodyError $raw
        if ($parsed) { return $parsed }
        if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
    }

    # 2. The exception's own Response stream.
    $ex = Get-AdcProperty $ErrorRecord 'Exception'
    if ($ex) {
        $resp = Get-AdcProperty $ex 'Response'
        if ($resp) {
            try {
                $content = Get-AdcProperty $resp 'Content'
                if ($content) {
                    $raw = $content.ReadAsStringAsync().GetAwaiter().GetResult()
                    $parsed = Get-AdcBodyError $raw
                    if ($parsed) { return $parsed }
                    if (-not [string]::IsNullOrWhiteSpace($raw)) { return $raw }
                }
            }
            catch { }
        }

        $inner = Get-AdcProperty $ex 'InnerException'
        if ($inner) {
            $m = Get-AdcProperty $inner 'Message' ''
            if (-not [string]::IsNullOrWhiteSpace($m)) { return $m }
        }

        return (Get-AdcProperty $ex 'Message' 'Unknown error.')
    }

    return "$ErrorRecord"
}

# ---------------------------------------------------------------------------
# App inventory
# ---------------------------------------------------------------------------
function Get-AdcWin32Apps {
    <#
        Reads every Win32 app in the tenant.

        mobileApps holds every app type, and only win32LobApp carries
        dependency and supersedence relationships, so the collection has to
        be narrowed to that type.

        The narrowing is done with $filter=isof(...) rather than a
        /microsoft.graph.win32LobApp cast segment: the Intune app service
        does not publish an OData route for the cast and answers it with
        400 "No method match route template".

        No $select is used either. On this collection the declared type is
        mobileApp, so selecting a win32LobApp-only property such as
        displayVersion is rejected as an unknown property. Graph returns
        the full objects, and the fields are picked off the response below.
    #>
    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$top=200"

    Write-AdcLog 'Reading Win32 app inventory...' 'INFO'
    $apps = @(Get-AdcGraphPaged -Uri $uri)
    Write-AdcLog "Found $($apps.Count) Win32 app(s)." 'OK'

    $map = @{}
    foreach ($a in $apps) {
        $id = Get-AdcProperty $a 'id' ''
        if (-not $id) { continue }

        # isof() also matches any subtype, so confirm the concrete type
        # before the app is treated as a Win32 app.
        $odataType = [string](Get-AdcProperty $a '@odata.type' '')
        if ($odataType -and $odataType -notmatch 'win32LobApp') { continue }

        $map[$id] = [pscustomobject]@{
            Id            = $id
            DisplayName   = Get-AdcProperty $a 'displayName' '(unnamed app)'
            Publisher     = Get-AdcProperty $a 'publisher' ''
            Version       = Get-AdcProperty $a 'displayVersion' ''
            Assigned      = $null   # filled in by Add-AdcAssignmentState
            Relationships = $null   # filled in on demand by Get-AdcRelationships
        }
    }
    return $map
}

function Add-AdcAssignmentState {
    <#
        Marks every app as assigned or unassigned.

        An unassigned dependency is the most common reason a dependent app
        never installs: Intune will only install a dependency it can resolve
        to an assignment, so a dependency sitting in the tenant with no
        groups targeted at it silently blocks the parent.
    #>
    param([Parameter(Mandatory)][hashtable]$AppMap)

    Write-AdcLog 'Reading app assignments...' 'INFO'

    # One call per app would be hundreds of round trips. $expand pulls the
    # assignments back with the app list in a single paged read instead.
    # Same isof() filter as the inventory call, and for the same reason.
    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$expand=assignments&`$top=200"

    $seen = 0
    foreach ($a in (Get-AdcGraphPaged -Uri $uri)) {
        $id = Get-AdcProperty $a 'id' ''
        if (-not $id -or -not $AppMap.ContainsKey($id)) { continue }
        $assignments = @(Get-AdcProperty $a 'assignments' @())
        $AppMap[$id].Assigned = ($assignments.Count -gt 0)
        $seen++
    }

    # Anything the expand did not return is treated as unassigned rather than
    # unknown - it keeps the report definite instead of hedging.
    foreach ($k in @($AppMap.Keys)) {
        if ($null -eq $AppMap[$k].Assigned) { $AppMap[$k].Assigned = $false }
    }

    $assignedCount = @($AppMap.Values | Where-Object { $_.Assigned }).Count
    Write-AdcLog "Assignment state resolved for $seen app(s): $assignedCount assigned, $($AppMap.Count - $assignedCount) unassigned." 'OK'
}

function Get-AdcRelationships {
    <#
        Reads the relationships declared on one app.

        The endpoint returns BOTH directions in one collection, and
        targetType says which way each edge points, from the perspective of
        the app being queried:

          child  - the target sits BELOW this app. This app depends on the
                   target, or supersedes it.
          parent - the target sits ABOVE this app. The target depends on
                   this app, or is superseded by it.

        This module answers "what depends on X", so only the parent edges
        are of interest - they point at the apps that would break if X were
        changed. The direction is selected by the caller.
    #>
    param(
        [Parameter(Mandatory)][string]$AppId,
        [ValidateSet('parent','child')][string]$Direction = 'parent',
        [switch]$IncludeSupersedence
    )

    $uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$AppId/relationships"
    $edges = New-Object System.Collections.Generic.List[object]

    foreach ($r in (Get-AdcGraphPaged -Uri $uri)) {
        $type = Get-AdcProperty $r '@odata.type' ''
        $kind =
            if     ($type -match 'Dependency')   { 'Dependency' }
            elseif ($type -match 'Supersedence') { 'Supersedence' }
            else                                 { 'Other' }

        if ($kind -eq 'Other') { continue }
        if ($kind -eq 'Supersedence' -and -not $IncludeSupersedence) { continue }

        $targetId   = Get-AdcProperty $r 'targetId' ''
        $sourceId   = Get-AdcProperty $r 'sourceId' ''
        $targetType = [string](Get-AdcProperty $r 'targetType' '')

        if (-not $targetId) { continue }

        # targetType is the reliable discriminator and is returned by
        # default. sourceId is only a fallback: Graph does not always
        # populate it, and when it is absent an inbound copy of the edge
        # cannot be told apart from an outbound one.
        if ($targetType) {
            if ($targetType -ne $Direction) { continue }
        }
        elseif ($sourceId -and $sourceId -ne $AppId) {
            continue
        }

        $edges.Add([pscustomobject]@{
            Kind           = $kind
            TargetId       = $targetId
            TargetName     = Get-AdcProperty $r 'targetDisplayName' '(unknown app)'
            TargetVersion  = Get-AdcProperty $r 'targetDisplayVersion' ''
            TargetPublisher = Get-AdcProperty $r 'targetPublisher' ''
            DependencyType = Get-AdcProperty $r 'dependencyType' ''
        }) | Out-Null
    }

    return $edges
}

# ---------------------------------------------------------------------------
# Dependent lookup
# ---------------------------------------------------------------------------
function Get-AdcAppLabel {
    # Prefers the inventory name over the name carried on the edge: the edge
    # copy is a snapshot and goes stale when an app is renamed.
    param([hashtable]$AppMap, [string]$AppId, [string]$Fallback)
    if ($AppMap.ContainsKey($AppId)) { return $AppMap[$AppId].DisplayName }
    if ([string]::IsNullOrWhiteSpace($Fallback)) { return "(app $AppId)" }
    return $Fallback
}

function Get-AdcDependentApps {
    <#
        Returns every app that depends on $RootId, directly or indirectly.

        This is a breadth-first walk UP the relationship graph. Breadth-first
        matters: it reaches each app by its shortest chain, so an app that is
        both a direct dependent and a distant one is reported at level 1,
        which is the level an operator would act on.

        $visited holds every app id already queued. It is what makes a cycle
        safe - a loop simply finds nothing new to add and the walk drains -
        and it also stops a diamond from being reported twice.

        One Graph call is made per app reached, so the cost scales with the
        size of the dependent tree, not with the size of the tenant.
    #>
    param(
        [Parameter(Mandatory)][hashtable]$AppMap,
        [Parameter(Mandatory)][string]$RootId,
        [switch]$IncludeSupersedence,
        $Progress = $null
    )

    $rows    = New-Object System.Collections.Generic.List[object]
    $visited = New-Object System.Collections.Generic.HashSet[string]
    $queue   = New-Object System.Collections.Generic.Queue[object]

    $visited.Add($RootId) | Out-Null
    $rootName = Get-AdcAppLabel $AppMap $RootId $null

    # Seed with the root. Chain is the human-readable path back to the root.
    $queue.Enqueue([pscustomobject]@{ Id = $RootId; Level = 0; Chain = @($rootName) })

    $examined = 0

    while ($queue.Count -gt 0) {
        $node = $queue.Dequeue()

        if ($node.Level -ge $Script:AdcMaxDepth) {
            Write-AdcLog "Stopping at level $($node.Level) - depth guard reached." 'WARN'
            continue
        }

        $examined++
        if ($Progress) { & $Progress $examined $queue.Count (Get-AdcAppLabel $AppMap $node.Id $null) }

        try {
            $edges = @(Get-AdcRelationships -AppId $node.Id -Direction 'parent' `
                                            -IncludeSupersedence:$IncludeSupersedence)
        }
        catch {
            Write-AdcLog "Could not read relationships for '$(Get-AdcAppLabel $AppMap $node.Id $null)': $(Get-AdcGraphError $_)" 'WARN'
            continue
        }

        foreach ($edge in $edges) {
            if ($visited.Contains($edge.TargetId)) { continue }
            $visited.Add($edge.TargetId) | Out-Null

            $level = $node.Level + 1
            $name  = Get-AdcAppLabel $AppMap $edge.TargetId $edge.TargetName
            $chain = @($node.Chain) + $name

            # Prefer the inventory record; fall back to the edge snapshot for
            # an app that is outside the Win32 list.
            $known     = $AppMap.ContainsKey($edge.TargetId)
            $version   = if ($known) { $AppMap[$edge.TargetId].Version }   else { $edge.TargetVersion }
            $publisher = if ($known) { $AppMap[$edge.TargetId].Publisher } else { $edge.TargetPublisher }

            $assigned =
                if (-not $known)                        { 'Unknown' }
                elseif ($AppMap[$edge.TargetId].Assigned) { 'Yes' }
                else                                    { 'No' }

            $row = New-Object EndpointguyAppDependentRow
            $row.DependentApp   = $name
            $row.Version        = $version
            $row.Publisher      = $publisher
            $row.Relationship   = $edge.Kind
            $row.Level          = if ($level -eq 1) { 'Direct' } else { "Indirect ($level)" }
            $row.DependencyType = $edge.DependencyType
            $row.Assigned       = $assigned
            $row.Chain          = ($chain -join ' <- ')
            $row.AppId          = $edge.TargetId
            $rows.Add($row) | Out-Null

            Write-AdcLog "  level $level | $name | $($edge.Kind)" 'DEBUG'

            $queue.Enqueue([pscustomobject]@{ Id = $edge.TargetId; Level = $level; Chain = $chain })
        }
    }

    Write-AdcLog "Examined $examined app(s); found $($rows.Count) dependent(s) of '$rootName'." 'OK'
    return $rows
}

# ---------------------------------------------------------------------------
# Window
# ---------------------------------------------------------------------------
function Show-AppDependencyCheckWindow {
    <#
    .SYNOPSIS
        Opens the App Dependency Check window.
    .PARAMETER Owner
        The Toolkit window, so this one centres on it and stays on top.
    .PARAMETER ShowConsole
        Opens with the diagnostics console visible. Leave it off for normal
        use; turn it on when troubleshooting a Graph error. F12 toggles it at
        any time.
    #>
    [CmdletBinding()]
    param(
        $Owner = $null,
        [string]$XamlPath,
        [switch]$ShowConsole
    )

    # --- load XAML ---------------------------------------------------------
    $xamlText = $AdcXamlString
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
            "The App Dependency Check XAML could not be loaded:`n`n$($_.Exception.Message)",
            'XAML load error','OK','Error') | Out-Null
        return
    }

    # --- resolve controls --------------------------------------------------
    $ui = @{}
    foreach ($n in @('CmbApp','ChkIncludeSupersedence','BtnScan','BtnReload',
                     'TxtSummary','GridFindings','FindingCount','ChkDirectOnly',
                     'StatusText','BtnExportCsv','BtnClose',
                     'TxtLog','ChkVerbose','ChkAutoScroll','BtnCopyLog','BtnSaveLog','BtnClearLog',
                     'ConsolePanel')) {
        $ctl = $win.FindName($n)
        if ($null -eq $ctl) {
            [System.Windows.MessageBox]::Show("Control '$n' was not found in the App Dependency Check XAML.",
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
        AppMap   = $null    # id -> app object, cached across scans
        Dependents = @()   # every dependent found by the last search
        Busy     = $false
        RootName   = ''
    }

    $findingRows = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $ui.GridFindings.ItemsSource = $findingRows

    $setStatus = {
        param([string]$Message)
        $ui.StatusText.Text = $Message
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    # --- diagnostics console ----------------------------------------------
    # The sink is what Write-AdcLog calls. Appending to a TextBox is cheap
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
    $Script:AdcLogSink    = $appendLog
    $Script:AdcLogVerbose = [bool]$ui.ChkVerbose.IsChecked

    if ($ShowConsole) {
        $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
        $win.Height = 1040
    }

    $setBusy = {
        param([bool]$Busy)
        $state.Busy = $Busy
        $win.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
        foreach ($b in @($ui.BtnScan, $ui.BtnReload, $ui.BtnExportCsv)) { $b.IsEnabled = -not $Busy }
        # The app list stays disabled until it has actually been populated.
        $ui.CmbApp.IsEnabled = (-not $Busy) -and ($null -ne $ui.CmbApp.ItemsSource)
        $win.Dispatcher.Invoke([action]{}, [Windows.Threading.DispatcherPriority]::Render)
    }

    $refreshGrid = {
        # Rebuilds the visible rows, honouring the 'Direct dependents only' toggle.
        $directOnly = [bool]$ui.ChkDirectOnly.IsChecked
        $findingRows.Clear()
        foreach ($r in $state.Dependents) {
            if (-not $directOnly -or $r.Level -eq 'Direct') { $findingRows.Add($r) }
        }

        $direct   = @($state.Dependents | Where-Object { $_.Level -eq 'Direct' }).Count
        $indirect = @($state.Dependents).Count - $direct
        $ui.FindingCount.Text = "$direct direct, $indirect indirect"
    }

    # --- app list ----------------------------------------------------------
    # The dropdown is the only way to choose scope, so it is filled as soon as
    # the window opens. Apps are listed A-Z with the version appended, because
    # the same app name commonly appears more than once in a tenant.
    $loadAppList = {
        param([bool]$ForceReload)

        & $setBusy $true
        try {
            if ($ForceReload -or $null -eq $state.AppMap) {
                & $setStatus 'Reading Win32 app inventory from Graph...'
                $state.AppMap = Get-AdcWin32Apps
                Add-AdcAssignmentState -AppMap $state.AppMap
            }

            # Keep the current selection across a reload where possible.
            $previousId = ''
            if ($ui.CmbApp.SelectedItem) { $previousId = [string]$ui.CmbApp.SelectedItem.Id }

            $items = New-Object System.Collections.ObjectModel.ObservableCollection[object]

            $placeholder = New-Object EndpointguyAppListItem
            $placeholder.Id    = ''
            $placeholder.Label = 'Select an app...'
            $items.Add($placeholder) | Out-Null

            # App names repeat in most tenants, so the version is appended only
            # when it is actually needed to tell two entries apart. Everything
            # else in the list stays as the plain app name.
            $nameCounts = @{}
            foreach ($app in $state.AppMap.Values) {
                $key = [string]$app.DisplayName
                if ($nameCounts.ContainsKey($key)) { $nameCounts[$key] = $nameCounts[$key] + 1 }
                else                              { $nameCounts[$key] = 1 }
            }

            # Sort-Object is culture-aware and case-insensitive, which is what
            # an operator scanning a list by eye expects.
            foreach ($app in ($state.AppMap.Values | Sort-Object DisplayName)) {
                $label = [string]$app.DisplayName
                if ($nameCounts[$label] -gt 1 -and
                    -not [string]::IsNullOrWhiteSpace($app.Version)) {
                    $label = "$label  (v$($app.Version))"
                }

                $item = New-Object EndpointguyAppListItem
                $item.Id    = $app.Id
                $item.Label = $label
                $items.Add($item) | Out-Null
            }

            # Set the paths before ItemsSource, or the first render shows the
            # object type name instead of the label.
            $ui.CmbApp.DisplayMemberPath = 'Label'
            $ui.CmbApp.SelectedValuePath = 'Id'
            $ui.CmbApp.ItemsSource       = $items
            $ui.CmbApp.SelectedIndex     = 0

            if ($previousId) {
                $match = @($items | Where-Object { $_.Id -eq $previousId })
                if ($match.Count -gt 0) { $ui.CmbApp.SelectedItem = $match[0] }
            }

            $ui.CmbApp.IsEnabled = $true

            $ui.TxtSummary.Text = "$($state.AppMap.Count) Win32 app(s) loaded, listed A-Z. Pick the app you are about to change, then select Find dependents."
            & $setStatus "App list ready - $($state.AppMap.Count) Win32 app(s)."
            Write-AdcLog "App list populated with $($state.AppMap.Count) app(s), sorted A-Z." 'OK'
        }
        catch {
            $detail = Write-AdcLogException $_ 'Loading the app list'
            & $setStatus "Could not load the app list: $detail"
            $ui.TxtSummary.Text = "The app list could not be loaded: $detail"
            [System.Windows.MessageBox]::Show($detail,'App Dependency Check','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- find dependents ---------------------------------------------------
    $doScan = {
        # The app list is loaded when the window opens and refreshed by
        # Reload app list; searching only ever reads what is already cached.
        if ($null -eq $state.AppMap) { & $loadAppList $false }
        if ($null -eq $state.AppMap) { return }

        $selected = $ui.CmbApp.SelectedItem
        if ($null -eq $selected -or [string]::IsNullOrWhiteSpace([string]$selected.Id)) {
            [System.Windows.MessageBox]::Show('Pick an app from the list first.',
                'App Dependency Check','OK','Warning') | Out-Null
            return
        }

        $rootId = [string]$selected.Id
        if (-not $state.AppMap.ContainsKey($rootId)) {
            & $setStatus 'That app is no longer in the cached list - select Reload app list.'
            return
        }

        $rootName = $state.AppMap[$rootId].DisplayName
        $state.RootName = $rootName

        & $setBusy $true
        try {
            $includeSup = [bool]$ui.ChkIncludeSupersedence.IsChecked
            Write-AdcLog "Finding apps that depend on '$rootName'; supersedence $(if ($includeSup) { 'included' } else { 'excluded' })." 'INFO'

            $progress = {
                param($done, $remaining, $name)
                & $setStatus "Walking the dependency graph - $done app(s) examined, $remaining queued: $name"
            }

            $found = @(Get-AdcDependentApps -AppMap $state.AppMap -RootId $rootId `
                                            -IncludeSupersedence:$includeSup -Progress $progress)

            # Direct first, then by depth, then alphabetically.
            $state.Dependents = @($found | Sort-Object `
                @{Expression = { if ($_.Level -eq 'Direct') { 0 } else { 1 } }}, `
                @{Expression = { $_.Level }}, DependentApp)

            & $refreshGrid

            $direct   = @($state.Dependents | Where-Object { $_.Level -eq 'Direct' }).Count
            $indirect = @($state.Dependents).Count - $direct

            if (@($state.Dependents).Count -eq 0) {
                $ui.TxtSummary.Text = "Nothing depends on '$rootName'. No other Win32 app lists it as a dependency$(if ($includeSup) { ' or supersedes it' }), so changing or removing it will not break another app through a relationship."
                & $setStatus "No apps depend on '$rootName'."
            }
            else {
                $ui.TxtSummary.Text = "$(@($state.Dependents).Count) app(s) depend on '$rootName' - $direct direct and $indirect indirect. These are the apps affected if it is changed, replaced or removed."
                & $setStatus "Found $(@($state.Dependents).Count) dependent(s) of '$rootName' - $direct direct, $indirect indirect."
            }
        }
        catch {
            $detail = Write-AdcLogException $_ 'Finding dependents'
            & $setStatus "Search failed: $detail"
            [System.Windows.MessageBox]::Show($detail,'App Dependency Check','OK','Error') | Out-Null
        }
        finally { & $setBusy $false }
    }

    # --- export ------------------------------------------------------------
    $doExport = {
        if (@($state.Dependents).Count -eq 0) {
            [System.Windows.MessageBox]::Show('Run a search first - there is nothing to export.',
                'App Dependency Check','OK','Information') | Out-Null
            return
        }

        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = 'CSV file (*.csv)|*.csv'
        $dlg.FileName = "AppDependents-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            # Exports what is on screen, so the Direct-only toggle is
            # reflected in the file rather than quietly ignored.
            @($findingRows) |
                Select-Object DependentApp, Version, Publisher, Relationship,
                              Level, DependencyType, Assigned, Chain, AppId |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8

            & $setStatus "Exported $(@($findingRows).Count) dependent(s) to $($dlg.FileName)"
            Write-AdcLog "Exported $(@($findingRows).Count) dependent(s) to $($dlg.FileName)" 'OK'
        }
        catch {
            Write-AdcLog "Export failed: $($_.Exception.Message)" 'ERROR'
            [System.Windows.MessageBox]::Show($_.Exception.Message,'Export failed','OK','Error') | Out-Null
        }
    }

    # --- events ------------------------------------------------------------
    $ui.ChkDirectOnly.Add_Checked({   & $refreshGrid })
    $ui.ChkDirectOnly.Add_Unchecked({ & $refreshGrid })

    $ui.BtnScan.Add_Click({   & $doScan })
    $ui.BtnReload.Add_Click({ & $loadAppList $true })

    $ui.ChkVerbose.Add_Checked({
        $Script:AdcLogVerbose = $true
        Write-AdcLog 'Verbose logging on - per-app relationship walks and Graph timings will be shown.' 'INFO'
    })
    $ui.ChkVerbose.Add_Unchecked({
        $Script:AdcLogVerbose = $false
        Write-AdcLog 'Verbose logging off.' 'INFO'
    })

    $ui.BtnClearLog.Add_Click({
        $ui.TxtLog.Clear()
        Write-AdcLog 'Log cleared.' 'INFO'
    })

    $ui.BtnCopyLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        [System.Windows.Clipboard]::SetText($ui.TxtLog.Text)
        & $setStatus 'Diagnostics log copied to the clipboard.'
    })

    $ui.BtnSaveLog.Add_Click({
        if ([string]::IsNullOrWhiteSpace($ui.TxtLog.Text)) { return }
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $file  = Join-Path ([Environment]::GetFolderPath('Desktop')) "AppDependencyCheck-$stamp.log"
        try {
            $ui.TxtLog.Text | Set-Content -Path $file -Encoding UTF8
            & $setStatus "Log saved to $file"
            Write-AdcLog "Log saved to $file" 'OK'
        }
        catch { Write-AdcLog "Could not save the log: $($_.Exception.Message)" 'ERROR' }
    })

    $ui.BtnExportCsv.Add_Click({ & $doExport })
    $ui.BtnClose.Add_Click({     $win.Close() })

    # Drop the sink so a closed window is never written to.
    $win.Add_Closed({ $Script:AdcLogSink = $null })

    # F12 toggles the console without reopening the window.
    $win.Add_KeyDown({
        if ($args[1].Key -ne 'F12') { return }
        if ($ui.ConsolePanel.Visibility -eq [System.Windows.Visibility]::Visible) {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Collapsed
            $win.Height = 880
        }
        else {
            $ui.ConsolePanel.Visibility = [System.Windows.Visibility]::Visible
            $win.Height = 1040
        }
    })

    # --- go ----------------------------------------------------------------
    Write-AdcLog 'App Dependency Check opened. This module only ever reads from Graph.' 'INFO'
    $gm = Get-Module Microsoft.Graph.Authentication | Select-Object -First 1
    $gv = if ($gm) { $gm.Version } else { 'not loaded' }
    Write-AdcLog "PowerShell $($PSVersionTable.PSVersion) | Graph module $gv" 'DEBUG'
    Write-AdcLog "Depth guard: $Script:AdcMaxDepth level(s)" 'DEBUG'

    & $setStatus 'Loading the app list...'

    # Fill the dropdown as soon as the window is up, so the operator never
    # has to type an app name to choose one.
    $win.Add_ContentRendered({ & $loadAppList $false })
    $win.ShowDialog() | Out-Null
}

# ---------------------------------------------------------------------------
# Standalone run: when this file is executed directly rather than dot-sourced,
# connect and open the window so the module can be tested on its own.
# ---------------------------------------------------------------------------
if ($MyInvocation.InvocationName -ne '.') {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -Scopes 'DeviceManagementApps.Read.All' -NoWelcome
    Show-AppDependencyCheckWindow -XamlPath $XamlPath -ShowConsole
}
