<#
.SYNOPSIS
    EndpointGuy Intune Toolkit configuration wizard.
.DESCRIPTION
    Standalone WPF wizard that writes ModuleConfig.psd1 beside this script.
#>
[CmdletBinding()]
param(
    [AllowEmptyString()][string]$ConfigPath = '',
    [switch]$ShowConsole
)

# Resolve the configuration beside this script after parameter binding.
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $scriptFile = $MyInvocation.MyCommand.Path
    $scriptFolder = if ([string]::IsNullOrWhiteSpace($scriptFile)) { (Get-Location).Path } else { [System.IO.Path]::GetDirectoryName($scriptFile) }
    if ([string]::IsNullOrWhiteSpace($scriptFolder)) { $scriptFolder = (Get-Location).Path }
    $ConfigPath = [System.IO.Path]::Combine($scriptFolder, 'ModuleConfig.psd1')
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$nl = [Environment]::NewLine

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# Keep the PowerShell console off the taskbar, matching Toolkit.ps1.
if (-not ('EndpointguyConfigConsole' -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class EndpointguyConfigConsole
{
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);
}
"@
}

$launchedDirectly = [bool](@([Environment]::GetCommandLineArgs()) -like "*$(Split-Path -Leaf $PSCommandPath)*")
$hostHwnd = [EndpointguyConfigConsole]::GetConsoleWindow()
if ($launchedDirectly -and -not $ShowConsole -and -not $env:ENDPOINTGUY_CONFIG_RELAUNCHED -and $hostHwnd -ne [IntPtr]::Zero) {
    $consoleClass = New-Object System.Text.StringBuilder 256
    [EndpointguyConfigConsole]::GetClassName($hostHwnd, $consoleClass, $consoleClass.Capacity) | Out-Null
    $conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'
    if ($consoleClass.ToString() -ne 'ConsoleWindowClass' -and (Test-Path $conhost)) {
        $relaunchArgs = @(
            "`"$PSHOME\powershell.exe`"", '-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass',
            '-WindowStyle', 'Hidden', '-File', "`"$PSCommandPath`"",
            '-ConfigPath', "`"$ConfigPath`""
        )
        $env:ENDPOINTGUY_CONFIG_RELAUNCHED = '1'
        Start-Process -FilePath $conhost -ArgumentList $relaunchArgs -WindowStyle Hidden
        exit 0
    }
}

$moduleDefinitions = @(
    [pscustomobject]@{ Key='CopyDeviceGroups';   Group='Device Actions'; Label='Copy Device Groups';       Description='Copy assigned security-group memberships to another device.' },
    [pscustomobject]@{ Key='RemoveDeviceGroups'; Group='Device Actions'; Label='Remove Device Groups';     Description='Remove a device from assigned security groups.' },
    [pscustomobject]@{ Key='AssetStatus';        Group='Device Actions'; Label='Update Asset Status';      Description='Write an approved lifecycle label to Intune Management Name.' },
    [pscustomobject]@{ Key='BulkAddToGroup';     Group='Bulk Actions';   Label='Bulk Add to Group';        Description='Add devices from a CSV file to one assigned group.' },
    [pscustomobject]@{ Key='BulkAssetStatus';    Group='Bulk Actions';   Label='Bulk Update Asset Status'; Description='Write one lifecycle label to multiple devices from a CSV.' },
    [pscustomobject]@{ Key='AppDependencyCheck'; Group='Reporting';      Label='App Dependency Check';     Description='Report which Win32 apps depend on a selected app.' },
    [pscustomobject]@{ Key='DevicePrepCheck';    Group='Reporting';      Label='Device Prep Check';        Description='Check whether a serial number is ready for Autopilot.' }
)
$defaultStatuses = @('Assigned','In-Stock','Retired','Recycled','Stolen','Legalhold','Lost')
$currentModules = @{}
foreach ($definition in $moduleDefinitions) { $currentModules[$definition.Key] = $true }
$currentStatuses = @($defaultStatuses)
$currentToolkitName = 'EndpointGuy Intune Toolkit'
$currentWebsiteText = 'endpointguy.com'
$currentWebsiteUrl = 'https://endpointguy.com'
$currentLogoFile = 'EGLogoNew.png'
$script:SelectedLogoSource = $null

if (Test-Path -LiteralPath $ConfigPath) {
    try {
        $current = Import-PowerShellDataFile -LiteralPath $ConfigPath
        if ($current.ContainsKey('Branding') -and $current.Branding -is [hashtable]) {
            if ($current.Branding.ContainsKey('ToolkitName') -and -not [string]::IsNullOrWhiteSpace([string]$current.Branding.ToolkitName)) { $currentToolkitName = [string]$current.Branding.ToolkitName }
            if ($current.Branding.ContainsKey('WebsiteText') -and -not [string]::IsNullOrWhiteSpace([string]$current.Branding.WebsiteText)) { $currentWebsiteText = [string]$current.Branding.WebsiteText }
            if ($current.Branding.ContainsKey('WebsiteUrl') -and -not [string]::IsNullOrWhiteSpace([string]$current.Branding.WebsiteUrl)) { $currentWebsiteUrl = [string]$current.Branding.WebsiteUrl }
            if ($current.Branding.ContainsKey('LogoFile') -and -not [string]::IsNullOrWhiteSpace([string]$current.Branding.LogoFile)) { $currentLogoFile = [string]$current.Branding.LogoFile }
        }
        if ($current.ContainsKey('Modules') -and $current.Modules -is [hashtable]) {
            foreach ($definition in $moduleDefinitions) {
                if ($current.Modules.ContainsKey($definition.Key)) { $currentModules[$definition.Key] = [bool]$current.Modules[$definition.Key] }
            }
        }
        if ($current.ContainsKey('AssetStatusValues')) {
            $loaded = @($current.AssetStatusValues) | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($loaded.Count -gt 0) { $currentStatuses = @($loaded) }
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(('The existing configuration could not be read. Built-in defaults will be used.' + $nl + $nl + $_.Exception.Message),'EndpointGuy Toolkit Configuration','OK','Warning') | Out-Null
    }
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="EndpointGuy Toolkit Configuration" Height="770" Width="820" WindowStartupLocation="CenterScreen"
 ResizeMode="NoResize" Background="#1F1F1F" FontFamily="Segoe UI" UseLayoutRounding="True">
 <Window.Resources>
  <SolidColorBrush x:Key="Card" Color="#2B2B2B"/><SolidColorBrush x:Key="Border" Color="#3D3D3D"/>
  <SolidColorBrush x:Key="Text" Color="#E8E8E8"/><SolidColorBrush x:Key="Subtle" Color="#A6A6A6"/>
  <SolidColorBrush x:Key="Accent" Color="#2B7CD3"/>
  <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{StaticResource Text}"/></Style>
  <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{StaticResource Text}"/><Setter Property="FontSize" Value="13"/><Setter Property="Margin" Value="0,5"/></Style>
  <Style TargetType="RadioButton"><Setter Property="Foreground" Value="{StaticResource Text}"/><Setter Property="FontSize" Value="14"/><Setter Property="Margin" Value="0,8"/></Style>
  <Style x:Key="NavButton" TargetType="Button"><Setter Property="MinWidth" Value="100"/><Setter Property="Height" Value="36"/><Setter Property="Margin" Value="10,0,0,0"/><Setter Property="Padding" Value="18,0"/><Setter Property="FontSize" Value="13"/><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Foreground" Value="#F0F0F0"/><Setter Property="Background" Value="#3A3A3A"/><Setter Property="BorderBrush" Value="#4D4D4D"/></Style>
  <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource NavButton}"><Setter Property="Background" Value="{StaticResource Accent}"/><Setter Property="BorderBrush" Value="{StaticResource Accent}"/><Setter Property="Foreground" Value="White"/></Style>
 </Window.Resources>
 <Grid Margin="24">
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Grid Grid.Row="0" Margin="0,0,0,18"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
   <StackPanel><TextBlock Text="Toolkit Configuration" FontSize="24" FontWeight="SemiBold" Foreground="White"/><TextBlock Text="Configure modules and environment-specific Asset Status labels." FontSize="12.5" Foreground="{StaticResource Subtle}" Margin="0,4,0,0"/></StackPanel>
   <Border Grid.Column="1" Background="#252F3A" BorderBrush="{StaticResource Accent}" BorderThickness="1" CornerRadius="14" Padding="12,5" VerticalAlignment="Center"><TextBlock x:Name="StepIndicator" Text="Step 1 of 4" Foreground="#6BB3E8" FontWeight="SemiBold"/></Border>
  </Grid>
  <Border Grid.Row="1" Background="{StaticResource Card}" BorderBrush="{StaticResource Border}" BorderThickness="1" CornerRadius="9" Padding="28">
   <Grid>
    <Grid x:Name="Page1"><StackPanel>
     <TextBlock Text="Use Intune Management Name as an Asset Status field?" FontSize="19" FontWeight="SemiBold" Foreground="White"/>
     <TextBlock Text="This is asked first because it determines whether the Asset Status modules and label list are needed." FontSize="13" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,8,0,22"/>
     <Border Background="#252525" BorderBrush="{StaticResource Border}" BorderThickness="1" CornerRadius="7" Padding="18"><StackPanel>
      <RadioButton x:Name="UseAssetYes" GroupName="UseAsset" Content="Yes - use Management Name for lifecycle / asset status labels"/>
      <TextBlock Text="You can choose the approved labels and enable either Asset Status module." Foreground="{StaticResource Subtle}" FontSize="12" TextWrapping="Wrap" Margin="25,-3,0,12"/>
      <RadioButton x:Name="UseAssetNo" GroupName="UseAsset" Content="No - do not use Management Name as an Asset Status field"/>
      <TextBlock Text="AssetStatus and BulkAssetStatus will be forced off, and label entry will be skipped." Foreground="{StaticResource Subtle}" FontSize="12" TextWrapping="Wrap" Margin="25,-3,0,0"/>
     </StackPanel></Border>
    </StackPanel></Grid>
    <Grid x:Name="Page2" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
     <StackPanel><TextBlock Text="Choose enabled modules" FontSize="19" FontWeight="SemiBold" Foreground="White"/><TextBlock x:Name="ModuleIntro" FontSize="13" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,8,0,14"/></StackPanel>
     <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"><StackPanel x:Name="ModuleList"/></ScrollViewer>
    </Grid>
    <Grid x:Name="Page3" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
     <StackPanel><TextBlock Text="Asset Status values" FontSize="19" FontWeight="SemiBold" Foreground="White"/><TextBlock x:Name="ReviewIntro" Text="Add, edit, or remove the approved lifecycle labels. At least one value is required." FontSize="13" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,8,0,14"/></StackPanel>
     <Border x:Name="StatusEditor" Grid.Row="1" Background="#252525" BorderBrush="{StaticResource Border}" BorderThickness="1" CornerRadius="7" Padding="18">
      <Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
       <TextBlock Grid.Row="0" Text="Approved values" FontSize="14" FontWeight="SemiBold"/>
       <TextBlock Grid.Row="1" Text="Values appear in this order and are written verbatim to Intune Management Name." FontSize="12" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,5,0,14"/>
       <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,6,0"><StackPanel x:Name="StatusRows"/></ScrollViewer>
       <Button x:Name="AddStatusButton" Grid.Row="3" Content="Add another value" Style="{StaticResource PrimaryButton}" HorizontalAlignment="Left" Margin="0,16,0,0"/>
      </Grid>
     </Border>
    </Grid>
    <Grid x:Name="Page4" Visibility="Collapsed"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
     <StackPanel><TextBlock Text="Toolkit branding" FontSize="19" FontWeight="SemiBold" Foreground="White"/><TextBlock Text="Customize the name, website link, and logo shown in the toolkit header." FontSize="13" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,8,0,14"/></StackPanel>
     <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto"><StackPanel>
      <TextBlock Text="Toolkit name" FontSize="12" FontWeight="SemiBold" Margin="0,0,0,5"/><TextBox x:Name="ToolkitNameText" Height="34" Padding="9,5" Background="#2E2E2E" Foreground="{StaticResource Text}" BorderBrush="#4D4D4D"/>
      <TextBlock Text="Website display text" FontSize="12" FontWeight="SemiBold" Margin="0,14,0,5"/><TextBox x:Name="WebsiteTextBox" Height="34" Padding="9,5" Background="#2E2E2E" Foreground="{StaticResource Text}" BorderBrush="#4D4D4D"/>
      <TextBlock Text="Website URL" FontSize="12" FontWeight="SemiBold" Margin="0,14,0,5"/><TextBox x:Name="WebsiteUrlBox" Height="34" Padding="9,5" Background="#2E2E2E" Foreground="{StaticResource Text}" BorderBrush="#4D4D4D"/>
      <TextBlock Text="Logo image" FontSize="12" FontWeight="SemiBold" Margin="0,14,0,5"/>
      <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBox x:Name="LogoPathBox" Grid.Column="0" Height="34" Padding="9,5" IsReadOnly="True" Background="#2E2E2E" Foreground="{StaticResource Subtle}" BorderBrush="#4D4D4D"/><Button x:Name="BrowseLogoButton" Grid.Column="1" Content="Browse..." Style="{StaticResource NavButton}" Height="34" Margin="10,0,0,0"/></Grid>
      <TextBlock Text="PNG, JPG, JPEG, or BMP. A selected image is copied into the toolkit folder as ToolkitLogo.png." FontSize="11.5" Foreground="{StaticResource Subtle}" TextWrapping="Wrap" Margin="0,6,0,0"/>
     </StackPanel></ScrollViewer>
    </Grid>
   </Grid>
  </Border>
  <Grid Grid.Row="2" Margin="0,18,0,0"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><Border Grid.Row="0" Background="#252525" BorderBrush="{StaticResource Border}" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="0,0,0,12"><TextBlock x:Name="ConfigPathText" Foreground="{StaticResource Subtle}" FontSize="11.5" TextWrapping="Wrap"/></Border><StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right"><Button x:Name="CancelButton" Content="Cancel" Style="{StaticResource NavButton}"/><Button x:Name="BackButton" Content="Back" Style="{StaticResource NavButton}" IsEnabled="False"/><Button x:Name="NextButton" Content="Next" Style="{StaticResource PrimaryButton}"/></StackPanel></Grid>
 </Grid>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$Page1=$window.FindName('Page1'); $Page2=$window.FindName('Page2'); $Page3=$window.FindName('Page3'); $Page4=$window.FindName('Page4')
$StepIndicator=$window.FindName('StepIndicator'); $UseAssetYes=$window.FindName('UseAssetYes'); $UseAssetNo=$window.FindName('UseAssetNo')
$ModuleList=$window.FindName('ModuleList'); $ModuleIntro=$window.FindName('ModuleIntro'); $StatusEditor=$window.FindName('StatusEditor')
$StatusRows=$window.FindName('StatusRows'); $AddStatusButton=$window.FindName('AddStatusButton'); $ReviewIntro=$window.FindName('ReviewIntro')
$ConfigPathText=$window.FindName('ConfigPathText'); $CancelButton=$window.FindName('CancelButton'); $BackButton=$window.FindName('BackButton'); $NextButton=$window.FindName('NextButton')
$ToolkitNameText=$window.FindName('ToolkitNameText'); $WebsiteTextBox=$window.FindName('WebsiteTextBox'); $WebsiteUrlBox=$window.FindName('WebsiteUrlBox')
$LogoPathBox=$window.FindName('LogoPathBox'); $BrowseLogoButton=$window.FindName('BrowseLogoButton')
$ToolkitNameText.Text = $currentToolkitName
$WebsiteTextBox.Text = $currentWebsiteText
$WebsiteUrlBox.Text = $currentWebsiteUrl
$LogoPathBox.Text = if ([string]::IsNullOrWhiteSpace($currentLogoFile)) { "No logo configured" } else { "Current: $currentLogoFile" }

function Update-StatusRemoveButtons {
    $canRemove = ($StatusRows.Children.Count -gt 1)
    foreach ($row in $StatusRows.Children) {
        if ($row.Children.Count -gt 1) { $row.Children[1].IsEnabled = $canRemove }
    }
}
function Add-StatusRow {
    param([AllowEmptyString()][string]$Value = '')
    $row = New-Object System.Windows.Controls.Grid
    $row.Margin = '0,0,0,12'
    [void]$row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = '*' }))
    [void]$row.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition -Property @{ Width = 'Auto' }))
    $textBox = New-Object System.Windows.Controls.TextBox
    $textBox.Text = $Value
    $textBox.Height = 38
    $textBox.Padding = '11,7'
    $textBox.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#2E2E2E'))
    $textBox.Foreground = $window.FindResource('Text')
    $textBox.BorderBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#4D4D4D'))
    $textBox.FontSize = 13.5
    [System.Windows.Controls.Grid]::SetColumn($textBox, 0)
    $remove = New-Object System.Windows.Controls.Button
    $remove.Content = 'Remove'
    $remove.Style = $window.FindResource('NavButton')
    $remove.MinWidth = 92
    $remove.Width = 92
    $remove.Height = 38
    $remove.Margin = '12,0,0,0'
    $remove.Tag = $row
    [System.Windows.Controls.Grid]::SetColumn($remove, 1)
    $row.Tag = $textBox
    [void]$row.Children.Add($textBox)
    [void]$row.Children.Add($remove)
    $remove.Add_Click({
        param($clickSource, $clickArgs)
        if ($StatusRows.Children.Count -gt 1) {
            [void]$StatusRows.Children.Remove($clickSource.Tag)
            Update-StatusRemoveButtons
        }
    })
    [void]$StatusRows.Children.Add($row)
    Update-StatusRemoveButtons
    if ([string]::IsNullOrWhiteSpace($Value)) { [void]$textBox.Focus() }
}
$ConfigPathText.Text = "Writes: $ConfigPath"
foreach ($status in $currentStatuses) { Add-StatusRow -Value $status }
$moduleChecks = @{}
$lastGroup = $null
foreach ($definition in $moduleDefinitions) {
    if ($definition.Group -ne $lastGroup) {
        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text=$definition.Group.ToUpperInvariant(); $label.FontSize=10.5; $label.FontWeight=[System.Windows.FontWeights]::Bold
        $label.Foreground=New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString('#6BB3E8'))
        $label.Margin=if ($null -eq $lastGroup) {'0,4,0,4'} else {'0,14,0,4'}
        [void]$ModuleList.Children.Add($label); $lastGroup=$definition.Group
    }
    $check=New-Object System.Windows.Controls.CheckBox
    $check.Content=$definition.Label; $check.ToolTip=$definition.Description; $check.IsChecked=[bool]$currentModules[$definition.Key]; $check.Tag=$definition.Key
    [void]$ModuleList.Children.Add($check); $moduleChecks[$definition.Key]=$check
}
if ($currentModules['AssetStatus'] -or $currentModules['BulkAssetStatus']) {$UseAssetYes.IsChecked=$true} else {$UseAssetNo.IsChecked=$true}
$script:ConfigPage=1; $script:UseAssetStatus=$true

function Set-AssetModuleState([bool]$Enabled) {
    foreach ($key in @('AssetStatus','BulkAssetStatus')) { $moduleChecks[$key].IsEnabled=$Enabled; if (-not $Enabled) {$moduleChecks[$key].IsChecked=$false} }
    $ModuleIntro.Text = if ($Enabled) {'Checked modules are loaded and shown. Unchecked modules are hidden and their Graph scopes are not requested.'} else {'Asset Status is not being used, so both Asset Status modules are locked off. Choose the remaining modules.'}
}
function Update-AssetStatusPage {
    $StatusEditor.Visibility = if ($script:UseAssetStatus) { 'Visible' } else { 'Collapsed' }
    $ReviewIntro.Text = if ($script:UseAssetStatus) { 'Add, edit, or remove the approved lifecycle labels. At least one value is required.' } else { 'Asset Status values are not required because Management Name will not be used for this purpose.' }
}
function Show-ConfigPage([ValidateRange(1,4)][int]$Number) {
    $script:ConfigPage=$Number
    $Page1.Visibility=if($Number-eq 1){'Visible'}else{'Collapsed'}
    $Page2.Visibility=if($Number-eq 2){'Visible'}else{'Collapsed'}
    $Page3.Visibility=if($Number-eq 3){'Visible'}else{'Collapsed'}
    $Page4.Visibility=if($Number-eq 4){'Visible'}else{'Collapsed'}
    $StepIndicator.Text="Step $Number of 4"
    $BackButton.IsEnabled=($Number-gt 1)
    $NextButton.Content=if($Number-eq 4){'Save configuration'}else{'Next'}
    if($Number-eq 3){Update-AssetStatusPage}
}
function Get-ValidatedStatuses {
    $result=New-Object System.Collections.Generic.List[string]
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $StatusRows.Children) {
        $value = ([string]$row.Tag.Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($value) -and $seen.Add($value)) { [void]$result.Add($value) }
    }
    return @($result)
}
function Test-BrandingInput {
    if ([string]::IsNullOrWhiteSpace($ToolkitNameText.Text)) { [System.Windows.MessageBox]::Show("Enter a toolkit name.","Branding required","OK","Warning") | Out-Null; return $false }
    if ([string]::IsNullOrWhiteSpace($WebsiteTextBox.Text)) { [System.Windows.MessageBox]::Show("Enter website display text.","Branding required","OK","Warning") | Out-Null; return $false }
    $uri = $null
    if (-not [System.Uri]::TryCreate($WebsiteUrlBox.Text.Trim(), [System.UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @("http","https")) { [System.Windows.MessageBox]::Show("Enter a complete website URL beginning with http:// or https://.","Branding required","OK","Warning") | Out-Null; return $false }
    return $true
}
function Save-ToolkitConfiguration {
    $statuses=@(); if($script:UseAssetStatus){$statuses=@(Get-ValidatedStatuses); if($statuses.Count-eq 0){[System.Windows.MessageBox]::Show('Enter at least one Asset Status value.','Configuration required','OK','Warning')|Out-Null; return}}
    if (-not (Test-BrandingInput)) { return }
    $lines=New-Object System.Collections.Generic.List[string]
    @('# ---------------------------------------------------------------------------','# EndpointGuy Intune Toolkit - module configuration','# Generated by Configure-Toolkit.ps1','# ---------------------------------------------------------------------------','@{','    Modules = @{') | ForEach-Object {[void]$lines.Add($_)}
    foreach($definition in $moduleDefinitions){$enabled=[bool]$moduleChecks[$definition.Key].IsChecked; if(-not $script:UseAssetStatus -and $definition.Key -in @('AssetStatus','BulkAssetStatus')){$enabled=$false}; $literal=if($enabled){'$true'}else{'$false'}; [void]$lines.Add(('        {0,-20} = {1};' -f $definition.Key,$literal))}
    [void]$lines.Add('    };')
    $brandName = $ToolkitNameText.Text.Trim().Replace("'","''")
    $brandText = $WebsiteTextBox.Text.Trim().Replace("'","''")
    $brandUrl = $WebsiteUrlBox.Text.Trim().Replace("'","''")
    $logoFile = if ($script:SelectedLogoSource) { 'ToolkitLogo.png' } else { $currentLogoFile }
    $logoLiteral = ([string]$logoFile).Replace("'","''")
    [void]$lines.Add('')
    [void]$lines.Add('    Branding = @{')
    [void]$lines.Add("        ToolkitName = '$brandName';")
    [void]$lines.Add("        WebsiteText = '$brandText';")
    [void]$lines.Add("        WebsiteUrl  = '$brandUrl';")
    [void]$lines.Add("        LogoFile    = '$logoLiteral';")
    [void]$lines.Add('    };')
    if($script:UseAssetStatus){[void]$lines.Add(''); [void]$lines.Add('    # Values are offered in this order and written verbatim to Intune Management Name.'); [void]$lines.Add('    AssetStatusValues = @('); foreach($status in $statuses){$escaped=$status.Replace("'","''"); [void]$lines.Add("        '$escaped'")}; [void]$lines.Add('    );')}
    [void]$lines.Add('}')
    $folder=Split-Path -Parent $ConfigPath; if([string]::IsNullOrWhiteSpace($folder)){$folder=$PSScriptRoot}; if(-not(Test-Path -LiteralPath $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
    $temp=Join-Path $folder (([IO.Path]::GetRandomFileName())+'.tmp'); $backup=$null
    try{if($script:SelectedLogoSource){Copy-Item -LiteralPath $script:SelectedLogoSource -Destination (Join-Path $folder 'ToolkitLogo.png') -Force}; $utf8=New-Object System.Text.UTF8Encoding($false); [IO.File]::WriteAllLines($temp,[string[]]$lines,$utf8); [void](Import-PowerShellDataFile -LiteralPath $temp); if(Test-Path -LiteralPath $ConfigPath){$backup="$ConfigPath.bak-$(Get-Date -Format 'yyyyMMdd-HHmmss')"; Copy-Item -LiteralPath $ConfigPath -Destination $backup -Force}; Move-Item -LiteralPath $temp -Destination $ConfigPath -Force; $message='Configuration saved successfully.'+$nl+$nl+$ConfigPath; if($backup){$message+=$nl+$nl+'Previous configuration backed up to:'+$nl+$backup}; [System.Windows.MessageBox]::Show($message,'EndpointGuy Toolkit Configuration','OK','Information')|Out-Null; $window.DialogResult=$true; $window.Close()}
    catch{if(Test-Path -LiteralPath $temp){Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue}; [System.Windows.MessageBox]::Show(('The configuration could not be saved.'+$nl+$nl+$_.Exception.Message),'Save failed','OK','Error')|Out-Null}
}
$BrowseLogoButton.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Title = 'Select toolkit logo'
    $dialog.Filter = 'Image files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|All files (*.*)|*.*'
    if ($dialog.ShowDialog() -eq $true) { $script:SelectedLogoSource = $dialog.FileName; $LogoPathBox.Text = $dialog.FileName }
})
$AddStatusButton.Add_Click({ Add-StatusRow -Value '' })
$CancelButton.Add_Click({$window.Close()})
$BackButton.Add_Click({if($script:ConfigPage-eq 4 -and -not $script:UseAssetStatus){Show-ConfigPage 2}elseif($script:ConfigPage-gt 1){Show-ConfigPage ($script:ConfigPage-1)}})
$NextButton.Add_Click({switch($script:ConfigPage){1{if(-not $UseAssetYes.IsChecked -and -not $UseAssetNo.IsChecked){[System.Windows.MessageBox]::Show('Choose Yes or No before continuing.','Configuration required','OK','Warning')|Out-Null; return}; $script:UseAssetStatus=[bool]$UseAssetYes.IsChecked; Set-AssetModuleState $script:UseAssetStatus; Show-ConfigPage 2} 2{if($script:UseAssetStatus){Show-ConfigPage 3}else{Show-ConfigPage 4}} 3{Show-ConfigPage 4} 4{Save-ToolkitConfiguration}}})
Show-ConfigPage 1
$Script:ConsoleHwnd = [EndpointguyConfigConsole]::GetConsoleWindow()
$ownLaunch = [bool](@([Environment]::GetCommandLineArgs()) -like "*$(Split-Path -Leaf $PSCommandPath)*")
if ($ownLaunch -and -not $ShowConsole -and $Script:ConsoleHwnd -ne [IntPtr]::Zero) {
    [EndpointguyConfigConsole]::ShowWindow($Script:ConsoleHwnd, 0) | Out-Null
}

try {
    $window.ShowDialog() | Out-Null
}
catch {
    if ($Script:ConsoleHwnd -ne [IntPtr]::Zero) { [EndpointguyConfigConsole]::ShowWindow($Script:ConsoleHwnd, 5) | Out-Null }
    [System.Windows.MessageBox]::Show(("The configuration wizard closed because of an error:`n`n" + $_.Exception.Message), "EndpointGuy Toolkit Configuration", "OK", "Error") | Out-Null
    throw
}
