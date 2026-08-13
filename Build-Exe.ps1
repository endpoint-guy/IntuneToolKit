<#
.SYNOPSIS
    Compiles EndpointguyToolkit-Standalone.ps1 into a single Windows executable.

.DESCRIPTION
    Wraps PS2EXE with the flags a WPF application requires:
      -STA        single-threaded apartment - MANDATORY for WPF, window will not
                  appear without it
      -noConsole  compiles as a GUI app so no console window sits behind the UI
      -DPIAware   correct scaling on high-DPI displays
      -x64        64-bit runtime

.EXAMPLE
    .\Build-Exe.ps1

.EXAMPLE
    .\Build-Exe.ps1 -IconFile .\toolkit.ico -Version 1.1.0.0

.NOTES
    Run from Windows PowerShell 5.1 in the folder containing the standalone script.
#>

[CmdletBinding()]
param(
    [string]$Source  = (Join-Path $PSScriptRoot 'EndpointguyToolkit-Standalone.ps1'),
    [string]$Output  = (Join-Path $PSScriptRoot 'EndpointguyToolkit.exe'),
    [string]$IconFile,
    [string]$Version = '1.0.0.0'
)

$ErrorActionPreference = 'Stop'

# --- 1. Ensure PS2EXE is available ------------------------------------------
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host 'PS2EXE not found. Installing from the PowerShell Gallery...' -ForegroundColor Yellow
    Install-Module ps2exe -Scope CurrentUser -Force
}
Import-Module ps2exe

# --- 2. Validate the source -------------------------------------------------
if (-not (Test-Path $Source)) { throw "Source script not found: $Source" }

# PS2EXE expects UTF8 or UTF16 encoded input. Re-save with a BOM if missing.
$bytes = [System.IO.File]::ReadAllBytes($Source)
if ($bytes.Length -lt 3 -or
    -not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
    Write-Host 'Source is missing a UTF-8 BOM. Re-saving with BOM...' -ForegroundColor Yellow
    $content = Get-Content -Path $Source -Raw
    [System.IO.File]::WriteAllText($Source, $content, (New-Object System.Text.UTF8Encoding $true))
}

# --- 3. Compile -------------------------------------------------------------
$params = @{
    inputFile   = $Source
    outputFile  = $Output
    noConsole   = $true      # GUI app - no console window
    STA         = $true      # REQUIRED for WPF
    DPIAware    = $true      # high-DPI scaling
    x64         = $true
    title       = 'Endpointguy Intune Toolkit'
    description = 'Device lookup and Intune administration tools'
    product     = 'Endpointguy Intune Toolkit'
    company     = 'endpointguy.com'
    copyright   = "(c) $(Get-Date -Format yyyy)"
    version     = $Version
}
if ($IconFile -and (Test-Path $IconFile)) { $params['iconFile'] = $IconFile }

Write-Host "Compiling -> $Output" -ForegroundColor Cyan
Invoke-PS2EXE @params

if (Test-Path $Output) {
    $sizeKb = [math]::Round((Get-Item $Output).Length / 1KB, 1)
    Write-Host ''
    Write-Host "Build succeeded: $Output ($sizeKb KB)" -ForegroundColor Green
    Write-Host 'Reminder: Microsoft.Graph.Authentication must be installed on any machine that runs it.' -ForegroundColor DarkGray
}
else {
    throw 'Build failed - no output produced.'
}
