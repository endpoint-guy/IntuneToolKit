# Endpoint-Guy Intune Toolkit

A Windows PowerShell 5.1 and WPF desktop toolkit for common Microsoft Intune and Microsoft Entra ID device-administration tasks through Microsoft Graph.

The toolkit provides a dark-themed operator interface for finding managed devices, reviewing device details, exporting search results, and safely managing assigned security-group memberships. Its current device-action modules support copying groups between devices, removing a device from groups, and adding many devices from a CSV file to one group.

> **Project status:** Phase 1. The device search and the three group-management workflows documented below are implemented. Other device actions shown as **Coming soon** are placeholders.

## Table of contents

- [Features](#features)
- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Microsoft Graph permissions](#microsoft-graph-permissions)
- [Installation](#installation)
- [Running the toolkit](#running-the-toolkit)
- [Usage](#usage)
- [Bulk-add CSV format](#bulk-add-csv-format)
- [Group eligibility rules](#group-eligibility-rules)
- [Repository layout](#repository-layout)
- [Running modules independently](#running-modules-independently)
- [Diagnostics and logging](#diagnostics-and-logging)
- [Safety and operational behavior](#safety-and-operational-behavior)
- [Troubleshooting](#troubleshooting)
- [Development notes](#development-notes)
- [Security considerations](#security-considerations)
- [Acknowledgments](#acknowledgments)

## Features

### Managed-device search

- Interactive Microsoft Graph sign-in with the connected account and tenant shown in the UI.
- Search Intune managed devices by:
  - Device name
  - Serial number
  - Primary username
- Contains matching by default, with an **Exact match only** option.
- Local managed-device cache for fast repeat searches.
- Manual **Refresh Device Cache** action.
- Search results include device name, serial number, primary user, operating system, OS version, compliance state, ownership, model, and last sync time.
- UTF-8 CSV export of the current search results.

### Copy Device Groups

Copies eligible Microsoft Entra security-group memberships from a selected source device to a target device.

- Uses the device selected in the main search grid as the source.
- Searches the local device cache for a target device; if no cache is supplied, the module can query Graph directly.
- Enumerates every group membership of the source device.
- Classifies each group as eligible or ineligible before any write occurs.
- Preselects every eligible assigned security group.
- Optionally shows ineligible groups and the reason each cannot be changed.
- Requires confirmation before copying.
- Adds memberships one at a time and reports **Added**, **Already a member**, or **Failed** per group.
- Finishes with a summary count for added, existing, and failed memberships.

### Remove Device Groups

Removes the selected device from eligible assigned Microsoft Entra security groups.

- Resolves the selected Intune managed device to its Entra device object.
- Enumerates and classifies all current group memberships.
- Preselects every eligible assigned security group for the common “remove from all writable groups” workflow.
- Optionally displays ineligible memberships without allowing them to be selected.
- Requires a confirmation prompt that names the device and selected group count; the default response is **No**.
- Processes removals one at a time and writes the result back to the grid.
- Finishes with counts for **Removed**, **Not a member**, and **Failed**.

### Bulk Add to Group

Adds device names from a CSV file to one eligible assigned security group.

- Reads device names from the first CSV column.
- Treats every nonblank line—including line 1—as data; the file does not require or skip a header.
- Ignores blank lines.
- Removes duplicate names case-insensitively and reports the number ignored.
- Preserves quoted first-column values that contain commas.
- Matches each name against the Intune managed-device cache, loading the cache once from Graph when necessary.
- Shows unmatched and ambiguous names before any change is made.
- Selects only names that match exactly one managed device and resolve to an Entra device object.
- Searches target groups by display-name prefix or accepts a pasted group object ID.
- Prevents selection of dynamic, synced, mail-enabled, Microsoft 365, and other non-writable groups.
- Requires confirmation naming the target group and number of devices; the default response is **No**.
- Adds devices one at a time and reports **Added**, **Already a member**, or **Failed**.
- Exports the result grid to UTF-8 CSV.

## How it works

1. The operator launches `Toolkit.ps1` in Windows PowerShell 5.1 with STA enabled.
2. The toolkit checks for `Microsoft.Graph.Authentication` and loads the action modules.
3. The operator signs in interactively to Microsoft Graph.
4. Managed devices are read from Intune and cached locally in the current process.
5. Device searches run against that cache for responsive filtering.
6. Group actions resolve the Intune managed-device record to the corresponding Entra device object.
7. Group metadata is evaluated before selection controls are enabled.
8. Confirmed writes are sent to Microsoft Graph one item at a time, preserving an individual outcome for every device or group.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- PowerShell must run in Single-Threaded Apartment mode (`-STA`) for WPF
- Microsoft Graph PowerShell authentication module:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

- A Microsoft Entra account allowed to consent to or use the requested delegated permissions
- Appropriate Intune and Entra administrative roles for the operations being performed
- Network access to `https://graph.microsoft.com`

> This project targets Windows PowerShell 5.1, not PowerShell 7, because the UI is implemented with Windows WPF assemblies.

## Microsoft Graph permissions

The main toolkit currently requests the following delegated scopes:

| Scope | Purpose |
|---|---|
| `DeviceManagementManagedDevices.ReadWrite.All` | Read Intune managed-device inventory and provide headroom for device-management actions. |
| `DeviceManagementConfiguration.Read.All` | Read Intune device-management configuration data used or anticipated by the toolkit. |
| `Device.Read.All` | Resolve managed devices to Microsoft Entra device objects. |
| `Group.Read.All` | Read group details and device group memberships. |
| `Group.ReadWrite.All` | Support group membership changes. |
| `GroupMember.ReadWrite.All` | Add and remove device objects from groups. |
| `User.Read.All` | Read user information associated with managed devices. |

When an action module is run independently, it requests `DeviceManagementManagedDevices.Read.All`, `Device.Read.All`, `Group.Read.All`, `Group.ReadWrite.All`, and `GroupMember.ReadWrite.All`.

> Tenant policy may require administrator consent. Use least privilege, review the requested scopes for your environment, and test with a non-production tenant first.

## Installation

### 1. Clone the repository

```powershell
git clone https://github.com/<your-org-or-user>/endpoint-guy-intune-toolkit.git
cd endpoint-guy-intune-toolkit
```

### 2. Install the Graph authentication module

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

### 3. Confirm the module is available

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication
```

## Running the toolkit

Users can run Endpoint-Guy Intune Toolkit in either of two supported ways:

> **A prebuilt `EndpointguyToolkit.exe` is included by default.** Most users can launch it immediately after installing the required Microsoft Graph module. `Build-Exe.ps1` is provided when you want to rebuild the executable from source, change its version or icon, or produce a fresh local build.

1. Run `Toolkit.ps1` directly with Windows PowerShell 5.1.
2. Use `Build-Exe.ps1` to compile the standalone source into `EndpointguyToolkit.exe`, then launch the executable.

Both options use the same WPF interface and require `Microsoft.Graph.Authentication` to be installed on the computer running the toolkit.

### Option 1: Run the PowerShell script

From the repository root:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1
```

To target a specific tenant during interactive sign-in:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
```

During UI development, load an external main-window XAML file instead of the embedded copy:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1 -XamlPath .\MainWindow.xaml
```

### Option 2: Build and run the executable

The repository already includes the default prebuilt `EndpointguyToolkit.exe`; rebuilding is optional. Run `Build-Exe.ps1` only when you need to regenerate or customize the executable.

`Build-Exe.ps1` wraps PS2EXE with the settings required by the WPF application:

- `STA` for the single-threaded apartment required by WPF
- `noConsole` so no console window appears behind the application
- `DPIAware` for correct high-DPI scaling
- `x64` for a 64-bit runtime
- Product metadata for **Endpointguy Intune Toolkit**

The build script installs the `ps2exe` module for the current user when it is not already available. It also ensures the source is saved with a UTF-8 BOM before compiling.

By default, the build expects:

- Source: `EndpointguyToolkit-Standalone.ps1`
- Output: `EndpointguyToolkit.exe`
- Version: `1.0.0.0`

Build with the defaults:

```powershell
.\Build-Exe.ps1
```

Build with an icon and custom version:

```powershell
.\Build-Exe.ps1 -IconFile .\toolkit.ico -Version 1.1.0.0
```

If your standalone source has a different name or location, specify it explicitly:

```powershell
.\Build-Exe.ps1 -Source .\Toolkit.ps1 -Output .\EndpointguyToolkit.exe
```

After a successful build, launch:

```powershell
.\EndpointguyToolkit.exe
```

> The executable packages the PowerShell/WPF entry script, but it does **not** remove the runtime requirement for `Microsoft.Graph.Authentication`. Install that module on every computer that runs the EXE.

When compiled, the toolkit uses the executable directory as its module root because `$PSScriptRoot` may not be available in the packaged process. Keep the `Modules` directory beside the executable unless the standalone build embeds those module scripts.

## Usage

### Connect and search for a device

1. Launch the toolkit.
2. Select **Connect to Microsoft Graph** and complete interactive sign-in.
3. Choose **Device Name**, **Serial Number**, or **Primary Username**.
4. Enter a search value.
5. Leave **Exact match only** cleared for a contains search, or select it for equality matching.
6. Select **Search**.
7. Select a result row to enable device actions.
8. Use **Export to CSV** to save the current result set.

### Copy groups to another device

1. Select the source device in the main result grid.
2. Select **Copy Device Groups**.
3. Search for and select the target device.
4. Review eligible and ineligible source memberships.
5. Adjust the selected eligible groups if necessary.
6. Select the copy action and review the confirmation prompt.
7. Review the per-group results and final summary.

### Remove a device from groups

1. Select the device in the main result grid.
2. Select **Remove Device Groups**.
3. Review the memberships. Eligible assigned groups are selected by default.
4. Clear any group that should remain unchanged.
5. Select the remove action and explicitly confirm the warning prompt.
6. Review the result for every selected group.

### Add devices from a CSV to one group

1. Select **Bulk Add to Group**.
2. Browse to a CSV containing device names in its first column.
3. Review matched, unmatched, ambiguous, blank, and duplicate entries.
4. Search for the destination group by name prefix or paste its object ID.
5. Select an eligible assigned security group.
6. Adjust the selected matched devices if necessary.
7. Select **Add devices to group** and explicitly confirm the warning prompt.
8. Review or export the per-device result grid.

## Bulk-add CSV format

The current parser has no header-row concept. Every nonblank line is read, including the first line.

Recommended file:

```csv
LAPTOP-001
LAPTOP-002
KIOSK-014
```

Rules:

- Put the device name in the first column.
- Do not add a header unless you intentionally want that header text to appear as an unmatched device.
- Extra columns are ignored.
- Blank lines are ignored.
- Duplicate names are ignored case-insensitively.
- Quoted first-column values are parsed with `ConvertFrom-Csv`, so a name containing a comma is preserved.
- A device is selectable only when the name matches exactly one managed device and that record resolves to an Entra device object.

## Group eligibility rules

Only assigned, cloud-managed security groups are writable in the UI.

| Group type or state | Selectable? | Reason |
|---|:---:|---|
| Assigned cloud security group | Yes | Membership can be directly changed. |
| Dynamic membership group | No | Membership is controlled by a rule. |
| Group with a membership rule | No | A membership rule is treated as authoritative even if `groupTypes` is incomplete. |
| Microsoft 365 / Unified group | No | Not treated as a writable device security group. |
| Mail-enabled security group | No | Not eligible for these device membership operations. |
| Distribution group | No | Not a security group for this workflow. |
| On-premises synchronized group | No | Membership is mastered outside Microsoft Entra ID. |
| Non-security group | No | The group is not an eligible security target. |

Ineligible groups remain visible when **Show ineligible groups** is enabled, but their check boxes are disabled.

## Repository layout

```text
.
├── Toolkit.ps1
├── EndpointguyToolkit-Standalone.ps1  # Default source consumed by Build-Exe.ps1
├── EndpointguyToolkit.exe             # Prebuilt executable included by default
├── Build-Exe.ps1                      # Builds the Windows GUI executable
├── MainWindow.xaml
└── Modules
    ├── CopyDeviceGroups
    │   ├── CopyDeviceGroups.ps1
    │   └── CopyDeviceGroups.xaml
    ├── RemoveDeviceGroups
    │   ├── RemoveDeviceGroups.ps1
    │   └── RemoveDeviceGroups.xaml
    └── BulkAddToGroup
        ├── BulkAddToGroup.ps1
        └── BulkAddToGroup.xaml
```

Each PowerShell action module contains an embedded XAML copy for self-contained execution and PS2EXE packaging. The matching external XAML file can be supplied through `-XamlPath` while iterating on the UI.

The main toolkit first probes the module subfolders shown above, then falls back to legacy flat module paths beside the entry script. For EXE distribution, keep the `Modules` folder beside `EndpointguyToolkit.exe` unless those modules are embedded in `EndpointguyToolkit-Standalone.ps1`.

## Running modules independently

The action scripts include standalone development entry points. When invoked directly rather than dot-sourced, they connect to Graph, prompt for the device information they need, and open with diagnostics available.

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .ModulesCopyDeviceGroupsCopyDeviceGroups.ps1
powershell.exe -STA -ExecutionPolicy Bypass -File .ModulesRemoveDeviceGroupsRemoveDeviceGroups.ps1
powershell.exe -STA -ExecutionPolicy Bypass -File .ModulesBulkAddToGroupBulkAddToGroup.ps1
```

For normal integrated operation, `Toolkit.ps1` dot-sources the modules and calls:

```powershell
Show-CopyDeviceGroupsWindow
Show-RemoveDeviceGroupsWindow
Show-BulkAddToGroupWindow
```

## Diagnostics and logging

The copy and remove modules include an expandable diagnostics console with:

- Timestamped log levels such as `INFO`, `OK`, `WARN`, `ERROR`, `GRAPH`, and `DEBUG`
- Optional verbose Graph timing and per-group decision details
- Auto-scroll
- Copy log
- Save log
- Clear log
- Exception type, inner-exception chain, error ID, category, invocation line, and script stack trace for failures

When no on-screen log sink is active, messages fall back to `Write-Verbose`.

The default embedded Bulk Add window logs through `Write-Verbose`. Keep the embedded and external Bulk Add XAML copies synchronized if you enable the external file’s additional diagnostics controls.

## Safety and operational behavior

- No group membership is changed during search or preview.
- Only eligible objects have enabled selection controls.
- Destructive or high-impact actions require explicit confirmation.
- Remove and bulk-add confirmations default to **No**.
- Writes occur one item at a time, preserving an individual result instead of failing the entire batch silently.
- “Already a member” and “Not a member” are handled as idempotent outcomes where possible.
- Graph paging follows `@odata.nextLink` so the modules do not silently stop at the first page.
- Missing Graph properties are probed defensively to support strict-mode execution and response differences.
- Errors attempt to surface the useful Microsoft Graph response body instead of only a generic HTTP status message.
- Dynamic and synchronized groups are never offered as writable targets.

## Troubleshooting

### `Microsoft.Graph.Authentication` is not installed

Install it for the current user and restart the toolkit:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

### The window does not open

- Confirm you are using Windows PowerShell 5.1.
- Launch with `-STA`.
- Confirm WPF assemblies are available on the machine.
- If using `-XamlPath`, confirm the path exists and the XAML contains every control expected by the script.

### Sign-in or consent fails

- Verify the account can sign in to the intended tenant.
- Use `-TenantId` when the account belongs to multiple tenants.
- Ask a tenant administrator to review or grant the delegated scopes.
- Confirm Conditional Access permits the sign-in context.

### A device is not found

- Refresh the managed-device cache.
- Search by serial number or primary username.
- Clear **Exact match only** for a contains search.
- Confirm the device exists in Intune and has an `azureADDeviceId` value.
- In Bulk Add, confirm the CSV has no unintended header and that the device name is in column 1.

### A group is visible but cannot be selected

This is expected for dynamic, rule-driven, Microsoft 365, mail-enabled, distribution, on-premises synchronized, or non-security groups. Review the **Eligibility** column for the exact reason.

### A membership change fails

- Open or enable diagnostics and repeat the operation.
- Review the Graph response, exception chain, and stack trace.
- Confirm the connected account has both the Graph consent and Entra role required to change memberships.
- Confirm the target is a cloud-managed assigned security group.
- Recheck that the device still exists as an Entra device object.

### The executable cannot find modules

Keep the `Modules` folder beside the executable using the documented subfolder layout, or update the build process to embed/copy the module scripts.

## Development notes

- The scripts use module-specific prefixes (`Cdg`, `Rdg`, and `Bag`) because all modules are dot-sourced into the same PowerShell session.
- Row types implement `INotifyPropertyChanged` so WPF grids repaint when selections and results change.
- Graph collection requests follow pagination links.
- Group details are fetched before eligibility decisions so membership rules are not missed when a membership response is incomplete.
- Managed-device records and Entra device objects are deliberately treated as different objects and resolved through the Entra device ID.
- External XAML is a development override; embedded XAML is the default runtime UI.
- Keep each external XAML file synchronized with its embedded counterpart before release.
- The current `BulkAddToGroup.ps1` parser reads every line, including line 1. Ensure any external Bulk Add XAML help text says the same; stale text claiming that the first line is skipped would be incorrect.
- The Graph request helpers accommodate module-version differences around `SkipHttpErrorCheck` and status-code capture.

## Security considerations

This toolkit can change Microsoft Entra group memberships and requests high-impact delegated permissions. Before production use:

- Review the code and requested scopes.
- Use a dedicated administrative account with the minimum required Entra and Intune roles.
- Test in a non-production tenant.
- Validate CSV contents before confirming bulk operations.
- Protect exported device and diagnostics data, which may contain device identifiers, usernames, tenant information, and Graph error details.
- Do not commit exported CSV files or saved diagnostic logs containing organizational data.
- Follow your organization’s change-management and audit requirements.

## Contributing

Issues and pull requests are welcome. For UI changes, update both the external XAML file and the embedded XAML string in the matching PowerShell script. Include test notes for device resolution, group eligibility, confirmation behavior, Graph pagination, and error handling.

## Acknowledgments

Endpoint-Guy Intune Toolkit was built with assistance from Claude by Anthropic. Claude was used to support development, documentation, troubleshooting, and code review; project decisions and responsibility for the final implementation remain with the project author.

## Disclaimer

This project is provided as an administrative tool. Review and test it in your environment before production use. Microsoft Intune, Microsoft Entra ID, Microsoft Graph, Windows PowerShell, and WPF are Microsoft technologies and trademarks.
<img width="268" height="32766" alt="image" src="https://github.com/user-attachments/assets/130bb7fe-ff52-478f-9e78-33bb9d03b3ce" />
