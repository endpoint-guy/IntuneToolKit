# Endpoint-Guy Intune Toolkit

Endpoint-Guy Intune Toolkit is a Windows PowerShell 5.1 and WPF application for Microsoft Intune and Microsoft Entra ID device administration through Microsoft Graph.

> **Recommended:** unblock the GitHub ZIP, then double-click `Run-Toolkit.bat`.

## Table of contents

- [Requirements](#requirements)
- [Modules](#modules)
- [User instructions](#user-instructions)
- [Microsoft Graph permissions](#microsoft-graph-permissions)
- [Using the toolkit](#using-the-toolkit)
- [Bulk-add CSV format](#bulk-add-csv-format)
- [App dependency checks](#app-dependency-checks)
- [Group eligibility](#group-eligibility)
- [Repository structure](#repository-structure)
- [Diagnostics and safety](#diagnostics-and-safety)
- [Development notes](#development-notes)
- [Acknowledgments](#acknowledgments)

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Network access to the PowerShell Gallery for first-time setup
- Network access to Microsoft Graph
- Appropriate Microsoft Intune and Entra roles

The toolkit automatically installs `Microsoft.Graph.Authentication` for the current user if it is missing.

## Modules

The toolkit consists of a core application plus optional action modules. The main application remains useful even when one or more module folders are not present; missing modules disable only their corresponding actions.

### Core toolkit — no action modules required

Without the attached action modules, `Toolkit.ps1` can still:

- Connect interactively to Microsoft Graph and display the connected account.
- Load and cache the Intune managed-device inventory.
- Search devices by device name, serial number, or primary username.
- Use contains matching or **Exact match only**.
- Display device name, serial number, user, operating system, OS version, compliance, ownership, model, and last sync.
- Refresh the local device cache.
- Select and review a managed device.
- Export the current search results to a UTF-8 CSV file.

### Copy Device Groups module

Location: `Modules\CopyDeviceGroups\CopyDeviceGroups.ps1`

- Copies eligible assigned security-group memberships from a selected source device to a target device.
- Shows eligible and ineligible memberships before writing.
- Preselects eligible groups and requires confirmation.
- Reports **Added**, **Already a member**, or **Failed** for each group.

### Remove Device Groups module

Location: `Modules\RemoveDeviceGroups\RemoveDeviceGroups.ps1`

- Removes the selected device from eligible assigned security groups.
- Preselects eligible memberships while leaving ineligible memberships read-only.
- Uses a confirmation prompt that defaults to **No**.
- Reports **Removed**, **Not a member**, or **Failed** for each group.

### Bulk Add to Group module

Location: `Modules\BulkAddToGroup\BulkAddToGroup.ps1`

- Reads device names from the first CSV column.
- Reads every nonblank line, including line 1, and ignores duplicate names.
- Shows matched, unmatched, and ambiguous names before writing.
- Adds matched devices to one eligible assigned security group.
- Uses a confirmation prompt that defaults to **No** and supports result export.

### App Dependency Check module

Location: `Modules\AppDependencyCheck\AppDependencyCheck.ps1`

- Maps Win32 app dependency and supersedence relationships.
- Flags circular references, chains deeper than five levels, and unassigned dependencies.
- Read-only: issues only GET requests and never writes to Intune.
- Lists every Win32 app A-Z in a drop-down; scan one app or all of them.
- Exports findings to CSV.

## User instructions

1. Download the toolkit ZIP from GitHub.
2. Extract the ZIP to a folder on your computer.
3. Review `Toolkit.ps1` and every PowerShell script in the `Modules` folder.
4. Double-click `Run-Toolkit.bat`.
5. On first launch, allow the toolkit to install `Microsoft.Graph.Authentication` for the current user if prompted.
6. Select **Connect to Graph** and complete the work or school account sign-in.

> **Production-use disclaimer:** Review and understand all included PowerShell scripts yourself before running this toolkit in a production environment. Confirm that the requested Microsoft Graph permissions, device actions, group operations, and automatic module installation meet your organization’s security, change-management, and compliance requirements.

`Run-Toolkit.bat` starts `Toolkit.ps1` in Windows PowerShell 5.1 with STA enabled and a process-only execution-policy bypass. It does not permanently change the user or computer execution policy.

## Microsoft Graph permissions

| Scope | Purpose |
|---|---|
| `DeviceManagementManagedDevices.ReadWrite.All` | Read and administer managed-device data. |
| `DeviceManagementConfiguration.Read.All` | Read Intune configuration data. |
| `Device.Read.All` | Resolve Entra device objects. |
| `Group.Read.All` | Read groups and memberships. |
| `Group.ReadWrite.All` | Support membership operations. |
| `GroupMember.ReadWrite.All` | Add and remove members. |
| `User.Read.All` | Read associated user information. |
| `DeviceManagementApps.Read.All` | Read Win32 app and relationship data. |

> Tenant policy can require administrator consent and suitable administrative roles.

## Using the toolkit

### Connect and search

1. Launch the toolkit.
2. Select **Connect to Graph**.
3. Complete sign-in.
4. Choose a search field and enter a term.
5. Select **Exact match only** if needed.
6. Search and select a result.

### Copy groups

1. Select the source device.
2. Select **Copy Device Groups**.
3. Choose the target.
4. Review eligible groups.
5. Confirm and review results.

### Remove groups

1. Select the device.
2. Select **Remove Device Groups**.
3. Adjust the selected memberships.
4. Confirm and review results.

### Bulk add

1. Select **Bulk Add to Group**.
2. Browse to the CSV.
3. Review matches.
4. Choose an eligible destination group.
5. Confirm and review or export results.

### App dependency check

1. Select **App Dependency Check**.
2. Pick an app from the drop-down, or leave it on **(All apps)**.
3. Select **Scan**.
4. Review the findings and export them if needed.

## Bulk-add CSV format

Every nonblank line—including line 1—is data. The first column supplies the device name.

```csv
LAPTOP-001
LAPTOP-002
KIOSK-014
```

- Do not include a header unless it should appear as an unmatched device.
- Extra columns are ignored.
- Duplicate names are ignored case-insensitively.
- A selectable row must match exactly one managed device with an Entra device object.

## App dependency checks

The App Dependency Check module is read-only. It issues only GET requests and never modifies apps, relationships, or assignments.

It reports three faults that stop a Win32 app chain from installing:

| Finding | Severity | Meaning |
|---|---|---|
| Circular reference | Critical | The chain returns to an app it has already visited. Intune cannot resolve it and the install never completes. |
| Chain too deep | Critical | The chain runs past the five levels Intune processes. Anything below level five is ignored at install time. |
| Unassigned dependency | Warning | A dependency has no assignment, so Intune has nothing to install and the parent app stays blocked. |

Supersedence relationships are included by default and can be excluded. Findings can be exported to CSV.

## Group eligibility

Only assigned, cloud-managed security groups are writable. Dynamic, rule-driven, Microsoft 365, mail-enabled, distribution, synchronized, and non-security groups are read-only.

## Repository structure

```text
.
├── Run-Toolkit.bat
├── Toolkit.ps1
├── MainWindow.xaml
└── Modules
    ├── CopyDeviceGroups
    │   ├── CopyDeviceGroups.ps1
    │   └── CopyDeviceGroups.xaml
    ├── RemoveDeviceGroups
    │   ├── RemoveDeviceGroups.ps1
    │   └── RemoveDeviceGroups.xaml
    ├── BulkAddToGroup
    │   ├── BulkAddToGroup.ps1
    │   └── BulkAddToGroup.xaml
    └── AppDependencyCheck
        ├── AppDependencyCheck.ps1
        └── AppDependencyCheck.xaml
```

PowerShell files contain embedded XAML for runtime use. External XAML supports UI development through `-XamlPath`.

## Diagnostics and safety

- Detailed diagnostics and Graph error information.
- Searches and previews do not change memberships.
- Ineligible rows cannot be selected.
- High-impact operations require confirmation.
- Writes occur one item at a time.
- Graph paging follows `@odata.nextLink`.
- App Dependency Check is read-only and issues only GET requests.

## Development notes

- Module functions use `Cdg`, `Rdg`, `Bag`, and `Adc` prefixes.
- Grid rows implement `INotifyPropertyChanged`.
- Intune and Entra device records are resolved separately.
- Keep embedded and external XAML synchronized.
- Bulk Add reads line 1 as data.

## Acknowledgments

Endpoint-Guy Intune Toolkit was built with assistance from Claude by Anthropic. Claude supported development, documentation, troubleshooting, and code review. Project decisions and responsibility remain with the project author.

## Disclaimer

Review every included PowerShell script before use, especially before running the toolkit in production. You are responsible for validating its behavior, permissions, security impact, exported data, and suitability for your environment.

