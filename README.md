# Endpoint-Guy Intune Toolkit

Endpoint-Guy Intune Toolkit is a Windows PowerShell 5.1 and WPF application for Microsoft Intune and Microsoft Entra ID device administration through Microsoft Graph.

> **Recommended:** unblock the GitHub ZIP, then double-click `Run-Toolkit.bat`.

## Table of contents

- [Features](#features)
- [Requirements](#requirements)
- [User instructions](#user-instructions)
- [Microsoft Graph permissions](#microsoft-graph-permissions)
- [Using the toolkit](#using-the-toolkit)
- [Bulk-add CSV format](#bulk-add-csv-format)
- [Group eligibility](#group-eligibility)
- [Repository structure](#repository-structure)
- [Diagnostics and safety](#diagnostics-and-safety)
- [Troubleshooting](#troubleshooting)
- [Development notes](#development-notes)
- [Acknowledgments](#acknowledgments)

## Features

### Managed-device search

- Interactive delegated Microsoft Graph sign-in.
- Search by device name, serial number, or primary username.
- Contains or exact matching.
- Local managed-device cache with manual refresh.
- Displays OS, compliance, ownership, model, user, and last sync.
- UTF-8 CSV export.

### Copy Device Groups

- Copies eligible assigned security-group memberships from a source device to a target device.
- Shows eligible and ineligible memberships.
- Requires confirmation.
- Reports **Added**, **Already a member**, or **Failed** per group.

### Remove Device Groups

- Removes a device from eligible assigned security groups.
- Preselects eligible memberships.
- Confirmation defaults to **No**.
- Reports **Removed**, **Not a member**, or **Failed** per group.

### Bulk Add to Group

- Reads device names from the first CSV column.
- Reads every nonblank line, including line 1.
- Ignores blank and duplicate names.
- Shows unmatched and ambiguous names before writing.
- Adds matched devices to one eligible group.
- Confirmation defaults to **No**.
- Exports results.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Network access to the PowerShell Gallery for first-time setup
- Network access to Microsoft Graph
- Appropriate Microsoft Intune and Entra roles

The toolkit automatically installs `Microsoft.Graph.Authentication` for the current user if it is missing.

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
    └── BulkAddToGroup
        ├── BulkAddToGroup.ps1
        └── BulkAddToGroup.xaml
```

PowerShell files contain embedded XAML for runtime use. External XAML supports UI development through `-XamlPath`.

## Diagnostics and safety

- Detailed diagnostics and Graph error information.
- Searches and previews do not change memberships.
- Ineligible rows cannot be selected.
- High-impact operations require confirmation.
- Writes occur one item at a time.
- Graph paging follows `@odata.nextLink`.

## Troubleshooting

### A module is not digitally signed

Windows may have marked the downloaded scripts as coming from the internet. In Windows PowerShell, run:

```powershell
Get-ChildItem ".\EndpointGuyIntuneToolkit" -Recurse -File | Unblock-File
```

Then double-click `Run-Toolkit.bat` again.

### First-time setup fails

- Confirm access to the PowerShell Gallery.
- Confirm proxy or policy does not block module installation.
- Restart after resolving connectivity.

### Sign-in fails

- Confirm account, tenant, Conditional Access, consent, and administrative roles.

### A device is not found

- Refresh the cache.
- Try serial number or username.
- Clear exact matching.
- Confirm the device has an Entra device ID.

## Development notes

- Module functions use `Cdg`, `Rdg`, and `Bag` prefixes.
- Grid rows implement `INotifyPropertyChanged`.
- Intune and Entra device records are resolved separately.
- Keep embedded and external XAML synchronized.
- Bulk Add reads line 1 as data.

## Acknowledgments

Endpoint-Guy Intune Toolkit was built with assistance from Claude by Anthropic. Claude supported development, documentation, troubleshooting, and code review. Project decisions and responsibility remain with the project author.

## Disclaimer

Review every included PowerShell script before use, especially before running the toolkit in production. You are responsible for validating its behavior, permissions, security impact, exported data, and suitability for your environment.




























<img width="268" height="32766" alt="image" src="https://github.com/user-attachments/assets/7d707dbf-ee15-4cd2-8124-249cd6c19fc9" />
