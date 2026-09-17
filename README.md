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
- [Asset statuses](#asset-statuses)
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
- Display device name, serial number, management name, user, operating system, OS version, compliance, ownership, model, and last sync.
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

- Lists every app that depends on a selected Win32 app, directly or indirectly.
- Shows the relationship type, depth, install behaviour, and the full chain.
- Read-only: issues only GET requests and never writes to Intune.
- Lists every Win32 app A-Z in a drop-down.
- Exports the list to CSV.

### Asset Status module

Location: `Modules\AssetStatus\AssetStatus.ps1`

- Sets the Intune Management name of the selected device to one of six lifecycle statuses: **In-Stock**, **Retired**, **Recycled**, **Stolen**, **Legalhold**, or **Lost**.
- Replaces the Management name entirely with the chosen status word.
- Shows the current Management name and a preview of the result before writing.
- Uses a confirmation prompt that names the device, shows the old and new name, and defaults to **No**.
- Acts on one device at a time and requires a device that has enrolled in Intune.
- Re-reads the device from Graph after a write so the window shows what Intune actually holds.

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
2. Pick the app you are about to change from the drop-down.
3. Select **Find dependents**.
4. Review the dependent apps and export the list if needed.

### Update asset status

1. Select the device in the search results.
2. Select **Update Asset Status**.
3. Review the current Management name shown for the device.
4. Pick one of the six statuses from the drop-down and check the preview.
5. Select **Update Management name**, then confirm at the prompt.

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

It answers one question: **which apps depend on this one?** Pick an app and the module walks the relationship graph upwards, listing every app that would be affected if it were changed, replaced, or removed.

| Column | Meaning |
|---|---|
| Dependent app | The app that depends on the selected app. |
| Relationship | `Dependency` or `Supersedence`. |
| Level | `Direct` for a first-level dependent, `Indirect (n)` when reached through another app. |
| Install | The Intune dependency behaviour: `autoInstall` or `detect`. |
| Assigned | Whether the dependent app is assigned to any group. |
| Chain | The full path from the selected app up to the dependent. |

Each app is reached by its shortest chain, so an app that is both a direct and an indirect dependent is reported as `Direct`. Supersedence is included by default and can be excluded. The list can be exported to CSV.

## Asset statuses

The Asset Status module writes one of six fixed lifecycle statuses to the Intune **Management name** (`managedDeviceName`) of the selected device. The status word replaces the Management name entirely.

| Status | Typical use |
|---|---|
| `In-Stock` | Held in stock and awaiting assignment. |
| `Retired` | Withdrawn from service. |
| `Recycled` | Sent for disposal or recycling. |
| `Stolen` | Reported stolen. |
| `Legalhold` | Retained for legal or investigative reasons. |
| `Lost` | Reported lost. |

Only these six values can be written; the status is chosen from a drop-down and free text is not accepted.

The Management name is a label only. Changing it does **not** change the device name, serial number, group memberships, or assignments, and it does **not** retire, wipe, or unenroll the device. Intune keeps no history of the previous Management name, so the old value cannot be restored from the toolkit — the confirmation prompt shows both the old and new name and defaults to **No**.

A device must be enrolled in Intune to have a Management name. A device that has been imported into Autopilot but has not yet enrolled has no `managedDevice` record, so the action reports this and stops.

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
    ├── AppDependencyCheck
    │   ├── AppDependencyCheck.ps1
    │   └── AppDependencyCheck.xaml
    └── AssetStatus
        ├── AssetStatus.ps1
        └── AssetStatus.xaml
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
- Asset Status replaces only the Management name and never retires, wipes, or unenrolls a device.

## Development notes

- Module functions use `Cdg`, `Rdg`, `Bag`, `Adc`, and `Ast` prefixes.
- Grid rows implement `INotifyPropertyChanged`.
- Intune and Entra device records are resolved separately.
- Keep embedded and external XAML synchronized.
- Bulk Add reads line 1 as data.

## Acknowledgments

Endpoint-Guy Intune Toolkit was built with assistance from Claude by Anthropic. Claude supported development, documentation, troubleshooting, and code review. Project decisions and responsibility remain with the project author.

## Disclaimer

Review every included PowerShell script before use, especially before running the toolkit in production. You are responsible for validating its behavior, permissions, security impact, exported data, and suitability for your environment.
