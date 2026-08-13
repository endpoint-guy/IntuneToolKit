# Endpoint-Guy Intune Toolkit

Endpoint-Guy Intune Toolkit is a Windows PowerShell 5.1 and WPF desktop application for Microsoft Intune and Microsoft Entra ID device administration through Microsoft Graph.

It provides a dark-themed operator interface for finding managed devices, exporting results, copying assigned security-group memberships between devices, removing a device from eligible groups, and adding devices from a CSV file to one group.

> **Two ways to run it:** the repository includes a prebuilt `EndpointguyToolkit.exe` for normal use, and it also includes `Toolkit.ps1` plus `Build-Exe.ps1` for users who prefer to run the script or rebuild the EXE themselves.

> **Project status:** The device search and the three group-management workflows documented below are implemented. Any controls marked **Coming soon** are placeholders for future functionality.

## Table of contents

- [Features](#features)
- [Requirements](#requirements)
- [User instructions](#user-instructions)
- [Quick start with the prebuilt EXE](#quick-start-with-the-prebuilt-exe)
- [Run from PowerShell](#run-from-powershell)
- [Build or rebuild the EXE](#build-or-rebuild-the-exe)
- [Microsoft Graph permissions](#microsoft-graph-permissions)
- [Using the toolkit](#using-the-toolkit)
- [Bulk-add CSV format](#bulk-add-csv-format)
- [Group eligibility](#group-eligibility)
- [Authentication behavior in the EXE](#authentication-behavior-in-the-exe)
- [Repository structure](#repository-structure)
- [Diagnostics and logging](#diagnostics-and-logging)
- [Safety controls](#safety-controls)
- [Troubleshooting](#troubleshooting)
- [Development notes](#development-notes)
- [Security considerations](#security-considerations)
- [Acknowledgments](#acknowledgments)

## Features

### Managed-device search

- Interactive delegated sign-in to Microsoft Graph.
- Displays the connected account and tenant status in the application.
- Searches Intune managed devices by:
  - Device name
  - Serial number
  - Primary username
- Uses contains matching by default.
- Supports **Exact match only** when an equality search is required.
- Caches the managed-device inventory locally for fast repeat searches.
- Provides a **Refresh Device Cache** action.
- Displays device name, serial number, primary user, operating system, OS version, compliance, ownership, model, and last sync time.
- Exports the current result set to a UTF-8 CSV file.

### Copy Device Groups

Copies eligible assigned security-group memberships from a selected source device to a target device.

- Uses the device selected in the main search results as the source.
- Searches for and selects a target managed device.
- Resolves Intune records to their corresponding Entra device objects.
- Reads all source-device group memberships, including paged Graph results.
- Classifies every group as eligible or ineligible before any write.
- Preselects eligible assigned security groups.
- Can display ineligible memberships and their reason.
- Requires confirmation before changes are made.
- Adds groups one at a time.
- Reports **Added**, **Already a member**, or **Failed** for each group.
- Displays an added/existing/failed summary when complete.

### Remove Device Groups

Removes the selected device from eligible assigned security groups.

- Reads and classifies every current group membership.
- Preselects all eligible assigned security groups for the common remove-all workflow.
- Disables selection for groups that cannot be directly modified.
- Allows the operator to clear any eligible groups that should remain.
- Requires confirmation naming the device and selected group count.
- The confirmation defaults to **No**.
- Removes memberships one at a time.
- Reports **Removed**, **Not a member**, or **Failed** for every selected group.

### Bulk Add to Group

Adds many devices listed in a CSV file to one eligible assigned security group.

- Reads device names from the first CSV column.
- Reads every nonblank line, including line 1.
- Does not assume or skip a header row.
- Ignores blank lines.
- Ignores duplicate names case-insensitively and reports the count.
- Shows matched, unmatched, and ambiguous names before any write.
- Selects only names that match exactly one managed device and resolve to an Entra device object.
- Searches destination groups by display-name prefix.
- Also accepts a pasted group object ID.
- Requires confirmation naming the group and selected device count.
- The confirmation defaults to **No**.
- Adds devices one at a time.
- Reports **Added**, **Already a member**, or **Failed** per device.
- Exports the result grid to UTF-8 CSV.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- Internet access to the PowerShell Gallery on first launch if `Microsoft.Graph.Authentication` is not already installed
- Network access to `https://graph.microsoft.com`
- A Microsoft Entra account permitted to use or consent to the requested delegated scopes
- Appropriate Microsoft Intune and Entra administrative roles for the selected operations
- PowerShell execution policy that permits local scripts when running or rebuilding from source

### Automatic first-time setup

When the toolkit starts, it checks for `Microsoft.Graph.Authentication`. If the module is missing, the toolkit:

1. Enables TLS 1.2 for the PowerShell Gallery connection.
2. Installs the NuGet package provider for the current user when required.
3. Installs `Microsoft.Graph.Authentication` with `-Scope CurrentUser`.
4. Imports the module and continues launching.

This process does not normally require administrator rights. The only manual configuration expected from the user is the execution-policy setting below.

Configure the current-user execution policy if permitted by your organization:

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

> Review execution-policy changes with your security team. Group Policy can override the current-user setting.

## User instructions

Follow this section when downloading the toolkit from GitHub. Windows marks downloaded ZIP files as coming from the internet. If the archive is extracted while still blocked, that mark can propagate to the included PowerShell modules and cause errors stating that a module is not digitally signed.

> Unblocking the downloaded archive removes its Mark-of-the-Web. It does **not** disable PowerShell execution policy globally and does not add an execution-policy bypass to the toolkit.

### Recommended: unblock before extracting

1. Download the toolkit ZIP from GitHub.
2. Open Windows PowerShell in the folder containing the downloaded ZIP.
3. Replace the example ZIP filename below if your downloaded file has a different name.
4. Run:

```powershell
Unblock-File -Path ".\IntuneToolKit-main.zip"
Expand-Archive -Path ".\IntuneToolKit-main.zip" -DestinationPath ".\EndpointGuyIntuneToolkit"
```

5. Open the extracted `EndpointGuyIntuneToolkit` folder.
6. Launch the included `EndpointguyToolkit.exe`.
7. If Windows Defender SmartScreen displays **Windows protected your PC**, verify that the file came from this project’s official GitHub release. For an unsigned community build, select **More info** and then **Run anyway** only if you trust the download.
8. On first launch, allow the toolkit to install `Microsoft.Graph.Authentication` for the current user if prompted.
9. Select **Connect to Graph** and complete the work or school account sign-in.

### If the ZIP was already extracted

Run the following command against the extracted toolkit directory:

```powershell
Get-ChildItem ".\EndpointGuyIntuneToolkit" -Recurse -File | Unblock-File
```

Then launch:

```powershell
.\EndpointGuyIntuneToolkit\EndpointguyToolkit.exe
```

### Run the source script instead

After unblocking the downloaded package, users can run the source version instead of the EXE:

```powershell
powershell.exe -NoProfile -STA -File .\EndpointGuyIntuneToolkit\Toolkit.ps1
```

If your organization requires `RemoteSigned`, configure it once for the current user as described in [Requirements](#requirements). A permanent `Bypass` policy is not recommended.

### Verify the files are unblocked

This command should return no output after the files are successfully unblocked:

```powershell
Get-ChildItem ".\EndpointGuyIntuneToolkit" -Recurse -File |
    Get-Item -Stream Zone.Identifier -ErrorAction SilentlyContinue
```

## Quick start with the prebuilt EXE

The repository includes a prebuilt `EndpointguyToolkit.exe`. Most users do not need to compile anything.

For a GitHub download, complete [User instructions](#user-instructions) first so Windows does not pass Mark-of-the-Web to the included module scripts.

1. Download or clone the repository.
2. Launch the application. If needed, it installs `Microsoft.Graph.Authentication` automatically for the current user.
3. Keep `EndpointguyToolkit.exe` and the `Modules` folder together in the repository layout.
4. Complete the first-time module setup prompt if it appears.
5. Select **Connect to Graph**.
6. Complete the Microsoft work or school account sign-in.
7. Search for a device and use the required action.

```powershell
.\EndpointguyToolkit.exe
```

> On first launch, the toolkit checks for `Microsoft.Graph.Authentication`. If it is missing, the toolkit installs it from the PowerShell Gallery for the current user without requiring administrator elevation.

## Run from PowerShell

Users who do not want to use the prebuilt executable can run `Toolkit.ps1` directly:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1
```

To target a specific tenant:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1 -TenantId "00000000-0000-0000-0000-000000000000"
```

To load the external main-window XAML during UI development:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File .\Toolkit.ps1 -XamlPath .\MainWindow.xaml
```

`-STA` is required because the interface is built with WPF.

## Build or rebuild the EXE

`Build-Exe.ps1` compiles the current `Toolkit.ps1` into `EndpointguyToolkit.exe`.

Rebuilding is optional. Use it when you:

- Modify `Toolkit.ps1`.
- Want a new local build.
- Change the application version.
- Add or replace the application icon.
- Need to replace the prebuilt EXE with an updated release.

### Default build

Run from Windows PowerShell 5.1 in the repository root:

```powershell
.\Build-Exe.ps1
```

Default inputs and output:

| Setting | Default |
|---|---|
| Source | `Toolkit.ps1` |
| Output | `EndpointguyToolkit.exe` |
| Version | `1.0.0.0` |

### Build with an icon and version

```powershell
.\Build-Exe.ps1 -IconFile .\toolkit.ico -Version 1.1.0.0
```

### Override source or output

```powershell
.\Build-Exe.ps1 -Source .\Toolkit.ps1 -Output .\EndpointguyToolkit.exe
```

### What the build script does

- Installs the `ps2exe` module for the current user when it is missing.
- Verifies that the source file exists.
- Ensures the source uses a PS2EXE-compatible UTF-8 BOM.
- Compiles as a GUI application with no console window.
- Enables STA for WPF.
- Enables high-DPI awareness.
- Targets a 64-bit runtime.
- Adds product, company, copyright, and version metadata.
- Confirms that the output file was created.

> Rebuilding replaces the output path you specify. Keep a backup or commit your current EXE before producing a new release build.

## Microsoft Graph permissions

The main toolkit requests these delegated scopes:

| Scope | Used for |
|---|---|
| `DeviceManagementManagedDevices.ReadWrite.All` | Reading Intune managed devices and supporting device-management operations. |
| `DeviceManagementConfiguration.Read.All` | Reading Intune device-management configuration data. |
| `Device.Read.All` | Resolving Intune records to Entra device objects. |
| `Group.Read.All` | Reading group details and memberships. |
| `Group.ReadWrite.All` | Supporting group membership changes. |
| `GroupMember.ReadWrite.All` | Adding and removing device objects from groups. |
| `User.Read.All` | Reading user information associated with managed devices. |

When an action module is run independently, it requests the scopes needed by that module, including `DeviceManagementManagedDevices.Read.All`, `Device.Read.All`, `Group.Read.All`, `Group.ReadWrite.All`, and `GroupMember.ReadWrite.All`.

> Tenant policy can require administrator consent. The connected account must also hold suitable directory and Intune roles; Graph consent alone does not grant an operator permission to administer every object.

## Using the toolkit

### Connect and search

1. Launch the EXE or run `Toolkit.ps1`.
2. Select **Connect to Graph**.
3. Complete delegated sign-in.
4. Select **Device Name**, **Serial Number**, or **Primary Username**.
5. Enter a search term.
6. Select **Exact match only** if required.
7. Select **Search**.
8. Select a result row to enable device-specific actions.

### Copy groups

1. Select the source device in the main results.
2. Select **Copy Device Groups**.
3. Search for and choose the target device.
4. Review eligible and ineligible groups.
5. Adjust the selected eligible groups.
6. Confirm the copy.
7. Review the result for every group.

### Remove groups

1. Select a device in the main results.
2. Select **Remove Device Groups**.
3. Review the preselected eligible memberships.
4. Clear groups that should remain.
5. Confirm the removal prompt.
6. Review the per-group results.

### Bulk add devices

1. Select **Bulk Add to Group**.
2. Browse to the CSV file.
3. Review all match states.
4. Search for the destination group or paste its object ID.
5. Select an eligible group.
6. Adjust the selected matched devices.
7. Confirm the bulk add.
8. Review or export the result grid.

## Bulk-add CSV format

The parser treats every nonblank line as data. It does not skip line 1 as a header.

```csv
LAPTOP-001
LAPTOP-002
KIOSK-014
```

Rules:

- Place the device name in the first column.
- Do not include a header unless you intentionally want it shown as an unmatched name.
- Extra columns are ignored.
- Blank lines are ignored.
- Duplicate names are ignored case-insensitively.
- Quoted first-column values containing commas are supported.
- A row is selectable only when it matches exactly one Intune managed device with a corresponding Entra device object.

## Group eligibility

Only assigned, cloud-managed security groups are writable.

| Group type or state | Writable? | Reason |
|---|:---:|---|
| Assigned cloud security group | Yes | Direct membership changes are supported. |
| Dynamic membership group | No | Membership is controlled by a rule. |
| Group with a membership rule | No | Rule-driven membership is authoritative. |
| Microsoft 365 / Unified group | No | Not treated as a writable device security group. |
| Mail-enabled security group | No | Not eligible for these device workflows. |
| Distribution group | No | Not an eligible security target. |
| On-premises synchronized group | No | Membership is mastered outside Entra ID. |
| Non-security group | No | Not a security-group target. |

Ineligible groups can be displayed for transparency, but their selection controls remain disabled.

## Authentication behavior in the EXE

The toolkit continues to use the standard interactive delegated `Connect-MgGraph` flow.

Recent Microsoft Graph PowerShell releases use Windows Web Account Manager (WAM) on Windows. WAM requires a parent window handle. Because the EXE is compiled by PS2EXE as a no-console GUI application, the toolkit creates a temporary hidden console HWND during sign-in so WAM has a valid parent handle.

During the sign-in flow:

1. The toolkit creates the temporary hidden authentication window only when the process does not already have a console HWND.
2. The main toolkit window minimizes so the WAM account and credential prompts cannot become trapped behind it.
3. `Connect-MgGraph` runs normally.
4. The temporary HWND is released.
5. The toolkit restores and activates its previous window state whether sign-in succeeds, fails, or is cancelled.

Running `Toolkit.ps1` from a normal PowerShell console already provides a console HWND, so no temporary one is created.

## Repository structure

```text
.
├── EndpointguyToolkit.exe             # Prebuilt executable for normal use
├── Toolkit.ps1                        # Main WPF application and default build source
├── Build-Exe.ps1                      # Optional PS2EXE build script
├── MainWindow.xaml                    # External UI override for development
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

The main and module PowerShell files contain embedded XAML for self-contained runtime use. The external XAML files are development overrides supplied with `-XamlPath`.

The main toolkit looks for modules in the documented `Modules` subfolders and also probes the legacy flat layout beside `Toolkit.ps1` or the EXE.

## Diagnostics and logging

The copy and remove modules provide diagnostics support including:

- Timestamped `INFO`, `OK`, `WARN`, `ERROR`, `GRAPH`, and `DEBUG` entries.
- Optional verbose Graph timings and eligibility decisions.
- Auto-scroll.
- Copy log.
- Save log.
- Clear log.
- Exception type and inner-exception details.
- Fully qualified error ID and category.
- Invocation line and script stack trace.

When no on-screen log sink is active, messages are sent to `Write-Verbose`.

## Safety controls

- Searches and previews do not change membership.
- Ineligible devices and groups cannot be selected for writes.
- Dynamic and synchronized groups are never offered as writable targets.
- Remove and bulk-add confirmation dialogs default to **No**.
- Writes are processed one item at a time.
- Every selected item receives an individual result.
- **Already a member** and **Not a member** are treated as safe idempotent outcomes where possible.
- Graph collection requests follow `@odata.nextLink`.
- Graph response properties are checked defensively.
- Error handling attempts to display the useful Graph response body rather than only a generic HTTP status.

## Troubleshooting

### A module cannot be loaded because it is not digitally signed

This commonly occurs when the GitHub ZIP was extracted before it was unblocked. Remove Mark-of-the-Web from the extracted files:

```powershell
Get-ChildItem ".\EndpointGuyIntuneToolkit" -Recurse -File | Unblock-File
```

Restart the toolkit afterward. This is preferable to changing the machine or user execution policy to `Bypass`.

### First-time Graph module setup fails

The toolkit installs `Microsoft.Graph.Authentication` automatically for the current user when it is missing. If installation fails:

- Confirm the computer can reach the PowerShell Gallery.
- Confirm PowerShell Gallery access is not blocked by a proxy or organizational policy.
- Confirm the current user can write to the current-user PowerShell module path.
- Restart the toolkit after connectivity or policy issues are resolved.

No administrator elevation is normally required because installation uses `-Scope CurrentUser`.

### The EXE does not contain recent source changes

The prebuilt EXE does not update when `Toolkit.ps1` changes. Rebuild it:

```powershell
.\Build-Exe.ps1
```

Then replace the old `EndpointguyToolkit.exe` with the newly generated file.

### The window does not open

- Confirm Windows PowerShell 5.1 is being used.
- When running the script, include `-STA`.
- If first-time setup failed, confirm the PowerShell Gallery is reachable and restart the toolkit.
- Confirm the source file was not modified with a text editor that damaged a PowerShell here-string terminator.
- When using `-XamlPath`, verify the XAML exists and contains all required named controls.

### Graph sign-in fails

- Confirm the account can sign in to the intended tenant.
- Use `-TenantId` if the account is associated with multiple tenants.
- Confirm Conditional Access allows the sign-in context.
- Confirm an administrator has granted any required consent.
- Update `Microsoft.Graph.Authentication` if the installed version has a known WAM defect.

### The WAM prompt appears behind the application

The current toolkit minimizes its main window during authentication and restores it afterward. If this still occurs, confirm that the EXE was rebuilt from the current `Toolkit.ps1` and that the old executable was replaced.

### A device is not found

- Refresh the device cache.
- Try serial number or primary username.
- Clear **Exact match only** for a contains search.
- Confirm the device exists in Intune.
- Confirm its managed-device record has an Entra device ID.
- For bulk add, confirm column 1 contains the device name and no unintended header.

### A group is visible but disabled

Review the **Eligibility** column. Dynamic, rule-driven, Microsoft 365, mail-enabled, distribution, synchronized, and non-security groups are intentionally not writable.

### The EXE cannot find an action module

Keep the `Modules` folder beside `EndpointguyToolkit.exe` using the structure shown above.

## Development notes

- The action scripts are dot-sourced into one PowerShell session.
- Module functions use `Cdg`, `Rdg`, and `Bag` prefixes to avoid collisions.
- Grid row types implement `INotifyPropertyChanged` so selections and results repaint correctly.
- Intune managed-device records and Entra device objects are treated as separate objects and explicitly resolved.
- Group details are read before eligibility decisions so membership rules are not missed.
- External XAML is intended for UI iteration; embedded XAML is the default runtime interface.
- Keep external and embedded XAML synchronized before a release.
- The Bulk Add parser reads line 1 as data. Ensure external XAML help text does not claim that line 1 is skipped.
- `Build-Exe.ps1` now compiles `Toolkit.ps1` directly.

## Security considerations

This application can change Entra group memberships and requests high-impact delegated permissions.

- Review the source and requested scopes before production use.
- Use the least-privileged Entra and Intune roles that support the required tasks.
- Test new builds in a non-production tenant.
- Review CSV contents before confirming bulk operations.
- Protect exported CSV files and diagnostic logs.
- Do not commit organizational device data, usernames, tenant details, or Graph error bodies to a public repository.
- Follow your organization’s change-management and audit requirements.

## Contributing

Issues and pull requests are welcome. For UI changes, update both the external XAML file and the embedded XAML in the matching PowerShell script. Include test notes for authentication, device resolution, group eligibility, pagination, confirmation behavior, and Graph error handling.

## Acknowledgments

Endpoint-Guy Intune Toolkit was built with assistance from Claude by Anthropic. Claude supported development, documentation, troubleshooting, and code review. Project decisions and responsibility for the final implementation remain with the project author.

## Disclaimer

This project is provided as an administrative tool. Review and test it in your environment before production use. Microsoft Intune, Microsoft Entra ID, Microsoft Graph, Windows PowerShell, and WPF are Microsoft technologies and trademarks.
<img width="268" height="32766" alt="image" src="https://github.com/user-attachments/assets/944e39bd-f981-4f6d-9b8d-ed21a707b56d" />
