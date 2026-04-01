# M365 License Management

Interactive PowerShell script to manage Microsoft 365 license assignments using the Microsoft Graph API.

## Prerequisites

- PowerShell 7+
- Microsoft.Graph module (`Install-Module Microsoft.Graph -Scope CurrentUser`)
- Permissions: `User.ReadWrite.All`, `Directory.Read.All`, `Organization.Read.All`

## Usage

### Assign a license

Select a license to filter users by, review the user list, then pick a target license to assign:

```powershell
.\Assign-M365License.ps1
```

### Remove a license

Select a license and remove it from all users who have it assigned:

```powershell
.\Assign-M365License.ps1 -Remove
```

### Dry run

Preview what would happen without making any changes:

```powershell
.\Assign-M365License.ps1 -WhatIf
.\Assign-M365License.ps1 -Remove -WhatIf
```

### Rollback

Each operation saves a rollback JSON file. Use it to reverse the operation:

```powershell
.\Assign-M365License.ps1 -Rollback .\rollback_20260401_120000.json
```

- Rolling back an **assignment** removes the license from those users.
- Rolling back a **removal** re-assigns the license to those users.

## Interactive Flow

1. Connect to Microsoft Graph
2. Display all tenant licenses with assigned/total counts
3. Select a license to filter users by
4. Review and confirm the list of matched users
5. **Assign mode**: select a target license to assign (shows available counts)
   **Remove mode** (`-Remove`): confirm removal of the filter license
6. Final confirmation, then execute with per-user status reporting
7. Rollback file saved automatically
