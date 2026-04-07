# M365 License Management

Interactive PowerShell scripts to manage Microsoft 365 license assignments using the Microsoft Graph API.

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

### Find users with one license but not another

Pick a "must have" and a "must NOT have" license, then list all enabled (non-blocked) users that match. Optionally export to CSV:

```powershell
.\Get-M365UsersWithoutLicense.ps1
```

### Dry run

Preview what would happen without making any changes:

```powershell
.\Assign-M365License.ps1 -WhatIf
.\Assign-M365License.ps1 -Remove -WhatIf
.\Compare-M365UserLicenses.ps1 -WhatIf
.\Get-M365UsersWithoutLicense.ps1 -WhatIf
```

### Rollback

Each operation saves a rollback JSON file. Use it to reverse the operation:

```powershell
.\Assign-M365License.ps1 -Rollback .\rollback_20260401_120000.json
.\Compare-M365UserLicenses.ps1 -Rollback .\rollback_copy_20260401_120000.json
```

- Rolling back an **assignment** removes the license from those users.
- Rolling back a **removal** re-assigns the license to those users.
- Rolling back a **copy** removes the copied licenses from the target user.

### Compare and copy licenses between users

Filter users by license, pick two to compare side-by-side, then copy all missing licenses from one to the other:

```powershell
.\Compare-M365UserLicenses.ps1
```

## Interactive Flows

### Assign-M365License.ps1

1. Connect to Microsoft Graph
2. Display all tenant licenses with assigned/total counts
3. Select a license to filter users by
4. Review and confirm the list of matched users
5. **Assign mode**: select a target license to assign (shows available counts)
   **Remove mode** (`-Remove`): confirm removal of the filter license
6. Final confirmation, then execute with per-user status reporting
7. Rollback file saved automatically

### Compare-M365UserLicenses.ps1

1. Connect to Microsoft Graph
2. Select a license to filter users by
3. Pick two users from the filtered list
4. View color-coded side-by-side license comparison
5. Choose copy direction (which user's licenses to copy to the other)
6. Confirm and execute
7. Rollback file saved automatically

### Get-M365UsersWithoutLicense.ps1

1. Connect to Microsoft Graph (read-only scopes)
2. Select the license users must have
3. Select the license users must NOT have
4. Fetch and filter: excludes blocked users and users who have the excluded license
5. Display matching users with summary counts
6. Optionally export results to CSV
