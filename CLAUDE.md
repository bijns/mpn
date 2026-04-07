# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This repo contains PowerShell scripts for Microsoft 365 tenant administration using the Microsoft Graph API. Scripts are interactive and designed for safe, auditable license management operations.

## Prerequisites

- PowerShell 7+
- Microsoft.Graph module (`Install-Module Microsoft.Graph -Scope CurrentUser`)
- Tenant permissions: `User.ReadWrite.All`, `Directory.Read.All`, `Organization.Read.All`

## Running

```powershell
# Assign a license interactively
.\Assign-M365License.ps1

# Remove a license from filtered users
.\Assign-M365License.ps1 -Remove

# Dry run (no changes)
.\Assign-M365License.ps1 -WhatIf

# Rollback a previous operation
.\Assign-M365License.ps1 -Rollback .\rollback_<timestamp>.json

# Compare two users' licenses and copy from one to the other
.\Compare-M365UserLicenses.ps1

# Find users with one license but not another (non-blocked only)
.\Get-M365UsersWithoutLicense.ps1

# Dry run
.\Compare-M365UserLicenses.ps1 -WhatIf
.\Get-M365UsersWithoutLicense.ps1 -WhatIf

# Rollback a copy operation
.\Compare-M365UserLicenses.ps1 -Rollback .\rollback_copy_<timestamp>.json
```

## Architecture

- **Assign-M365License.ps1**: Single self-contained script. Uses `Get-MgSubscribedSku` for license inventory, `Get-MgUser` with OData filter for user lookups, and `Set-MgUserLicense` for assignments/removals. Every mutating operation saves a JSON rollback file automatically.
- **Compare-M365UserLicenses.ps1**: Filter users by license, pick two to compare side-by-side, then copy all missing licenses from one to the other. Saves a rollback file for reversibility.
- **Get-M365UsersWithoutLicense.ps1**: Read-only report script. Pick a "must have" and "must NOT have" license, lists all enabled (non-blocked) users matching those criteria. Optional CSV export.
- **rollback_*.json**: Auto-generated rollback files containing affected user IDs, SKU info, and action type (`remove`, assign, or `copy`). Used by the `-Rollback` parameter to reverse operations.

Each script is self-contained with its own copies of `Show-Menu` and `Confirm-Prompt` helpers (no shared module). Graph API pattern: `Get-MgSubscribedSku` for license inventory, `Get-MgUser -Filter "assignedLicenses/any(...)"` with `-ConsistencyLevel eventual` for user lookups, `Set-MgUserLicense` for mutations.

## Conventions

- All scripts must be interactive with confirmation prompts before any mutating operation.
- Always support `-WhatIf` for dry runs.
- Always generate rollback files for reversibility.
- Use `Connect-MgGraph` with explicit scopes (least privilege).
- Scripts are self-contained — duplicate shared helpers rather than extracting a module.
