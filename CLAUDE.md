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
```

## Architecture

- **Assign-M365License.ps1**: Single self-contained script. Uses `Get-MgSubscribedSku` for license inventory, `Get-MgUser` with OData filter for user lookups, and `Set-MgUserLicense` for assignments/removals. Every mutating operation saves a JSON rollback file automatically.
- **rollback_*.json**: Auto-generated rollback files containing affected user IDs, SKU info, and action type (`remove` or assign). Used by the `-Rollback` parameter to reverse operations.

## Conventions

- All scripts must be interactive with confirmation prompts before any mutating operation.
- Always support `-WhatIf` for dry runs.
- Always generate rollback files for reversibility.
- Use `Connect-MgGraph` with explicit scopes (least privilege).
