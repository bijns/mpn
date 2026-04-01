<#
.SYNOPSIS
    Interactively filter M365 users by license and assign a different license.

.DESCRIPTION
    Connects to Microsoft Graph, lets you pick a license to filter users by,
    then pick a target license to assign to those users. Saves a rollback file
    for each operation.

.PARAMETER Remove
    Remove the selected filter license from the filtered users instead of assigning a new one.

.PARAMETER Rollback
    Path to a rollback JSON file to reverse a previous assignment.

.PARAMETER WhatIf
    Show what would happen without making changes.

.EXAMPLE
    .\Assign-M365License.ps1
    .\Assign-M365License.ps1 -Remove
    .\Assign-M365License.ps1 -Rollback .\rollback_20260401_120000.json
    .\Assign-M365License.ps1 -WhatIf
#>
[CmdletBinding()]
param(
    [switch]$Remove,
    [string]$Rollback,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# --- Helper: numbered menu selection ---
function Show-Menu {
    param(
        [string]$Title,
        [array]$Items,
        [scriptblock]$FormatItem
    )
    Write-Host "`n$Title" -ForegroundColor Cyan
    Write-Host ("-" * $Title.Length) -ForegroundColor Cyan
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $line = & $FormatItem $Items[$i]
        Write-Host ("  [{0,3}] {1}" -f ($i + 1), $line)
    }
    Write-Host ""
    do {
        $input = Read-Host "Select a number (1-$($Items.Count))"
        $num = $input -as [int]
    } while ($num -lt 1 -or $num -gt $Items.Count)
    return $Items[$num - 1]
}

function Confirm-Prompt {
    param([string]$Message)
    $answer = Read-Host "$Message (Y/N)"
    return $answer -eq 'Y' -or $answer -eq 'y'
}

# --- Connect to Microsoft Graph ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All", "Organization.Read.All" -NoWelcome

# --- Rollback mode ---
if ($Rollback) {
    if (-not (Test-Path $Rollback)) {
        Write-Host "Rollback file not found: $Rollback" -ForegroundColor Red
        exit 1
    }

    $rollbackData = Get-Content $Rollback -Raw | ConvertFrom-Json
    $skuId = $rollbackData.SkuId
    $skuName = $rollbackData.SkuName
    $users = $rollbackData.Users
    $wasRemoval = $rollbackData.Action -eq "remove"

    if ($wasRemoval) {
        Write-Host "`nRollback: Re-assign license '$skuName' to $($users.Count) user(s)" -ForegroundColor Cyan
    } else {
        Write-Host "`nRollback: Remove license '$skuName' from $($users.Count) user(s)" -ForegroundColor Cyan
    }
    Write-Host ("-" * 60) -ForegroundColor Cyan
    foreach ($u in $users) {
        Write-Host "  $($u.DisplayName) ($($u.UserPrincipalName))"
    }

    if (-not (Confirm-Prompt "`nProceed with rollback?")) {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }

    $success = 0
    $failed = 0
    foreach ($u in $users) {
        try {
            if ($wasRemoval) {
                if ($WhatIf) {
                    Write-Host "  [WhatIf] Would re-assign '$skuName' to $($u.UserPrincipalName)" -ForegroundColor DarkGray
                } else {
                    Set-MgUserLicense -UserId $u.UserId -AddLicenses @(@{SkuId = $skuId }) -RemoveLicenses @()
                    Write-Host "  Re-assigned to $($u.UserPrincipalName)" -ForegroundColor Green
                }
            } else {
                if ($WhatIf) {
                    Write-Host "  [WhatIf] Would remove '$skuName' from $($u.UserPrincipalName)" -ForegroundColor DarkGray
                } else {
                    Set-MgUserLicense -UserId $u.UserId -AddLicenses @() -RemoveLicenses @($skuId)
                    Write-Host "  Removed from $($u.UserPrincipalName)" -ForegroundColor Green
                }
            }
            $success++
        } catch {
            Write-Host "  FAILED for $($u.UserPrincipalName): $_" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "`nRollback complete: $success succeeded, $failed failed." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Normal mode: Step 1 - Fetch licenses ---
Write-Host "Fetching tenant licenses..." -ForegroundColor Yellow
$skus = Get-MgSubscribedSku | Sort-Object SkuPartNumber

if ($skus.Count -eq 0) {
    Write-Host "No licenses found in tenant." -ForegroundColor Red
    exit 1
}

# --- Step 2: Select filter license ---
$filterSku = Show-Menu -Title "Select a license to FILTER users by" -Items $skus -FormatItem {
    param($sku)
    $enabled = $sku.PrepaidUnits.Enabled
    $consumed = $sku.ConsumedUnits
    "{0,-50} ({1}/{2} assigned)" -f $sku.SkuPartNumber, $consumed, $enabled
}

Write-Host "`nSelected filter: $($filterSku.SkuPartNumber)" -ForegroundColor Green
Write-Host "Fetching users with this license..." -ForegroundColor Yellow

$users = Get-MgUser -All -Filter "assignedLicenses/any(l:l/skuId eq $($filterSku.SkuId))" `
    -Property Id, DisplayName, UserPrincipalName, AssignedLicenses `
    -ConsistencyLevel eventual -CountVariable userCount

if ($users.Count -eq 0) {
    Write-Host "No users found with license '$($filterSku.SkuPartNumber)'." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 3: Confirm user list ---
Write-Host "`nUsers with '$($filterSku.SkuPartNumber)' ($($users.Count) total):" -ForegroundColor Cyan
Write-Host ("-" * 60) -ForegroundColor Cyan
foreach ($u in $users) {
    Write-Host ("  {0,-35} {1}" -f $u.DisplayName, $u.UserPrincipalName)
}

if (-not (Confirm-Prompt "`nContinue with these $($users.Count) user(s)?")) {
    Write-Host "Cancelled." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Remove mode: remove the filter license from these users ---
if ($Remove) {
    Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
    Write-Host "  License to REMOVE: $($filterSku.SkuPartNumber)" -ForegroundColor Red
    Write-Host "  Users affected   : $($users.Count)"
    if ($WhatIf) {
        Write-Host "  Mode             : WhatIf (no changes)" -ForegroundColor DarkGray
    }
    Write-Host "===================" -ForegroundColor Cyan

    if (-not (Confirm-Prompt "`nRemove '$($filterSku.SkuPartNumber)' from $($users.Count) user(s)?")) {
        Write-Host "Cancelled." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        exit 0
    }

    $successUsers = @()
    $failed = 0
    foreach ($u in $users) {
        try {
            if ($WhatIf) {
                Write-Host "  [WhatIf] Would remove '$($filterSku.SkuPartNumber)' from $($u.UserPrincipalName)" -ForegroundColor DarkGray
            } else {
                Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses @($filterSku.SkuId)
                Write-Host "  Removed from $($u.UserPrincipalName)" -ForegroundColor Green
            }
            $successUsers += @{
                UserId            = $u.Id
                DisplayName       = $u.DisplayName
                UserPrincipalName = $u.UserPrincipalName
            }
        } catch {
            Write-Host "  FAILED for $($u.UserPrincipalName): $_" -ForegroundColor Red
            $failed++
        }
    }

    # Save rollback file (rollback = re-assign the removed license)
    if ($successUsers.Count -gt 0 -and -not $WhatIf) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $rollbackPath = Join-Path (Get-Location) "rollback_remove_$timestamp.json"
        $rollbackData = @{
            Timestamp = (Get-Date -Format "o")
            Action    = "remove"
            SkuId     = $filterSku.SkuId
            SkuName   = $filterSku.SkuPartNumber
            Users     = $successUsers
        }
        $rollbackData | ConvertTo-Json -Depth 3 | Set-Content $rollbackPath -Encoding UTF8
        Write-Host "`nRollback file saved: $rollbackPath" -ForegroundColor Yellow
    }

    Write-Host "`nDone: $($successUsers.Count) succeeded, $failed failed." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 4: Select target license ---
$targetSku = Show-Menu -Title "Select a license to ASSIGN" -Items $skus -FormatItem {
    param($sku)
    $enabled = $sku.PrepaidUnits.Enabled
    $consumed = $sku.ConsumedUnits
    $available = $enabled - $consumed
    "{0,-50} ({1} available of {2})" -f $sku.SkuPartNumber, $available, $enabled
}

$available = $targetSku.PrepaidUnits.Enabled - $targetSku.ConsumedUnits
Write-Host "`nSelected target: $($targetSku.SkuPartNumber) ($available available)" -ForegroundColor Green

if ($available -lt $users.Count) {
    Write-Host "WARNING: Only $available licenses available but $($users.Count) users selected!" -ForegroundColor Red
    if (-not (Confirm-Prompt "Continue anyway?")) {
        Write-Host "Cancelled." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        exit 0
    }
}

# Check for users who already have the target license
$alreadyAssigned = @()
$toAssign = @()
foreach ($u in $users) {
    if ($u.AssignedLicenses.SkuId -contains $targetSku.SkuId) {
        $alreadyAssigned += $u
    } else {
        $toAssign += $u
    }
}

if ($alreadyAssigned.Count -gt 0) {
    Write-Host "`nSkipping $($alreadyAssigned.Count) user(s) who already have '$($targetSku.SkuPartNumber)':" -ForegroundColor Yellow
    foreach ($u in $alreadyAssigned) {
        Write-Host "  $($u.DisplayName) ($($u.UserPrincipalName))" -ForegroundColor DarkGray
    }
}

if ($toAssign.Count -eq 0) {
    Write-Host "`nAll users already have this license. Nothing to do." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 5: Final confirmation and assign ---
Write-Host "`n===== SUMMARY =====" -ForegroundColor Cyan
Write-Host "  Filter license : $($filterSku.SkuPartNumber)"
Write-Host "  Target license : $($targetSku.SkuPartNumber)"
Write-Host "  Users to assign: $($toAssign.Count)"
if ($WhatIf) {
    Write-Host "  Mode           : WhatIf (no changes)" -ForegroundColor DarkGray
}
Write-Host "===================" -ForegroundColor Cyan

if (-not (Confirm-Prompt "`nAssign '$($targetSku.SkuPartNumber)' to $($toAssign.Count) user(s)?")) {
    Write-Host "Cancelled." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

$successUsers = @()
$failed = 0
foreach ($u in $toAssign) {
    try {
        if ($WhatIf) {
            Write-Host "  [WhatIf] Would assign '$($targetSku.SkuPartNumber)' to $($u.UserPrincipalName)" -ForegroundColor DarkGray
        } else {
            Set-MgUserLicense -UserId $u.Id -AddLicenses @(@{SkuId = $targetSku.SkuId }) -RemoveLicenses @()
            Write-Host "  Assigned to $($u.UserPrincipalName)" -ForegroundColor Green
        }
        $successUsers += @{
            UserId            = $u.Id
            DisplayName       = $u.DisplayName
            UserPrincipalName = $u.UserPrincipalName
        }
    } catch {
        Write-Host "  FAILED for $($u.UserPrincipalName): $_" -ForegroundColor Red
        $failed++
    }
}

# Save rollback file
if ($successUsers.Count -gt 0 -and -not $WhatIf) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $rollbackPath = Join-Path (Get-Location) "rollback_$timestamp.json"
    $rollbackData = @{
        Timestamp = (Get-Date -Format "o")
        SkuId     = $targetSku.SkuId
        SkuName   = $targetSku.SkuPartNumber
        FilterSku = $filterSku.SkuPartNumber
        Users     = $successUsers
    }
    $rollbackData | ConvertTo-Json -Depth 3 | Set-Content $rollbackPath -Encoding UTF8
    Write-Host "`nRollback file saved: $rollbackPath" -ForegroundColor Yellow
}

Write-Host "`nDone: $($successUsers.Count) succeeded, $failed failed." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
