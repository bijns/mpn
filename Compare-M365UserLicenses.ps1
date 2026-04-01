<#
.SYNOPSIS
    Compare licenses between two M365 users and copy licenses from one to the other.

.DESCRIPTION
    Connects to Microsoft Graph, lets you pick a license to filter users by,
    then select two users to compare. Shows a side-by-side license comparison
    and lets you copy all licenses from one user to the other.

.PARAMETER Rollback
    Path to a rollback JSON file to reverse a previous copy operation.

.PARAMETER WhatIf
    Show what would happen without making changes.

.EXAMPLE
    .\Compare-M365UserLicenses.ps1
    .\Compare-M365UserLicenses.ps1 -Rollback .\rollback_copy_20260401_120000.json
    .\Compare-M365UserLicenses.ps1 -WhatIf
#>
[CmdletBinding()]
param(
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

# --- Build SKU lookup (used by both rollback and normal mode) ---
Write-Host "Fetching tenant licenses..." -ForegroundColor Yellow
$skus = Get-MgSubscribedSku | Sort-Object SkuPartNumber

if ($skus.Count -eq 0) {
    Write-Host "No licenses found in tenant." -ForegroundColor Red
    exit 1
}

$skuLookup = @{}
foreach ($sku in $skus) {
    $skuLookup[$sku.SkuId] = $sku.SkuPartNumber
}

# --- Rollback mode ---
if ($Rollback) {
    if (-not (Test-Path $Rollback)) {
        Write-Host "Rollback file not found: $Rollback" -ForegroundColor Red
        exit 1
    }

    $rollbackData = Get-Content $Rollback -Raw | ConvertFrom-Json

    Write-Host "`nRollback: Remove copied licenses from '$($rollbackData.TargetUserName)'" -ForegroundColor Cyan
    Write-Host ("-" * 60) -ForegroundColor Cyan
    Write-Host "  Source was : $($rollbackData.SourceUserName)"
    Write-Host "  Target was : $($rollbackData.TargetUserName)"
    Write-Host "  Licenses to remove:" -ForegroundColor Yellow
    foreach ($lic in $rollbackData.LicensesAdded) {
        $name = $skuLookup[$lic.SkuId]
        if (-not $name) { $name = $lic.SkuName }
        Write-Host "    - $name"
    }

    if (-not (Confirm-Prompt "`nProceed with rollback?")) {
        Write-Host "Cancelled." -ForegroundColor Yellow
        exit 0
    }

    $success = 0
    $failed = 0
    foreach ($lic in $rollbackData.LicensesAdded) {
        $name = $skuLookup[$lic.SkuId]
        if (-not $name) { $name = $lic.SkuName }
        try {
            if ($WhatIf) {
                Write-Host "  [WhatIf] Would remove '$name' from $($rollbackData.TargetUserName)" -ForegroundColor DarkGray
            } else {
                Set-MgUserLicense -UserId $rollbackData.TargetUserId -AddLicenses @() -RemoveLicenses @($lic.SkuId)
                Write-Host "  Removed '$name'" -ForegroundColor Green
            }
            $success++
        } catch {
            Write-Host "  FAILED to remove '$name': $_" -ForegroundColor Red
            $failed++
        }
    }

    Write-Host "`nRollback complete: $success succeeded, $failed failed." -ForegroundColor Cyan
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Normal mode: Step 1 - Select filter license ---
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

if ($users.Count -lt 2) {
    Write-Host "Need at least 2 users with this license to compare. Found $($users.Count)." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 2: Select two users to compare ---
$userItems = $users | Sort-Object DisplayName

$userA = Show-Menu -Title "Select FIRST user to compare" -Items $userItems -FormatItem {
    param($u)
    "{0,-35} {1}" -f $u.DisplayName, $u.UserPrincipalName
}

$userB = Show-Menu -Title "Select SECOND user to compare" -Items $userItems -FormatItem {
    param($u)
    "{0,-35} {1}" -f $u.DisplayName, $u.UserPrincipalName
}

if ($userA.Id -eq $userB.Id) {
    Write-Host "You selected the same user twice. Nothing to compare." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 3: Compare licenses side by side ---
$licensesA = @{}
foreach ($lic in $userA.AssignedLicenses) {
    $name = $skuLookup[$lic.SkuId]
    if ($name) { $licensesA[$lic.SkuId] = $name }
}

$licensesB = @{}
foreach ($lic in $userB.AssignedLicenses) {
    $name = $skuLookup[$lic.SkuId]
    if ($name) { $licensesB[$lic.SkuId] = $name }
}

$allSkuIds = @($licensesA.Keys) + @($licensesB.Keys) | Sort-Object -Unique

$nameA = $userA.DisplayName
$nameB = $userB.DisplayName

Write-Host "`n===== LICENSE COMPARISON =====" -ForegroundColor Cyan
Write-Host ("{0,-50} {1,-15} {2,-15}" -f "License", $nameA, $nameB) -ForegroundColor Cyan
Write-Host ("-" * 80) -ForegroundColor Cyan

foreach ($skuId in $allSkuIds) {
    $name = if ($licensesA.ContainsKey($skuId)) { $licensesA[$skuId] } else { $licensesB[$skuId] }
    $hasA = if ($licensesA.ContainsKey($skuId)) { "YES" } else { "-" }
    $hasB = if ($licensesB.ContainsKey($skuId)) { "YES" } else { "-" }

    $color = "White"
    if ($hasA -eq "YES" -and $hasB -eq "YES") { $color = "Green" }
    elseif ($hasA -eq "YES") { $color = "Yellow" }
    else { $color = "Magenta" }

    Write-Host ("{0,-50} {1,-15} {2,-15}" -f $name, $hasA, $hasB) -ForegroundColor $color
}

Write-Host ""
Write-Host "  Green   = both have it" -ForegroundColor Green
Write-Host "  Yellow  = only $nameA" -ForegroundColor Yellow
Write-Host "  Magenta = only $nameB" -ForegroundColor Magenta

# --- Step 4: Choose copy direction ---
$onlyA = $allSkuIds | Where-Object { $licensesA.ContainsKey($_) -and -not $licensesB.ContainsKey($_) }
$onlyB = $allSkuIds | Where-Object { $licensesB.ContainsKey($_) -and -not $licensesA.ContainsKey($_) }

if ($onlyA.Count -eq 0 -and $onlyB.Count -eq 0) {
    Write-Host "`nBoth users have identical licenses. Nothing to copy." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

$options = @()
if ($onlyA.Count -gt 0) {
    $options += [PSCustomObject]@{
        Label      = "Copy $nameA -> $nameB ($($onlyA.Count) license(s))"
        SourceName = $nameA
        SourceId   = $userA.Id
        TargetName = $nameB
        TargetId   = $userB.Id
        SkuIds     = $onlyA
    }
}
if ($onlyB.Count -gt 0) {
    $options += [PSCustomObject]@{
        Label      = "Copy $nameB -> $nameA ($($onlyB.Count) license(s))"
        SourceName = $nameB
        SourceId   = $userB.Id
        TargetName = $nameA
        TargetId   = $userA.Id
        SkuIds     = $onlyB
    }
}

$choice = Show-Menu -Title "Choose copy direction" -Items $options -FormatItem {
    param($opt) $opt.Label
}

# --- Step 5: Show what will be copied and confirm ---
Write-Host "`n===== COPY SUMMARY =====" -ForegroundColor Cyan
Write-Host "  Source: $($choice.SourceName)" -ForegroundColor Green
Write-Host "  Target: $($choice.TargetName)" -ForegroundColor Yellow
Write-Host "  Licenses to add:" -ForegroundColor Cyan
foreach ($skuId in $choice.SkuIds) {
    $name = $skuLookup[$skuId]
    Write-Host "    + $name"
}
if ($WhatIf) {
    Write-Host "  Mode: WhatIf (no changes)" -ForegroundColor DarkGray
}
Write-Host "========================" -ForegroundColor Cyan

if (-not (Confirm-Prompt "`nProceed?")) {
    Write-Host "Cancelled." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 6: Assign licenses ---
$successLicenses = @()
$failed = 0
foreach ($skuId in $choice.SkuIds) {
    $name = $skuLookup[$skuId]
    try {
        if ($WhatIf) {
            Write-Host "  [WhatIf] Would assign '$name' to $($choice.TargetName)" -ForegroundColor DarkGray
        } else {
            Set-MgUserLicense -UserId $choice.TargetId -AddLicenses @(@{SkuId = $skuId }) -RemoveLicenses @()
            Write-Host "  Assigned '$name' to $($choice.TargetName)" -ForegroundColor Green
        }
        $successLicenses += @{
            SkuId   = $skuId
            SkuName = $name
        }
    } catch {
        Write-Host "  FAILED to assign '$name': $_" -ForegroundColor Red
        $failed++
    }
}

# --- Save rollback file ---
if ($successLicenses.Count -gt 0 -and -not $WhatIf) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $rollbackPath = Join-Path (Get-Location) "rollback_copy_$timestamp.json"
    $rollbackData = @{
        Timestamp      = (Get-Date -Format "o")
        Action         = "copy"
        SourceUserId   = $choice.SourceId
        SourceUserName = $choice.SourceName
        TargetUserId   = $choice.TargetId
        TargetUserName = $choice.TargetName
        LicensesAdded  = $successLicenses
    }
    $rollbackData | ConvertTo-Json -Depth 3 | Set-Content $rollbackPath -Encoding UTF8
    Write-Host "`nRollback file saved: $rollbackPath" -ForegroundColor Yellow
}

Write-Host "`nDone: $($successLicenses.Count) license(s) copied, $failed failed." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
