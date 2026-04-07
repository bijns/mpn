<#
.SYNOPSIS
    Find enabled M365 users who have one license but lack another.

.DESCRIPTION
    Connects to Microsoft Graph, lets you pick a "must have" license and a
    "must NOT have" license, then lists all non-blocked users that match.
    Results can be exported to CSV.

.PARAMETER WhatIf
    Show selected criteria without querying users (useful to verify selections).

.EXAMPLE
    .\Get-M365UsersWithoutLicense.ps1
    .\Get-M365UsersWithoutLicense.ps1 -WhatIf
#>
[CmdletBinding()]
param(
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
Connect-MgGraph -Scopes "Directory.Read.All", "Organization.Read.All" -NoWelcome

# --- Fetch tenant licenses ---
Write-Host "Fetching tenant licenses..." -ForegroundColor Yellow
$skus = Get-MgSubscribedSku | Sort-Object SkuPartNumber

if ($skus.Count -eq 0) {
    Write-Host "No licenses found in tenant." -ForegroundColor Red
    exit 1
}

# --- Step 1: Select the license users MUST have ---
$mustHaveSku = Show-Menu -Title "Select the license users MUST have" -Items $skus -FormatItem {
    param($sku)
    $enabled = $sku.PrepaidUnits.Enabled
    $consumed = $sku.ConsumedUnits
    "{0,-50} ({1}/{2} assigned)" -f $sku.SkuPartNumber, $consumed, $enabled
}

Write-Host "`nMust have: $($mustHaveSku.SkuPartNumber)" -ForegroundColor Green

# --- Step 2: Select the license users must NOT have ---
$mustNotHaveSku = Show-Menu -Title "Select the license users must NOT have" -Items $skus -FormatItem {
    param($sku)
    $enabled = $sku.PrepaidUnits.Enabled
    $consumed = $sku.ConsumedUnits
    "{0,-50} ({1}/{2} assigned)" -f $sku.SkuPartNumber, $consumed, $enabled
}

Write-Host "Must NOT have: $($mustNotHaveSku.SkuPartNumber)" -ForegroundColor Red

if ($mustHaveSku.SkuId -eq $mustNotHaveSku.SkuId) {
    Write-Host "`nYou selected the same license for both. No users can match." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- WhatIf: just show criteria ---
if ($WhatIf) {
    Write-Host "`n[WhatIf] Would search for users where:" -ForegroundColor DarkGray
    Write-Host "  HAS        : $($mustHaveSku.SkuPartNumber)" -ForegroundColor DarkGray
    Write-Host "  DOES NOT have: $($mustNotHaveSku.SkuPartNumber)" -ForegroundColor DarkGray
    Write-Host "  Sign-in    : Enabled (not blocked)" -ForegroundColor DarkGray
    Disconnect-MgGraph | Out-Null
    exit 0
}

# --- Step 3: Fetch users with the "must have" license ---
Write-Host "`nFetching users with '$($mustHaveSku.SkuPartNumber)'..." -ForegroundColor Yellow

$allUsers = Get-MgUser -All `
    -Filter "assignedLicenses/any(l:l/skuId eq $($mustHaveSku.SkuId))" `
    -Property Id, DisplayName, UserPrincipalName, AssignedLicenses, AccountEnabled `
    -ConsistencyLevel eventual -CountVariable userCount

# --- Step 4: Filter out blocked users and users who have the excluded license ---
$results = $allUsers | Where-Object {
    $_.AccountEnabled -eq $true -and
    ($_.AssignedLicenses.SkuId -notcontains $mustNotHaveSku.SkuId)
}

$blocked = ($allUsers | Where-Object { $_.AccountEnabled -ne $true }).Count
$hasExcluded = ($allUsers | Where-Object {
    $_.AccountEnabled -eq $true -and
    ($_.AssignedLicenses.SkuId -contains $mustNotHaveSku.SkuId)
}).Count

# --- Step 5: Display results ---
Write-Host "`n===== FILTER SUMMARY =====" -ForegroundColor Cyan
Write-Host "  HAS            : $($mustHaveSku.SkuPartNumber) ($($allUsers.Count) total)"
Write-Host "  DOES NOT have  : $($mustNotHaveSku.SkuPartNumber) (excluded $hasExcluded)"
Write-Host "  Blocked users  : excluded $blocked"
Write-Host "  Matching users : $($results.Count)" -ForegroundColor Green
Write-Host "==========================" -ForegroundColor Cyan

if ($results.Count -eq 0) {
    Write-Host "`nNo users match the criteria." -ForegroundColor Yellow
    Disconnect-MgGraph | Out-Null
    exit 0
}

$sorted = $results | Sort-Object DisplayName

Write-Host ("`n{0,-35} {1}" -f "Display Name", "User Principal Name") -ForegroundColor Cyan
Write-Host ("-" * 70) -ForegroundColor Cyan
foreach ($u in $sorted) {
    Write-Host ("{0,-35} {1}" -f $u.DisplayName, $u.UserPrincipalName)
}

# --- Step 6: Optional CSV export ---
if (Confirm-Prompt "`nExport results to CSV?") {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path (Get-Location) "users_has_${timestamp}.csv"
    $sorted | Select-Object DisplayName, UserPrincipalName, Id | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported to: $csvPath" -ForegroundColor Green
}

Write-Host "`nDone." -ForegroundColor Cyan
Disconnect-MgGraph | Out-Null
