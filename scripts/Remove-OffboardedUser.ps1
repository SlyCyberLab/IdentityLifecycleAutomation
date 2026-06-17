<#
.SYNOPSIS
    Automated offboarding script for slytech.us

.DESCRIPTION
    Polls SharePoint list for pending offboarding requests.
    Disables AD user, removes from all groups, moves to Disabled OU,
    removes M365 license, revokes Entra ID sessions,
    sends confirmation email to manager.

.NOTES
    Author:     SlyCyberLab
    Repo:       https://github.com/SlyCyberLab/IdentityLifecycleAutomation
    Schedule:   Runs every 15 minutes via Task Scheduler on DC01

.REQUIREMENTS
    - Domain Admin or Account Operators on DC01
    - Microsoft.Graph PowerShell module
    - Microsoft.PowerShell.SecretManagement module
    - Microsoft.PowerShell.SecretStore module
    - SharePoint list with offboarding request schema
#>

# -------------------------------------------------------
# Configuration
# -------------------------------------------------------
$TenantId       = "6f5c979b-2a52-4309-bcb8-02e039c8fcc6"
$ClientId       = "a86a6e33-b4bb-4781-bab9-b81d9807959e"  # App registration client ID
$Domain         = "slytech.us"
$ListName       = "OffboardingRequests"
$DisabledOU     = "OU=Disabled,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
$LogPath        = "C:\Logs\HireAutomation\offboarding.log"
$TranscriptPath = "C:\Logs\HireAutomation\Transcript-Offboard-$((Get-Date).ToString('yyyyMMdd-HHmmss')).log"

# -------------------------------------------------------
# Logging
# -------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Write-Output $entry
    Add-Content -Path $LogPath -Value $entry
}

# -------------------------------------------------------
# Secret retrieval
# -------------------------------------------------------
function Get-GraphSecret {
    Get-Secret -Name "GraphClientSecret" -AsPlainText
}

# -------------------------------------------------------
# Request validation
# -------------------------------------------------------
function Test-OffboardingRequest {
    param($Fields)

    $errors = @()

    $requiredFields = @("UPN", "ManagerEmail", "LastWorkingDay")

    foreach ($field in $requiredFields) {
        if ([string]::IsNullOrWhiteSpace($Fields.$field)) {
            $errors += "Missing required field: $field"
        }
    }

    if ($Fields.UPN -and $Fields.UPN -notmatch "^[^@]+@[^@]+\.[^@]+$") {
        $errors += "Invalid UPN format: '$($Fields.UPN)'"
    }

    if ($Fields.ManagerEmail -and $Fields.ManagerEmail -notmatch "^[^@]+@[^@]+\.[^@]+$") {
        $errors += "Invalid manager email address: '$($Fields.ManagerEmail)'"
    }

    if ($Fields.LastWorkingDay) {
        try {
            [datetime]::Parse($Fields.LastWorkingDay) | Out-Null
        } catch {
            $errors += "LastWorkingDay '$($Fields.LastWorkingDay)' is not a valid date"
        }
    }

    if ($errors.Count -gt 0) {
        throw "Validation failed: $($errors -join '; ')"
    }
}

# -------------------------------------------------------
# Connect to Microsoft Graph
# -------------------------------------------------------
function Connect-Graph {
    Write-Log "Connecting to Microsoft Graph..."
    $clientSecret = Get-GraphSecret
    Connect-MgGraph `
        -TenantId     $TenantId `
        -ClientId     $ClientId `
        -ClientSecret $clientSecret `
        -NoWelcome
    Write-Log "Connected to Microsoft Graph."
}

# -------------------------------------------------------
# Disable AD account and remove from all groups
# -------------------------------------------------------
function Disable-HireADUser {
    param([string]$UPN)

    $username = $UPN.Split("@")[0]
    $adUser   = Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Properties MemberOf -ErrorAction SilentlyContinue

    if (-not $adUser) {
        throw "AD user not found for UPN: $UPN"
    }

    # Disable the account
    Disable-ADAccount -Identity $adUser.SamAccountName
    Write-Log "AD account disabled: $UPN"

    # Remove from all groups except Domain Users (primary group, cannot be removed)
    foreach ($groupDN in $adUser.MemberOf) {
        try {
            Remove-ADGroupMember -Identity $groupDN -Members $adUser.SamAccountName -Confirm:$false
            Write-Log "Removed $username from group: $groupDN"
        } catch {
            Write-Log "Could not remove $username from group $groupDN - $_" "WARN"
        }
    }

    # Move to Disabled OU
    Move-ADObject -Identity $adUser.DistinguishedName -TargetPath $DisabledOU
    Write-Log "Moved $UPN to Disabled OU: $DisabledOU"

    # Append OFFBOARDED to description for audit trail
    Set-ADUser -Identity $adUser.SamAccountName `
        -Description "OFFBOARDED: $(Get-Date -Format 'yyyy-MM-dd') - Account disabled by Identity Lifecycle Automation"

    Write-Log "AD offboarding complete for: $UPN"
}

# -------------------------------------------------------
# Revoke all Entra ID sessions
# -------------------------------------------------------
function Revoke-EntraSessions {
    param([string]$UPN)

    Write-Log "Revoking Entra ID sessions for $UPN..."

    $user = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue

    if (-not $user) {
        Write-Log "User $UPN not found in Entra ID. Sessions cannot be revoked." "WARN"
        return
    }

    Revoke-MgUserSignInSession -UserId $UPN
    Write-Log "All Entra ID sessions revoked for: $UPN"
}

# -------------------------------------------------------
# Remove M365 license
# -------------------------------------------------------
function Remove-UserLicense {
    param([string]$UPN)

    Write-Log "Removing M365 licenses from $UPN..."

    $user = Get-MgUser -UserId $UPN -Property "assignedLicenses" -ErrorAction SilentlyContinue

    if (-not $user) {
        Write-Log "User $UPN not found in Entra ID. License removal skipped." "WARN"
        return
    }

    if ($user.AssignedLicenses.Count -eq 0) {
        Write-Log "No licenses assigned to $UPN. Nothing to remove."
        return
    }

    $licenseSkuIds = $user.AssignedLicenses | Select-Object -ExpandProperty SkuId

    Set-MgUserLicense `
        -UserId         $UPN `
        -AddLicenses    @() `
        -RemoveLicenses $licenseSkuIds

    Write-Log "Removed $($licenseSkuIds.Count) license(s) from $UPN"
}

# -------------------------------------------------------
# Trigger Entra Connect delta sync
# -------------------------------------------------------
function Invoke-EntraSync {
    Write-Log "Triggering Entra Connect delta sync..."
    Import-Module ADSync -ErrorAction SilentlyContinue
    Start-ADSyncSyncCycle -PolicyType Delta
    Write-Log "Delta sync triggered."
}

# -------------------------------------------------------
# Send offboarding confirmation to manager
# -------------------------------------------------------
function Send-OffboardingConfirmation {
    param($ManagerEmail, $DisplayName, $UPN, $LastWorkingDay)

    $subject = "Offboarding Complete: $DisplayName"
    $body    = @"
Hello,

The offboarding process for the following employee has been completed.

Employee:        $DisplayName
Username:        $UPN
Last Working Day: $LastWorkingDay
Offboarded On:   $(Get-Date -Format 'yyyy-MM-dd HH:mm')

Actions completed:
- AD account disabled
- Removed from all security and distribution groups
- Account moved to Disabled OU
- All Microsoft 365 sessions revoked
- Microsoft 365 licenses reclaimed
- Entra ID account will reflect changes within 30 minutes

If you need access to this user's mailbox or files, please submit a separate access request.

This message was generated automatically by the SlyTech Identity Lifecycle Automation system.
"@

    $message = @{
        Message = @{
            Subject      = $subject
            Body         = @{ ContentType = "Text"; Content = $body }
            ToRecipients = @(@{ EmailAddress = @{ Address = $ManagerEmail } })
        }
    }

    Send-MgUserMail -UserId "admin@slytechlab.onmicrosoft.com" -BodyParameter $message
    Write-Log "Offboarding confirmation sent to $ManagerEmail for $UPN"
}

# -------------------------------------------------------
# Update SharePoint item status
# -------------------------------------------------------
function Update-RequestStatus {
    param($SiteId, $ListId, $ItemId, $Status)

    $fields = @{
        Status        = $Status
        CompletedDate = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
    }
    Update-MgSiteListItem -SiteId $SiteId -ListId $ListId -ListItemId $ItemId -Fields $fields
    Write-Log "SharePoint item $ItemId updated to status: $Status"
}

# -------------------------------------------------------
# Main execution
# -------------------------------------------------------
New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force | Out-Null
Start-Transcript -Path $TranscriptPath -Append

Write-Log "===== Identity Lifecycle Automation - Offboarding ====="
Write-Log "Script started on $env:COMPUTERNAME"

try {
    Connect-Graph

    $siteId  = (Get-MgSite -SiteId "slytechlab.sharepoint.com:/sites/IT").Id
    $list    = Get-MgSiteList -SiteId $siteId | Where-Object { $_.DisplayName -eq $ListName }
    $pending = Get-MgSiteListItem -SiteId $siteId -ListId $list.Id -ExpandProperty Fields |
               Where-Object { $_.Fields.AdditionalData.Status -eq "Pending" }

    if ($pending.Count -eq 0) {
        Write-Log "No pending offboarding requests found. Exiting."
        Stop-Transcript
        exit 0
    }

    Write-Log "Found $($pending.Count) pending offboarding request(s)"

    foreach ($item in $pending) {
        $fields = $item.Fields.AdditionalData

        Write-Log "--- Processing offboarding request ID: $($item.Id) ---"

        try {
            Test-OffboardingRequest -Fields $fields

            $upn            = $fields.UPN
            $managerEmail   = $fields.ManagerEmail
            $lastWorkingDay = $fields.LastWorkingDay
            $displayName    = $fields.DisplayName

            # Step 1: Disable AD account, remove groups, move to Disabled OU
            Disable-HireADUser -UPN $upn

            # Step 2: Sync changes to Entra ID
            Invoke-EntraSync

            # Step 3: Revoke all active sessions immediately
            Revoke-EntraSessions -UPN $upn

            # Step 4: Remove M365 licenses
            Remove-UserLicense -UPN $upn

            # Step 5: Notify manager
            Send-OffboardingConfirmation `
                -ManagerEmail   $managerEmail `
                -DisplayName    $displayName `
                -UPN            $upn `
                -LastWorkingDay $lastWorkingDay

            # Step 6: Update SharePoint
            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.Id `
                -Status  "Completed"

            Write-Log "Successfully offboarded: $upn" "SUCCESS"
        }
        catch {
            Write-Log "Failed to process offboarding request $($item.Id): $_" "ERROR"
            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.Id `
                -Status  "Failed"
        }
    }
}
catch {
    Write-Log "Critical error: $_" "ERROR"
    Stop-Transcript
    exit 1
}

Write-Log "===== Offboarding run complete ====="
Stop-Transcript
