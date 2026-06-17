<#
.SYNOPSIS
    Automated new hire onboarding script for slytech.us

.DESCRIPTION
    Polls SharePoint list for pending new hire requests.
    Creates AD user, syncs to Entra ID, assigns M365 license,
    sends credentials to manager via email.

.NOTES
    Author:     SlyCyberLab
    Repo:       https://github.com/SlyCyberLab/IdentityLifecycleAutomation
    Schedule:   Runs every 15 minutes via Task Scheduler on DC01

.REQUIREMENTS
    - Domain Admin or Account Operators on DC01
    - Microsoft.Graph PowerShell module
    - Microsoft.PowerShell.SecretManagement module
    - Microsoft.PowerShell.SecretStore module
    - SharePoint list with new hire request schema
    - M365 Business Premium license available in tenant
#>

# -------------------------------------------------------
# Configuration
# -------------------------------------------------------
$TenantId       = "6f5c979b-2a52-4309-bcb8-02e039c8fcc6"
$ClientId       = ""  # App registration client ID
$LicenseSkuId   = ""  # M365 Business Premium SKU ID
$Domain         = "slytech.us"
$ListName       = "NewHireRequests"
$LogPath        = "C:\Logs\HireAutomation\onboarding.log"
$TranscriptPath = "C:\Logs\HireAutomation\Transcript-$((Get-Date).ToString('yyyyMMdd-HHmmss')).log"

# -------------------------------------------------------
# Department to OU and Group mapping
# -------------------------------------------------------
$DepartmentMap = @{
    "IT" = @{
        OU     = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("IT-Staff", "File-IT")
    }
    "Sales" = @{
        OU     = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("Sales-Staff", "File-Sales")
    }
    "Security" = @{
        OU     = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("IT-Staff", "Security-Staff")
    }
    "HR" = @{
        OU     = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("Sales-Staff")
    }
}

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
# Password generation
# -------------------------------------------------------
function New-TempPassword {
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%'
    return (-join (1..16 | ForEach-Object {
        $chars[(Get-Random -Maximum $chars.Length)]
    }))
}

# -------------------------------------------------------
# Department validation
# -------------------------------------------------------
function Test-Department {
    param([string]$Department)
    if (-not $DepartmentMap.ContainsKey($Department)) {
        $validDepts = $DepartmentMap.Keys -join ", "
        throw "Unknown department: '$Department'. Valid options: $validDepts"
    }
}

# -------------------------------------------------------
# Request validation - collects all errors before throwing
# -------------------------------------------------------
function Test-NewHireRequest {
    param($Fields)

    $errors = @()

    $requiredFields = @("FirstName", "LastName", "Department", "JobTitle", "ManagerEmail")

    foreach ($field in $requiredFields) {
        if ([string]::IsNullOrWhiteSpace($Fields.$field)) {
            $errors += "Missing required field: $field"
        }
    }

    if ($Fields.ManagerEmail -and $Fields.ManagerEmail -notmatch "^[^@]+@[^@]+\.[^@]+$") {
        $errors += "Invalid manager email address: '$($Fields.ManagerEmail)'"
    }

    if ($Fields.StartDate) {
        try {
            $date = [datetime]::Parse($Fields.StartDate)
            if ($date -lt (Get-Date).Date) {
                $errors += "StartDate '$($Fields.StartDate)' is in the past"
            }
        } catch {
            $errors += "StartDate '$($Fields.StartDate)' is not a valid date"
        }
    }

    if ($Fields.Department -and -not [string]::IsNullOrWhiteSpace($Fields.Department)) {
        try {
            Test-Department $Fields.Department
        } catch {
            $errors += $_.Exception.Message
        }
    }

    if ($errors.Count -gt 0) {
        $summary = $errors -join "; "
        throw "Validation failed: $summary"
    }
}

# -------------------------------------------------------
# License availability check
# -------------------------------------------------------
function Test-LicenseAvailability {
    param([string]$SkuId)

    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $SkuId }

    if (-not $sku) {
        throw "License SKU '$SkuId' not found in tenant. Run Get-MgSubscribedSku to list available SKUs."
    }

    $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits

    if ($available -lt 1) {
        throw "No available licenses for '$($sku.SkuPartNumber)'. Consumed: $($sku.ConsumedUnits) of $($sku.PrepaidUnits.Enabled)."
    }

    Write-Log "License check passed: $($sku.SkuPartNumber) - $available unit(s) remaining"
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
# Unique username generation
# -------------------------------------------------------
function Get-UniqueSamAccountName {
    param([string]$FirstName, [string]$LastName)

    $base    = "$($FirstName.Substring(0,1).ToLower())$($LastName.ToLower())"
    $sam     = $base
    $counter = 1

    while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
        Write-Log "Username '$sam' already exists, trying '$base$counter'" "WARN"
        $sam = "$base$counter"
        $counter++
    }

    return $sam
}

# -------------------------------------------------------
# Create AD user
# -------------------------------------------------------
function New-HireADUser {
    param($FirstName, $LastName, $Department, $JobTitle, $TempPassword)

    $username    = Get-UniqueSamAccountName -FirstName $FirstName -LastName $LastName
    $displayName = "$FirstName $LastName"
    $upn         = "$username@$Domain"
    $mapping     = $DepartmentMap[$Department]

    Write-Log "Creating AD user: $upn in OU: $($mapping.OU)"

    $securePassword = ConvertTo-SecureString $TempPassword -AsPlainText -Force

    New-ADUser `
        -SamAccountName        $username `
        -UserPrincipalName     $upn `
        -GivenName             $FirstName `
        -Surname               $LastName `
        -DisplayName           $displayName `
        -Title                 $JobTitle `
        -Department            $Department `
        -AccountPassword       $securePassword `
        -ChangePasswordAtLogon $true `
        -Enabled               $true `
        -Path                  $mapping.OU

    foreach ($group in $mapping.Groups) {
        Add-ADGroupMember -Identity $group -Members $username
        Write-Log "Added $username to group: $group"
    }

    Write-Log "AD user created successfully: $upn"
    return $upn
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
# Wait for user to appear in Entra ID
# -------------------------------------------------------
function Wait-ForEntraUser {
    param([string]$UPN, [int]$MaxWaitMinutes = 15)

    $timeout = (Get-Date).AddMinutes($MaxWaitMinutes)

    do {
        $user = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue
        if ($user) {
            Write-Log "User $UPN confirmed in Entra ID."
            return $user
        }
        Write-Log "Waiting for Entra sync... ($([math]::Round(($timeout - (Get-Date)).TotalMinutes, 1)) min remaining)"
        Start-Sleep -Seconds 30
    } while ((Get-Date) -lt $timeout)

    throw "Timed out waiting for $UPN to appear in Entra ID after $MaxWaitMinutes minutes."
}

# -------------------------------------------------------
# Assign M365 license
# -------------------------------------------------------
function Set-UserLicense {
    param([string]$UPN)

    Test-LicenseAvailability -SkuId $LicenseSkuId

    Set-MgUserLicense `
        -UserId         $UPN `
        -AddLicenses    @(@{ SkuId = $LicenseSkuId }) `
        -RemoveLicenses @()

    Write-Log "M365 license assigned to $UPN"
}

# -------------------------------------------------------
# Send credentials email to manager
# -------------------------------------------------------
function Send-CredentialsEmail {
    param($ManagerEmail, $FirstName, $LastName, $UPN, $TempPassword, $Department, $StartDate)

    $subject = "New Hire Account Ready: $FirstName $LastName"
    $body    = @"
Hello,

The account for your new employee has been provisioned.

Employee:      $FirstName $LastName
Username:      $UPN
Department:    $Department
Start Date:    $StartDate

Temporary Password:
$TempPassword

The employee will be required to change their password at first sign-in.
The employee must also complete MFA registration during first login.

Microsoft 365 Portal: https://portal.office.com

Please share these credentials with the new hire securely before their start date.
Do not forward this email externally.

NOTE: Sending credentials via email is acceptable for this lab environment.
In production, use a secure credential delivery method such as an encrypted
messaging platform or privileged access workstation.

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
    Write-Log "Credentials email sent to $ManagerEmail for $UPN"
}

# -------------------------------------------------------
# Update SharePoint item status
# -------------------------------------------------------
function Update-RequestStatus {
    param($SiteId, $ListId, $ItemId, $Status, $UPN)

    $fields = @{
        Status         = $Status
        ProvisionedUPN = $UPN
        CompletedDate  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
    }
    Update-MgSiteListItem -SiteId $SiteId -ListId $ListId -ListItemId $ItemId -Fields $fields
    Write-Log "SharePoint item $ItemId updated to status: $Status"
}

# -------------------------------------------------------
# Main execution
# -------------------------------------------------------
New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force | Out-Null
Start-Transcript -Path $TranscriptPath -Append

Write-Log "===== Identity Lifecycle Automation - New Hire Onboarding ====="
Write-Log "Script started on $env:COMPUTERNAME"

try {
    Connect-Graph

    $siteId  = (Get-MgSite -SiteId "slytechlab.sharepoint.com:/sites/IT").Id
    $list    = Get-MgSiteList -SiteId $siteId | Where-Object { $_.DisplayName -eq $ListName }
    $pending = Get-MgSiteListItem -SiteId $siteId -ListId $list.Id -ExpandProperty Fields |
               Where-Object { $_.Fields.AdditionalData.Status -eq "Pending" }

    if ($pending.Count -eq 0) {
        Write-Log "No pending requests found. Exiting."
        Stop-Transcript
        exit 0
    }

    Write-Log "Found $($pending.Count) pending request(s)"

    foreach ($item in $pending) {
        $fields = $item.Fields.AdditionalData

        Write-Log "--- Processing request ID: $($item.Id) ---"

        try {
            Test-NewHireRequest -Fields $fields

            $tempPassword = New-TempPassword

            $upn = New-HireADUser `
                -FirstName    $fields.FirstName `
                -LastName     $fields.LastName `
                -Department   $fields.Department `
                -JobTitle     $fields.JobTitle `
                -TempPassword $tempPassword

            Invoke-EntraSync

            Wait-ForEntraUser -UPN $upn

            Set-UserLicense -UPN $upn

            Send-CredentialsEmail `
                -ManagerEmail $fields.ManagerEmail `
                -FirstName    $fields.FirstName `
                -LastName     $fields.LastName `
                -UPN          $upn `
                -TempPassword $tempPassword `
                -Department   $fields.Department `
                -StartDate    $fields.StartDate

            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.Id `
                -Status  "Completed" `
                -UPN     $upn

            Write-Log "Successfully onboarded: $upn" "SUCCESS"
        }
        catch {
            Write-Log "Failed to process request $($item.Id): $_" "ERROR"
            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.Id `
                -Status  "Failed" `
                -UPN     ""
        }
    }
}
catch {
    Write-Log "Critical error: $_" "ERROR"
    Stop-Transcript
    exit 1
}

Write-Log "===== Onboarding run complete ====="
Stop-Transcript
