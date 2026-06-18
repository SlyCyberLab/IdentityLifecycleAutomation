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
    - SharePoint list: NewHireRequests
    - M365 Business Premium license available in tenant
#>

# -------------------------------------------------------
# Configuration
# -------------------------------------------------------
$TenantId       = "6f5c979b-2a52-4309-bcb8-02e039c8fcc6"
$ClientId       = "a86a6e33-b4bb-4781-bab9-b81d9807959e"
$LicenseSkuId   = "00e1ec7b-e4a3-40d1-9441-b69b597ab222"
$Domain         = "slytech.us"
$ListName       = "NewHireRequests"
$LogPath        = "C:\Logs\HireAutomation\onboarding.log"
$TranscriptPath = "C:\Logs\HireAutomation\Transcript-$((Get-Date).ToString('yyyyMMdd-HHmmss')).log"

# -------------------------------------------------------
# Department to OU and Group mapping
# Uses actual AD groups from slytech.us domain
# -------------------------------------------------------
$DepartmentMap = @{
    "IT" = @{
        OU     = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("IT-Users", "FileShare-IT-RW")
    }
    "Sales" = @{
        OU     = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("Sales-Users", "FileShare-Sales-RW")
    }
    "Security" = @{
        OU     = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("IT-Users", "IT-Admins")
    }
    "HR" = @{
        OU     = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"
        Groups = @("Sales-Users", "HR")
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
    return [string](-join (1..16 | ForEach-Object {
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
# Request validation
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
        try { Test-Department $Fields.Department }
        catch { $errors += $_.Exception.Message }
    }

    if ($errors.Count -gt 0) {
        throw "Validation failed: $($errors -join '; ')"
    }
}

# -------------------------------------------------------
# License availability check
# -------------------------------------------------------
function Test-LicenseAvailability {
    param([string]$SkuId)

    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $SkuId }

    if (-not $sku) {
        throw "License SKU '$SkuId' not found in tenant."
    }

    $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits

    if ($available -lt 1) {
        throw "No available licenses for '$($sku.SkuPartNumber)'. Consumed: $($sku.ConsumedUnits) of $($sku.PrepaidUnits.Enabled)."
    }

    Write-Log "License check passed: $($sku.SkuPartNumber) - $available unit(s) remaining"
}

# -------------------------------------------------------
# Connect to Microsoft Graph using client secret (no WAM)
# -------------------------------------------------------
function Connect-Graph {
    Write-Log "Connecting to Microsoft Graph..."
    $clientSecret = Get-GraphSecret
    $secureSecret = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $credential   = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome -ErrorAction Stop
    Write-Log "Connected to Microsoft Graph."
}

# -------------------------------------------------------
# Generate unique SamAccountName
# Returns guaranteed [string] - never an array
# -------------------------------------------------------
function Get-UniqueSamAccountName {
    param(
        [string]$FirstName,
        [string]$LastName
    )

    $base    = [string]("$($FirstName.Substring(0,1))$LastName").ToLower() -replace '\s', ''
    $sam     = $base
    $counter = 1

    while ($true) {
        $existing = Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue
        if (-not $existing) { break }
        Write-Log "Username '$sam' already exists, trying '$base$counter'" "WARN"
        $sam = [string]"$base$counter"
        $counter++
    }

    return [string]$sam
}

# -------------------------------------------------------
# Create AD user
# Skips creation if user already exists in AD
# Uses $script:ProvisionedUPN to avoid return value type issues
# -------------------------------------------------------
function New-HireADUser {
    param(
        [string]$FirstName,
        [string]$LastName,
        [string]$Department,
        [string]$JobTitle,
        [string]$TempPassword
    )

    $script:ProvisionedUPN = $null
    $displayName           = "$FirstName $LastName"
    $mapping               = $DepartmentMap[$Department]
    $baseSam               = "$($FirstName.Substring(0,1))$LastName".ToLower() -replace '\s', ''

    $existingUser = Get-ADUser -Filter "SamAccountName -eq '$baseSam'" -Properties UserPrincipalName -ErrorAction SilentlyContinue

    if ($existingUser) {
        $script:ProvisionedUPN = "$($existingUser.UserPrincipalName)"
        Write-Log "AD user already exists: $($script:ProvisionedUPN) - skipping AD creation." "WARN"
        return
    }

    $username = Get-UniqueSamAccountName -FirstName $FirstName -LastName $LastName
    $script:ProvisionedUPN = "$username@$Domain"

    Write-Log "Creating AD user: $($script:ProvisionedUPN) in OU: $($mapping.OU)"

    $securePassword = ConvertTo-SecureString $TempPassword -AsPlainText -Force

    New-ADUser `
        -Name                  $displayName `
        -SamAccountName        $username `
        -UserPrincipalName     $script:ProvisionedUPN `
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
        try {
            Add-ADGroupMember -Identity $group -Members $username
            Write-Log "Added $username to group: $group"
        } catch {
            Write-Log "Could not add $username to group $group - $_" "WARN"
        }
    }

    Write-Log "AD user created successfully: $($script:ProvisionedUPN)"
}

# -------------------------------------------------------
# Trigger Entra Connect delta sync
# -------------------------------------------------------
function Invoke-EntraSync {
    Write-Log "Triggering Entra Connect delta sync..."
    Import-Module ADSync -ErrorAction SilentlyContinue
    Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
    Write-Log "Delta sync triggered."
}

# -------------------------------------------------------
# Wait for user to appear in Entra ID
# Polls every 30 seconds for up to 30 minutes
# -------------------------------------------------------
function Wait-ForEntraUser {
    param(
        [string]$UPN,
        [int]$MaxWaitMinutes = 30
    )

    $timeout = (Get-Date).AddMinutes($MaxWaitMinutes)

    do {
        $user = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue
        if ($user) {
            Write-Log "User $UPN confirmed in Entra ID."
            return $user
        }
        $minutesLeft = [math]::Round(($timeout - (Get-Date)).TotalMinutes, 1)
        Write-Log "Waiting for Entra sync... ($minutesLeft min remaining)"
        Start-Sleep -Seconds 30
    } while ((Get-Date) -lt $timeout)

    throw "Timed out waiting for $UPN to appear in Entra ID after $MaxWaitMinutes minutes."
}

# -------------------------------------------------------
# Assign M365 license
# Skips if license already assigned
# -------------------------------------------------------
function Set-UserLicense {
    param([string]$UPN, [string]$UsageLocation = "US")

    Test-LicenseAvailability -SkuId $LicenseSkuId

    $user = Get-MgUser -UserId $UPN -Property "assignedLicenses,usageLocation" -ErrorAction SilentlyContinue
    if ($user.AssignedLicenses | Where-Object { $_.SkuId -eq $LicenseSkuId }) {
        Write-Log "License already assigned to $UPN - skipping."
        return
    }

    # Usage location is required before any license can be assigned
    if ([string]::IsNullOrWhiteSpace($user.UsageLocation)) {
        Write-Log "Setting usage location to '$UsageLocation' for $UPN"
        Update-MgUser -UserId $UPN -UsageLocation $UsageLocation -ErrorAction Stop | Out-Null

        # Wait for usage location to propagate before assigning license
        $locationSet = $false
        for ($i = 0; $i -lt 12; $i++) {
            Start-Sleep -Seconds 5
            $check = Get-MgUser -UserId $UPN -Property "usageLocation" -ErrorAction SilentlyContinue
            if (-not [string]::IsNullOrWhiteSpace($check.UsageLocation)) {
                $locationSet = $true
                Write-Log "Usage location confirmed for $UPN after $(($i + 1) * 5) seconds"
                break
            }
        }
        if (-not $locationSet) {
            throw "Usage location did not propagate for $UPN within 60 seconds."
        }
    }

    Set-MgUserLicense `
        -UserId         $UPN `
        -AddLicenses    @(@{ SkuId = $LicenseSkuId }) `
        -RemoveLicenses @() `
        -ErrorAction Stop | Out-Null

    Write-Log "M365 license assigned to $UPN"
}

# -------------------------------------------------------
# Send credentials email to manager
# -------------------------------------------------------
function Send-CredentialsEmail {
    param(
        [string]$ManagerEmail,
        [string]$FirstName,
        [string]$LastName,
        [string]$UPN,
        [string]$TempPassword,
        [string]$Department,
        [string]$StartDate
    )

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

    Send-MgUserMail -UserId "admin@slytechlab.onmicrosoft.com" -BodyParameter $message | Out-Null
    Write-Log "Credentials email sent to $ManagerEmail for $UPN"
}

# -------------------------------------------------------
# Update SharePoint item status via raw Graph API
# -------------------------------------------------------
function Update-RequestStatus {
    param(
        [string]$SiteId,
        [string]$ListId,
        [string]$ItemId,
        [string]$Status,
        [string]$UPN = ""
    )

    $fields = @{
        Status         = $Status
        ProvisionedUPN = $UPN
        CompletedDate  = (Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ")
    }

    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/lists/$ListId/items/$ItemId/fields" `
        -Body ($fields | ConvertTo-Json) `
        -ContentType "application/json" | Out-Null

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

    $siteId = (Get-MgSite -SiteId "slytechlab.sharepoint.com:/sites/IT").Id
    $list   = Get-MgSiteList -SiteId $siteId | Where-Object { $_.DisplayName -eq $ListName }

    $result  = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$($list.Id)/items?expand=fields"
    # Force array so .Count reflects record count, not property count on a single object
    $pending = @($result.value | Where-Object { $_.fields.Status -eq "Pending" })

    if ($pending.Count -eq 0) {
        Write-Log "No pending requests found. Exiting."
        Stop-Transcript
        exit 0
    }

    Write-Log "Found $($pending.Count) pending request(s)"

    foreach ($item in $pending) {
        $fields = $item.fields

        Write-Log "--- Processing request ID: $($item.id) ---"

        try {
            Test-NewHireRequest -Fields $fields

            $tempPassword = New-TempPassword

            New-HireADUser `
                -FirstName    ([string]$fields.FirstName) `
                -LastName     ([string]$fields.LastName) `
                -Department   ([string]$fields.Department) `
                -JobTitle     ([string]$fields.JobTitle) `
                -TempPassword $tempPassword

            $upn = $script:ProvisionedUPN
            Write-Log "Provisioned UPN: $upn"

            if ([string]::IsNullOrWhiteSpace($upn)) {
                throw "UPN is empty after AD provisioning - cannot continue."
            }

            # Check if user already exists in Entra before syncing
            $entraUser = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
            if ($entraUser) {
                Write-Log "User $upn already exists in Entra ID - skipping sync and wait."
            } else {
                Invoke-EntraSync
                Wait-ForEntraUser -UPN $upn
            }

            Set-UserLicense -UPN $upn

            Send-CredentialsEmail `
                -ManagerEmail ([string]$fields.ManagerEmail) `
                -FirstName    ([string]$fields.FirstName) `
                -LastName     ([string]$fields.LastName) `
                -UPN          $upn `
                -TempPassword $tempPassword `
                -Department   ([string]$fields.Department) `
                -StartDate    ([string]$fields.StartDate)

            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.id `
                -Status  "Completed" `
                -UPN     $upn

            Write-Log "Successfully onboarded: $upn" "SUCCESS"
        }
        catch {
            Write-Log "Failed to process request $($item.id): $_" "ERROR"
            Update-RequestStatus `
                -SiteId  $siteId `
                -ListId  $list.Id `
                -ItemId  $item.id `
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