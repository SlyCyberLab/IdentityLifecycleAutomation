#Requires -Module Pester
<#
.SYNOPSIS
    Pester tests for Remove-OffboardedUser.ps1

.NOTES
    Author:  SlyCyberLab
    Repo:    https://github.com/SlyCyberLab/IdentityLifecycleAutomation
    Run:     Invoke-Pester ./tests/Remove-OffboardedUser.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Mock external dependencies
    function global:Get-ADUser { param($Filter, $Properties) return $null }
    function global:Disable-ADAccount { param($Identity) }
    function global:Remove-ADGroupMember { param($Identity, $Members, $Confirm) }
    function global:Move-ADObject { param($Identity, $TargetPath) }
    function global:Set-ADUser { param($Identity, $Description) }
    function global:Get-MgUser { param($UserId, $Property) return $null }
    function global:Revoke-MgUserSignInSession { param($UserId) }
    function global:Set-MgUserLicense { param($UserId, $AddLicenses, $RemoveLicenses) }
    function global:Send-MgUserMail { param($UserId, $BodyParameter) }
    function global:Connect-MgGraph { }
    function global:Get-Secret { return "MockSecret" }
    function global:Start-ADSyncSyncCycle { }
    function global:Import-Module { }
    function global:Write-Log { param($Message, $Level) Write-Host "[$Level] $Message" }
    function global:Add-Content { }

    $global:DisabledOU = "OU=Disabled,OU=Users,OU=SLYTECH,DC=slytech,DC=us"

    . {
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
                try { [datetime]::Parse($Fields.LastWorkingDay) | Out-Null }
                catch { $errors += "LastWorkingDay '$($Fields.LastWorkingDay)' is not a valid date" }
            }
            if ($errors.Count -gt 0) {
                throw "Validation failed: $($errors -join '; ')"
            }
        }

        function Remove-UserLicense {
            param([string]$UPN)
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
            Set-MgUserLicense -UserId $UPN -AddLicenses @() -RemoveLicenses $licenseSkuIds
        }

        function Revoke-EntraSessions {
            param([string]$UPN)
            $user = Get-MgUser -UserId $UPN -ErrorAction SilentlyContinue
            if (-not $user) {
                Write-Log "User $UPN not found in Entra ID. Sessions cannot be revoked." "WARN"
                return
            }
            Revoke-MgUserSignInSession -UserId $UPN
        }
    }
}

# -------------------------------------------------------
Describe "Test-OffboardingRequest" {
# -------------------------------------------------------

    It "passes for a complete valid request" {
        $fields = @{
            UPN             = "jblake@slytech.us"
            ManagerEmail    = "manager@slytech.us"
            LastWorkingDay  = (Get-Date).AddDays(1).ToString("yyyy-MM-dd")
            DisplayName     = "Jordan Blake"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Not -Throw
    }

    It "throws when UPN is missing" {
        $fields = @{
            UPN            = ""
            ManagerEmail   = "manager@slytech.us"
            LastWorkingDay = "2026-06-30"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*UPN*"
    }

    It "throws when UPN format is invalid" {
        $fields = @{
            UPN            = "notavalidupn"
            ManagerEmail   = "manager@slytech.us"
            LastWorkingDay = "2026-06-30"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Invalid UPN*"
    }

    It "throws when ManagerEmail is missing" {
        $fields = @{
            UPN            = "jblake@slytech.us"
            ManagerEmail   = ""
            LastWorkingDay = "2026-06-30"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*ManagerEmail*"
    }

    It "throws when ManagerEmail is malformed" {
        $fields = @{
            UPN            = "jblake@slytech.us"
            ManagerEmail   = "notanemail"
            LastWorkingDay = "2026-06-30"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Invalid manager email*"
    }

    It "throws when LastWorkingDay is missing" {
        $fields = @{
            UPN            = "jblake@slytech.us"
            ManagerEmail   = "manager@slytech.us"
            LastWorkingDay = ""
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*LastWorkingDay*"
    }

    It "throws when LastWorkingDay is not a valid date" {
        $fields = @{
            UPN            = "jblake@slytech.us"
            ManagerEmail   = "manager@slytech.us"
            LastWorkingDay = "not-a-date"
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*not a valid date*"
    }

    It "throws and reports all errors when multiple fields are invalid" {
        $fields = @{
            UPN            = ""
            ManagerEmail   = ""
            LastWorkingDay = ""
        }
        { Test-OffboardingRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Validation failed*"
    }
}

# -------------------------------------------------------
Describe "Remove-UserLicense" {
# -------------------------------------------------------

    It "skips license removal when user is not found in Entra ID" {
        Mock Get-MgUser { return $null }
        Mock Set-MgUserLicense { }
        { Remove-UserLicense -UPN "unknown@slytech.us" } | Should -Not -Throw
        Should -Invoke Set-MgUserLicense -Times 0
    }

    It "skips license removal when user has no licenses" {
        Mock Get-MgUser {
            return [PSCustomObject]@{
                AssignedLicenses = @()
            }
        }
        Mock Set-MgUserLicense { }
        { Remove-UserLicense -UPN "jblake@slytech.us" } | Should -Not -Throw
        Should -Invoke Set-MgUserLicense -Times 0
    }

    It "calls Set-MgUserLicense when user has licenses" {
        Mock Get-MgUser {
            return [PSCustomObject]@{
                AssignedLicenses = @(
                    [PSCustomObject]@{ SkuId = "test-sku-id" }
                )
            }
        }
        Mock Set-MgUserLicense { }
        Remove-UserLicense -UPN "jblake@slytech.us"
        Should -Invoke Set-MgUserLicense -Times 1
    }
}

# -------------------------------------------------------
Describe "Revoke-EntraSessions" {
# -------------------------------------------------------

    It "skips session revocation when user is not found in Entra ID" {
        Mock Get-MgUser { return $null }
        Mock Revoke-MgUserSignInSession { }
        { Revoke-EntraSessions -UPN "unknown@slytech.us" } | Should -Not -Throw
        Should -Invoke Revoke-MgUserSignInSession -Times 0
    }

    It "calls Revoke-MgUserSignInSession when user exists" {
        Mock Get-MgUser {
            return [PSCustomObject]@{ Id = "test-id"; UserPrincipalName = "jblake@slytech.us" }
        }
        Mock Revoke-MgUserSignInSession { }
        Revoke-EntraSessions -UPN "jblake@slytech.us"
        Should -Invoke Revoke-MgUserSignInSession -Times 1
    }
}
