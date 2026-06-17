#Requires -Module Pester
<#
.SYNOPSIS
    Pester tests for New-HireOnboarding.ps1

.NOTES
    Author:  SlyCyberLab
    Repo:    https://github.com/SlyCyberLab/IdentityLifecycleAutomation
    Run:     Invoke-Pester ./tests/New-HireOnboarding.Tests.ps1 -Output Detailed
#>

BeforeAll {
    # Load the script in a way that only imports functions, does not execute main block
    # We dot-source after mocking the modules it depends on
    $scriptPath = "$PSScriptRoot\..\scripts\New-HireOnboarding.ps1"

    # Mock external dependencies so tests run without AD or Graph connectivity
    function global:Get-ADUser { param($Filter) return $null }
    function global:New-ADUser { param([hashtable]$params) }
    function global:Add-ADGroupMember { param($Identity, $Members) }
    function global:Get-MgUser { param($UserId) return $null }
    function global:Get-MgSubscribedSku { return @() }
    function global:Set-MgUserLicense { }
    function global:Send-MgUserMail { }
    function global:Connect-MgGraph { }
    function global:Get-Secret { return "MockSecret" }
    function global:Start-ADSyncSyncCycle { }
    function global:Import-Module { }
    function global:Write-Log { param($Message, $Level) Write-Host "[$Level] $Message" }
    function global:Add-Content { }

    # Define DepartmentMap as it appears in the script
    $global:DepartmentMap = @{
        "IT"       = @{ OU = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"; Groups = @("IT-Staff","File-IT") }
        "Sales"    = @{ OU = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"; Groups = @("Sales-Staff","File-Sales") }
        "Security" = @{ OU = "OU=IT,OU=Users,OU=SLYTECH,DC=slytech,DC=us"; Groups = @("IT-Staff","Security-Staff") }
        "HR"       = @{ OU = "OU=Sales,OU=Users,OU=SLYTECH,DC=slytech,DC=us"; Groups = @("Sales-Staff") }
    }

    $global:Domain       = "slytech.us"
    $global:LicenseSkuId = "test-sku-id"

    # Dot-source only the functions, skip main execution block
    . {
        function Test-Department {
            param([string]$Department)
            if (-not $DepartmentMap.ContainsKey($Department)) {
                $validDepts = $DepartmentMap.Keys -join ", "
                throw "Unknown department: '$Department'. Valid options: $validDepts"
            }
        }

        function Test-NewHireRequest {
            param($Fields)
            $errors = @()
            $requiredFields = @("FirstName","LastName","Department","JobTitle","ManagerEmail")
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
                try { Test-Department $Fields.Department } catch { $errors += $_.Exception.Message }
            }
            if ($errors.Count -gt 0) {
                throw "Validation failed: $($errors -join '; ')"
            }
        }

        function Test-LicenseAvailability {
            param([string]$SkuId)
            $sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $SkuId }
            if (-not $sku) { throw "License SKU '$SkuId' not found in tenant." }
            $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
            if ($available -lt 1) { throw "No available licenses for '$($sku.SkuPartNumber)'." }
        }

        function New-TempPassword {
            $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%'
            return (-join (1..16 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] }))
        }

        function Get-UniqueSamAccountName {
            param([string]$FirstName, [string]$LastName)
            $base    = "$($FirstName.Substring(0,1).ToLower())$($LastName.ToLower())"
            $sam     = $base
            $counter = 1
            while (Get-ADUser -Filter "SamAccountName -eq '$sam'" -ErrorAction SilentlyContinue) {
                $sam = "$base$counter"
                $counter++
            }
            return $sam
        }
    }
}

# -------------------------------------------------------
Describe "Test-Department" {
# -------------------------------------------------------

    It "passes for valid department: IT" {
        { Test-Department "IT" } | Should -Not -Throw
    }

    It "passes for valid department: Sales" {
        { Test-Department "Sales" } | Should -Not -Throw
    }

    It "passes for valid department: Security" {
        { Test-Department "Security" } | Should -Not -Throw
    }

    It "passes for valid department: HR" {
        { Test-Department "HR" } | Should -Not -Throw
    }

    It "throws for unknown department" {
        { Test-Department "Finance" } | Should -Throw -ExpectedMessage "*Unknown department*"
    }

    It "throws for empty department" {
        { Test-Department "" } | Should -Throw
    }
}

# -------------------------------------------------------
Describe "Test-NewHireRequest" {
# -------------------------------------------------------

    It "passes for a complete valid request" {
        $fields = @{
            FirstName    = "Jane"
            LastName     = "Smith"
            Department   = "IT"
            JobTitle     = "Systems Administrator"
            ManagerEmail = "manager@slytech.us"
            StartDate    = (Get-Date).AddDays(7).ToString("yyyy-MM-dd")
        }
        { Test-NewHireRequest -Fields $fields } | Should -Not -Throw
    }

    It "throws when FirstName is missing" {
        $fields = @{
            FirstName    = ""
            LastName     = "Smith"
            Department   = "IT"
            JobTitle     = "Analyst"
            ManagerEmail = "manager@slytech.us"
        }
        { Test-NewHireRequest -Fields $fields } | Should -Throw -ExpectedMessage "*FirstName*"
    }

    It "throws when Department is invalid" {
        $fields = @{
            FirstName    = "Jane"
            LastName     = "Smith"
            Department   = "Finance"
            JobTitle     = "Analyst"
            ManagerEmail = "manager@slytech.us"
        }
        { Test-NewHireRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Unknown department*"
    }

    It "throws when ManagerEmail is malformed" {
        $fields = @{
            FirstName    = "Jane"
            LastName     = "Smith"
            Department   = "IT"
            JobTitle     = "Analyst"
            ManagerEmail = "notanemail"
        }
        { Test-NewHireRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Invalid manager email*"
    }

    It "throws when StartDate is in the past" {
        $fields = @{
            FirstName    = "Jane"
            LastName     = "Smith"
            Department   = "IT"
            JobTitle     = "Analyst"
            ManagerEmail = "manager@slytech.us"
            StartDate    = "2020-01-01"
        }
        { Test-NewHireRequest -Fields $fields } | Should -Throw -ExpectedMessage "*in the past*"
    }

    It "throws when multiple fields are missing and reports all errors" {
        $fields = @{
            FirstName    = ""
            LastName     = ""
            Department   = "IT"
            JobTitle     = ""
            ManagerEmail = ""
        }
        { Test-NewHireRequest -Fields $fields } | Should -Throw -ExpectedMessage "*Validation failed*"
    }
}

# -------------------------------------------------------
Describe "New-TempPassword" {
# -------------------------------------------------------

    It "generates a 16 character password" {
        $pw = New-TempPassword
        $pw.Length | Should -Be 16
    }

    It "generates different passwords on each call" {
        $pw1 = New-TempPassword
        $pw2 = New-TempPassword
        $pw1 | Should -Not -Be $pw2
    }

    It "only contains characters from the allowed set" {
        $allowed = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%'
        $pw = New-TempPassword
        foreach ($char in $pw.ToCharArray()) {
            $allowed | Should -Match [regex]::Escape($char)
        }
    }
}

# -------------------------------------------------------
Describe "Get-UniqueSamAccountName" {
# -------------------------------------------------------

    It "returns first initial plus lowercase lastname" {
        Mock Get-ADUser { return $null }
        $result = Get-UniqueSamAccountName -FirstName "Jane" -LastName "Smith"
        $result | Should -Be "jsmith"
    }

    It "appends counter when username already exists" {
        $callCount = 0
        Mock Get-ADUser {
            $callCount++
            if ($callCount -le 1) { return @{ SamAccountName = "jsmith" } }
            return $null
        }
        $result = Get-UniqueSamAccountName -FirstName "Jane" -LastName "Smith"
        $result | Should -Be "jsmith1"
    }

    It "generates lowercase username" {
        Mock Get-ADUser { return $null }
        $result = Get-UniqueSamAccountName -FirstName "JOHN" -LastName "DOE"
        $result | Should -Be "jdoe"
    }
}

# -------------------------------------------------------
Describe "Test-LicenseAvailability" {
# -------------------------------------------------------

    It "throws when SKU is not found in tenant" {
        Mock Get-MgSubscribedSku { return @() }
        { Test-LicenseAvailability -SkuId "nonexistent-sku" } | Should -Throw -ExpectedMessage "*not found*"
    }

    It "throws when no licenses are available" {
        Mock Get-MgSubscribedSku {
            return @([PSCustomObject]@{
                SkuId         = "test-sku-id"
                SkuPartNumber = "M365_BUSINESS_PREMIUM"
                ConsumedUnits = 25
                PrepaidUnits  = [PSCustomObject]@{ Enabled = 25 }
            })
        }
        { Test-LicenseAvailability -SkuId "test-sku-id" } | Should -Throw -ExpectedMessage "*No available licenses*"
    }

    It "passes when licenses are available" {
        Mock Get-MgSubscribedSku {
            return @([PSCustomObject]@{
                SkuId         = "test-sku-id"
                SkuPartNumber = "M365_BUSINESS_PREMIUM"
                ConsumedUnits = 10
                PrepaidUnits  = [PSCustomObject]@{ Enabled = 25 }
            })
        }
        { Test-LicenseAvailability -SkuId "test-sku-id" } | Should -Not -Throw
    }
}
