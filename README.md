# Identity Lifecycle Automation

Automated new hire onboarding and employee offboarding for the slytech.us hybrid environment. Built on PowerShell, Microsoft Graph, SharePoint, and Power Automate. Runs on DC01 via Task Scheduler every 15 minutes.

Part of the [SlyTech Hybrid Cloud & Security Lab Series](https://blog.slytech.us).

---

## Architecture

```
HR fills Power Apps form
  → SharePoint List (Status: Pending)
    → Task Scheduler on DC01 (every 15 min)
      → PowerShell script polls SharePoint via Graph API
        → Creates/disables AD user
        → Triggers Entra Connect delta sync
        → Assigns/removes M365 license via Graph API
        → Sends email to manager via Graph API
        → Updates SharePoint item (Status: Completed)
```

---

## Repository Structure

```
IdentityLifecycleAutomation/
├── scripts/
│   ├── New-HireOnboarding.ps1       # Onboarding script
│   └── Remove-OffboardedUser.ps1    # Offboarding script
├── tests/
│   ├── New-HireOnboarding.Tests.ps1
│   └── Remove-OffboardedUser.Tests.ps1
├── screenshots/                     # Blog post screenshots
└── README.md
```

---

## Prerequisites

### On DC01 (Windows Server 2025)

```powershell
# Install required PowerShell modules
Install-Module Microsoft.Graph -Force
Install-Module Microsoft.PowerShell.SecretManagement -Force
Install-Module Microsoft.PowerShell.SecretStore -Force
Install-Module Pester -Force
```

### In Azure / Entra ID

- App registration with the following Graph API application permissions:
  - `User.Read.All`
  - `User.ReadWrite.All`
  - `Directory.Read.All`
  - `Directory.ReadWrite.All`
  - `Mail.Send`
  - `Sites.ReadWrite.All`
  - `Organization.Read.All`
- Entra Connect installed and syncing to `slytech.us`
- M365 Business Premium licenses available in tenant

### SharePoint

Two lists required under your SharePoint site:

**NewHireRequests** — columns:

| Column | Type | Required |
|---|---|---|
| FirstName | Single line of text | Yes |
| LastName | Single line of text | Yes |
| Department | Choice (IT, Sales, Security, HR) | Yes |
| JobTitle | Single line of text | Yes |
| ManagerEmail | Single line of text | Yes |
| StartDate | Date | Yes |
| Status | Choice (Pending, Completed, Failed, ValidationFailed) | Yes |
| ProvisionedUPN | Single line of text | No |
| CompletedDate | Date and Time | No |

**OffboardingRequests** — columns:

| Column | Type | Required |
|---|---|---|
| UPN | Single line of text | Yes |
| DisplayName | Single line of text | Yes |
| ManagerEmail | Single line of text | Yes |
| LastWorkingDay | Date | Yes |
| Status | Choice (Pending, Completed, Failed) | Yes |
| CompletedDate | Date and Time | No |

---

## Setup

### 1. Store the app registration secret in SecretStore

```powershell
# Run once on DC01
Register-SecretVault -Name LocalStore -ModuleName Microsoft.PowerShell.SecretStore
Set-Secret -Name "GraphClientSecret" -Secret "your-client-secret-here"
```

### 2. Update configuration in both scripts

Open each script and update the configuration block at the top:

```powershell
$TenantId     = "your-tenant-id"
$ClientId     = "your-app-registration-client-id"
$LicenseSkuId = "your-m365-license-sku-id"  # onboarding script only
```

To find your license SKU ID:

```powershell
Connect-MgGraph -Scopes "Organization.Read.All"
Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId
```

### 3. Create Task Scheduler jobs on DC01

**Onboarding:**

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument "-NonInteractive -File C:\Scripts\New-HireOnboarding.ps1"
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "IdentityLifecycle-Onboarding" `
    -Action $action -Trigger $trigger -RunLevel Highest -Force
```

**Offboarding:**

```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" `
    -Argument "-NonInteractive -File C:\Scripts\Remove-OffboardedUser.ps1"
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 15) -Once -At (Get-Date)
Register-ScheduledTask -TaskName "IdentityLifecycle-Offboarding" `
    -Action $action -Trigger $trigger -RunLevel Highest -Force
```

---

## Department Mapping

Department selection on the Power Apps form drives OU placement and group assignment automatically.

| Department | OU | Groups |
|---|---|---|
| IT | `OU=IT,OU=Users,OU=SLYTECH` | IT-Staff, File-IT |
| Sales | `OU=Sales,OU=Users,OU=SLYTECH` | Sales-Staff, File-Sales |
| Security | `OU=IT,OU=Users,OU=SLYTECH` | IT-Staff, Security-Staff |
| HR | `OU=Sales,OU=Users,OU=SLYTECH` | Sales-Staff |

---

## Offboarding Actions

When an offboarding request is processed the following actions run in order:

1. AD account disabled
2. User removed from all security and distribution groups
3. Account moved to `OU=Disabled,OU=Users,OU=SLYTECH`
4. Description updated with offboard date and audit note
5. Entra Connect delta sync triggered
6. All active Entra ID / M365 sessions revoked immediately
7. All M365 licenses removed
8. Confirmation email sent to manager
9. SharePoint item updated to Completed

---

## Logging

Each script run writes to two log destinations:

- **Structured log:** `C:\Logs\HireAutomation\onboarding.log` / `offboarding.log`
- **Transcript:** `C:\Logs\HireAutomation\Transcript-YYYYMMDD-HHmmss.log`

Log format:

```
2026-06-16 23:15:36 [INFO] Starting Identity Lifecycle Automation - New Hire Onboarding
2026-06-16 23:15:37 [INFO] Found 1 pending request(s)
2026-06-16 23:15:38 [INFO] Creating AD user: jsmith@slytech.us in OU=IT,OU=Users,OU=SLYTECH
2026-06-16 23:15:39 [INFO] Added jsmith to group: IT-Staff
2026-06-16 23:15:40 [SUCCESS] Successfully onboarded: jsmith@slytech.us
```

---

## Running Tests

```powershell
# Install Pester if not already installed
Install-Module Pester -Force

# Run all tests
Invoke-Pester ./tests/ -Output Detailed

# Run onboarding tests only
Invoke-Pester ./tests/New-HireOnboarding.Tests.ps1 -Output Detailed

# Run offboarding tests only
Invoke-Pester ./tests/Remove-OffboardedUser.Tests.ps1 -Output Detailed
```

---

## Security Notes

- Client secret stored in PowerShell SecretStore, never hardcoded
- Temporary passwords expire on first login (`ChangePasswordAtLogon = $true`)
- Credentials sent via email for lab purposes only. In production use an encrypted messaging platform or privileged access workstation for credential delivery
- All actions logged with timestamps for audit trail
- Offboarding revokes sessions immediately, not at next sync cycle

---

## Related Posts

- [IAM Lab](https://blog.slytech.us/blog/iam-lab)
- [PAM Lab](https://blog.slytech.us/blog/pam-lab)
- [Hybrid Identity with Entra Connect](https://blog.slytech.us/blog/entra-connect-hybrid-identity)
- [Intune and Defender for Endpoint](https://blog.slytech.us/blog/intune-defender-endpoint-management)

---

*Part of the [SlyCyberLab](https://github.com/SlyCyberLab) homelab series.*
