# Danmade Patch Agent Admin Guide

## Purpose

Danmade Patch Agent runs unattended `winget` upgrades on domain-managed Windows endpoints. It is designed for Group Policy scheduled task deployment, Domain CA script signing, and Wazuh collection through Windows Event Log and JSONL file monitoring.

The agent is separate from the interactive workstation updater. It does not use prompts, grids, UAC relaunches, or non-silent installer retries.

Machine mode skips packages that are clearly installed per-user. User mode skips packages that are not clearly per-user. Packages whose scope cannot be determined are left to machine mode so the user task does not repeatedly attempt machine-scope upgrades without elevation.

## Files

- `danmade-patch-agent.ps1`: signed endpoint agent.
- `danmade-patch-agent.policy.json`: optional policy file copied to endpoints.
- `danmade-patch-agent.policy.sample.json`: sample policy to copy and edit.

Recommended SYSVOL source:

```text
\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\
```

Recommended endpoint target:

```text
C:\ProgramData\DanmadePatchAgent\
```

Set endpoint ACLs so `SYSTEM` and `Administrators` have full control. Keep the signed script and policy read/execute for `Users`; grant `Users` modify only on the `Logs` and `Events` subfolders if the user-context scheduled task is enabled.

## Policy

The agent uses policy in this order:

1. `-PolicyPath` command-line argument.
2. `C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.policy.json`.
3. Built-in defaults.

Default behavior:

- Agent enabled.
- Include unknown versions.
- No allow list.
- No block list.
- Maximum two retries after the first attempt.
- No maintenance window restriction.
- Reboots are reported only.
- Event Log and JSONL reporting enabled.
- Winget source reset/update repair enabled.
- Logs retained for 30 days.

Copy `windows-update-scripts/danmade-patch-agent.policy.sample.json` to:

```text
\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\danmade-patch-agent.policy.json
```

Then edit package allow/block lists and maintenance window values for the OU being targeted.

## Signing With The Domain CA

Use an internal AD CS code-signing certificate. The signing certificate must include the Code Signing enhanced key usage.

1. On the CA or admin workstation, open `certtmpl.msc`.
2. Duplicate the built-in `Code Signing` template.
3. Name it `Danmade PowerShell Code Signing`.
4. Confirm the template includes Code Signing EKU.
5. On Security, grant `Read` and `Enroll` to the AD group allowed to sign the agent.
6. On the CA, open `certsrv.msc`, right-click Certificate Templates, choose New, and issue the new template.
7. On the signing workstation, request the certificate through `certmgr.msc` or `certreq`.
8. Sign the script:

```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
  Where-Object { $_.Subject -like '*Danmade*' -or $_.EnhancedKeyUsageList.FriendlyName -contains 'Code Signing' } |
  Sort-Object NotAfter -Descending |
  Select-Object -First 1

Set-AuthenticodeSignature `
  -FilePath .\danmade-patch-agent.ps1 `
  -Certificate $cert `
  -TimestampServer 'http://timestamp.digicert.com'
```

If external timestamping is not permitted, omit `-TimestampServer`. Without timestamping, the script signature may stop validating after the signing certificate expires.

Verify:

```powershell
Get-AuthenticodeSignature .\danmade-patch-agent.ps1 | Format-List Status,SignerCertificate,TimeStamperCertificate
```

Expected status is `Valid`.

## Trust And Execution Policy GPO

Use a computer GPO linked to the workstation/server OU.

1. Open `gpmc.msc`.
2. Create or edit `Danmade Patch Agent - Trust`.
3. Import the Domain Root CA certificate into:

```text
Computer Configuration
  Policies
    Windows Settings
      Security Settings
        Public Key Policies
          Trusted Root Certification Authorities
```

4. Import the code-signing publisher certificate into:

```text
Computer Configuration
  Policies
    Windows Settings
      Security Settings
        Public Key Policies
          Trusted Publishers
```

5. Set PowerShell execution policy:

```text
Computer Configuration
  Policies
    Administrative Templates
      Windows Components
        Windows PowerShell
          Turn on Script Execution
```

Set it to `Allow only signed scripts`.

This maps to the intended `AllSigned` trust model for managed endpoints.

## File Deployment GPO

Create or edit `Danmade Patch Agent - Files`.

Deploy the signed script:

```text
Computer Configuration
  Preferences
    Windows Settings
      Files
```

Create an `Update` item:

- Source: `\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\danmade-patch-agent.ps1`
- Destination: `C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1`

Create another `Update` item:

- Source: `\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\danmade-patch-agent.policy.json`
- Destination: `C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.policy.json`

Also create the folder:

```text
Computer Configuration
  Preferences
    Windows Settings
      Folders
```

- Action: `Update`
- Path: `C:\ProgramData\DanmadePatchAgent`

Create these subfolders as well:

```text
C:\ProgramData\DanmadePatchAgent\Logs
C:\ProgramData\DanmadePatchAgent\Events
```

Use Group Policy Preferences security settings or a startup ACL script so:

- `C:\ProgramData\DanmadePatchAgent`: `SYSTEM` and `Administrators` full control; `Users` read and execute.
- `C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1`: `Users` read and execute only.
- `C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.policy.json`: `Users` read only.
- `C:\ProgramData\DanmadePatchAgent\Logs` and `Events`: `Users` modify if using the user-context task.

## Machine Scheduled Task GPO

Create or edit `Danmade Patch Agent - Machine Task`.

Path:

```text
Computer Configuration
  Preferences
    Control Panel Settings
      Scheduled Tasks
```

Create a new scheduled task:

- Name: `Danmade Patch Agent - Machine`
- Run whether user is logged on or not.
- Run with highest privileges.
- User: `NT AUTHORITY\SYSTEM`
- Hidden: enabled.
- Configure for: Windows 10 or later.
- Trigger: daily at `2:00 AM`.
- Random delay: `2 hours`.
- Stop task if it runs longer than `4 hours`.

Action:

```text
Program/script:
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

Arguments:
-NoProfile -ExecutionPolicy AllSigned -File "C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1" -Mode Machine
```

## User Scheduled Task GPO

Use this task to cover per-user winget and Store packages.

Path:

```text
User Configuration
  Preferences
    Control Panel Settings
      Scheduled Tasks
```

Create a new scheduled task:

- Name: `Danmade Patch Agent - User`
- Run only when user is logged on.
- Hidden: enabled.
- Configure for: Windows 10 or later.
- Trigger: at logon.
- Optional trigger: daily, run only if idle.
- Stop task if it runs longer than `2 hours`.

Action:

```text
Program/script:
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

Arguments:
-NoProfile -ExecutionPolicy AllSigned -File "C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1" -Mode User
```

## Wazuh Collection

The agent writes both Windows Event Log records and JSONL events.

Event Log:

- Log: `Application`
- Source: `DanmadePatchAgent`
- Event IDs:
  - `5000`: run started
  - `5001`: run completed
  - `5100`: package succeeded
  - `5101`: package skipped
  - `5200`: recovery attempted
  - `5300`: restart required
  - `5400`: final package failure
  - `5500`: winget repair result
  - `5600`: preflight or agent health failure

Wazuh agent config for Event Log:

```xml
<localfile>
  <location>Application</location>
  <log_format>eventchannel</log_format>
  <query>Event[System[Provider[@Name='DanmadePatchAgent']]]</query>
</localfile>
```

JSONL path:

```text
C:\ProgramData\DanmadePatchAgent\Events\patch-agent.jsonl
```

Wazuh agent config for JSONL:

```xml
<localfile>
  <location>C:\ProgramData\DanmadePatchAgent\Events\patch-agent.jsonl</location>
  <log_format>json</log_format>
  <label key="@source">danmade-patch-agent</label>
</localfile>
```

Suggested Wazuh custom rules:

```xml
<group name="danmade,patching,winget,">
  <rule id="110500" level="3">
    <field name="@source">danmade-patch-agent</field>
    <description>Danmade Patch Agent event</description>
  </rule>

  <rule id="110530" level="7">
    <if_sid>110500</if_sid>
    <field name="status">RestartRequired</field>
    <description>Danmade Patch Agent package update requires restart</description>
  </rule>

  <rule id="110540" level="10">
    <if_sid>110500</if_sid>
    <field name="status">Failed</field>
    <description>Danmade Patch Agent package update failed after retries</description>
  </rule>

  <rule id="110560" level="10">
    <if_sid>110500</if_sid>
    <field name="status">WingetNotFound|WingetInfoFailed|PreflightFailed|UnhandledAgentError</field>
    <description>Danmade Patch Agent health or preflight failure</description>
  </rule>
</group>
```

## Validation

On a pilot workstation:

```powershell
Get-AuthenticodeSignature C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1

powershell.exe -NoProfile -ExecutionPolicy AllSigned `
  -File C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1 `
  -Mode Machine `
  -WhatIf
```

Confirm:

- The scheduled tasks appear after `gpupdate /force`.
- `Get-AuthenticodeSignature` returns `Valid`.
- `Application` contains `DanmadePatchAgent` events.
- `C:\ProgramData\DanmadePatchAgent\Events\patch-agent.jsonl` receives one JSON object per line during real runs.
- Wazuh receives either the Event Log records, JSONL records, or both.
- Machine-mode package events do not include clearly per-user installs; user-mode package events include only clearly per-user installs.

## Operational Notes

- The agent does not silently install or repair App Installer in v1. If `winget` is missing or broken, it reports a preflight failure for remediation.
- Installer exit code `3010` is treated as restart-required, not a final package failure.
- A failed package is retried only up to `maxRetries`; successful packages are never retried.
- `rebootPolicy` is `ReportOnly` in v1. Use Wazuh or a separate endpoint management workflow to coordinate reboots.
