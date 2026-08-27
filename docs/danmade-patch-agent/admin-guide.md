# Danmade Patch Agent Admin Guide

## Purpose

Danmade Patch Agent runs unattended `winget` upgrades on domain-managed Windows endpoints. It is designed for Group Policy scheduled task deployment, Domain CA script signing, and Wazuh collection through Windows Event Log and JSONL file monitoring.

The agent is separate from the interactive workstation updater. It does not use prompts, grids, UAC relaunches, or non-silent installer retries.

Machine mode skips packages that are clearly installed per-user. User mode skips packages that are not clearly per-user. Packages whose scope cannot be determined are left to machine mode so the user task does not repeatedly attempt machine-scope upgrades without elevation.

## Files

- `danmade-patch-agent.ps1`: signed endpoint agent.
- `Test-DanmadePatchAgentHealth.ps1`: signed, report-only stale-process health probe.
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
- Only MSI exit codes listed in `retryableInstallerExitCodes` are retried; `1603` is recorded for review instead of retried blindly.
- Major-version upgrades are deferred by default. Use a package rule with `allowMajorVersion: true` only after package-specific validation.
- Each mode writes its read-only Winget upgrade inventory under `Inventory\` so differences between SYSTEM and logged-on-user views can be investigated.
- Each run has a bounded `maxRunMinutes` deadline. Winget child-process timeouts are clamped to the remaining run time.
- An inapplicable Winget offer is deferred for seven days when the package ID, installed version, available version, and source remain unchanged. The optional `inapplicableUpgradeDeferralDays` policy value is bounded to 1 through 30 days.
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

### Package Rules

Use `packageRules` for packages that need a stricter change path than ordinary unattended patching:

```json
"packageRules": [
  {
    "id": "OpenJS.NodeJS",
    "mode": "ManualOnly",
    "allowMajorVersion": false,
    "processNames": ["node.exe"]
  },
  {
    "id": "Microsoft.PowerShell",
    "mode": "QuiescenceRequired",
    "installerType": "wix",
    "processNames": ["pwsh.exe"]
  }
]
```

- `Auto`: eligible for unattended patching, subject to the major-version gate.
- `ManualOnly`: records `DeferredManualOnly` without launching an installer.
- `QuiescenceRequired`: records `DeferredProcessInUse` while any configured process is running; it never terminates those processes.
- `installerType`: optionally constrains the Winget installer type for the package. Use this only where the Winget manifest provides a validated alternative, such as `wix` for machine-scope PowerShell.

The agent records `DeferredMajorVersion` when the first numeric version component changes. This is the default for all packages. A package rule must explicitly set `allowMajorVersion` to `true` to permit an unattended major upgrade.

Use `retryableInstallerExitCodes` only for known transient installer results, such as `1618` when another installation is in progress. Do not add generic fatal result `1603` without package-specific evidence.

The agent verifies only a completed upgrade by running a complete read-only `winget upgrade --output json` inventory refresh and searching for the package ID. It must never use `winget upgrade --id <id>` as a verification query because that command launches an installer.

Winget can return exit code `0` while its command output describes a failed acquisition. HTTP 5xx download failures are recorded as `TransientDownloadFailure`, retried only within the existing `maxRetries` budget, and never sent to inventory verification. The event includes the download URL, HTTP status/HRESULT-derived status, and the package attempt evidence path. A `No applicable upgrade found` result is recorded as `DeferredNoApplicableUpgrade`, not a failed update; its saved evidence captures the installed and available versions, source, and diagnostic output reference. Repeated runs record `DeferredNoApplicableUpgradeCached` until either those values change or the deferral expires.

When an installer reports `3010` or `1641`, emits an explicit restart-required message, or introduces a new reboot marker, the agent records `RestartRequiredPendingVerification`. It writes the normal pending user-notification state and durable restart-verification state under `State\pending-restart-verification.json`. On the first run after reboot, the agent re-lists the package inventory before attempting the package again. It records `RestartVerificationResolved` only if the package is no longer offered; a package still offered after reboot becomes `RestartVerificationFailedStillOffered`.

Each package attempt writes a JSON evidence record under `Logs\` with the Winget output, structured command and installer outcomes, selected installer type, and before/after reboot-marker snapshots. The installer parser honors final MSI success markers such as `Installation completed successfully` and `Installation success or error status: 0`; a continued custom-action error is not a terminal MSI failure. Genuine terminal installer `1603` and `1605` results remain failures. The agent does not treat generic, pre-existing pending file-renames as proof that a package requires a restart.

The agent keeps restart-verification state and collection boundaries compatible with Windows PowerShell 5.1, which is the runtime used by the scheduled tasks. Missing `PendingFileRenameOperations` registry values are treated as an empty marker rather than an error.

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

## Health Scheduled Task GPO

Create a computer-side task named `Danmade Patch Agent - Health` that runs as `NT AUTHORITY\SYSTEM` shortly before the daily Machine task, for example daily at `1:45 AM`. Use the same signed SYSVOL script and policy path:

```text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
-NoProfile -ExecutionPolicy AllSigned -File "\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\Test-DanmadePatchAgentHealth.ps1" -PolicyPath "\\<domain>\SYSVOL\<domain>\scripts\danmade-patch-agent\danmade-patch-agent.policy.json"
```

The health task reads `maxRunMinutes`, the agent's `State\active-run.json`, and only PowerShell processes whose command line invokes `danmade-patch-agent.ps1` in Machine mode. `StaleAgentProcessDetected` is evidence only: the task does not stop a process, disable a task, or change any unrelated scheduled task. Investigate the endpoint before any termination or servicing action.

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
  - `5601`: stale agent process evidence

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
Get-AuthenticodeSignature C:\ProgramData\DanmadePatchAgent\Test-DanmadePatchAgentHealth.ps1

powershell.exe -NoProfile -ExecutionPolicy AllSigned `
  -File C:\ProgramData\DanmadePatchAgent\danmade-patch-agent.ps1 `
  -Mode Machine `
  -WhatIf
```

Confirm:

- The scheduled tasks appear after `gpupdate /force`.
- `Danmade Patch Agent - Health` runs before the Machine task and has no automatic remediation action.
- `Get-AuthenticodeSignature` returns `Valid`.
- `Application` contains `DanmadePatchAgent` events.
- `C:\ProgramData\DanmadePatchAgent\Events\patch-agent.jsonl` receives one JSON object per line during real runs.
- Wazuh receives either the Event Log records, JSONL records, or both.
- Machine-mode package events do not include clearly per-user installs; user-mode package events include only clearly per-user installs.

## Operational Notes

- The agent does not silently install or repair App Installer in v1. If `winget` is missing or broken, it reports a preflight failure for remediation.
- Installer exit codes `3010` and `1641`, explicit Winget restart output, and a new reboot marker are treated as restart-required pending verification, not an immediate final package failure.
- A failed package is retried only up to `maxRetries`; successful packages are never retried.
- LibreOffice-style HTTP 5xx download failures are acquisition failures, not post-upgrade verification failures. Retest only after the configured retry is exhausted and the download evidence is reviewed.
- A Temurin MSI log can contain a continued custom-action `1603` before a final success marker. Treat the final installer outcome and inventory verification as authoritative.
- `Microsoft.Azd` can be inapplicable because of manifest applicability constraints. Capture one verbose Winget diagnostic before changing package policy; the agent's bounded deferral prevents daily failure notifications.
- For `StaleAgentProcessDetected`, collect the active-run state, task history, and matching process command line first. The health task intentionally does not terminate the process; endpoint remediation is a separately approved action.
- `rebootPolicy` is `ReportOnly` in v1. Use Wazuh or a separate endpoint management workflow to coordinate reboots.
