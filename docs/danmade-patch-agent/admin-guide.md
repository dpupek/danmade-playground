# Danmade Patch Agent Admin Guide

Date: 2026-10-16

Operator Level: Windows-first beginner operator (little Windows domain experience and no Linux experience)

Target Environment: nexport.local auto patch test pilot

## Purpose

Danmade Patch Agent runs unattended `winget` upgrades on domain-managed Windows endpoints. It is designed for Group Policy scheduled task deployment, Domain CA script signing, and Wazuh collection through Windows Event Log and JSONL file monitoring.

The agent is separate from the interactive workstation updater. It does not use prompts, grids, UAC relaunches, or non-silent installer retries.

Machine mode skips packages that are clearly installed per-user. User mode skips packages that are not clearly per-user. Packages whose scope cannot be determined are left to machine mode so the user task does not repeatedly attempt machine-scope upgrades without elevation.

## Files

- `danmade-patch-agent.ps1`: signed endpoint agent.
- `Test-DanmadePatchAgentHealth.ps1`: signed, report-only stale-process health probe.
- `Invoke-DanmadePatchAgentCached.ps1`: local bootstrap validating the release signer and policy hash.
- `Install-DanmadePatchAgentCache.ps1`: elevated release staging and optional task-action migration.
- `danmade-patch-agent.policy.json`: required policy file copied to endpoints.
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
If neither path supplies a valid policy, the agent records `PolicyLoadFailedNoPatching`, exits `2`, and preserves existing restart/notification state. Built-in defaults fill optional fields only; they are never an unattended fallback for missing policy. A policy must supply Boolean `enabled`/`includeUnknown`, array allow/block lists and package rules, and integer retry/duration limits.

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

The agent verifies only a confirmed completed upgrade using read-only scope-filtered inventory. It must never use `winget upgrade --id <id>` as a verification query because that command launches an installer. JSON is preferred when supported; an explicit unsupported-option result permits table fallback. Command failure, timeout, unknown schema, and unrecognized table rows are inconclusive, never an empty inventory. `InventoryInconclusive` stops patching without clearing restart state. A newer offer is not failure when the installed numeric version has reached the original target.

Winget can return exit code `0` while its command output describes a failed acquisition. HTTP 5xx download failures are recorded as `TransientDownloadFailure`, retried only within the existing `maxRetries` budget, and never sent to inventory verification. The event includes the download URL, HTTP status/HRESULT-derived status, and the package attempt evidence path. A `No applicable upgrade found` result is recorded as `DeferredNoApplicableUpgrade`, not a failed update; its saved evidence captures the installed and available versions, source, and diagnostic output reference. Repeated runs record `DeferredNoApplicableUpgradeCached` until either those values change or the deferral expires.

When an installer reports `3010` or `1641`, emits an explicit restart-required message, or introduces a new reboot marker, the agent records `RestartRequiredPendingVerification`. It writes the normal pending user-notification state and durable restart-verification state under `State\pending-restart-verification.json`. On the first run after reboot, the agent re-lists the package inventory before attempting the package again. It records `RestartVerificationResolved` only if the package is no longer offered; a package still offered after reboot becomes `RestartVerificationFailedStillOffered`.

Each package attempt writes a JSON evidence record under `Logs\` with the Winget output, structured command and installer outcomes, selected installer type, and before/after reboot-marker snapshots. The installer parser honors final MSI success markers such as `Installation completed successfully` and `Installation success or error status: 0`; a continued custom-action error is not a terminal MSI failure. Genuine terminal installer `1603` and `1605` results remain failures. The agent does not treat generic, pre-existing pending file-renames as proof that a package requires a restart.

### Reliability Outcomes

- `DeferredInstallerTechnologyMismatch`: Edge or another application has incompatible installed/offered technology. Requires administrator investigation; never automatically uninstall. Cached records preserve the reason and version/source/scope evidence.
- `DeferredNoAvailableUpgrade`: scoped command did not perform an upgrade even though discovery offered one. Investigate Copilot/Terminal ownership and provisioning before changing scope. This is unresolved applicability, not successful patching.
- `TransientDownloadFailure`: includes HTTP 429/5xx, timeout, DNS, reset, and connection failures. Backoff and jitter respect the retry/run budget. Retry only before installer-start evidence appears.
- `InstallationOutcomeUncertain`: an installer may outlive its winget parent. No immediate retry or verification. Saved evidence suppresses reattempts through the applicability cooldown; administrator review is required before it expires.
- `UnconfirmedUpgradeOutcome`: zero exit without actual installation-success evidence. No verification and no blind retry.
- `VerificationInconclusive` / `InventoryInconclusive`: command or parsing uncertainty. Preserve restart evidence and investigate the recorded query artifact.
- `SkippedConcurrentRun`: another run owns the log-root lock. Active-run cleanup is ownership checked; durable active/deferral/restart writes are atomic.

MSI terminal records are evaluated chronologically. A later terminal failure wins over earlier success, while explicitly continued custom actions remain nonfatal. Explicit `3010`/`1641` restart results remain restart-required. A fatal installer failure takes precedence over incidental new reboot markers. Each retry has a unique installer log and attempt JSON with arguments, scope, attempt number, and verification outcome. Inventory query evidence follows normal retention.

Run the tracked [function tests](../../windows-update-scripts/tests/Test-PatchAgentFunctions.ps1), [fake-winget harness](../../windows-update-scripts/tests/Test-PatchAgentHardening.ps1), and [cache dispatch tests](../../windows-update-scripts/tests/Test-PatchAgentCache.ps1) before signing a release. The harness uses mocked endpoint observations by default; passing it does not prove real SYSTEM execution or endpoint rollout.

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

## Protected Local Release Deployment

Goal of this section: remove the runtime SYSVOL dependency while retaining signer trust and the existing pilot scope.

Operator note: stop and contact Dan Pupek / Nexport Systems Architecture if signature checks fail, the parent folder allows non-admin writes, or any patch task is running.

1. Sign and timestamp all changed/new PS1 files, including both cache helpers. Verify `Valid` signatures from the approved signer. Keep the existing pilot policy unchanged.
2. Stage the files in the authoritative SYSVOL release folder. From an elevated endpoint terminal, run the [cache installer](../../windows-update-scripts/Install-DanmadePatchAgentCache.ps1) first with `-WhatIf` and the approved signer thumbprint.
3. Run the same command without `-WhatIf`, with `-UpdateTasks -Confirm:$false` only after reviewing the four planned patch-agent task names. It preserves principals and triggers and never updates unrelated tasks.
4. Update the existing pilot GPO action paths to the stable local paths below. Otherwise Group Policy Preferences will restore the old UNC actions. Increment the computer-side AD/GPT version together; retain the user-side version.
5. Refresh policy and read back task actions, signatures, protected ACLs, and next-window logs before widening scope.

Preflight command, from the signed release directory:

```powershell
.\Install-DanmadePatchAgentCache.ps1 -SourcePath '\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent' -ExpectedSignerThumbprint 'DB13837BA676E68359DC408F476C7DD1F57AFF41' -UpdateTasks -WhatIf
```

After reviewing the preflight and confirming the release/GPO change window:

```powershell
.\Install-DanmadePatchAgentCache.ps1 -SourcePath '\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent' -ExpectedSignerThumbprint 'DB13837BA676E68359DC408F476C7DD1F57AFF41' -UpdateTasks -Confirm:$false
```

The installer protects `C:\ProgramData\DanmadePatchAgent\Package`, retains immutable releases for rollback, validates copied signatures and hashes, and atomically publishes `current.json`. SYSTEM and Administrators may write; Users may read/execute only. It refuses a writable parent or reparse-point targets. Existing log/state permissions are not broadened. The bootstrap writes `LocalBootstrapFailed` locally if the release cannot be validated. Machine/Health failures go to ProgramData; user/notification failures go to the user's LocalAppData.

Expected publication output: `Status=LocalReleasePublished`, a release directory, `TasksUpdated=4`, and the policy SHA-256. A successful staging command is not endpoint acceptance.

Rollback: use the protected backup path returned in `PreviousPointer` to restore `current.json` after verifying its release still has valid signatures. First-install output has no prior pointer. Do not remove an active release or weaken execution policy.

## Legacy File Deployment GPO

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
-NoProfile -NonInteractive -ExecutionPolicy AllSigned -File "C:\ProgramData\DanmadePatchAgent\Package\Invoke-DanmadePatchAgentCached.ps1" -Mode Machine
```

## Health Scheduled Task GPO

Create a computer-side task named `Danmade Patch Agent - Health` that runs as `NT AUTHORITY\SYSTEM` shortly before the daily Machine task, for example daily at `1:45 AM`. Use the signed local bootstrap:

```text
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
-NoProfile -NonInteractive -ExecutionPolicy AllSigned -File "C:\ProgramData\DanmadePatchAgent\Package\Invoke-DanmadePatchAgentCached.ps1" -Mode Health
```

The health task reads `maxRunMinutes`, active-run state, task results, and bounded recent logs. It recognizes direct agent and cached-bootstrap Machine processes. It records `HealthCheckCompleted`, `StaleAgentProcessDetected`, `OrphanedActiveRunDetected`, `AgentTaskNeedsReview`, and `AgentRunEvidenceMissing` (no terminal record within 26 hours). These are report-only: no termination or task disabling. An absent health heartbeat is not proof of health.

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
C:\Windows\System32\wscript.exe

Arguments:
//B //NoLogo "C:\ProgramData\DanmadePatchAgent\Package\Start-DanmadePatchAgentHidden.vbs" user
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

## Final Validation

On a pilot workstation:

```powershell
Get-AuthenticodeSignature C:\ProgramData\DanmadePatchAgent\Package\Invoke-DanmadePatchAgentCached.ps1

powershell.exe -NoProfile -NonInteractive -ExecutionPolicy AllSigned `
  -File C:\ProgramData\DanmadePatchAgent\Package\Invoke-DanmadePatchAgentCached.ps1 `
  -Mode Health `
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

### October 16 Reliability Rules

- For scheduled `powershell.exe -File` cache installation, explicitly use `-Unattended`; Windows PowerShell 5.1 cannot bind `-Confirm:$false` through that native entry point. AllSigned and signer validation remain required. This option suppresses confirmation only; it neither elevates nor bypasses trust checks.
- Review parent ACLs before approving `-SecureParent`. It protects only `C:\ProgramData\DanmadePatchAgent` against broad Everyone/Authenticated Users/Users/Domain Users write grants, retains read/execute and other principals, and saves an ACL receipt. The installer refuses an unsafe parent by default. No log/state contents are removed. Unsupported or noncanonical ACLs require administrator review; do not weaken the protection check.

- `DeferredInstalledPackageNotFound` preserves WinGet's native `0x8A150014` evidence for administrator repair; it is not a successful installation and does not ask the user to restart. Do not uninstall PowerShell or remove its Wix/quiescence restriction automatically.
- Final Burn `i007` exits are interpreted as decimal or hex. `0xbc2` is restart-required success; HRESULT-wrapped MSI `1603` remains fatal. WinGet restart-to-finish differs from restart-before-install failure.
- `NativeUpdater` leaves installation to the product's native updater and reports the installed/offered version as `DeferredNativeUpdater` with administrator follow-up. Confirm the native updater is functioning; this is not a compliance-success exemption.
- Add both `OpenJS.NodeJS` and `OpenJS.NodeJS.LTS` as `ManualOnly` to protect the development runtime.
- Optional `releaseHolds` on a package rule contain `version`, `scope` (`machine` or `user`), `installerSha256` (64 hex characters), `until` (ISO date/time), and `reason`. The agent inspects the selected installer hash without installing. A matching known-bad hash yields `DeferredKnownBadRelease`; failed or ambiguous metadata yields `DeferredReleaseHoldUnverified`. A corrected hash releases the hold immediately, even at the same version. Review holds before expiry; do not disable installer hashing or target verification.
- Upgrades pin `--version` to the inspected target. Post-reboot installed versions at or above that recorded target count as resolved even if a newer release is offered.
- Failure notifications and cached deferrals clear only when successful inventory or an exact, unique, same-scope uninstall record proves the target is installed. Pending reboot notifications require recorded reboot verification; a clean run must not discard an older mixed-outcome restart.
- An enabled maintenance window with `machineOnly: true` limits Machine starts without suppressing normal User logon updates. The pilot catch-up window is 02:00-04:00 local time. No new package attempt begins outside it; an installer already started may finish afterward. No automatic reboot is performed.
- Local releases expire for Machine/User patching seven days after `current.json` `verifiedAt`. A daily signed cache refresh revalidates source hashes/signers and atomically refreshes the pointer; unchanged releases reuse the same directory. Health/Notification remain available to report stale policy. `LocalPolicyRefreshRequired` and `LocalBootstrapFailed` require operator review, not an execution-policy bypass.

For the separately tracked SYSTEM API evaluation, see the [Microsoft.WinGet.Client iteration](reliability-plan.md#future-iteration-microsoftwingetclient). It is not enabled by this release.

- The agent does not silently install or repair App Installer in v1. If `winget` is missing or broken, it reports a preflight failure for remediation.
- Installer exit codes `3010` and `1641`, explicit Winget restart output, and a new reboot marker are treated as restart-required pending verification, not an immediate final package failure.
- A failed package is retried only up to `maxRetries`; successful packages are never retried.
- LibreOffice-style HTTP 5xx download failures are acquisition failures, not post-upgrade verification failures. Retest only after the configured retry is exhausted and the download evidence is reviewed.
- A Temurin MSI log can contain a continued custom-action `1603` before a final success marker. Treat the final installer outcome and inventory verification as authoritative.
- `Microsoft.Azd` can be inapplicable because of manifest applicability constraints. Capture one verbose Winget diagnostic before changing package policy; the agent's bounded deferral prevents daily failure notifications.
- For `StaleAgentProcessDetected`, collect the active-run state, task history, and matching process command line first. The health task intentionally does not terminate the process; endpoint remediation is a separately approved action.
- `rebootPolicy` is `ReportOnly` in v1. Use Wazuh or a separate endpoint management workflow to coordinate reboots.
