# Patch Agent Reliability Hardening

Date: 2026-10-16

Operator Level: Windows-first beginner operator (little Windows domain experience and no Linux experience)

Target Environment: Existing nexport.local auto patch test pilot only

## Historical Baseline Before Hardening

The deployed agent and health scripts match source and have valid signatures.
October 6 readback found recurring Edge installer-technology mismatches, daily
Copilot/Terminal no-upgrade outcomes, and a Chrome download timeout classified as
successful installation. Laptop 03 has startup failures without new agent logs;
SCM-LAPTOP2 has an agent process from October 4. Health shares the SYSVOL startup
dependency. The current inventory reader can mistake failure for an empty list,
and unavailable policy falls back to broad defaults.

## Iterations

### October 16 Outcome And Launch Corrections

- [x] Preserve native exit codes with an owned process, concurrent stream copies, UTF-8 capture and bounded output draining.
- [x] Parse final Burn decimal/hex exits, distinguish restart-to-finish from restart-before-install, and compare post-reboot inventory with the recorded target.
- [x] Defer installed-package identity mismatches for administrator review without uninstalling PowerShell or removing its Wix restriction.
- [x] Add NativeUpdater reporting, Node LTS ManualOnly coverage and version/scope/hash-bounded release holds. Pin upgrades to the inspected version.
- [x] Reconcile externally satisfied targets; preserve unresolved and pre-reboot notifications, including mixed outcomes.
- [x] Add seven-day cache verification expiry and idempotent refresh; retain report-only Health/Notification when patching stops for stale policy.
- [x] Complete final regression run, code review and timestamped release signing.
- [x] Publish to the existing 17-member pilot; configure Health catch-up and the user-selected 02:00-04:00 Machine catch-up window in GPO revision 21.
- [ ] Complete protected release/local-task/SYSTEM Health acceptance across all 17 members. Seven now have full readback; DEVWKS03 has a valid cache but incomplete task migration, and nine lack authenticated remoting. See the October 16 automatic GPO convergence checkpoint below.
- [ ] Observe next-window package outcomes without forced upgrades or reboots.

### Earlier Hardening

- [x] Classify no-action, acquisition, and installer outcomes; preserve unique attempt evidence.
- [x] Fail closed on missing policy and inconclusive inventory; protect concurrent run state.
- [x] Add protected local package staging and report-only launch/run health evidence.
- [x] Run deterministic parsing, fake-winget, state, and health regression tests.
- [x] Update operator guidance and record local verification results.
- [x] Sign and timestamp four changed/new PS1 release files with the existing approved domain signer; verify Valid signatures.
- [x] Publish validated signed release to the existing SYSVOL pilot folder; preserve live policy and membership.
- [x] Configure the pilot GPO for protected cache staging and local actions, and verify automatic two-pass migration on five endpoints plus refresh convergence on two previously aligned endpoints. Full 17-member acceptance remains open above.
- [ ] Prove endpoint execution and next-window package outcomes before expansion.

## Decisions And Boundaries

No automatic uninstall, process termination, package blocking changes, or scope
expansion. Failed policy means no patching; no permissive fallback. Applicability
deferrals remain version/source/scope keyed and time bounded. Startup helpers were
initially optional; the October 16 revision 22 checkpoint supersedes that state
with verified local actions and SYSTEM Health on seven endpoints. Full pilot
coverage, interactive behavior and next-window package outcomes remain acceptance
gates.

## October 7 Publication

Published agent, report-only health script, hidden launcher, cache installer, and
cached bootstrap to `\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent`.
Readback on the domain path, OKCDC02, and OKCDC03 confirmed all five hashes match
source; all four PS1 signatures are Valid, timestamped, and use the approved signer.
The live policy remains SHA256
`510081AD96ABD6642D25A342064F2E0D1512EAD5E5A66822BF875E2F8A1CC2AE`.
GPO version 18, GPT.INI, task XML, and all 17 pilot members remain unchanged.

Pre-publication checks passed: 38 source-function checks, 20 fake-winget/health
scenarios, and six cache-dispatch cases. The sixth case proves WhatIf hash
validation and dispatch suppression under Windows PowerShell 5.1. Publication
preflight found the path-based Get-FileHash helper suppresses its ProviderPath
reader under WhatIf; both cache helpers now hash a read-only stream instead.
The published installer passed an AllSigned WhatIf run against the live release,
without changing endpoint files or tasks.

Original rollback copies are under
`E:\Sandbox\danmade-playground\.scratch\publication-20261007-085627-f30debfe\backup`.
The final release manifest and replica readback are under
`E:\Sandbox\danmade-playground\.scratch\publication-20261007-085908-b01c561b`.
No package task was forced, no processes were stopped, and no GPO actions changed.

## Final Validation

October 16 repository checkpoint: reran all five tracked suites under Windows
PowerShell 5.1 before commit/sync: 38 function checks, 26 reliability checks,
28 fake-winget/health scenarios, nine cache-dispatch cases and six pure ACL cases
passed (107 total). Nine changed/new PS1 files parsed without errors, the sample
policy parsed successfully, and all five release PS1 signatures were Valid and
timestamped. Scoped sensitive-pattern scanning found no private keys or literal
credential/token patterns. These local checks do not broaden the seven-device
live task/Health coverage recorded below. Scratch receipts remain local-only.

Local tests must exercise actual source functions, including captured Edge,
Copilot/Terminal, and Chrome outputs and MSI success followed by failure. Live
acceptance requires registered local task actions, successful health heartbeat,
complete agent runs, preserved package restrictions, and no unexpected windows.

Locally verified on October 6 under Windows PowerShell 5.1:

- 38 source-function checks passed, including positive table parsing, empty/failing inventory, policy rejection, atomic replacement, locking, terminal MSI chronology, target-version verification, and timeout no-retry behavior.
- 20 named fake-winget/health scenarios passed, including the captured Edge/Copilot/Chrome outcomes, MSI continued success/later failure/restart, retained retry evidence, and cached-bootstrap stale-process detection.
- Five cache-dispatch cases passed (signer API mocked); invalid signer/hash/path never dispatched, and failure evidence stayed local.
- Existing interactive updater fallback-parser regression passed.
- Seven changed/new PS1 files parsed with zero errors. Release signature readback confirmed Valid and timestamped for agent, health, bootstrap, and cache installer using signer DB13837BA676E68359DC408F476C7DD1F57AFF41.
- Cache installer preflight passed with actual signed files and sample policy under WhatIf. Protected read/execute ACL construction was checked without applying it to an endpoint.
- Git whitespace check passed with CRLF-aware settings required for Authenticode-generated signature blocks.

Limitations: default fake-winget machine identity is mocked. Real SYSTEM task execution, cache installation/ACL readback, GPO migration, signer/WDAC acceptance on endpoints, and next-maintenance-window package outcomes remain deployment gates. SYSVOL publication is complete; endpoint task/GPO mutation and process containment were not performed.

## October 16 Publication And Endpoint Validation

The current agent SHA256 is
`EE28999184C61A7B983E2C41378F867885FAF4989E32FB8C48D63A2672296504`;
policy SHA256 is
`12C2B091646F4B42DFE936D990C180F97DF700AA3C506CE41F8E31FEB3CC957C`.
All five release PS1 files are Valid and timestamped with the approved domain
signer. Seven payload hashes match on OKCDC02 and OKCDC03; AD/GPT revisions are
21 on both, with user revision unchanged. Membership remains 17.

Local checks passed: 38 source-function cases, 26 native-process/state/release-hold
cases, 28 fake-winget/health scenarios, nine cache-dispatch cases, six pure ACL
cases and the existing interactive fallback-parser regression. Real published
metadata checks hold the bad VS Code 1.142.0 installer in both scopes.
Review retained fatal-installer precedence, scope gates and target verification.

NEX-LAPTOP-01 and NEX-LAPTOP-05 now have protected local releases and local
actions for the existing four agent tasks. Parent Users write grants are absent;
Users retains read access and Package is SYSTEM/Administrators writable only.
Two unchanged refreshes retain the same release directory. The signed installer
supports explicit -Unattended and -SecureParent; inherited and explicit ACL
normalization is tested without changing test-host permissions.

Actual SYSTEM Health and Refresh Local Cache tasks both completed with result 0
on both laptops. Health records retain genuine historical Machine failures and
stale-run findings; successful Health is not successful application patching.
The refresh tasks use the approved daily/boot/Group Policy event triggers.
Machine/Health catch-up and offline-capable hidden user task settings were aligned
directly on these two endpoints. No package task or reboot was forced.

At the initial revision 21 checkpoint, automatic preference registration remained an acceptance gap: the clients cached
the new task payload and processed the Scheduled Tasks extension, but existing
actions remained UNC and the new refresh task was absent. The two endpoint
actions/settings and native refresh-task registration were therefore applied
directly, preserving principals and triggers. Native Task Scheduler registration
requires its XML namespace; this does not by itself prove the GPP root cause.
Do not mark GPO convergence or the other 15 endpoints effective from these checks.
DEVWKS03's current operator token is not elevated; its manual gpupdate client
stalled and only that newly launched client was stopped. No agent was stopped.

Release, rollback, replica, native task and endpoint receipts:
`E:\Sandbox\danmade-playground\.scratch\outcome-fixes-20261016`.
Next at that checkpoint: prove automatic GPO convergence, remaining protected caches and hidden
interactive launches, then observe the next maintenance window. The VS Code
hold expires October 23 and must be reviewed before that date.

## October 16 Automatic GPO Convergence

Baseline readback on NEX-LAPTOP-03 proved GPP had registered the SYSTEM cache
refresh task automatically, but it had never run. Its daily trigger began the
next day; boot/background-policy triggers did not promptly bootstrap a manual
refresh. Cache-file filters correctly retained UNC actions until provisioning.
The earlier 30-second gpupdate client wait also preceded actual CSE completion.
No GPP XML namespace or principal change was needed.

- [x] Publish computer revision 22 on OKCDC02/03, preserving user revision 0, the OU link, Apply filtering and all 17 auto patch test members. Only the refresh task's trigger set/change timestamp changed: registration delay 15 seconds and manual/background computer-policy events 8004/8006, retaining daily/boot triggers.
- [x] Prove automatic cache staging on NEX-LAPTOP-03, DEVWKS01, DEVWKS05, NEX-LAPTOP-02 and NEX-LAPTOP-06 without direct task-action edits. Existing laptop 01/05 caches also refresh automatically. Source scripts, signatures and package policy remain unchanged.
- [x] Prove second-pass GPP migration on NEX-LAPTOP-03: all four actions local, SYSTEM Machine/Health with catch-up, limited hidden WScript User/Notification without network-only startup. Matching release hashes, Valid signatures, protected cache ACLs and the 02:00-04:00 machine-only window were read back. Its actual local SYSTEM Health completed with result 0 at 13:59:54 Central; a historical Machine-failure finding remains.
- [x] Complete second-pass action and report-only Health readback across all seven remotely reachable devices: DEVWKS01, DEVWKS05, NEX-LAPTOP-01, NEX-LAPTOP-02, NEX-LAPTOP-03, NEX-LAPTOP-05 and NEX-LAPTOP-06. All four actions are local; hashes/signatures, protected cache ACLs, principals, hidden user settings and the selected window match. Actual SYSTEM Health returned 0 with fresh records on all seven. Findings respectively remain 0, 1, 3, 2, 1, 2 and 0; successful reporting is not resolution of those findings. Five devices migrated automatically in this pass; laptop 01/05 were previously aligned directly and received revision 22 refresh settings through GPP.
- [ ] Observe the next real Machine window and normal User/Notification logon/unlock behavior. Seven of 17 endpoints have full task/Health readback; DEVWKS03 has a matching cache but incomplete task acceptance, and nine lack authenticated remoting. Do not equate inaccessible endpoints with agent failures.
- [ ] Complete DEVWKS03 task migration/readback in an authorized elevated context or subsequent normal computer policy cycle. Its background process staged a matching Valid cache at 13:59:42 Central, but visible user actions remain UNC and this operator cannot inspect SYSTEM tasks.
- [ ] Obtain authenticated readback from DEV-DARR, DEV-WKS02, SCM-LAPTOP1, DEV-SCOTT, DEV-PUPEK2, DEV-THAI, DEVWKS04, SCM-LAPTOP2 and NEX-LAPTOP-04. The last two report Kerberos 0x80090322; failed WinRM does not prove an agent failure or an offline endpoint.

Some gpupdate clients exceeded their bounded wait despite completed computer
policy/Scheduled Tasks extension events. Only those newly launched clients were
stopped; task/cache evidence establishes targeted convergence, not blanket GPO
health. Existing access-check errors remain outside this narrow trigger change.
Newly returning devices intentionally need two normal computer-policy cycles:
first stage the signed cache, then pass the file-exists filters and migrate actions.
No recursive policy-refresh task was introduced.
No patch task, installer or reboot was forced. Actual hidden interactive behavior
and next-window application outcomes remain unverified by these Health checks.

Rollback and live readback receipts:
`E:\Sandbox\SysOps-And-CyberSecurity\.scratch\patch-agent-gpo-convergence-20261016`.
ScheduledTasks.xml SHA256 on both DCs:
`A69EF63B948249392DEE1A0D05D2E905F42685AD8C5819F18F067A3658BA2684`.

## Future Iteration: Microsoft.WinGet.Client

This follow-up is not part of the October 16 runtime change. Microsoft's
[SYSTEM-context guidance](https://learn.microsoft.com/en-us/windows/package-manager/winget/troubleshooting#system-context)
states that the packaged WinGet CLI is unsupported as LocalSystem and identifies
Microsoft.WinGet.Client for machine-wide applications in SYSTEM context.

- [ ] Pin and stage a reviewed module version and dependencies in an isolated pilot; do not install the latest module dynamically during an agent run.
- [ ] Exercise structured package objects and upgrade results under real SYSTEM. Reproduce PowerShell identity correlation, .NET restart, Edge installer technology and unavailable-source cases.
- [ ] Preserve exact ID/source/scope, manual-only/release-hold policy, major-version/quiescence gates, retry limits, deadline, persistent restart state and target-version verification.
- [ ] Prove machine/user separation, hidden execution, AllSigned trust, local startup and rollback. Machine-wide SYSTEM support does not establish support for logged-out user profiles.
- [ ] Compare module outcomes and evidence with the CLI path before migration. Keep the current release and module experiment independently reversible.

Acceptance requires actual endpoint behavior, not merely a module import or fewer
text parsers. No module is installed as part of this release.
