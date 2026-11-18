# Patch Agent Reliability Hardening

Date: 2026-11-06

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

### November 6 Read-Only Acceptance Review

- [x] Collect eight of 17 pilot members at 11:17 Central: seven use the November 4 main-package release; all eight cached signatures are Valid. Both DC agent/policy hashes still match publication. Bounded logs show no parse errors, active-run records, matching live agent processes or recent LocalBootstrapFailure events.
- [x] Observe real VerifiedNotOffered acceptance for DEVWKS03 WinDbg and NEX-LAPTOP-02 Terminal. NEX-LAPTOP-05 completed its 02:07 Eastern Machine assessment with three verified successes; its earlier Health overdue warning predates this run. This is sampled acceptance, not full pilot completion.
- [ ] First implement separate application-restart pending handling: Terminal on DEVWKS01/03 and ChatGPT on NEX-LAPTOP-01 return success plus "Restart the application to complete the upgrade", then VerificationFailedStillOffered. Preserve strict Main target proof; persist bounded per-user/package/source/target state, request close/reopen and suppress repeated reinstall attempts. Do not turn application restart into OS reboot or assume installation is current. Add deterministic AAAA cases for app-restart output, target convergence, unchanged target, timeout and ordinary installer failure.
- [ ] Diagnose NEX-LAPTOP-03 refresh launch before its seven-day cache gate expires around November 11 at 04:44 Central. Latest refresh 10:28 returns 0xFFFD0000; cache is still prior 33BE7036CE8BCD89FB0A09EB604E1CFA4A498388BE0C4398E97E3DA264AC424C, last verified November 4. Capture exact SYSTEM startup stderr/file-open evidence; current administrative UNC reads cannot establish root cause. Extend report-only Health to include Refresh Local Cache failures and aging successful verification before expiry; it currently ignores that task and reports zero findings here.
- [ ] Correct conservative target proof for Zoom.Zoom.EXE 7.2.5 (51704): System.Version.TryParse rejects the raw target before installed-registration comparison. NEX-LAPTOP-02 has rebooted but retains RestartVerificationInconclusive/admin-remediation evidence. Use product-aware version/build tests and same-scope proof; retain raw values and distinguish post-reboot review from another reboot request. Actual target installation remains unverified. Reconcile older .NET notices only against a real Machine assessment or authoritative installed proof.
- [ ] Diagnose overnight availability on laptops 02/06 without widening the 02:00-04:00 local Machine window. Latest runs skip outside that window; last real assessments are November 1/5 respectively. One recovered laptop 05 night does not establish a permanent fix. Coordinate pending OS restarts on DEVWKS03/05 and laptop 05; keep Azd/Rustup applicability and native-updater/Unknown-scope coverage separate.
- [ ] Recover visibility for nine members: DNS unresolved for SCM-LAPTOP2/laptop 04 and WinRM collection failed for seven others. Do not expand rollout from an eight-member sample.

Detailed evidence and ordered recommendations are in the SysOps pilot runbook's
November 6 section. Local receipts are under
E:\Sandbox\SysOps-And-CyberSecurity\.scratch\patch-agent-review-20261106-171712,
including analysis-summary.json and classifier-readback.json. Pure production
function replay confirms the application-restart gap and unparseable Zoom target;
no agent entry point or installer ran. The 174-check release receipt remains
historical; this documentation-only review makes no new code/test/deployment claim.
No endpoint state, policy, GPO, membership, signed code, task or network changes,
commit or push occurred. Repo docs and Quick Memory apply; no exact execution
case or Vanta target was identified, DNS-only case 255215 remains unchanged, and
marketplace synchronization is inapplicable because no skill/plugin changed.

### November 4 Authorized Main-Package Release

- [x] Recheck tested unsigned source SHA256 0F96618A97EB6B71328E06C72A1525761612C7647EF40E2CDAC745ED0859BFD6, 174-check validation receipt, approved signer/private key and both DC baselines.
- [x] Sign/timestamp the agent with SHA256, approved signer DB13837BA676E68359DC408F476C7DD1F57AFF41 and DigiCert timestamp; independently verify Valid and zero parser errors. Exact staged release passed Windows PowerShell 5.1 AllSigned cache-installer WhatIf without changing local cache.
- [x] At 17:30 Central back up all seven existing payload files, publish only the agent with atomic replacement, and verify both replicas. All five PS1 signatures are Valid/timestamped; policy, other payload hashes, ACLs, GPT.INI/ScheduledTasks XML and GPO AD revision 22/0 are unchanged. Pilot membership remains 17.
- [x] Record release/rollback receipts and update handoff/Quick Memory. No task start, cache refresh, package installation, GPO change, restart, commit or push occurred.
- [ ] Observe natural endpoint cache convergence and a real User assessment proving main-package scope/version behavior. Laptop 03's intermittent refresh-launch failure remains a separate open diagnostic; publication does not fix or prove that path.

This supersedes the local unsigned checkpoint below. Published agent SHA256 is
2FCCB1A6E7E5DDB1F673E4F0CA23C7374A11D6B67D0A1794FF9839978B476EB0.
Live policy remains 27C20CBF5C0ECD03B9180457BFA08B6C8BE24A9BCBA331FF1CCEF0CF44F150FD.
Publication manifest and verified rollback backup are source-local under
`.scratch\publication-20261104-173035-e95598b6`; manifest status is
PublishedAndVerifiedOnBothDCs. Candidate, baseline/post-replica hashes/ACLs,
scope-before/after and release-validation.json are in
`.scratch\appx-main-release-20261104`. These ignored receipts are not committed.

Rollback needs owner approval: verify current agent identity, stage and atomically
restore only the prior agent from that backup (SHA256
33BE7036CE8BCD89FB0A09EB604E1CFA4A498388BE0C4398E97E3DA264AC424C),
then verify both DCs and coordinate cache convergence. Stop if backup/current
hashes or signatures differ. Policy, other payloads and endpoint state need no
rollback for this agent-only publication. No exact execution case or Vanta target
applies; repo docs and Quick Memory are the tracking surfaces. No reusable
skill/plugin changed, so marketplace synchronization is not applicable.

### November 4 Main-Package Fix And Refresh Diagnosis

Baseline: the November 3 release queries Main and Bundle, then treats two healthy
records of the same family as ambiguous. Terminal's bundle version differs from
the installed main version. Cache refresh launches fail before a successful cache
receipt on DEVWKS03; laptop 03 also has intermittent 0xFFFD0000 exits.

- [x] Restrict installed identity/version evidence to healthy main packages; exclude bundle-only and unknown package-type records and preserve true multiple-main ambiguity. Query Main only, keep current-user-only scope and exclude IsBundle=true or missing IsBundle metadata.
- [x] Add eight AAAA regressions for real-shaped bundle/main pairs, distinct versions, bundle-only/failed main and target proof. All seven Windows PowerShell 5.1 suites pass: 38 function, 29 reliability, 31 fake-winget/Health, nine cache dispatch, six ACL, 13 pilot and 48 coverage scenarios, 174 total. Parser and whitespace checks pass; coverage was rerun after the final argument-guard assertion.
- [x] Diagnose domain-alias/DC referral and exact UNC launch/signature evidence without starting cache/patch tasks or changing trust, GPO, DFS or endpoint state. DEVWKS03 now has the November 3 cache: installed 10:12, verified 11:44 Central; successful SYSTEM task events at 10:12/11:42/11:44 predate this diagnosis. Current domain alias/direct DC reads agree and are Valid; native read-only DFS cache API selects OKCDC03, unchanged before/after. AllSigned WhatIf succeeds via all three paths in a Windows PowerShell-compatible child environment. No DFS flush or repair was performed.
- [x] Review the minimal semantic change against the signed publication, update the admin guide and record findings/limitations. Local current-user function-only read selects exactly one healthy Terminal main at 1.24.12741.0, returns User scope and rejects target 1.25.2900.0. No installer/agent entry point ran.
- [x] Sign/timestamp and publish the correction after authorization at 17:30 Central, as recorded in the release checkpoint above. 0F96618A97EB6B71328E06C72A1525761612C7647EF40E2CDAC745ED0859BFD6 identifies the historical tested unsigned source; current signed publication is 2FCCB1A6E7E5DDB1F673E4F0CA23C7374A11D6B67D0A1794FF9839978B476EB0. Policy remains unchanged; real User acceptance is pending. No commit/push occurred.
- [ ] Capture exact SYSTEM launch stderr/file access on laptop 03 before changing refresh design. Its latest 12:33 refresh still returns 0xFFFD0000 while cache remains current and Valid, last verified 04:44. The sampled failing interval has no matching PowerShell engine-start event; consistent with a pre-entry UNC/File launch failure, not a proven installer or certificate defect. Remote-admin UNC failures are limited by SMB double-hop and cannot prove SYSTEM failure. Existing machine-account SYSVOL sessions likewise do not prove exact script access at failure time.

Regression transcript: source-local `.scratch\appx-main-fix-20261104-123155\regressions.txt`.
Diagnosis receipts: SysOps `.scratch\cache-launch-diagnosis-20261104-183311` and
`cache-launch-diagnosis-20261104-183549`. The first probe's security-module error
was a diagnostic artifact from a raw PowerShell 7 ProcessStartInfo module path;
repeating with the registered Windows PowerShell module path only in the owned
child yielded three exit-zero AllSigned WhatIf reads. No persistent environment
or execution-policy change occurred; do not call that artifact the task's root cause.
Laptop 03's existing SYSTEM network report shows healthy DC locator/secure channel,
but its DC-name list was passed as one comma-joined string; that report's resulting
invalid-name DNS row is not a DC outage. Its SMB multichannel warnings explicitly
say extra-channel failures need not be application failures. Original cause of
DEVWKS03 recovery and laptop 03's exact startup failure remain unproven.

### November 4 Read-Only Outcome Check

- [x] Review seven of 17 pilot members and verify both DC published hashes/agent signature. Six endpoint caches are current; DEVWKS03 remains on the old release. Six overnight Machine assessments yield eight VerifiedNotOffered-backed successes and no new package failures in the bounded sample.
- [x] Confirm Teams/native-updater dispositions and new assessment/coverage fields on live runs. Laptop 06 completed overnight after its earlier Health overdue warning. Laptop 05 remains overdue after a daytime skip; no new reboot verifies the retained DEVWKS03/05 targets.
- [ ] Fix Main+Bundle AppX handling with real-shaped fixtures: laptop 01 production function-only readback returns two healthy pinned Terminal registrations and Unknown scope. Bundle version 3001.24.12741.0 differs from main 1.24.12741.0; never pick the largest bundle/main version as installed-target proof. Preserve genuine ambiguity and User-only guards. No code/signature/policy change is authorized by this log review.
- [ ] Diagnose cache refresh launch failure: DEVWKS03 scheduler event 201 proves repeated 0xFFFD0000, including the daily 01:30 run; current cached package/Health execution still works. Laptop 03's latest refresh also failed but previous refreshes succeeded and its cache is current. Do not infer a missing task from the filtered local token.
- [ ] Diagnose laptop 05 overnight availability and unmatched ChatGPT identity; named-user AppX query finds no pinned ChatGPT family and matching ARP query is empty. Preserve Unknown until authoritative identity proof; it is not confirmed to be the bundle issue.
- [ ] Complete approved restart coordination, Azd/Rustup applicability diagnostics and ten-member reachability. No task, install, cache refresh, reboot, GPO or source edit was performed during this investigation.

Detailed findings/receipts are in the SysOps pilot runbook's November 4 section and
local `.scratch\patch-agent-review-20261104-131703` under that workspace.
Earlier six-suite/local implementation/publication checkpoints below are historical;
the November 3 signatures and live release remain unchanged. This review does not
complete pilot acceptance, signing/deployment of a bundle fix or source-control sync.

### November 3 Authorized Signing And Publication

- [x] Verify tested source hashes, approved signer/private key, both DC release baselines and current 17-member pilot scope.
- [x] Sign and timestamp the changed agent; independently verify Valid signer and parse final bytes. Use SHA256, DigiCert timestamping and the existing approved CurrentUser certificate with private key, thumbprint DB13837BA676E68359DC408F476C7DD1F57AFF41 (expires July 6, 2027). No PFX/key export or import was needed.
- [x] Back up all seven live release files with verified hashes; merge only Teams, Teams Free, Terminal, WinDbg and ChatGPT sample rules into the existing policy. Preserve all prior rules/holds and every other top-level field, including includeUnknown=true, maxRetries=1, seven blocks and Machine-only 02:00-04:00 window.
- [x] At 16:10 Central publish only agent/policy with atomic per-file replacement, hash/signature and ACL checks; independently verify all seven release hashes and five Valid timestamped PS1 signatures on both DCs. GPO AD versions remain 22/0 and GPT.INI/ScheduledTasks XML hashes are unchanged.
- [x] Record publication receipts and rollback, distinguish published from endpoint-cache convergence, and update handoff/Quick Memory. No refresh, patch, health, notification or restart task was forced; no GPO write, commit or push occurred.
- [ ] Observe natural cache refresh and real next-window/user/reboot outcomes. Publication is verified, but endpoint activation and full pilot acceptance are not yet verified.

This supersedes the unsigned/unpublished local checkpoint below. Published agent
SHA256: 33BE7036CE8BCD89FB0A09EB604E1CFA4A498388BE0C4398E97E3DA264AC424C.
Published policy SHA256: 27C20CBF5C0ECD03B9180457BFA08B6C8BE24A9BCBA331FF1CCEF0CF44F150FD.
Other release files are unchanged. Evidence/backup/staged release are local-only in
`.scratch\publication-20261103-outcome-coverage`; manifest.json records before/after
hashes, original ACLs, promoted files and publication time; replica-readback.json and
scope-readback.json record independent DC and group/GPO checks. The signed cache
installer passed Windows PowerShell 5.1 AllSigned WhatIf against the exact candidate
release before publication without modifying local cache. DEVWKS03's earlier cache
pointer still names the prior release; no new endpoint effectiveness is claimed.

Rollback requires owner approval: retain the backup, restore only the two changed
files through verified staging/atomic replacement, using baseline agent
14687A9AABA0F424FB8C00C407D41ACB909AEB8C3B314A51529E17AF98228D9C and policy
5D6FA4DF387B10296B1D9A7E6DFDFF273BFC9CCCBCFB35E3492EC5813B2B8FE1,
then verify signatures/hashes on both DCs and coordinate cache convergence. Stop
if live bytes differ from this release or backup hashes/signatures do not match.
Do not delete pending notification/restart state or broaden execution trust.

### November 3 Verification And Coverage Implementation

Baseline: post-reboot verification accepts an absent upgrade offer without installed
registration proof. Scope discovery picks the first best display-name registry
match and does not query current-user AppX registration. Teams build changes are
held by the generic major guard. Terminal summaries lack distinct policy/scope
disposition counts, making NoPackagesSelected easy to misread as full currency.

- [x] Require positive installed-target evidence after a later boot; retain inconclusive pending state and existing genuine-failure behavior. Immediate post-install verification is unchanged.
- [x] Add exact policy-pinned current-user AppX family matching and reject ambiguous registry matches without weakening scope guards. Sample mappings cover Terminal, WinDbg and ChatGPT; Keeper remains unconfigured pending authoritative identity.
- [x] Add explicit Teams/Teams Free NativeUpdater sample-policy rules; preserve other products' major-version restrictions. Native-updater delegation remains unresolved administrator-owned evidence, not compliance success.
- [x] Add additive assessment outcome and coverage counts while retaining terminal statuses and separating completed inventory from package currency.
- [x] Add AAAA deterministic regressions and fake-winget integration cases; all seven Windows PowerShell 5.1 suites pass: 38 function, 29 reliability, 31 fake-winget/Health, nine cache dispatch, six ACL, 13 pilot and 40 coverage scenarios (166 total). Parser and sample-policy JSON checks pass.
- [x] Review final changes and update the next-agent handoff; Markdown lint and whitespace checks pass. Signing and pilot publication remain separate gates; do not publish changed unsigned bytes.

Validation is local only. The fake-winget harness uses mocked SYSTEM identity and
does not install real packages. The edited agent's stale signature block was removed;
Get-AuthenticodeSignature reports NotSigned. Existing published/cached releases and
live policy remain unchanged. Follow-up release gates: sign and timestamp with the
approved certificate, independently verify Valid signatures, merge sample rules into
the live policy without overwriting its other values, publish and refresh the pilot
cache only when authorized, then verify real outcomes. No commit or push occurred.
No reusable skill/plugin changed, so marketplace synchronization is not applicable.
Repo docs and Quick Memory apply; no script-execution FogBugz case or Vanta target
was identified. Case 255215 remains the separate DNS dependency/observation case.

### November 3 Latest Pilot Review And Handoff

Baseline: read-only collection at 14:50-14:53 Central returned seven of the current
17 members: DEVWKS01/03/05 and NEX-LAPTOP-01/03/05/06. All seven cached agents
match the October 27 published hash and have Valid signatures. No task, installer,
cache refresh or reboot was forced. Full fleet acceptance remains open.

- [x] Correlate latest Machine/user terminals, Health, recent attempt JSON, saved notification/restart/deferral state, task conditions and targeted registry versions. No latest package-failure event, malformed JSONL or active-run/process evidence was collected.
- [x] Supersede the old DEVWKS03 stale-cache concern: correct signed release, pointer verified November 2 and November 3 Machine completion with Azure CLI VerifiedNotOffered. Elevated SYSTEM-task XML/profile coverage is still limited by the filtered local token, not proof that tasks are missing.
- [x] Record current Machine outcomes: DEVWKS01 and laptop 01 Completed; DEVWKS03/05 CompletedRestartRequired; laptop 03 NoPackagesSelected; laptops 05/06 skipped outside the maintenance window and report AgentAssessmentOverdue.
- [x] Distinguish Storage Explorer's old hash failures from current state. DEVWKS05 now registers 1.47.0 and cleared its failure notice November 2; update actor unverified. LibreOffice has successful final MSI evidence with restart-to-finish, not a current download failure.
- [ ] Restore overnight assessment coverage on laptop 05 (last real assessment October 30, Eastern time) and laptop 06 (November 1, Central). Both have StartWhenAvailable=true and WakeToRun=false. Investigate power/availability and scheduler evidence; preserve the approved 02:00-04:00 endpoint-local window. A daytime successful skip is not an assessment.
- [ ] Coordinate separately approved restarts and verify activation: DEVWKS03 Java pending since October 30; DEVWKS05 PowerToys since October 29 and LibreOffice since October 31. Boot identities predate detection. Installed registry targets alone do not authorize clearing pending state.
- [ ] Diagnose Teams/native-updater coverage or propose a narrowly approved product rule with tests: the first-numeric-component major check defers build 26260 to 26276. Do not disable major guards globally.
- [ ] Diagnose NoApplicableInstaller in the affected context: DEVWKS03 Microsoft.Azd 1.34.200 to 1.37.300 cached until November 5; DEVWKS05 user Rustlang.Rustup 1.29.1 to 1.29.2 until November 8. Both exit 0x8A150010 before installer start; exact registration/applicability causes remain unproven.
- [ ] Reconcile Unknown user scope for Windows Terminal, WinDbg, Keeper Desktop and ChatGPT using authoritative package registration and deterministic fixtures. Keep the User-only guard; NoPackagesSelected is not full compliance.
- [ ] Reconcile ten unavailable members. NEX-LAPTOP-02/04 and SCM-LAPTOP2 have no A record after approved stale cleanup; seven other members fail WinRM. Their latest state is unknown. Verify reconnect identity/registration rather than recreating old wrong addresses or weakening Kerberos.
- [ ] Preserve bad-hash holds, evaluate generic integrity-failure classification/quarantine as a future tested change, and complete November 13 DNS observation in case 255215. Microsoft.WinGet.Client and outside-window read-only post-boot reconciliation remain future iterations.

The canonical detailed next-agent queue, full population, expected release hashes,
state paths, evidence limitations and authorization boundaries are in the SysOps
[latest pilot handoff](E:/Sandbox/SysOps-And-CyberSecurity/docs/runbooks/nexport-danmade-patch-agent-gpo-pilot-2026-06-23.md#november-3-latest-pilot-review-and-next-agent-handoff).
Raw receipts are local-only under
`E:\Sandbox\SysOps-And-CyberSecurity\.scratch\patch-agent-review-20261103-205036`;
they are not committed or attached and must be recollected elsewhere. Earlier
source-control baselines remain 0d5f725 (agent) and 2dac218 (SysOps). No production
source/signature/policy change or new full-suite test run occurred for this docs
handoff; the preceding same-day six-suite checkpoint passed 124 checks/scenarios.

### November 3 Progress And Source-Control Checkpoint

- [x] Reconcile the October 27 outcome/Health fixes with source and published bytes; retain NoApplicableInstaller administrator deferrals, bad-hash quarantine after review deadlines, Chrome-only major upgrades, SQL Developer ManualOnly and exact-target failure-notice transfer.
- [x] Rerun all six suites under Windows PowerShell 5.1: 38 function checks, 29 reliability checks, 29 fake-winget/Health scenarios, nine cache-dispatch checks, six pure ACL checks and 13 pilot regressions, 124 total. Eight script parser checks and sample-policy JSON validation also passed. Tests used local fixtures; no real installer, diagnostic SYSTEM task or endpoint restart was invoked.
- [x] Verify all five release PS1 signatures remain Valid, timestamped and signed by the approved domain publisher. Fresh authenticated readback on both DCs confirms the October 27 agent, Health and policy hashes unchanged.
- [x] Record selective DNS correction and aging observation in [case 255215](https://darwin-global.fogbugz.com/f/cases/255215/), Active and assigned to Daniel Pupek. Five confirmed obsolete VPN A records were removed; nexport.local aging is enabled with 3-day no-refresh and 7-day refresh. Both DNS servers retain automatic scavenging disabled. This shared-zone correction is not patch-agent rollout expansion.
- [ ] Review timestamp registration, static exclusions and unexplained old records on or after November 13 in case 255215. Future scavenging/additional deletions require separate approval; rapid VPN address reuse remains a separate registration/ownership concern.
- [ ] Investigate the current DEVWKS03 domain-based SYSVOL access limitation: Get-FileHash through nexport.local returns "The request is not supported" while both direct DC shares return the expected agent hash. The referral target could not be captured because dfsutil.exe is unavailable. Do not infer a server outage or flush/change DFS/network configuration from this observation alone.
- [ ] Complete elevated local task-XML and full 17-member acceptance, then evaluate actual next-window package outcomes and pending restart targets. The later 14:50 review confirms DEVWKS03's correct cache and Machine/User execution; the earlier cache concern is superseded. Seven remote Refresh/Health successes retain their October 27 evidence date. Ten members remain unverified in the latest read-only review.
- [ ] Review the four VS Code release-hold deadlines due November 3; unchanged bad hashes remain quarantined after the deadline. Azure Developer CLI applicability and future Microsoft.WinGet.Client evaluation remain separate work.

Current-checkpoint receipts are local-only under
`E:\Sandbox\SysOps-And-CyberSecurity\.scratch\progress-sync-20261103-201551`.
No new publication, forced patch task, cache refresh, policy change, package update
or reboot was performed for this source-control checkpoint. Source signing bytes
are preserved; documentation changes do not require re-signing release scripts.

### October 27 Pilot Outcome Corrections

Baseline: the October 26 pilot review found expired VS Code bad-hash holds,
no-installer applicability results presented as failed installations, SQL base
installer offers against a serviced instance, Chrome milestone deferrals and
Health confusing recent skips with recent assessments. Preserve the existing
17-member pilot, seven blocked IDs and endpoint-local 02:00-04:00 Machine window.

- [x] Distinguish native 0x8A150010 from identity mismatch/no-upgrade; retain bounded administrator deferrals with original attempt evidence and no installer retry/verification.
- [x] Keep unchanged known-bad hashes quarantined after their review deadline; allow a corrected hash immediately and report overdue review.
- [x] Add Chrome-only major-version permission and SQL Server 2025 Developer ManualOnly policy coverage.
- [x] Transfer only exact ID/source/target failure notices to administrator evidence for these deferrals; preserve unrelated failures and all restart notices.
- [x] Separate current/historical incomplete runs and assessment/skip timestamps; include selection deferrals and new/carried restart counts in terminal records.
- [x] Validate deterministic regressions without installing applications or creating diagnostic SYSTEM tasks.
- [x] Sign and publish; verify the protected release and actual SYSTEM Refresh/Health result 0 on DEVWKS01, DEVWKS05 and NEX-LAPTOP-01/02/03/05/06. Local DEVWKS03 refresh remains pending authorized elevated task execution; nine other members lack authenticated readback.
- [ ] Observe actual package outcomes in the next approved maintenance window; coordinate required restarts separately.
- [ ] Complete Azure Developer CLI native applicability diagnostics before targeted repair. Laptop 06 has a later successful User assessment on October 26, so its October 23 source error is historical; diagnose in real user context only if it recurs.
- [ ] Finish pilot DNS/asset reconciliation. The five approved stale VPN A records were corrected November 3 and 3 + 7 day zone aging is under observation in case 255215; remaining asset ownership and reconnect registration are not complete. Do not weaken Kerberos or broaden firewall exceptions.

Read-only post-boot reconciliation outside the installation window remains a
future iteration; this release does not clear saved restart targets on deployment.

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

November 6 read-only validation: seven of eight readable caches are on the latest
release, all eight signatures are Valid and both DC publication hashes still
match. WinDbg on DEVWKS03 and Terminal on laptop 02 have real verified upgrade
outcomes. App-restart handling, laptop 03 refresh/Health visibility, Zoom compound
version proof and nine unverified members remain acceptance gaps. No production
code changed or new release was published in this review.

Latest November 4, 17:30 Central: main-package correction is signed/timestamped and
published agent-only, SHA256 2FCCB1A6E7E5DDB1F673E4F0CA23C7374A11D6B67D0A1794FF9839978B476EB0,
verified on both DCs. All seven release hashes agree; all five PS1 signatures are
Valid/timestamped. Existing policy/ACLs/GPO bytes and 22/0 AD revision, plus the
17-member pilot, are preserved. Staged AllSigned WhatIf passed. Prior local-only
unsigned statements below are historical. Actual cache/User acceptance remains
pending natural refresh; no task/installer/reboot/GPO change or commit/push.

November 4 follow-up: main-only AppX correction is locally implemented and reviewed;
174 regression scenarios/checks pass, plus real current-user function-only proof.
Edited agent is NotSigned and unpublished; existing live/cached release remains
Valid. DEVWKS03 cache converged before this diagnosis and its current UNC paths,
DFS target and AllSigned WhatIf checks pass. Laptop 03 startup failure remains
open pending exact SYSTEM evidence. No refresh/patch task, GPO, trust, network,
cache or reboot mutation was performed by this work; no commit/push occurred.

November 4 read-only check: six new signed caches among seven readable members,
six overnight assessments and eight verification-backed successes; no new package
failure in the bounded sample. DEVWKS03 cache refresh and laptop 05 overdue coverage
remain unresolved. The Main+Bundle scope defect is reproduced with live registration
and production function-only reads on laptop 01; a code fix/test/release is pending.
No new reboot validates post-reboot target proof, and ten members lack readback.

Latest November 3 authorized publication: signed agent and merged live policy
match the new hashes above on both DCs. All five release PS1 files are Valid and
timestamped, all seven release file hashes agree, GPO versions remain 22/0 and
the group remains 17 members. Cache installer AllSigned WhatIf passed. Published
files retain their original ACLs. Real cache convergence and package/reboot outcomes
remain pending natural scheduled execution. Earlier unsigned statements below are
historical local validation, not current source-signature state. No commit/push.

Subsequent November 3 local implementation: all seven Windows PowerShell 5.1 suites
passed, 166 checks/scenarios total. Four changed PS1 files parse, sample JSON loads,
all four changed Markdown files pass lint, and git diff --check is clean in both
workspaces. The final regression transcript is source-local
`.scratch\verification-coverage-20261103-160032\regressions.txt`.
Edited source agent SHA256 is
0ED3A30E5D01668465D3AD0AE3BADB903C90EE52DD0B13A42A013D0A4CFBF573,
NotSigned; sample policy SHA256 is
1160FCA1226D0C6B8323CA9F62655C7C90C930843D30A14D909D3CE8E0441C46.
These are not deployed release hashes. Signing, policy reconciliation, publication,
real SYSTEM execution and next-window/reboot acceptance remain pending. No live
mutation, commit or push occurred. The earlier review/publication evidence below
describes the unchanged deployed release, not these edited local bytes.

Latest November 3 review: seven of 17 members returned readable Machine/user logs
and current cached-agent hashes/Valid signatures; five Machine lanes assessed
November 3 and two are overdue after daytime skips. No latest package-failure
events or active-run/process evidence were collected. Two endpoints still need
restart verification, and ten members lack current readback. DEVWKS03's previous
cache concern is superseded; elevated task-XML acceptance remains distinct. No
code, signing, policy or endpoint change occurred for this handoff.

November 3: all six local suites passed again, 124 checks/scenarios total; eight
script parser checks and sample-policy JSON validation passed. All five release
signatures remain Valid and timestamped. Fresh direct-DC/guest readback confirms
the agent, Health and live policy hashes match the October 27 publication below.
The domain-alias SYSVOL read failed in this operator session, despite successful
direct DC reads; no domain-path availability or fresh endpoint acceptance is
claimed. DNS change case 255215 remains Active for the November 13 observation
review, with server automatic scavenging disabled.

October 27: all six tracked suites passed under Windows PowerShell 5.1: 38 function,
29 reliability, 29 fake-winget/Health, nine cache-dispatch, six pure ACL and 13
pilot regression checks/scenarios (124 total). All five release PS1 signatures are
Valid and timestamped with the approved domain signer. Published agent SHA256
`14687A9AABA0F424FB8C00C407D41ACB909AEB8C3B314A51529E17AF98228D9C`
and policy SHA256
`5D6FA4DF387B10296B1D9A7E6DFDFF273BFC9CCCBCFB35E3492EC5813B2B8FE1`
match the domain path, both DC replicas and all seven refreshed endpoints. GPO
revision 22 and all 17 members are unchanged; no task payload change was needed.
The four exact VS Code hold review dates are November 3; unchanged bad hashes stay
held after review is overdue. Machine/User/Notification tasks and restarts were
not forced. Next-window package behavior remains unproven by this deployment.
Local publication, rollback, endpoint and remoting receipts are under
`E:\Sandbox\SysOps-And-CyberSecurity\.scratch\patch-agent-fixes-20261027`.

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
