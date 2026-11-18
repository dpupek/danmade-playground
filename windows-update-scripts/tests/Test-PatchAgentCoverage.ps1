<#
.SYNOPSIS
Tests installed-target proof, current-user package identity and coverage dispositions.
.DESCRIPTION
Loads production functions without running the agent. Registry, AppX and boot observations
are synthetic. Only this run's .scratch directory is written; no installers or tasks run.
.EXAMPLE
powershell.exe -NoProfile -ExecutionPolicy Bypass -File windows-update-scripts/tests/Test-PatchAgentCoverage.ps1
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scratch = Join-Path $repo ('.scratch\coverage-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'windows-update-scripts\danmade-patch-agent.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Production agent does not parse.' }
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Parent.Parent -eq $ast }, $true)) { Invoke-Expression $node.Extent.Text }
function Assert-Equal { param($Actual, $Expected, [string]$Message) if ($Actual -ne $Expected) { throw "$Message Expected=$Expected Actual=$Actual" } }
function Get-CachedUninstallEntries { $script:Entries }
function Test-IsSystemAccount { $script:IsSystem }
function Get-AppxPackage {
  param($PackageTypeFilter, $ErrorAction)
  $script:AppxCalls++
  Assert-Equal ($PackageTypeFilter -join ',') 'Main' 'Installed application evidence must query main packages only.'
  Assert-Equal $args.Count 0 'Current-user query must not add AllUsers or another User argument.'
  if ($script:AppxFails) { throw 'Synthetic AppX query failure.' }
  $script:Appx
}
function Get-SystemBootTime { [datetime]'2026-11-03' }
function Write-AgentEvent { param($EventId, $EntryType, $Fields) $script:Events.Add($Fields) }
function Move-FailureNotificationToAdminDeferral { param($Package, $Reason) }
function Get-RunningRuleProcesses { param($Rule) @([pscustomobject]@{ProcessName='test'}) }
function New-Registration {
  param([string]$Scope, [string]$Version = '1.1.0')
  [pscustomobject]@{ DisplayName='Test App'; DisplayNameNorm=(Normalize-DisplayText 'Test App'); DisplayVersion=$Version; PSPath="Registry::HKEY_${Scope}\Software\Uninstall\Test"; InstallLocation=''; UninstallString='' }
}
$Mode = 'Machine'; $LogRoot = $scratch; $script:IsSystem = $false
$script:Policy = New-DefaultPolicy
$script:Entries = @(); $script:Appx = @(); $script:AppxFails = $false; $script:AppxCalls = 0
$script:CachedUserAppxPackages = $null; $script:ApplicabilityDeferredCount = 0
$script:InapplicableUpgradeStatePath = Join-Path $scratch 'State\inapplicable-upgrades.json'
$script:Events = New-Object 'System.Collections.Generic.List[object]'
$count = 0

# Arrange: absent, ambiguous and wrong-scope inventory must never resolve a target.
$record = [pscustomobject]@{id='Test.App'; name='Test App'; available='1.1.0'; source='winget'; bootTimeAtDetection='2026-11-01'; evidencePath='fixture'}
$offer = [pscustomobject]@{Id='Test.App'; Name='Test App'; Installed='1.1.0'; Available='1.2.0'; Source='winget'; InstallScope='Machine'}
foreach ($case in @(
  @{Name='absent'; Entries=@(); Offers=@(); Expected=$false},
  @{Name='exact machine ARP'; Entries=@(New-Registration LOCAL_MACHINE); Offers=@(); Expected=$true},
  @{Name='wrong scope ARP'; Entries=@(New-Registration CURRENT_USER); Offers=@(); Expected=$false},
  @{Name='older ARP'; Entries=@(New-Registration LOCAL_MACHINE '1.0.0'); Offers=@(); Expected=$false},
  @{Name='unknown ARP version'; Entries=@(New-Registration LOCAL_MACHINE 'Unknown'); Offers=@(); Expected=$false},
  @{Name='ambiguous ARP'; Entries=@((New-Registration LOCAL_MACHINE),(New-Registration CURRENT_USER)); Offers=@(); Expected=$false},
  @{Name='scoped newer offer'; Entries=@(); Offers=@($offer); Expected=$true},
  @{Name='duplicate offers'; Entries=@(); Offers=@($offer,$offer); Expected=$false},
  @{Name='wrong scope offer'; Entries=@(); Offers=@([pscustomobject]@{Id='Test.App';Installed='1.1.0';Source='winget';InstallScope='User'}); Expected=$false},
  @{Name='wrong source offer'; Entries=@(); Offers=@([pscustomobject]@{Id='Test.App';Installed='1.1.0';Source='other';InstallScope='Machine'}); Expected=$false}
)) {
  $script:Entries = $case.Entries
  # Initial Assert.
  Assert-Equal $record.available '1.1.0' 'Target fixture changed.'
  # Act.
  $satisfied = Test-RecordedTargetInstalled -Package $record -AvailablePackages $case.Offers
  # Assert.
  Assert-Equal $satisfied $case.Expected "Installed target proof: $($case.Name)."
  $count++
}

# Arrange: a later boot with no installed proof must preserve pending state.
$script:Entries = @()
$statePath = Get-RestartVerificationStatePath
New-Item -ItemType Directory -Path (Split-Path $statePath) -Force | Out-Null
foreach ($case in @(
  @{Name='absent after reboot'; Boot='2026-11-01'; Entries=@(); Offers=@(); Pending=1; Failed=0; Resolved=0; Status='RestartVerificationInconclusive'},
  @{Name='installed before reboot'; Boot='2026-11-04'; Entries=@(New-Registration LOCAL_MACHINE); Offers=@(); Pending=1; Failed=0; Resolved=0; Status='DeferredRestartVerificationPendingReboot'},
  @{Name='installed after reboot'; Boot='2026-11-01'; Entries=@(New-Registration LOCAL_MACHINE); Offers=@(); Pending=0; Failed=0; Resolved=1; Status='RestartVerificationResolved'},
  @{Name='still older after reboot'; Boot='2026-11-01'; Entries=@(); Offers=@([pscustomobject]@{Id='Test.App';Name='Test App';Installed='1.0.0';Available='1.1.0';Source='winget';InstallScope='Machine'}); Pending=0; Failed=1; Resolved=0; Status='RestartVerificationFailedStillOffered'}
)) {
  $script:Entries = $case.Entries; $script:Events.Clear()
  $record.bootTimeAtDetection = $case.Boot
  @{packages=@($record)} | ConvertTo-Json -Depth 5 | Set-Content $statePath
  # Initial Assert.
  Assert-Equal (Test-Path $statePath) $true 'Restart state must exist before reconciliation.'
  # Act.
  $result = Resolve-PendingRestartVerification -AvailablePackages $case.Offers
  # Assert.
  Assert-Equal $result.PendingRestartCount $case.Pending "$($case.Name) pending count."
  Assert-Equal @($result.FailedPackages).Count $case.Failed "$($case.Name) failure count."
  Assert-Equal @($result.ResolvedPackageIds).Count $case.Resolved "$($case.Name) resolved count."
  Assert-Equal $script:Events[0].status $case.Status "$($case.Name) evidence."
  Assert-Equal (Test-Path $statePath) ($case.Pending -gt 0) "$($case.Name) retained state."
  $count++
}

# Arrange: only a policy-pinned, healthy, unique current-user identity proves AppX scope.
$Mode = 'User'; $script:Entries = @()
$family = 'Test.App_8wekyb3d8bbwe'
$rule = [pscustomobject]@{id='Test.App';mode='Auto';appxPackageFamilyNames=@($family)}
$script:Policy.packageRules = @(Convert-ToPackageRules @($rule))
$registered = [pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false}
$bundle = [pscustomobject]@{PackageFamilyName=$family;Version='3001.1.0.0';Status=0;IsBundle=$true;IsFramework=$false;IsResourcePackage=$false}
foreach ($case in @(
  @{Name='healthy'; Apps=@($registered); Fails=$false; System=$false; Expected='User'},
  @{Name='wrong publisher'; Apps=@([pscustomobject]@{PackageFamilyName='Test.App_aaaaaaaaaaaaa';Version='1.1.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='duplicate registrations'; Apps=@($registered,$registered); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='main plus bundle'; Apps=@($registered,$bundle); Fails=$false; System=$false; Expected='User'},
  @{Name='bundle before main'; Apps=@($bundle,$registered); Fails=$false; System=$false; Expected='User'},
  @{Name='bundle only'; Apps=@($bundle); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='two distinct main versions'; Apps=@($registered,[pscustomobject]@{PackageFamilyName=$family;Version='1.2.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='unknown package type'; Apps=@([pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=0;IsFramework=$false;IsResourcePackage=$false}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='unhealthy'; Apps=@([pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=1;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='framework'; Apps=@([pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=0;IsBundle=$false;IsFramework=$true;IsResourcePackage=$false}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='resource'; Apps=@([pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$true}); Fails=$false; System=$false; Expected='Unknown'},
  @{Name='query failed'; Apps=@($registered); Fails=$true; System=$false; Expected='Unknown'},
  @{Name='SYSTEM'; Apps=@($registered); Fails=$false; System=$true; Expected='Unknown'}
)) {
  $script:Appx=$case.Apps; $script:AppxFails=$case.Fails; $script:IsSystem=$case.System; $script:CachedUserAppxPackages=$null; $script:AppxCalls=0
  # Initial Assert.
  Assert-Equal $script:AppxCalls 0 'AppX query must not run during fixture setup.'
  # Act.
  $scope = Resolve-PackageInstallScope -Package $offer
  # Assert.
  Assert-Equal $scope $case.Expected "AppX scope: $($case.Name)."
  Assert-Equal (Test-RecordedTargetInstalled $record @()) ($case.Expected -eq 'User') "AppX installed target: $($case.Name)."
  Assert-Equal $script:AppxCalls $(if ($case.System) {0} else {1}) 'AppX lookup must cache success and empty/failure results.'
  $count++
}

# Arrange: an exact family may identify scope without proving the recorded version.
$Mode='User'; $script:IsSystem=$false; $script:Entries=@(); $script:AppxFails=$false; $script:CachedUserAppxPackages=$null
$script:Appx=@([pscustomobject]@{PackageFamilyName=$family;Version='1.0.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false})
Assert-Equal $script:Appx[0].Version '1.0.0' 'Registration must be below the recorded target.'
# Act.
$proof=Test-RecordedTargetInstalled $record @()
# Assert.
Assert-Equal $proof $false 'Pinned family is identity proof, not target-version proof.'
$count++

# Arrange: reproduce Terminal's actual main/bundle version shape without real AppX access.
$script:Appx=@(
  [pscustomobject]@{PackageFamilyName=$family;Version='3001.24.12741.0';Status=0;IsBundle=$true;IsFramework=$false;IsResourcePackage=$false},
  [pscustomobject]@{PackageFamilyName=$family;Version='1.24.12741.0';Status=0;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false}
)
foreach ($case in @(
  @{Target='1.25.2900.0';Expected=$false},
  @{Target='1.24.12741.0';Expected=$true}
)) {
  $script:CachedUserAppxPackages=$null
  $targetRecord=[pscustomobject]@{id='Test.App';name='Test App';available=$case.Target;source='winget'}
  # Initial Assert.
  Assert-Equal ([version]$script:Appx[0].Version -gt [version]$case.Target) $true 'Bundle version must look newer than either target.'
  Assert-Equal ([version]$script:Appx[1].Version -ge [version]$case.Target) $case.Expected 'Main package must independently establish the expected target proof.'
  # Act.
  $proof=Test-RecordedTargetInstalled $targetRecord @()
  # Assert.
  Assert-Equal $proof $case.Expected 'Only main version may establish installed-target proof.'
  Assert-Equal @($script:CachedUserAppxPackages).Count 1 'Bundle must not participate in retained package identity.'
  Assert-Equal $script:CachedUserAppxPackages[0].IsBundle $false 'Cached evidence must be a main package.'
  $count++
}

# Arrange: a healthy bundle cannot hide a failed main registration.
$script:Appx=@($bundle,[pscustomobject]@{PackageFamilyName=$family;Version='1.1.0';Status=1;IsBundle=$false;IsFramework=$false;IsResourcePackage=$false})
$script:CachedUserAppxPackages=$null
Assert-Equal $script:Appx[0].IsBundle $true 'Healthy packaging fixture must be a bundle.'
# Act.
$scope=Resolve-PackageInstallScope $offer
# Assert.
Assert-Equal $scope 'Unknown' 'Unhealthy main package cannot be replaced by bundle identity proof.'
$count++

# Arrange: a healthy registration cannot supply an unconfigured identity mapping.
$script:Policy.packageRules=@(); $script:AppxCalls=0; $script:CachedUserAppxPackages=$null
Assert-Equal @($script:Policy.packageRules).Count 0 'Policy must have no family mapping.'
# Act.
$scope=Resolve-PackageInstallScope $offer
# Assert.
Assert-Equal $scope 'Unknown' 'Unconfigured families must not be guessed by display name.'
Assert-Equal $script:AppxCalls 0 'Unconfigured families do not need an AppX query.'
$script:Policy.packageRules=@(Convert-ToPackageRules @($rule))
$count++

# Arrange: machine mode must not use the operator's current-user registration.
$Mode='Machine'; $script:IsSystem=$false; $script:AppxCalls=0; $script:Appx=@($registered); $script:CachedUserAppxPackages=$null
Assert-Equal @($script:Entries).Count 0 'Machine fixture has no registry proof.'
# Act.
$scope = Resolve-PackageInstallScope $offer
# Assert.
Assert-Equal $scope 'Unknown' 'Current-user registration cannot prove machine scope.'
Assert-Equal $script:AppxCalls 0 'Machine mode must not query current-user AppX.'
$count++

# Arrange: tied registry matches must be order independent, even with a pinned AppX rule.
$Mode='User'; $script:CachedUserAppxPackages=$null; $script:Appx=@($registered)
foreach ($entries in @(@((New-Registration LOCAL_MACHINE),(New-Registration CURRENT_USER)), @((New-Registration CURRENT_USER),(New-Registration LOCAL_MACHINE)))) {
  $script:Entries=$entries
  Assert-Equal $script:Entries.Count 2 'Ambiguous registry fixture requires both registrations.'
  # Act.
  $scope = Resolve-PackageInstallScope $offer
  # Assert.
  Assert-Equal $scope 'Unknown' 'Tied ARP matches must not choose enumeration order.'
  $count++
}

# Arrange: unrelated ARP entries must not defeat an exact configured AppX identity.
$script:Entries=@([pscustomobject]@{DisplayName='Unrelated Product';DisplayNameNorm='unrelated product';PSPath='Registry::HKEY_LOCAL_MACHINE\Uninstall\Other';InstallLocation='';UninstallString='';DisplayVersion='9.0.0'})
$script:CachedUserAppxPackages=$null; $script:Appx=@($registered)
Assert-Equal $script:Entries[0].DisplayName 'Unrelated Product' 'ARP fixture must be unrelated.'
# Act.
$scope=Resolve-PackageInstallScope $offer
# Assert.
Assert-Equal $scope 'User' 'An unmatched ARP entry must not be treated as a best match.'
$count++

# Arrange: simultaneous machine ARP and user AppX ownership is ambiguous.
$script:Entries=@(New-Registration LOCAL_MACHINE)
Assert-Equal $script:Entries.Count 1 'Conflict requires a machine registration.'
# Act.
$scope=Resolve-PackageInstallScope $offer
# Assert.
Assert-Equal $scope 'Unknown' 'Do not silently choose between machine ARP and user AppX.'
$count++

# Arrange: policy requires full literal family identities, not wildcard/name matching.
foreach ($identity in @('Test.App_*', 'Test.App', '*_8wekyb3d8bbwe')) {
  $invalidRule = [pscustomobject]@{id='Test.App';mode='Auto';appxPackageFamilyNames=@($identity)}
  Assert-Equal $invalidRule.appxPackageFamilyNames[0] $identity 'Invalid identity fixture changed.'
  # Act.
  $rejected=$false
  try { $null=Convert-ToPackageRules @($invalidRule) } catch { $rejected=$true }
  # Assert.
  Assert-Equal $rejected $true 'Wildcard/incomplete family must be rejected.'
  $count++
}

# Arrange: mixed selection dispositions retain completed-assessment semantics without compliance claims.
$script:Policy=New-DefaultPolicy; $script:Policy.allowedPackageIds=@('Test.Manual','Test.Native','Test.Major','Test.Busy','Test.Unknown','Test.WrongScope','Test.Restart','Test.Blocked','Test.Selected')
$script:Policy.blockedPackageIds=@('Test.Blocked')
$script:Policy.packageRules=@(Convert-ToPackageRules @(
  [pscustomobject]@{id='Test.Manual';mode='ManualOnly'},
  [pscustomobject]@{id='Test.Native';mode='NativeUpdater'},
  [pscustomobject]@{id='Test.Busy';mode='QuiescenceRequired';processNames=@('test.exe')}
))
$packages=@(foreach ($name in @('Manual','Native','Major','Busy','Unknown','WrongScope','Restart','Blocked','NotAllowed','Selected')) {
  [pscustomobject]@{Id="Test.$name";Name=$name;Installed='1.0.0';Available=$(if ($name -eq 'Major') {'2.0.0'} else {'1.1.0'});Source='winget';InstallScope=$(if ($name -eq 'Unknown') {'Unknown'} elseif ($name -eq 'WrongScope') {'Machine'} else {'User'})}
})
Assert-Equal $packages.Count 10 'Coverage fixture must account for ten offers.'
# Act.
$selected=@(Select-PolicyPackages $packages @('Test.Restart'))
# Assert.
Assert-Equal $selected.Count 1 'Exactly one package must remain eligible.'
foreach ($key in @('manualOnlyCount','nativeUpdaterCount','majorVersionDeferredCount','processInUseCount','scopeUnknownCount','scopeMismatchCount','restartDeferredCount','blockedCount','notAllowedCount')) { Assert-Equal $script:SelectionCounts[$key] 1 "$key must count exactly one disposition." }
Assert-Equal $script:SelectionDeferredCount 5 'Existing deferred semantics must be retained.'
$count++

foreach ($case in @(
  @{Discovered=0;Selected=0;Failed=0;Deferred=0;Carried=0;Outcome='NoUpdatesOffered';Coverage='NoUnresolvedOffers'},
  @{Discovered=1;Selected=0;Failed=0;Deferred=1;Carried=0;Outcome='NoEligibleUpdates';Coverage='UnresolvedOffers'},
  @{Discovered=1;Selected=1;Failed=0;Deferred=0;Carried=0;Outcome='UpdatesProcessed';Coverage='NoUnresolvedOffers'},
  @{Discovered=1;Selected=1;Failed=1;Deferred=0;Carried=0;Outcome='UpdatesProcessed';Coverage='UnresolvedOffers'},
  @{Discovered=0;Selected=0;Failed=0;Deferred=0;Carried=1;Outcome='NoUpdatesOffered';Coverage='UnresolvedOffers'}
)) {
  # Arrange / Initial Assert.
  foreach ($key in @($script:SelectionCounts.Keys)) { $script:SelectionCounts[$key]=0 }
  if ($case.Discovered -eq 1 -and $case.Selected -eq 0) { $script:SelectionCounts.manualOnlyCount=1 }
  $summary=@{Succeeded=$case.Selected-$case.Failed;Failed=$case.Failed;Deferred=$case.Deferred;RestartRequired=0}
  Assert-Equal $summary.RestartRequired 0 'No new restart was requested in summary fixture.'
  # Act.
  $fields=Get-AssessmentSummaryFields -DiscoveredCount $case.Discovered -SelectedCount $case.Selected -Summary $summary -CarriedRestartCount $case.Carried
  # Assert.
  Assert-Equal $fields.assessmentCompleted $true 'Assessment completion must not be redefined as full coverage.'
  Assert-Equal $fields.assessmentOutcome $case.Outcome 'Assessment disposition mismatch.'
  Assert-Equal $fields.coverageStatus $case.Coverage 'Coverage claim mismatch.'
  Assert-Equal $fields.pendingRestartCount $case.Carried 'Carried restart count lost.'
  Assert-Equal $fields.manualOnlyCount $(if ($case.Discovered -eq 1 -and $case.Selected -eq 0) {1} else {0}) 'Selection counts must belong to the current assessment.'
  $count++
}

# Arrange: the sample policy delegates Teams without relaxing Node/SQL major-version safeguards.
$sample=Join-Path $repo 'windows-update-scripts\danmade-patch-agent.policy.sample.json'
Assert-Equal (Test-Path $sample) $true 'Sample policy must exist.'
# Act.
$PolicyPath=$sample
$loaded=Import-AgentPolicy
# Assert.
foreach ($id in @('Microsoft.Teams','Microsoft.Teams.Free')) { Assert-Equal @($loaded.packageRules | Where-Object id -eq $id)[0].mode 'NativeUpdater' 'Teams must use product-specific updater delegation.' }
foreach ($id in @('OpenJS.NodeJS','OpenJS.NodeJS.LTS','Microsoft.SQLServer.2025.Developer')) { Assert-Equal @($loaded.packageRules | Where-Object id -eq $id)[0].mode 'ManualOnly' 'Existing product safeguards changed.' }
$count++

# Arrange: reject a scalar identity field at the actual policy-loading boundary.
$badPolicy=Get-Content $sample -Raw | ConvertFrom-Json
$badPolicy.packageRules[0].appxPackageFamilyNames='MSTeams_8wekyb3d8bbwe'
$PolicyPath=Join-Path $scratch 'scalar-family-policy.json'
$badPolicy | ConvertTo-Json -Depth 8 | Set-Content $PolicyPath
Assert-Equal (Test-Path $PolicyPath) $true 'Invalid policy fixture must exist.'
# Act.
$rejected=$false
try { $null=Import-AgentPolicy } catch { $rejected=$true }
# Assert.
Assert-Equal $rejected $true 'Scalar family policy must fail closed.'
$count++
foreach ($case in @(
  @{Installed='7.2.5 (51704)';Scope='CURRENT_USER';Expected=$true},
  @{Installed='7.2.5.51704';Scope='CURRENT_USER';Expected=$true},
  @{Installed='7.2.5 (51703)';Scope='CURRENT_USER';Expected=$false},
  @{Installed='7.2.5';Scope='CURRENT_USER';Expected=$false},
  @{Installed='7.2.5 (51704) extra';Scope='CURRENT_USER';Expected=$false},
  @{Installed='7.2.5 (51704)';Scope='LOCAL_MACHINE';Expected=$false}
)) {
  # Arrange / Initial Assert.
  $zoom = [pscustomobject]@{id='Zoom.Zoom.EXE';name='Test App';source='winget';available='7.2.5 (51704)'}
  $script:Entries = @(New-Registration $case.Scope $case.Installed)
  Assert-Equal $script:Entries[0].DisplayVersion $case.Installed 'Raw installed evidence must be retained.'
  # Act.
  $proved = Test-RecordedTargetInstalled $zoom @()
  # Assert.
  Assert-Equal $proved $case.Expected "Zoom target proof: $($case.Installed) / $($case.Scope)."
  $count++
}

# Arrange: a real reboot occurred, but installed target evidence is inconclusive.
$zoom | Add-Member NoteProperty bootTimeAtDetection '2026-11-01'
$zoom | Add-Member NoteProperty evidencePath 'saved-zoom-attempt.json'
$script:Entries = @()
@{packages=@($zoom)} | ConvertTo-Json -Depth 5 | Set-Content $statePath
$noticePath = Get-NotificationStatePath
$other = [pscustomobject]@{id='Test.Unrelated';source='winget';available='2.0.0'}
@{status='CompletedRestartRequired';failedPackages=@();restartRequiredPackages=@($zoom,$other);summary=@{failed=0;restartRequired=2}} | ConvertTo-Json -Depth 5 | Set-Content $noticePath
$review = Resolve-PendingRestartVerification @()
Assert-Equal $review.ReviewPackages.Count 1 'Zoom should require post-reboot review.'
# Act.
Resolve-RecordedPackageState -AvailablePackages @() -ReviewRestartPackages $review.ReviewPackages
# Assert.
$notice = Get-Content -LiteralPath $noticePath -Raw | ConvertFrom-Json
Assert-Equal $notice.restartRequiredPackages.Count 1 'Only the exact post-reboot notice should be removed.'
Assert-Equal $notice.restartRequiredPackages[0].id $other.id 'Unrelated reboot notice must be retained.'
Assert-Equal (Test-Path $statePath) $true 'Admin evidence must remain pending.'
$count++

# Arrange / Initial Assert: reproduce laptop 02's exact registration after reboot.
$script:Entries = @(New-Registration CURRENT_USER '7.2.5 (51704)')
$script:Events.Clear()
Assert-Equal (Test-Path $statePath) $true 'Previously inconclusive Zoom record must still exist.'
# Act.
$zoomRecheck = Resolve-PendingRestartVerification @()
# Assert.
Assert-Equal ($zoomRecheck.ResolvedPackageIds -join ',') 'Zoom.Zoom.EXE' 'The exact build should resolve the record.'
Assert-Equal (Test-Path $statePath) $false 'Proven Zoom target should clear pending verification.'
Assert-Equal $script:Events[0].status 'RestartVerificationResolved' 'Zoom resolution evidence is missing.'
$count++

$script:Policy = New-DefaultPolicy
$script:Policy.packageRules = @(Convert-ToPackageRules @([pscustomobject]@{id='Test.App';mode='Auto';appxPackageFamilyNames=@($family)}))
$script:IsSystem = $false
foreach ($case in @(
  @{Name='current Main';Version='1.1.0';Bundle=$false;Days=1;Count=0;Status='ApplicationRestartVerificationResolved'},
  @{Name='stale Main';Version='1.0.0';Bundle=$false;Days=1;Count=1;Status=''},
  @{Name='bundle only';Version='9.1.0';Bundle=$true;Days=1;Count=1;Status=''},
  @{Name='expired';Version='1.0.0';Bundle=$false;Days=-1;Count=0;Status='ApplicationRestartVerificationExpired'}
)) {
  # Arrange / Initial Assert.
  $script:CachedUserAppxPackages = $null
  $script:AppxFails = $false
  $script:Appx = @([pscustomobject]@{Status=0;IsFramework=$false;IsResourcePackage=$false;IsBundle=$case.Bundle;PackageFamilyName=$family;Version=$case.Version})
  $script:Events.Clear()
  $pendingPath = Get-ApplicationRestartStatePath
  $app = [pscustomobject]@{id='Test.App';name='Test App';source='winget';available='1.1.0';evidencePath='fixture';runId='original';expiresAt=(Get-Date).AddDays($case.Days).ToString('o')}
  @{packages=@($app)} | ConvertTo-Json -Depth 5 | Set-Content $pendingPath
  Assert-Equal @(Get-ApplicationRestartState).Count 1 'Pending application evidence must exist.'
  # Act: no offer is not by itself installed proof.
  $left = @(Resolve-PendingApplicationRestart @())
  # Assert.
  Assert-Equal $left.Count $case.Count $case.Name
  Assert-Equal (Test-Path $pendingPath) ($case.Count -gt 0) 'Activation state persistence mismatch.'
  Assert-Equal (($script:Events | ForEach-Object { $_.status }) -join ',') $case.Status 'Activation evidence mismatch.'
  $count++
}

# Arrange / Initial Assert: one current activation in a disposable user profile state directory.
$app.expiresAt = (Get-Date).AddDays(1).ToString('o')
@{packages=@($app)} | ConvertTo-Json -Depth 5 | Set-Content $pendingPath
Assert-Equal (Test-Path $pendingPath) $true 'Notification fixture is missing.'
# Act: preview the actual notification without showing a popup or acknowledging it.
$notifier = Join-Path $repo 'windows-update-scripts\Show-DanmadePatchAgentNotification.ps1'
$preview = @(& $notifier -MachineStatePath (Join-Path $scratch 'absent.json') -UserStatePath $noticePath -AckRoot (Join-Path $scratch 'acks') -WhatIf)
# Assert.
$applicationPreview = @($preview | Where-Object { $_.PSObject.Properties['Title'] -and $_.Title -eq 'Reopen an application to finish updating' })
Assert-Equal $applicationPreview.Count 1 'Application notice missing.'
Assert-Equal ($applicationPreview[0].Message -match 'close and reopen') $true 'Notice must request an application reopen.'
Assert-Equal (Test-Path (Join-Path $scratch 'acks')) $false 'Preview must not acknowledge.'
$count++

foreach ($case in @(
  @{Name='failed-and-aging';AgeHours=125;State='Ready';Result=4294770688;Failed=1;Aging=1;Expired=0},
  @{Name='running-with-old-result';AgeHours=2;State='Running';Result=4294770688;Failed=0;Aging=0;Expired=0},
  @{Name='healthy';AgeHours=1;State='Ready';Result=0;Failed=0;Aging=0;Expired=0},
  @{Name='expired';AgeHours=169;State='Ready';Result=0;Failed=0;Aging=0;Expired=1},
  @{Name='not-yet-run';AgeHours=1;State='Ready';Result=267011;Failed=0;Aging=0;Expired=0}
)) {
  # Arrange / Initial Assert: the cached Health script uses a synthetic parent pointer and tasks.
  $root = Join-Path $scratch $case.Name
  $release = Join-Path $root 'release'
  New-Item -ItemType Directory -Path $release -Force | Out-Null
  $health = Join-Path $release 'Test-DanmadePatchAgentHealth.ps1'
  Copy-Item -LiteralPath (Join-Path $repo "windows-update-scripts\Test-DanmadePatchAgentHealth.ps1") -Destination $health
  @{verifiedAt=(Get-Date).AddHours(-$case.AgeHours).ToString('o')} | ConvertTo-Json | Set-Content (Join-Path $root 'current.json')
  $taskPath = Join-Path $root 'tasks.json'
  $processPath = Join-Path $root 'processes.json'
  @{Name='Danmade Patch Agent - Refresh Local Cache';State=$case.State;LastTaskResult=$case.Result;LastRunTime=(Get-Date).ToString('o')} | ConvertTo-Json | Set-Content $taskPath
  '[]' | Set-Content $processPath
  $healthLog = Join-Path $root 'Events\patch-agent.jsonl'
  Assert-Equal (Test-Path $healthLog) $false 'Health fixture must start without log history.'
  # Act.
  & $health -LogRoot $root -ProcessSnapshotPath $processPath -TaskSnapshotPath $taskPath
  # Assert.
  $events = @(Get-Content -LiteralPath $healthLog | ForEach-Object { $_ | ConvertFrom-Json })
  Assert-Equal @($events | Where-Object status -eq 'CacheRefreshTaskFailed').Count $case.Failed "$($case.Name): completed refresh failure."
  Assert-Equal @($events | Where-Object status -eq 'LocalPolicyVerificationAging').Count $case.Aging "$($case.Name): early aging warning."
  Assert-Equal @($events | Where-Object status -eq 'LocalPolicyRefreshRequired').Count $case.Expired "$($case.Name): expired cache warning."
  if ($case.Failed) {
    $failure = $events | Where-Object status -eq 'CacheRefreshTaskFailed'
    Assert-Equal $failure.lastTaskResult $case.Result 'Preserve scheduler result.'
    Assert-Equal $failure.response 'ReportOnlyNoTermination' 'Health may only report.'
  }
  $count++
}
Write-Output "Passed $count coverage and installed-target regression scenarios. Artifacts: $scratch"
