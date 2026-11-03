<# Deterministic October pilot regressions. No real installer, task or endpoint mutation. #>
[CmdletBinding(SupportsShouldProcess=$true)]param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scratch=Join-Path $repo ('.scratch\pilot-fixes-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $scratch
function Import-Functions([string]$Path){
  $tokens=$null;$errors=$null
  $ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
  if($errors.Count){throw 'Source must parse before test execution.'}
  foreach($node in $ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Parent.Parent -eq $ast},$true)){
    Set-Item -Path ('Function:script:'+$node.Name) -Value ([scriptblock]::Create($node.Body.Extent.Text.TrimStart('{').TrimEnd('}')))
  }
}
function Assert-Equal($Actual,$Expected,[string]$Message){if($Actual -ne $Expected){throw "$Message Expected=$Expected Actual=$Actual"}}
Import-Functions (Join-Path $repo 'windows-update-scripts\danmade-patch-agent.ps1')
$count=0
foreach($case in @(
  @{Code='8A150010';Text='No applicable installer found; see logs for more details.';Status='NoApplicableInstaller'},
  @{Code='00000000';Text='No applicable installer found; see logs for more details.';Status='NoApplicableInstaller'},
  @{Code='8A150014';Text='No installed package found matching input criteria.';Status='InstalledPackageNotFound'},
  @{Code='8A15002B';Text='No applicable upgrade found.';Status='NoApplicableUpgrade'},
  @{Code='8A150010';Text='Starting package install';Status='WingetFailed'}
)){
  # Arrange / initial assert.
  $inputResult=[pscustomobject]@{ExitCode=[Convert]::ToInt32($case.Code,16);TimedOut=$false;StdOut=$case.Text;StdErr=''}
  Assert-Equal ([string]::IsNullOrWhiteSpace($inputResult.StdOut)) $false 'Captured output must exist.'
  # Act.
  $result=Get-WingetCommandOutcome $inputResult
  # Assert.
  Assert-Equal $result.Status $case.Status 'Applicability codes must stay distinct.'
  Assert-Equal $result.IsRetryable $false 'Applicability is not a download retry.'
  $count++
}

# Arrange: Chrome-only major permission must not relax developer runtime rules.
$Mode='Machine';$LogRoot=$scratch;$script:Policy=New-DefaultPolicy
$script:Policy.packageRules=@(
  [pscustomobject]@{id='Google.Chrome';mode='Auto';allowMajorVersion=$true},
  [pscustomobject]@{id='Google.Chrome.EXE';mode='Auto';allowMajorVersion=$true},
  [pscustomobject]@{id='OpenJS.NodeJS.LTS';mode='ManualOnly';allowMajorVersion=$false},
  [pscustomobject]@{id='Microsoft.SQLServer.2025.Developer';mode='ManualOnly';allowMajorVersion=$false}
)
$script:CapturedEvents=@()
function Write-AgentEvent {param($EventId,$EntryType,$Fields) $script:CapturedEvents += [pscustomobject]$Fields}
$packages=@('Google.Chrome','Google.Chrome.EXE','OpenJS.NodeJS.LTS','Microsoft.SQLServer.2025.Developer','Test.Major')|ForEach-Object {[pscustomobject]@{Id=$_;Name=$_;Installed='155.0.1';Available='156.0.1';Source='winget';InstallScope='Machine'}}
Assert-Equal $packages.Count 5 'Selection fixture is incomplete.'
# Act.
$selected=@(Select-PolicyPackages -Packages $packages)
# Assert.
Assert-Equal $selected.Count 2 'Only Chrome IDs may cross major versions.'
Assert-Equal @($selected|Where-Object Id -notlike 'Google.Chrome*').Count 0 'Unrelated package escaped its gate.'
Assert-Equal $script:SelectionDeferredCount 3 'Selection-time deferrals must be counted.'
$count++

# Arrange / initial assert: preserve unrelated failures and pending reboot while moving one false user nag.
$script:InapplicableUpgradeStatePath=Join-Path $scratch 'inapplicable.json'
$notification=Get-NotificationStatePath
$null=New-Item -ItemType Directory -Path (Split-Path $notification) -Force
$target=[pscustomobject]@{Id='Microsoft.Azd';Name='Azure Developer CLI';Installed='1.34.200';Available='1.37.200';Source='winget'}
@{status='CompletedWithFailures';summary=@{failed=2;restartRequired=1};failedPackages=@(@{id='Microsoft.Azd';source='winget';available='1.37.200'},@{id='Test.RealFailure';source='winget';available='2.0'});restartRequiredPackages=@(@{id='Test.Pending';source='winget';available='1.1'})}|ConvertTo-Json -Depth 6|Set-Content $notification
Assert-Equal (Test-Path $notification) $true 'Mixed notification fixture must exist.'
# Act.
Move-FailureNotificationToAdminDeferral -Package $target -Reason NoApplicableInstaller -EvidencePath 'captured-attempt.json'
# Assert.
$state=Get-Content $notification -Raw|ConvertFrom-Json
Assert-Equal @($state.failedPackages).Count 1 'Only the reclassified target may leave user failure state.'
Assert-Equal $state.failedPackages[0].id 'Test.RealFailure' 'Real unrelated failure was cleared.'
Assert-Equal @($state.restartRequiredPackages).Count 1 'Pending reboot was discarded.'
$admin=@(Get-InapplicableUpgradeState)[0]
Assert-Equal $admin.reason 'NoApplicableInstaller' 'Administrator reason was lost.'
Assert-Equal @($admin.priorNotificationFailure).Count 1 'Original failure evidence was discarded.'
$count++

# Arrange / initial assert: a transferred known-bad notice must not become a generic cooldown.
$admin.reason='KnownBadRelease'
Save-InapplicableUpgradeState -Packages @($admin)
Assert-Equal $admin.reason 'KnownBadRelease' 'Quarantine fixture must be explicit.'
# Act.
$cached=Get-InapplicableUpgradeDeferral -Package $target
# Assert: corrected installer hashes remain eligible through the hash guard.
Assert-Equal ($null -eq $cached) $true 'A quarantine notice must not override corrected-hash release.'
$count++

Import-Functions (Join-Path $repo 'windows-update-scripts\Test-DanmadePatchAgentHealth.ps1')
$now=Get-Date
function New-Record($Id,$Status,$Time){[pscustomobject]@{mode='Machine';runId=$Id;status=$Status;timestamp=$Time.ToString('o')}}

# Arrange / initial assert: historical June-style incomplete record plus fresh completed assessment.
$records=@((New-Record 'historical' 'RunStarted' $now.AddDays(-100)),(New-Record 'fresh' 'CompletedRestartRequired' $now.AddHours(-1)))
Assert-Equal $records.Count 2 'Historical fixture must retain both records.'
# Act.
$health=Get-AgentLogHealth -Records $records -Processes @() -ActiveRunId ''
# Assert.
Assert-Equal @($health.Incomplete).Count 0 'Historical record must not be a current hang.'
Assert-Equal @($health.HistoricalIncomplete).Count 1 'Historical evidence must remain visible.'
Assert-Equal $health.LastAssessment $now.AddHours(-1) 'Latest real assessment was lost.'
$count++

# Arrange / initial assert: recent unfinished execution after the most recent assessment.
$records+=New-Record 'current' 'RunStarted' $now.AddMinutes(-30)
Assert-Equal $records[-1].runId 'current' 'Current run fixture missing.'
# Act.
$health=Get-AgentLogHealth -Records $records -Processes @() -ActiveRunId 'current'
# Assert.
Assert-Equal @($health.Incomplete).Count 1 'A genuine incomplete run must not be retired.'
Assert-Equal $health.Incomplete[0].RunId 'current' 'Wrong run identified as incomplete.'
$count++

# Arrange / initial assert: completed assessment six days ago followed only by skipped attempts.
$old=$now.AddDays(-6)
$records=@((New-Record 'assessed' 'NoPackagesSelected' $old),(New-Record 'skip' 'SkippedOutsideMaintenanceWindow' $now.AddMinutes(-1)))
Assert-Equal (($now-$old).TotalHours -gt 26) $true 'Assessment must be overdue.'
# Act.
$health=Get-AgentLogHealth -Records $records -Processes @() -ActiveRunId ''
# Assert.
Assert-Equal $health.LastAssessment $old 'Skipped attempts must not refresh assessment age.'
Assert-Equal $health.LastSkipped $now.AddMinutes(-1) 'Latest skip must be reported separately.'
$count++

# Arrange / initial assert: run without a terminal record remains live even after another completion.
$start=$now.AddHours(-3)
$records=@((New-Record 'live' 'RunStarted' $start),(New-Record 'other' 'Completed' $now.AddHours(-1)))
$processes=@([pscustomobject]@{CreationDate=$start.ToString('o')})
Assert-Equal $processes.Count 1 'Matching live process must exist.'
# Act.
$health=Get-AgentLogHealth -Records $records -Processes $processes -ActiveRunId ''
# Assert.
Assert-Equal @($health.Incomplete).Count 1 'Later assessment must not conceal a matching live execution.'
$count++

# Arrange / initial assert: actual Health control flow gets an old assessment and a current skip.
$healthRoot=Join-Path $scratch 'health'
$null=New-Item -ItemType Directory -Path (Join-Path $healthRoot 'Events') -Force
$events=Join-Path $healthRoot 'Events\patch-agent.jsonl'
@((New-Record 'assessed' 'NoPackagesSelected' $old),(New-Record 'skip' 'SkippedOutsideMaintenanceWindow' $now.AddMinutes(-1)))|ForEach-Object {$_|ConvertTo-Json -Compress}|Set-Content $events
'[]'|Set-Content (Join-Path $healthRoot 'processes.json')
Assert-Equal (Test-Path $events) $true 'Health log fixture must exist.'
# Act.
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'windows-update-scripts\Test-DanmadePatchAgentHealth.ps1') -LogRoot $healthRoot -ProcessSnapshotPath (Join-Path $healthRoot 'processes.json') | Out-Null
# Assert.
Assert-Equal $LASTEXITCODE 0 'Report-only Health failed to run.'
$events=@(Get-Content $events|ForEach-Object {$_|ConvertFrom-Json})
Assert-Equal @($events|Where-Object status -eq 'AgentAssessmentOverdue').Count 1 'Overdue assessment must generate evidence.'
Assert-Equal @($events|Where-Object status -eq 'AgentRunEvidenceMissing').Count 0 'Current attempt evidence must remain distinct.'
$count++
[pscustomobject]@{Status='Passed';Cases=$count;Scratch=$scratch}|ConvertTo-Json
