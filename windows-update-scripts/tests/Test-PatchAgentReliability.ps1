<# Regression tests for native exit codes, Burn outcomes, release holds and retained state.
   Loads production functions only. All generated files remain in .scratch; no installers run. #>
[CmdletBinding(SupportsShouldProcess = $true)]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$scratch = Join-Path $repo ('.scratch\reliability-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $scratch | Out-Null
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'windows-update-scripts\danmade-patch-agent.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Production agent does not parse.' }
foreach ($node in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Parent.Parent -eq $ast }, $true)) { Invoke-Expression $node.Extent.Text }
function Assert-Equal { param($Actual, $Expected, [string]$Message) if ($Actual -ne $Expected) { throw "$Message Expected=$Expected Actual=$Actual" } }
$count = 0

# Arrange: a harmless physical child, not a mocked process result.
$script:WingetCommand = Join-Path $scratch 'native-probe.exe'
Add-Type -OutputAssembly $script:WingetCommand -OutputType ConsoleApplication -TypeDefinition @'
using System;
using System.Threading;
public class NativeProbe {
  public static int Main(string[] args) {
    Console.OutputEncoding = new System.Text.UTF8Encoding(false);
    if (args[0] == "wait") { Thread.Sleep(10000); return 0; }
    if (args[0] == "streams") {
      for (int i=0;i<2000;i++) { Console.WriteLine("output-"+i); Console.Error.WriteLine("error-"+i); }
      Console.WriteLine("caf\u00e9"); return 37;
    }
    return int.Parse(args[0]);
  }
}
'@
$script:RunDeadline = (Get-Date).AddMinutes(10)
Assert-Equal (Test-Path $script:WingetCommand) $true 'Probe must exist.'
foreach ($code in @(0,37,3010,-1,-1978335212,-1978334967)) {
  # Act.
  $result = Invoke-WingetProcess -Arguments @([string]$code) -TimeoutSeconds 5
  # Assert.
  Assert-Equal $result.ExitCode $code 'Native exit code was lost.'
  Assert-Equal $result.TimedOut $false 'Native child unexpectedly timed out.'
  $count++
}
$result = Invoke-WingetProcess -Arguments @('streams') -TimeoutSeconds 5
Assert-Equal $result.ExitCode 37 'Large stream exit code was lost.'
Assert-Equal ($result.StdOut.Contains('output-1999')) $true 'stdout was truncated.'
Assert-Equal ($result.StdErr.Contains('error-1999')) $true 'stderr was truncated.'
Assert-Equal ($result.StdOut.Contains(('caf' + [char]233))) $true 'UTF-8 output was damaged.'
$count++
$clock = [Diagnostics.Stopwatch]::StartNew()
$result = Invoke-WingetProcess -Arguments @('wait') -TimeoutSeconds 1
Assert-Equal $result.TimedOut $true 'Timeout must not report success.'
Assert-Equal ($clock.Elapsed.TotalSeconds -lt 8) $true 'Output drain must remain bounded.'
$count++

foreach ($case in @(
  @{Text='[BF38:78F0]i007: Exit code: 0xbc2, restarting: No'; Code=3010},
  @{Text='[BF38:78F0]i007: Exit code: 0x80070643, restarting: No'; Code=1603},
  @{Text='[BF38:78F0]i007: Exit code: 0x645, restarting: No'; Code=1605},
  @{Text="CustomAction returned actual error code 1603 but will be translated to success due to continue marking`ni007: Exit code: 0x0, restarting: No"; Code=0},
  @{Text="i007: Exit code: 0x0, restarting: No`nInstallation success or error status: 1603."; Code=1603}
)) {
  Assert-Equal ([string]::IsNullOrEmpty($case.Text)) $false 'Burn fixture must exist.'
  $result = ConvertFrom-InstallerLogText -Text $case.Text
  Assert-Equal $result.ExitCode $case.Code 'Burn terminal result mismatch.'
  $count++
}
foreach ($case in @(
  @{Code='8A150010';Status='NoApplicableInstaller'},
  @{Code='8A150014';Status='InstalledPackageNotFound'},
  @{Code='8A15008E';Status='InstallerTechnologyMismatch'},
  @{Code='8A15002B';Status='NoApplicableUpgrade'},
  @{Code='8A150109';Status='RestartRequired'},
  @{Code='8A15010B';Status='RestartRequired'},
  @{Code='8A15010A';Status='WingetFailed'}
)) {
  $inputResult = [pscustomobject]@{ ExitCode=[Convert]::ToInt32($case.Code,16); TimedOut=$false; StdOut=''; StdErr='' }
  Assert-Equal ($inputResult.ExitCode -eq 0) $false 'HRESULT fixture must be nonzero.'
  $result = Get-WingetCommandOutcome -UpgradeResult $inputResult
  Assert-Equal $result.Status $case.Status 'Native HRESULT classification mismatch.'
  $count++
}

# Arrange: the reboot completed the requested version, but a newer release now exists.
$Mode = 'Machine'; $LogRoot = $scratch; $script:RunId = 'reliability'; $script:Policy = New-DefaultPolicy
$script:SchemaVersion = '1.0'; $script:EventLogUnavailable = $null
$script:JsonlPath = Join-Path $scratch 'events.jsonl'
$script:Policy.wazuhReporting.eventLog = $false
Initialize-AgentStorage
function Get-SystemBootTime { [datetime]'2026-10-16' }
$record = [pscustomobject]@{ id='Test.Runtime'; name='Runtime'; available='8.0.32'; source='winget'; bootTimeAtDetection='2026-10-01'; evidencePath='fixture' }
$offer = [pscustomobject]@{ Id='Test.Runtime'; Installed='8.0.32'; Available='8.0.33'; Source='winget'; InstallScope='Machine' }
New-Item -ItemType Directory -Path (Split-Path (Get-RestartVerificationStatePath)) -Force | Out-Null
@{packages=@($record)} | ConvertTo-Json -Depth 5 | Set-Content (Get-RestartVerificationStatePath)
Assert-Equal (Test-Path (Get-RestartVerificationStatePath)) $true 'Pending reboot fixture must exist.'
$resolved = Resolve-PendingRestartVerification -AvailablePackages @($offer)
Assert-Equal @($resolved.FailedPackages).Count 0 'A newer offer must not fail an already installed target.'
Assert-Equal ($resolved.ResolvedPackageIds -contains 'Test.Runtime') $true 'Resolved target must be returned.'
$count++

# Arrange: exact ARP name/scope proves an externally updated failure, not a pending reboot.
function Get-CachedUninstallEntries { @([pscustomobject]@{ DisplayNameNorm='edge'; DisplayVersion='155.0.0.0'; PSPath='HKEY_LOCAL_MACHINE'; InstallLocation=''; UninstallString='' }) }
$oldFailure = [pscustomobject]@{ id='Microsoft.Edge'; name='Edge'; available='154.0.0.0'; source='winget' }
$oldRestart = [pscustomobject]@{ id='Test.Pending'; name='Pending'; available='1.1.0'; source='winget' }
Write-NotificationState -Status CompletedWithFailures -Summary @{Failed=1;RestartRequired=1} -FailedPackages @($oldFailure) -RestartRequiredPackages @($oldRestart)
Assert-Equal (Test-Path (Get-NotificationStatePath)) $true 'Mixed notification fixture must exist.'
Resolve-RecordedPackageState -AvailablePackages @()
Clear-NotificationState -PreserveRestartRequired
$state = Get-Content (Get-NotificationStatePath) -Raw | ConvertFrom-Json
Assert-Equal @($state.failedPackages).Count 0 'Externally resolved target must stop nagging.'
Assert-Equal @($state.restartRequiredPackages).Count 1 'Pending reboot must survive reconciliation.'
Write-NotificationState -Status CompletedWithFailures -Summary @{Failed=1} -FailedPackages @($oldFailure) -RestartRequiredPackages @()
$state = Get-Content (Get-NotificationStatePath) -Raw | ConvertFrom-Json
Assert-Equal @($state.restartRequiredPackages).Count 1 'New failure must retain an older reboot notice.'
$count++
Resolve-RecordedPackageState -AvailablePackages @() -PostBootRestartPackages @($oldRestart)
Assert-Equal (Test-Path (Get-NotificationStatePath)) $false 'Confirmed resolved targets should clear state.'
$count++

# Arrange: hash-scoped bad release; a corrected manifest at the same version must be eligible.
$badHash = 'a' * 64
$script:Policy.packageRules = @([pscustomobject]@{ id='Test.Held'; mode='Auto'; allowMajorVersion=$false; processNames=@(); installerType=''; releaseHolds=@([pscustomobject]@{version='1.1.0';scope='machine';installerSha256=$badHash;until=(Get-Date).AddDays(7).ToString('o');reason='Fixture bad release'}) })
$package = [pscustomobject]@{ Id='Test.Held'; Installed='1.0.0'; Available='1.1.0'; Source='winget' }
$script:ProbeHash = $badHash; $script:ShowExit = 0
function Invoke-WingetProcess { param($Arguments,$TimeoutSeconds) Assert-Equal $Arguments[0] 'show' 'Hold may only inspect metadata.'; [pscustomobject]@{ExitCode=$script:ShowExit;TimedOut=$false;StdOut=('  Installer SHA256: ' + $script:ProbeHash);StdErr=''} }
Assert-Equal @($script:Policy.packageRules).Count 1 'Hold fixture must be present.'
$result = Get-PackageReleaseHold -Package $package
Assert-Equal $result.Status 'DeferredKnownBadRelease' 'Known bad hash must be held.'
$script:ProbeHash = 'b' * 64
$result = Get-PackageReleaseHold -Package $package
Assert-Equal $result.Status 'Clear' 'Corrected manifest hash must release the hold.'
$script:ShowExit = 42
$result = Get-PackageReleaseHold -Package $package
Assert-Equal $result.Status 'DeferredReleaseHoldUnverified' 'Unknown metadata must not bypass hold.'
$count += 3

# Arrange / initial assert: the same bad bytes remain after the review deadline.
$script:ShowExit = 0; $script:ProbeHash = $badHash
$script:Policy.packageRules[0].releaseHolds[0].until = (Get-Date).AddDays(-1).ToString('o')
Assert-Equal ([datetime]$script:Policy.packageRules[0].releaseHolds[0].until -lt (Get-Date)) $true 'Review fixture must be overdue.'
# Act.
$result = Get-PackageReleaseHold -Package $package
# Assert: overdue review does not approve known-bad bytes; corrected bytes are still eligible.
Assert-Equal $result.Status 'DeferredKnownBadRelease' 'Expired review must not release unchanged bad bytes.'
Assert-Equal $result.ReviewOverdue $true 'Overdue review must be visible.'
$script:ProbeHash = 'c' * 64
$result = Get-PackageReleaseHold -Package $package
Assert-Equal $result.Status 'Clear' 'Expired review must not block a corrected hash.'
$count += 2

$script:Policy.maintenanceWindow = [pscustomobject]@{enabled=$true;start='invalid';end='04:00';machineOnly=$true}
$Mode = 'User'
Assert-Equal (Test-MaintenanceWindow) $true 'Machine window must not suppress user logon updates.'
$Mode = 'Machine'; $caught = $false
try { Test-MaintenanceWindow | Out-Null } catch { $caught = $true }
Assert-Equal $caught $true 'Invalid machine window must fail closed.'
$count++
[pscustomobject]@{Status='Passed';Cases=$count;ScratchRoot=$scratch} | ConvertTo-Json
