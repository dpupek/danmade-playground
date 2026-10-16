<# Deterministic source-function tests; all writes stay in .scratch. No endpoint changes. #>
[CmdletBinding(SupportsShouldProcess = $true)]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$path = Join-Path $repoRoot 'windows-update-scripts\danmade-patch-agent.ps1'
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Agent parse failed.' }
foreach ($function in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Parent.Parent -eq $ast }, $true)) { Invoke-Expression $function.Extent.Text }
function Assert-Equal { param($Actual, $Expected, [string]$Message) if ($Actual -ne $Expected) { throw "$Message Expected=$Expected Actual=$Actual" } }
$count = 0

foreach ($case in @(
  @{Text='Installation completed successfully.';Code=0},
  @{Text="Installation completed successfully.`nInstaller return code: 1603";Code=1603},
  @{Text="CustomAction returned actual error code 1603 but will be translated to success due to continue marking`nInstallation success or error status: 0.";Code=0},
  @{Text='Installation success or error status: 1605.';Code=1605},
  @{Text="Installation success or error status: 0.`nMainEngineThread is returning 3010";Code=3010},
  @{Text='Errorcode: -2147024809';Code=-2147024809},
  @{Text='CustomAction returned actual error code 1605';Code=1605}
)) {
  # Arrange / initial assert.
  Assert-Equal ([string]::IsNullOrEmpty($case.Text)) $false 'MSI fixture must exist.'
  # Act.
  $outcome = ConvertFrom-InstallerLogText $case.Text
  # Assert.
  Assert-Equal $outcome.ExitCode $case.Code 'MSI terminal outcome mismatch.'
  $count++
}

foreach ($case in @(
  @{Text='A newer version was found, but the install technology is different from the current version installed.';Status='InstallerTechnologyMismatch';Retry=$false},
  @{Text='No available upgrade found. No newer package versions are available from the configured sources.';Status='NoAvailableUpgrade';Retry=$false},
  @{Text='InternetReadFile() failed. 0x80072ee2 : unknown error';Status='TransientDownloadFailure';Retry=$true},
  @{Text='Download request status is not success. 0x801901ad';Status='TransientDownloadFailure';Retry=$true},
  @{Text='Download request status is not success. 0x80190194';Status='DownloadFailure';Retry=$false},
  @{Text='Starting package install. InternetReadFile() failed. 0x80072ee2';Status='DownloadFailure';Retry=$false},
  @{Text='';Status='UnconfirmedOutcome';Retry=$false}
)) {
  $inputResult = [pscustomobject]@{ ExitCode=0; TimedOut=$false; StdOut=$case.Text; StdErr='' }
  Assert-Equal $inputResult.ExitCode 0 'Fixture must reproduce zero exit.'
  $outcome = Get-WingetCommandOutcome $inputResult
  Assert-Equal $outcome.Status $case.Status 'Command outcome mismatch.'
  Assert-Equal $outcome.IsRetryable $case.Retry 'Unsafe retry classification.'
  $count++
}

$LogRoot = Join-Path (Join-Path $repoRoot '.scratch') ('function-tests-' + [guid]::NewGuid().ToString('N'))
$Mode = 'Machine'; $script:RunId='function-test'; $script:RunLock=$null; $script:Policy=New-DefaultPolicy
$script:RunDeadline=(Get-Date).AddMinutes(120)
Initialize-AgentStorage
$WhatIfPreference=$true
function Add-PackageScope { param($Package) $Package | Add-Member InstallScope 'Machine' -Force; $Package }
function Invoke-WingetProcess { param($Arguments, $TimeoutSeconds) $script:InventoryCalls++; return $script:InventoryResults[$script:InventoryCalls-1] }
function New-Result { param($Output, $Exit=0, $Timeout=$false) [pscustomobject]@{StdOut=$Output;StdErr='';ExitCode=$Exit;TimedOut=$Timeout} }
foreach ($case in @(
  @{Results=@((New-Result '[]'));Status='Empty';Calls=1},
  @{Results=@((New-Result '{"unexpected":[]}'));Status='Failed';Calls=1},
  @{Results=@((New-Result '[]' 42));Status='Failed';Calls=1},
  @{Results=@((New-Result '' 0 $true));Status='Failed';Calls=1},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result 'No available upgrade found.'));Status='Empty';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result 'No available upgrade found.' ([Convert]::ToInt32('8A15002B',16))));Status='Empty';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result 'No available upgrade found.' 42));Status='Failed';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result 'No installed package found matching input criteria.' ([Convert]::ToInt32('8A150014',16))));Status='Empty';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result "No available upgrade found.`nFailed when searching source: winget"));Status='Failed';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result "Name  Id  Version  Available  Source`n------------------------------------`ncURL cURL.cURL 1.0.0 1.1.0 winget`n1 upgrades available."));Status='Succeeded';Calls=2},
  @{Results=@((New-Result 'Argument name was not recognized: --output' 1),(New-Result 'An unrecognized localized table'));Status='Failed';Calls=2}
)) {
  $script:InventoryCalls=0; $script:InventoryResults=$case.Results
  Assert-Equal $script:InventoryCalls 0 'Inventory must not have run yet.'
  $outcome=Get-WingetInventory
  Assert-Equal $outcome.Status $case.Status 'Inventory result mismatch.'
  Assert-Equal $script:InventoryCalls $case.Calls 'Unsafe inventory fallback.'
  $count++
}

$WhatIfPreference=$false
$PolicyPath=Join-Path $LogRoot 'missing-policy.json'
Assert-Equal (Test-Path $PolicyPath) $false 'Missing-policy fixture must be missing.'
$caught=$false
try { $null=Import-AgentPolicy } catch { $caught=$true }
Assert-Equal $caught $true 'Missing policy must fail closed.'
$count++
Assert-Equal (Enter-AgentRunLock) $true 'First run must acquire the lock.'
$lock=$script:RunLock
try {
  Assert-Equal (Enter-AgentRunLock) $false 'Concurrent run must not acquire the lock.'
  Write-ActiveRunState
  Write-ActiveRunState
  $script:RunId='other-run'
  Clear-ActiveRunState
  Assert-Equal (Test-Path $script:ActiveRunStatePath) $true 'Another run must not clear owned state.'
  $script:RunId='function-test'
  Clear-ActiveRunState
  Assert-Equal (Test-Path $script:ActiveRunStatePath) $false 'Owner must clear its state.'
} finally { $lock.Dispose() }
$count++

foreach ($row in @(
  @{Text='Test App Test.Package 1.0.0 1.1.0 winget';Id='Test.Package'},
  @{Text='Non Dotted App    curl    Unknown    8.0.0    winget';Id='curl'},
  @{Text='Application 2026  Test.App  1.0.0  1.2.0  winget';Id='Test.App'}
)) {
  Assert-Equal ([string]::IsNullOrEmpty($row.Text)) $false 'Table fixture must exist.'
  $parsed=ConvertFrom-WingetUpgradeTableRow $row.Text
  Assert-Equal $parsed.Id $row.Id 'Table parser lost package identity.'
  $count++
}

foreach ($badPolicy in @('{}','{"enabled":"false"}','{"enabled":true,"includeUnknown":true,"allowedPackageIds":[],"blockedPackageIds":[],"packageRules":[],"maxRetries":"1","maxRunMinutes":120}')) {
  $PolicyPath=Join-Path $LogRoot 'invalid-policy.json'
  $badPolicy | Set-Content $PolicyPath
  $caught=$false
  try { $null=Import-AgentPolicy } catch { $caught=$true }
  Assert-Equal $caught $true 'Malformed policy must fail closed.'
  $count++
}

$rawPath=Join-Path $LogRoot 'captured-output.txt'
'plain command output' | Set-Content $rawPath
$captured=Read-ProcessText $rawPath
Assert-Equal ($null -eq $captured.PSObject.Properties['PSPath']) $true 'Captured output must not retain file provider metadata.'
$count++

# Arrange / initial assert: timeout must not consume additional installer attempts.
$script:SchemaVersion='1.0';$script:Policy.wazuhReporting.eventLog=$false
$script:Policy.maxRetries=1;$script:UpgradeCalls=0
function Get-RebootMarkerSnapshot { [pscustomobject]@{cbsRebootPending=$false;windowsUpdateRebootPending=$false;pendingFileRenameCount=0;bootTime=$null} }
function Invoke-WingetProcess {param($Arguments,$TimeoutSeconds) $script:UpgradeCalls++; New-Result 'Starting package install' -1 $true}
$package=[pscustomobject]@{Id='Test.Timeout';Name='Timeout';Installed='1.0.0';Available='1.1.0';Source='winget'}
Assert-Equal $script:UpgradeCalls 0 'No command should precede timeout test.'
# Act.
$result=Invoke-PackageUpgrade $package
# Assert.
Assert-Equal $result 'Failed' 'Timeout must remain an uncertain failure.'
Assert-Equal $script:UpgradeCalls 1 'Timeout must never retry the installer.'
$null=Invoke-PackageUpgrade $package
Assert-Equal $script:UpgradeCalls 1 'Persisted uncertain outcome must suppress subsequent installation.'
$count++

# Arrange / initial assert: genuine installer failure must win over a new reboot marker.
$script:MarkerCalls=0;$script:UpgradeCalls=0
function Get-RebootMarkerSnapshot { $script:MarkerCalls++; [pscustomobject]@{cbsRebootPending=$false;windowsUpdateRebootPending=$false;pendingFileRenameCount=$script:MarkerCalls-1;bootTime=$null} }
function Get-InstallerOutcomeFromLog {param($Path) ConvertFrom-InstallerLogText 'Installer return code: 1603'}
function Invoke-WingetProcess {param($Arguments,$TimeoutSeconds) $script:UpgradeCalls++; New-Result 'Successfully installed' 0}
$package.Id='Test.FatalWithMarker'
Assert-Equal $script:MarkerCalls 0 'Reboot fixture must start clean.'
# Act.
$result=Invoke-PackageUpgrade $package
# Assert.
Assert-Equal $result 'Failed' 'Fatal installer result cannot become restart-required.'
Assert-Equal (Test-Path (Get-RestartVerificationStatePath)) $false 'Fatal failure must not create restart state.'
$count++
# Arrange / initial assert: a newer offer appears after the requested version is installed.
function Get-WingetUpgradeList { param([switch]$FailOnError) return @($script:VerificationPackages) }
$targetPackage=[pscustomobject]@{Id='Test.Version';Available='1.1.0';Source='winget'}
$script:VerificationPackages=@([pscustomobject]@{Id='Test.Version';Installed='1.1.0';Available='1.2.0';Source='winget'})
Assert-Equal $script:VerificationPackages[0].Installed $targetPackage.Available 'Target version must be installed initially.'
# Act / assert.
$verification=Test-WingetPackageStillOffered $targetPackage
Assert-Equal $verification.Status 'VerifiedNotOffered' 'Newer offer must not invalidate a reached target.'
$count++
$script:VerificationPackages[0].Installed='1.0.0'
Assert-Equal $script:VerificationPackages[0].Installed '1.0.0' 'Unpatched fixture version must be below target.'
$verification=Test-WingetPackageStillOffered $targetPackage
Assert-Equal $verification.Status 'VerificationFailedStillOffered' 'Unreached target must remain a verification failure.'
$count++
[pscustomobject]@{Status='Passed'; Cases=$count; ScratchRoot=$LogRoot} | ConvertTo-Json
