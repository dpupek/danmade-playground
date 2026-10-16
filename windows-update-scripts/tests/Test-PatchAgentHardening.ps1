[CmdletBinding()]
param([switch]$RunMachineAsSystem)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$agentPath = Join-Path $repoRoot 'windows-update-scripts\danmade-patch-agent.ps1'
$healthScriptPath = Join-Path $repoRoot 'windows-update-scripts\Test-DanmadePatchAgentHealth.ps1'
$runRoot = Join-Path (Join-Path $repoRoot '.scratch') ('patch-agent-hardening-{0}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$fakeBin = Join-Path $runRoot 'fake-bin'
$captures = Join-Path $runRoot 'captures'
New-Item -ItemType Directory -Path $fakeBin, $captures -Force | Out-Null

# Keep production control flow; replace only host observations in this scratch copy.
$originalAgentPath = $agentPath
$agentPath = Join-Path $runRoot 'test-agent.ps1'
$sourceText = Get-Content -LiteralPath $originalAgentPath -Raw
$sourceText = ($sourceText -split '# SIG # Begin signature block', 2)[0]
$hostSeams = @'
function Resolve-PackageInstallScope { param($Package) return $Mode }
function Test-IsSystemAccount { return $true }
function Initialize-EventSource { }
function Wait-DownloadRetry { param($Attempt) return $true }
function Get-RebootMarkerSnapshot { [pscustomobject]@{ cbsRebootPending=$false; windowsUpdateRebootPending=$false; pendingFileRenameCount=0; bootTime=[datetime]'2026-01-01' } }
function Get-SystemBootTime { [datetime]'2026-01-01' }
function Get-CachedUninstallEntries { @() }
'@
$sourceText.Replace('$script:EventLogUnavailable = $null', $hostSeams + [Environment]::NewLine + '$script:EventLogUnavailable = $null') | Set-Content -LiteralPath $agentPath -Encoding UTF8

function Assert-True {
  param(
    [bool]$Condition,
    [string]$Message
  )
  if (-not $Condition) { throw $Message }
}

function Write-FakeWinget {
  $exePath = Join-Path $fakeBin 'winget.exe'
  $source = @'
using System;
using System.IO;
using System.Linq;

public static class FakeWinget
{
  private static string Escape(string value)
  {
    if (value == null) { return ""; }
    return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
  }

  private static bool Has(string[] args, string value)
  {
    return args.Any(a => string.Equals(a, value, StringComparison.OrdinalIgnoreCase));
  }

  private static int IndexOf(string[] args, string value)
  {
    for (int i = 0; i < args.Length; i++)
    {
      if (string.Equals(args[i], value, StringComparison.OrdinalIgnoreCase)) { return i; }
    }
    return -1;
  }

  public static int Main(string[] args)
  {
    string capture = Environment.GetEnvironmentVariable("FAKE_WINGET_CAPTURE");
    string scenario = Environment.GetEnvironmentVariable("FAKE_WINGET_SCENARIO") ?? "";
    string statePath = Environment.GetEnvironmentVariable("FAKE_WINGET_STATE") ?? "";
    string packageId = Environment.GetEnvironmentVariable("FAKE_WINGET_PACKAGE_ID") ?? "Test.ScopePackage";
    string packageName = Environment.GetEnvironmentVariable("FAKE_WINGET_PACKAGE_NAME") ?? "Test Scope Package";
    string installedVersion = Environment.GetEnvironmentVariable("FAKE_WINGET_INSTALLED_VERSION") ?? "1.0.0";
    string availableVersion = Environment.GetEnvironmentVariable("FAKE_WINGET_AVAILABLE_VERSION") ?? "1.1.0";
    string source = "winget";

    if (!string.IsNullOrWhiteSpace(capture))
    {
      string serializedArgs = string.Join(",", args.Select(a => "\"" + Escape(a) + "\""));
      File.AppendAllText(capture, "{\"Scenario\":\"" + Escape(scenario) + "\",\"Args\":[" + serializedArgs + "]}" + Environment.NewLine);
    }

    if (args.Length > 0 && args[0] == "--info")
    {
      Console.WriteLine("Windows Package Manager v.fake");
      return 0;
    }

    if (args.Length >= 2 && args[0] == "source" && args[1] == "update")
    {
      Console.WriteLine("Done");
      return 0;
    }

    if (args.Length > 0 && args[0] == "show") {
      Console.WriteLine("  Installer SHA256: " + new string(scenario == "ReleaseHoldCorrected" ? 'b' : 'a', 64));
      return scenario == "ReleaseHoldUnreadable" ? 42 : 0;
    }

    if (args.Length > 0 && args[0] == "upgrade")
    {
      int logIndex = IndexOf(args, "--log");
      bool hasOutputJson = Has(args, "--output") && Has(args, "json");
      if (hasOutputJson)
      {
        int listCount = 0;
        if (!string.IsNullOrWhiteSpace(statePath) && File.Exists(statePath)) { int.TryParse(File.ReadAllText(statePath), out listCount); }
        listCount++;
        if (!string.IsNullOrWhiteSpace(statePath)) { File.WriteAllText(statePath, listCount.ToString()); }
        if (scenario == "InventoryFailure") { Console.WriteLine("Source unavailable"); return 42; }
        if (scenario == "VerificationMalformed" && listCount > 1) { Console.WriteLine("{\"unexpectedSchema\":[]}"); return 0; }
        if (scenario != "StillOffered" && scenario != "NoApplicableUpgrade" && listCount > 1) { Console.WriteLine("[]"); return 0; }
        Console.WriteLine("[{\"PackageIdentifier\":\"" + Escape(packageId) + "\",\"PackageName\":\"" + Escape(packageName) + "\",\"InstalledVersion\":\"" + Escape(installedVersion) + "\",\"AvailableVersion\":\"" + Escape(availableVersion) + "\",\"Source\":\"" + source + "\"}]");
        return 0;
      }

      if (logIndex >= 0 && logIndex + 1 < args.Length)
      {
        string logPath = args[logIndex + 1];
        Directory.CreateDirectory(Path.GetDirectoryName(logPath));
        if (scenario == "InstalledPackageNotFound") { Console.WriteLine("No installed package found matching input criteria."); return unchecked((int)0x8A150014); }
        if (scenario == "NativeRestart") { Console.WriteLine("Restart your PC to finish installation."); return unchecked((int)0x8A150109); }
        if (scenario == "BurnRestart") { File.WriteAllText(logPath,"[BF38:78F0]i007: Exit code: 0xbc2, restarting: No"); Console.WriteLine("Restart your PC to finish installation."); return 0; }
        if (scenario == "TechnologyMismatch") { Console.WriteLine("A newer version was found, but the install technology is different from the current version installed. Please uninstall the package and install the newer version."); return 0; }
        if (scenario == "NoAvailableUpgrade") { Console.WriteLine("No available upgrade found.\nNo newer package versions are available from the configured sources."); return 0; }
        if (scenario == "DownloadTimeout") { Console.WriteLine("Downloading https://download.example.invalid/chrome.exe\nInternetReadFile() failed.\n0x80072ee2 : unknown error"); return 0; }
        if (scenario == "DownloadHttp500")
        {
          Console.WriteLine("Downloading https://download.example.invalid/LibreOffice.msi");
          Console.WriteLine("An unexpected error occurred while executing the command:");
          Console.WriteLine("Download request status is not success.");
          Console.WriteLine("0x801901f4 : Internal server error (500).");
          return 0;
        }
        if (scenario == "NoApplicableUpgrade")
        {
          Console.WriteLine("No applicable upgrade found.");
          Console.WriteLine("A newer package version is available in a configured source, but it does not apply to your system or requirements.");
          return 0;
        }
        if (scenario == "Installer1605")
        {
          File.WriteAllText(logPath, "CustomAction returned actual error code 1605");
        }
        else if (scenario == "Installer1603")
        {
          File.WriteAllText(logPath, "Installer failed with exit code: 1603");
        }
        else if (scenario == "Installer1603ContinuedSuccess")
        {
          File.WriteAllText(logPath, "CustomAction Wix4RemoveFoldersEx_X64 returned actual error code 1603 but will be translated to success due to continue marking\r\nMSI completed. Installation completed successfully.\r\nInstallation success or error status: 0.");
        }
        else if (scenario == "SuccessThen1603") { File.WriteAllText(logPath, "Installation completed successfully.\r\nInstallation success or error status: 1603."); }
        else if (scenario == "Installer3010") { File.WriteAllText(logPath, "Installation success or error status: 0.\r\nMainEngineThread is returning 3010"); }
        else
        {
          File.WriteAllText(logPath, "Installer completed successfully");
        }
        Console.WriteLine("Successfully installed");
        if (scenario == "Installer1605" || scenario == "Installer1603") { return 1; }
        return 0;
      }

      if (scenario == "StillOffered")
      {
        Console.WriteLine("Name               Id                 Version Available Source");
        Console.WriteLine("--------------------------------------------------------------");
        Console.WriteLine(packageName + " " + packageId + " 1.0.0   2.0.0     " + source);
        Console.WriteLine("1 upgrades available.");
        return 0;
      }

      if (scenario == "VerificationNonZero")
      {
        Console.WriteLine("Verification failed");
        return 42;
      }

      Console.WriteLine("No available upgrade found.");
      return 0;
    }

    return 0;
  }
}
'@

  Add-Type -TypeDefinition $source -Language CSharp -OutputAssembly $exePath -OutputType ConsoleApplication
}

function Write-Policy {
  param(
    [string]$Path,
    [string]$AllowedId,
    [object[]]$PackageRules = @(),
    [int]$MaxRetries = 0,
    [int[]]$RetryableInstallerExitCodes = @(1618)
  )

  [pscustomobject]@{
    enabled = $true
    includeUnknown = $true
    allowedPackageIds = @($AllowedId)
    blockedPackageIds = @()
    packageRules = @($PackageRules)
    maxRetries = $MaxRetries
    retryableInstallerExitCodes = @($RetryableInstallerExitCodes)
    maxRunMinutes = 120
    maintenanceWindow = [pscustomobject]@{
      enabled = $false
      start = '02:00'
      end = '04:00'
    }
    rebootPolicy = 'ReportOnly'
    wazuhReporting = [pscustomobject]@{
      eventLog = $false
      jsonl = $true
    }
    wingetSourceRepair = $true
    logRetentionDays = 30
  } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Invoke-AgentScenario {
  param(
    [string]$Name,
    [ValidateSet('Machine', 'User')]
    [string]$Mode,
    [string]$Scenario,
    [string]$PackageId,
    [string]$PackageName,
    [object[]]$PackageRules = @(),
    [int]$MaxRetries = 0,
    [int[]]$RetryableInstallerExitCodes = @(1618),
    [string]$InstalledVersion = '1.0.0',
    [string]$AvailableVersion = '1.1.0',
    [string]$ScenarioRoot
  )

  if ([string]::IsNullOrWhiteSpace($ScenarioRoot)) { $ScenarioRoot = Join-Path $runRoot $Name }
  $scenarioRoot = $ScenarioRoot
  $logRoot = Join-Path $scenarioRoot 'agent'
  $policyPath = Join-Path $scenarioRoot 'policy.json'
  $capturePath = Join-Path $captures "$Name.jsonl"
  $statePath = Join-Path $captures "$Name.state"
  $wrapperPath = Join-Path $scenarioRoot 'run-agent.cmd'
  New-Item -ItemType Directory -Path $scenarioRoot -Force | Out-Null
  Write-Policy -Path $policyPath -AllowedId $PackageId -PackageRules $PackageRules -MaxRetries $MaxRetries -RetryableInstallerExitCodes $RetryableInstallerExitCodes

  if ($Mode -eq 'Machine' -and $RunMachineAsSystem) {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
      return [pscustomobject]@{
        Name = $Name
        ExitCode = $null
        Events = @()
        Calls = @()
        LogRoot = $logRoot
        Skipped = $true
        SkipReason = 'Current token is not elevated for temporary SYSTEM scheduled tasks.'
      }
    }
    @"
@echo off
set "PATH=$fakeBin;%PATH%"
set "FAKE_WINGET_SCENARIO=$Scenario"
set "FAKE_WINGET_PACKAGE_ID=$PackageId"
set "FAKE_WINGET_PACKAGE_NAME=$PackageName"
set "FAKE_WINGET_INSTALLED_VERSION=$InstalledVersion"
set "FAKE_WINGET_AVAILABLE_VERSION=$AvailableVersion"
set "FAKE_WINGET_CAPTURE=$capturePath"
set "FAKE_WINGET_STATE=$statePath"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "$agentPath" -PolicyPath "$policyPath" -Mode $Mode -LogRoot "$logRoot" -RunId "$Name"
exit /b %ERRORLEVEL%
"@ | Set-Content -LiteralPath $wrapperPath -Encoding ASCII

    $taskName = "Danmade Patch Agent Scratch Test $Name"
    $action = New-ScheduledTaskAction -Execute $wrapperPath
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(30)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest
    try {
      try {
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Force | Out-Null
      } catch {
        if (($_ | Out-String) -match 'Access is denied|0x80070005|PermissionDenied') {
          return [pscustomobject]@{
            Name = $Name
            ExitCode = $null
            Events = @()
            Calls = @()
            LogRoot = $logRoot
            Skipped = $true
            SkipReason = 'Current token cannot register temporary SYSTEM scheduled tasks.'
          }
        }
        throw
      }
      Start-ScheduledTask -TaskName $taskName
      $deadline = (Get-Date).AddSeconds(90)
      do {
        Start-Sleep -Milliseconds 500
        $info = Get-ScheduledTaskInfo -TaskName $taskName
        $task = Get-ScheduledTask -TaskName $taskName
      } while ($task.State -ne 'Ready' -and (Get-Date) -lt $deadline)
      $exitCode = [int](Get-ScheduledTaskInfo -TaskName $taskName).LastTaskResult
    } finally {
      Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    }
  } else {
    $oldPath = $env:Path
    $oldScenario = $env:FAKE_WINGET_SCENARIO
    $oldPackageId = $env:FAKE_WINGET_PACKAGE_ID
    $oldPackageName = $env:FAKE_WINGET_PACKAGE_NAME
    $oldCapture = $env:FAKE_WINGET_CAPTURE
    $oldState = $env:FAKE_WINGET_STATE
    $oldInstalledVersion = $env:FAKE_WINGET_INSTALLED_VERSION
    $oldAvailableVersion = $env:FAKE_WINGET_AVAILABLE_VERSION
    try {
      $env:Path = "$fakeBin;$oldPath"
      $env:FAKE_WINGET_SCENARIO = $Scenario
      $env:FAKE_WINGET_PACKAGE_ID = $PackageId
      $env:FAKE_WINGET_PACKAGE_NAME = $PackageName
      $env:FAKE_WINGET_CAPTURE = $capturePath
      $env:FAKE_WINGET_STATE = $statePath
      $env:FAKE_WINGET_INSTALLED_VERSION = $InstalledVersion
      $env:FAKE_WINGET_AVAILABLE_VERSION = $AvailableVersion

      & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $agentPath -PolicyPath $policyPath -Mode $Mode -LogRoot $logRoot -RunId $Name | Out-Null
      $exitCode = $LASTEXITCODE
    } finally {
      $env:Path = $oldPath
      $env:FAKE_WINGET_SCENARIO = $oldScenario
      $env:FAKE_WINGET_PACKAGE_ID = $oldPackageId
      $env:FAKE_WINGET_PACKAGE_NAME = $oldPackageName
      $env:FAKE_WINGET_CAPTURE = $oldCapture
      $env:FAKE_WINGET_STATE = $oldState
      $env:FAKE_WINGET_INSTALLED_VERSION = $oldInstalledVersion
      $env:FAKE_WINGET_AVAILABLE_VERSION = $oldAvailableVersion
    }
  }

  $jsonlPath = Join-Path $logRoot 'Events\patch-agent.jsonl'
  $events = if (Test-Path -LiteralPath $jsonlPath) {
    @(Get-Content -LiteralPath $jsonlPath | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json })
  } else {
    @()
  }
  $calls = if (Test-Path -LiteralPath $capturePath) {
    @(Get-Content -LiteralPath $capturePath | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json })
  } else {
    @()
  }

  [pscustomobject]@{
    Name = $Name
    ExitCode = $exitCode
    Events = $events
    Calls = $calls
    LogRoot = $logRoot
    ActiveRunStatePresent = Test-Path -LiteralPath (Join-Path $logRoot 'State\active-run.json')
    Skipped = $false
    SkipReason = $null
  }
}

Write-FakeWinget

$tokens = $null
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($agentPath, [ref]$tokens, [ref]$errors)
Assert-True -Condition ($errors.Count -eq 0) -Message "Agent parse errors: $($errors | Out-String)"

$machineSuccess = Invoke-AgentScenario -Name 'machine-success' -Mode Machine -Scenario Success -PackageId 'Test.ScopePackage' -PackageName 'Test Scope Package'
if (-not $machineSuccess.Skipped) {
  Assert-True -Condition ($machineSuccess.ExitCode -eq 0) -Message 'Machine success scenario should exit 0.'
  Assert-True -Condition (@($machineSuccess.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 1) -Message 'Verification must not invoke a named winget upgrade.'
  Assert-True -Condition (@($machineSuccess.Calls | Where-Object { $_.Args -contains '--scope' -and $_.Args -contains 'machine' }).Count -ge 3) -Message 'Machine upgrade and inventory did not pass --scope machine.'
  Assert-True -Condition (@($machineSuccess.Calls | Where-Object { $_.Args -contains '--output' -and $_.Args -contains 'json' -and $_.Args -notcontains '--id' }).Count -ge 2) -Message 'Machine verification did not re-list the read-only inventory.'
  Assert-True -Condition (@($machineSuccess.Events | Where-Object { $_.status -eq 'Succeeded' }).Count -eq 1) -Message 'Machine success scenario did not log Succeeded.'
}

$userSuccess = Invoke-AgentScenario -Name 'user-success' -Mode User -Scenario Success -PackageId 'cURL.cURL' -PackageName 'cURL'
Assert-True -Condition ($userSuccess.ExitCode -eq 0) -Message 'User success scenario should exit 0.'
Assert-True -Condition (@($userSuccess.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 1) -Message 'User verification invoked a named winget upgrade.'
Assert-True -Condition (@($userSuccess.Calls | Where-Object { $_.Args -contains '--scope' -and $_.Args -contains 'user' }).Count -ge 3) -Message 'User upgrade and inventory did not pass --scope user.'
Assert-True -Condition (-not $userSuccess.ActiveRunStatePresent) -Message 'Completed agent run left an active-run state record.'

$stillOffered = Invoke-AgentScenario -Name 'user-still-offered' -Mode User -Scenario StillOffered -PackageId 'cURL.cURL' -PackageName 'cURL'
Assert-True -Condition ($stillOffered.ExitCode -eq 1) -Message 'Still-offered scenario should exit 1.'
Assert-True -Condition (@($stillOffered.Events | Where-Object { $_.status -eq 'VerificationFailedStillOffered' }).Count -eq 1) -Message 'Still-offered scenario did not log VerificationFailedStillOffered.'
Assert-True -Condition (@($stillOffered.Events | Where-Object { $_.status -eq 'CompletedWithFailures' }).Count -eq 1) -Message 'Still-offered scenario did not log CompletedWithFailures.'

$installer1605 = Invoke-AgentScenario -Name 'user-installer-1605' -Mode User -Scenario Installer1605 -PackageId 'cURL.cURL' -PackageName 'cURL'
Assert-True -Condition ($installer1605.ExitCode -eq 1) -Message 'Installer 1605 scenario should exit 1.'
Assert-True -Condition (@($installer1605.Events | Where-Object { $_.status -eq 'Failed' -and $_.installerExitCode -eq 1605 }).Count -eq 1) -Message 'Installer 1605 scenario did not log Failed with installerExitCode 1605.'

$installer1603 = Invoke-AgentScenario -Name 'user-installer-1603' -Mode User -Scenario Installer1603 -PackageId 'cURL.cURL' -PackageName 'cURL' -MaxRetries 1
Assert-True -Condition ($installer1603.ExitCode -eq 1) -Message 'Installer 1603 scenario should exit 1.'
Assert-True -Condition (@($installer1603.Events | Where-Object { $_.status -eq 'Failed' -and $_.installerExitCode -eq 1603 -and $_.failureReason -eq 'InstallerFatalNoRetry' }).Count -eq 1) -Message 'Installer 1603 scenario did not log the terminal fatal failure reason.'
Assert-True -Condition (@($installer1603.Events | Where-Object { $_.status -eq 'RetryScheduled' }).Count -eq 0) -Message 'Installer 1603 should not schedule a retry.'
Assert-True -Condition (@($installer1603.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 1) -Message 'Installer 1603 should attempt exactly one named upgrade.'

$continuedInstallerSuccess = Invoke-AgentScenario -Name 'user-installer-1603-continued-success' -Mode User -Scenario Installer1603ContinuedSuccess -PackageId 'cURL.cURL' -PackageName 'cURL'
Assert-True -Condition ($continuedInstallerSuccess.ExitCode -eq 0) -Message 'Continued MSI custom-action 1603 scenario should exit 0.'
Assert-True -Condition (@($continuedInstallerSuccess.Events | Where-Object { $_.status -eq 'Succeeded' -and $_.installerExitCode -eq 0 }).Count -eq 1) -Message 'Final MSI success must override a continued custom-action 1603.'
Assert-True -Condition (@($continuedInstallerSuccess.Events | Where-Object { $_.status -eq 'Failed' }).Count -eq 0) -Message 'Continued MSI custom-action 1603 must not be reported as failed.'

$downloadFailure = Invoke-AgentScenario -Name 'user-download-http-500' -Mode User -Scenario DownloadHttp500 -PackageId 'cURL.cURL' -PackageName 'cURL' -MaxRetries 1
Assert-True -Condition ($downloadFailure.ExitCode -eq 1) -Message 'HTTP 500 download scenario should exit 1 after the retry budget is exhausted.'
Assert-True -Condition (@($downloadFailure.Events | Where-Object { $_.status -eq 'TransientDownloadFailureRetryScheduled' }).Count -eq 1) -Message 'HTTP 500 download scenario did not schedule exactly one retry.'
Assert-True -Condition (@($downloadFailure.Events | Where-Object { $_.status -eq 'TransientDownloadFailure' -and $_.httpStatus -eq 500 }).Count -eq 1) -Message 'HTTP 500 download scenario did not record terminal transient download evidence.'
Assert-True -Condition (@($downloadFailure.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 2) -Message 'HTTP 500 download scenario should make exactly two named upgrade attempts.'
Assert-True -Condition (@($downloadFailure.Calls | Where-Object { $_.Args -contains '--output' -and $_.Args -contains 'json' }).Count -eq 1) -Message 'HTTP 500 download scenario must not perform post-upgrade inventory verification.'

$noApplicableRoot = Join-Path $runRoot 'user-no-applicable-upgrade'
$noApplicable = Invoke-AgentScenario -Name 'user-no-applicable-upgrade' -Mode User -Scenario NoApplicableUpgrade -PackageId 'cURL.cURL' -PackageName 'cURL' -ScenarioRoot $noApplicableRoot
Assert-True -Condition ($noApplicable.ExitCode -eq 0) -Message 'No-applicable-upgrade scenario should exit 0.'
Assert-True -Condition (@($noApplicable.Events | Where-Object { $_.status -eq 'DeferredNoApplicableUpgrade' }).Count -eq 1) -Message 'No-applicable-upgrade scenario did not record a bounded deferral.'
Assert-True -Condition (@($noApplicable.Events | Where-Object { $_.status -eq 'VerificationFailedStillOffered' }).Count -eq 0) -Message 'No-applicable-upgrade scenario must not be treated as verification failure.'
Assert-True -Condition (-not (Test-Path -LiteralPath (Join-Path $noApplicable.LogRoot 'State\pending-notification.json'))) -Message 'No-applicable-upgrade scenario must not create notification state.'
$noApplicableCached = Invoke-AgentScenario -Name 'user-no-applicable-upgrade' -Mode User -Scenario NoApplicableUpgrade -PackageId 'cURL.cURL' -PackageName 'cURL' -ScenarioRoot $noApplicableRoot
Assert-True -Condition (@($noApplicableCached.Events | Where-Object { $_.status -eq 'DeferredNoApplicableUpgradeCached' }).Count -eq 1) -Message 'No-applicable-upgrade cache did not suppress the second attempt.'
Assert-True -Condition (@($noApplicableCached.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 1) -Message 'No-applicable-upgrade cache should prevent a second named upgrade attempt.'

$manualOnly = Invoke-AgentScenario -Name 'manual-only' -Mode User -Scenario Success -PackageId 'OpenJS.NodeJS' -PackageName 'Node.js' -PackageRules @([pscustomobject]@{ id='OpenJS.NodeJS'; mode='ManualOnly'; allowMajorVersion=$false; processNames=@('node.exe') }) -InstalledVersion '24.18.0' -AvailableVersion '26.4.0'
Assert-True -Condition ($manualOnly.ExitCode -eq 0) -Message 'Manual-only scenario should exit 0.'
Assert-True -Condition (@($manualOnly.Events | Where-Object { $_.status -eq 'DeferredManualOnly' }).Count -eq 1) -Message 'Manual-only scenario did not log DeferredManualOnly.'
Assert-True -Condition (@($manualOnly.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 0) -Message 'Manual-only package attempted an upgrade.'

$majorVersion = Invoke-AgentScenario -Name 'major-version' -Mode User -Scenario Success -PackageId 'Test.MajorPackage' -PackageName 'Test Major Package' -InstalledVersion '1.0.0' -AvailableVersion '2.0.0'
Assert-True -Condition ($majorVersion.ExitCode -eq 0) -Message 'Major-version defer scenario should exit 0.'
Assert-True -Condition (@($majorVersion.Events | Where-Object { $_.status -eq 'DeferredMajorVersion' }).Count -eq 1) -Message 'Major-version scenario did not log DeferredMajorVersion.'
Assert-True -Condition (@($majorVersion.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 0) -Message 'Major-version package attempted an upgrade.'

$quiescence = Invoke-AgentScenario -Name 'quiescence' -Mode User -Scenario Success -PackageId 'Test.QuiescencePackage' -PackageName 'Test Quiescence Package' -PackageRules @([pscustomobject]@{ id='Test.QuiescencePackage'; mode='QuiescenceRequired'; allowMajorVersion=$true; processNames=@('powershell.exe') })
Assert-True -Condition ($quiescence.ExitCode -eq 0) -Message 'Quiescence scenario should exit 0.'
Assert-True -Condition (@($quiescence.Events | Where-Object { $_.status -eq 'DeferredProcessInUse' }).Count -eq 1) -Message 'Quiescence scenario did not log DeferredProcessInUse.'
Assert-True -Condition (@($quiescence.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq 0) -Message 'Quiescence-required package attempted an upgrade.'

foreach ($case in @(
  @{ Name='edge-technology-mismatch'; Scenario='TechnologyMismatch'; Status='DeferredInstallerTechnologyMismatch'; ExitCode=0; Attempts=1 },
  @{ Name='copilot-no-upgrade'; Scenario='NoAvailableUpgrade'; Status='DeferredNoAvailableUpgrade'; ExitCode=0; Attempts=1 },
  @{ Name='chrome-download-timeout'; Scenario='DownloadTimeout'; Status='TransientDownloadFailure'; ExitCode=1; Attempts=2 },
  @{ Name='msi-later-failure'; Scenario='SuccessThen1603'; Status='Failed'; ExitCode=1; Attempts=1 },
  @{ Name='msi-restart'; Scenario='Installer3010'; Status='RestartRequiredPendingVerification'; ExitCode=0; Attempts=1 },
  @{ Name='inventory-failure'; Scenario='InventoryFailure'; Status='InventoryInconclusive'; ExitCode=2; Attempts=0 },
  @{ Name='verification-malformed'; Scenario='VerificationMalformed'; Status='VerificationInconclusive'; ExitCode=1; Attempts=1 },
  @{ Name='installed-package-not-found'; Scenario='InstalledPackageNotFound'; Status='DeferredInstalledPackageNotFound'; ExitCode=0; Attempts=1 },
  @{ Name='native-restart'; Scenario='NativeRestart'; Status='RestartRequiredPendingVerification'; ExitCode=0; Attempts=1 },
  @{ Name='burn-restart'; Scenario='BurnRestart'; Status='RestartRequiredPendingVerification'; ExitCode=0; Attempts=1 }
)) {
  # Arrange / initial assert: the fake case declares one terminal outcome and bounded attempts.
  Assert-True ($case.Attempts -le 2) 'Invalid case retry budget.'
  # Act: run the actual agent control flow against fake winget.
  $result = Invoke-AgentScenario -Name $case.Name -Mode User -Scenario $case.Scenario -PackageId 'Test.Package' -PackageName 'Test Package' -MaxRetries 1
  # Assert: exact status, exit, and installer invocation count.
  Assert-True ($result.ExitCode -eq $case.ExitCode) "$($case.Name): unexpected exit $($result.ExitCode)."
  Assert-True (@($result.Events | Where-Object status -eq $case.Status).Count -eq 1) "$($case.Name): missing $($case.Status)."
  Assert-True (@($result.Calls | Where-Object { $_.Args -contains '--id' }).Count -eq $case.Attempts) "$($case.Name): wrong attempt count."
  if ($case.Scenario -in @('TechnologyMismatch','NoAvailableUpgrade','DownloadTimeout')) {
    Assert-True (@($result.Calls | Where-Object { $_.Args -contains '--output' }).Count -eq 1) "$($case.Name): unexpected post-upgrade verification."
  }
}
$lts = Invoke-AgentScenario -Name 'node-lts-manual-only' -Mode User -Scenario Success -PackageId 'OpenJS.NodeJS.LTS' -PackageName 'Node LTS' -PackageRules @([pscustomobject]@{id='OpenJS.NodeJS.LTS';mode='ManualOnly'})
Assert-True (@($lts.Events | Where-Object status -eq 'DeferredManualOnly').Count -eq 1) 'Node LTS must honor ManualOnly.'
Assert-True (@($lts.Calls | Where-Object { $_.Args -contains '--log' }).Count -eq 0) 'Node LTS must never invoke an installer.'
$native = Invoke-AgentScenario -Name 'edge-native-updater' -Mode User -Scenario Success -PackageId 'Microsoft.Edge' -PackageName 'Edge' -PackageRules @([pscustomobject]@{id='Microsoft.Edge';mode='NativeUpdater'})
Assert-True (@($native.Events | Where-Object status -eq 'DeferredNativeUpdater').Count -eq 1) 'Native updater must remain visible in compliance evidence.'
Assert-True (@($native.Calls | Where-Object { $_.Args -contains '--log' }).Count -eq 0) 'Native updater must not invoke winget installer.'
foreach ($fixture in @(@{Scenario='ReleaseHold';Status='DeferredKnownBadRelease';Attempts=0},@{Scenario='ReleaseHoldUnreadable';Status='DeferredReleaseHoldUnverified';Attempts=0},@{Scenario='ReleaseHoldCorrected';Status='Succeeded';Attempts=1})) {
  $rule = [pscustomobject]@{id='Test.Held';mode='Auto';releaseHolds=@([pscustomobject]@{version='1.1.0';scope='user';installerSha256=('a'*64);until=(Get-Date).AddDays(7).ToString('o');reason='Bad fixture'})}
  Assert-True ($rule.releaseHolds.Count -eq 1) 'Release hold fixture missing.'
  $result = Invoke-AgentScenario -Name $fixture.Scenario -Mode User -Scenario $fixture.Scenario -PackageId 'Test.Held' -PackageName 'Held' -PackageRules @($rule)
  Assert-True (@($result.Events | Where-Object status -eq $fixture.Status).Count -eq 1) 'Release hold outcome mismatch.'
  Assert-True (@($result.Calls | Where-Object { $_.Args -contains '--log' }).Count -eq $fixture.Attempts) 'Release hold allowed an unsafe installer attempt.'
  if ($fixture.Attempts -eq 1) {
    $call = @($result.Calls | Where-Object { $_.Args -contains '--log' })[0]
    Assert-True ($call.Args -contains '--version' -and $call.Args -contains '1.1.0') 'Upgrade must pin the inspected target version.'
  }
}
$attemptFiles = @(Get-ChildItem -LiteralPath (Join-Path $downloadFailure.LogRoot 'Logs') -Filter '*.attempt.json')
Assert-True ($attemptFiles.Count -eq 2) 'Download retries must preserve both unique evidence files.'
$attempts = @($attemptFiles | ForEach-Object { Get-Content $_.FullName -Raw | ConvertFrom-Json })
Assert-True (@($attempts | Select-Object -ExpandProperty attempt -Unique).Count -eq 2) 'Evidence must identify each retry number.'

$healthRoot = Join-Path $runRoot 'health'
$healthPolicy = Join-Path $healthRoot 'policy.json'
$healthSnapshot = Join-Path $healthRoot 'processes.json'
New-Item -ItemType Directory -Path $healthRoot -Force | Out-Null
Write-Policy -Path $healthPolicy -AllowedId 'Test.HealthPackage'
$sentinel = Start-Process -FilePath powershell.exe -ArgumentList '-NoProfile', '-Command', 'Start-Sleep -Seconds 30' -WindowStyle Hidden -PassThru
try {
  @([pscustomobject]@{
    ProcessId = $sentinel.Id
    Name = 'powershell.exe'
    CommandLine = 'powershell.exe -NoProfile -File "\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent\danmade-patch-agent.ps1" -Mode Machine'
    CreationDate = (Get-Date).AddHours(-3).ToString('o')
  }) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $healthSnapshot -Encoding UTF8
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $healthScriptPath -PolicyPath $healthPolicy -LogRoot $healthRoot -ProcessSnapshotPath $healthSnapshot | Out-Null
  Assert-True -Condition ($LASTEXITCODE -eq 0) -Message 'Health scenario should exit 0.'
  Assert-True -Condition (-not $sentinel.HasExited) -Message 'Health scenario must not terminate the matching process.'
  $healthEvents = @(Get-Content -LiteralPath (Join-Path $healthRoot 'Events\patch-agent.jsonl') | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-Json })
  Assert-True -Condition (@($healthEvents | Where-Object { $_.status -eq 'StaleAgentProcessDetected' -and $_.processId -eq $sentinel.Id }).Count -eq 1) -Message 'Health scenario did not report the stale matching process.'
  Assert-True (@($healthEvents | Where-Object status -eq 'HealthCheckCompleted').Count -eq 1) 'Health heartbeat was not recorded.'
  $tasksSnapshot = Join-Path $healthRoot 'tasks.json'
  @([pscustomobject]@{Name='Danmade Patch Agent - Machine';State='Ready';LastRunTime=(Get-Date).ToString('o');LastTaskResult=4294770688}) | ConvertTo-Json | Set-Content $tasksSnapshot
  New-Item -ItemType Directory -Path (Join-Path $healthRoot 'State') -Force | Out-Null
  [pscustomobject]@{processId=999999;runId='orphan';startedAt=(Get-Date).AddDays(-2).ToString('o')} | ConvertTo-Json | Set-Content (Join-Path $healthRoot 'State\active-run.json')
  @([pscustomobject]@{ProcessId=$sentinel.Id;Name='powershell.exe';CommandLine='powershell.exe -File Invoke-DanmadePatchAgentCached.ps1 -Mode Machine';CreationDate=(Get-Date).AddHours(-3).ToString('o')}) | ConvertTo-Json | Set-Content $healthSnapshot
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $healthScriptPath -PolicyPath $healthPolicy -LogRoot $healthRoot -ProcessSnapshotPath $healthSnapshot -TaskSnapshotPath $tasksSnapshot | Out-Null
  Assert-True ($LASTEXITCODE -eq 0) 'Health task-readback scenario failed.'
  $healthEvents = @(Get-Content (Join-Path $healthRoot 'Events\patch-agent.jsonl') | ForEach-Object { $_ | ConvertFrom-Json })
  Assert-True (@($healthEvents | Where-Object status -eq 'OrphanedActiveRunDetected').Count -eq 1) 'Health did not report orphaned active state.'
  Assert-True (@($healthEvents | Where-Object status -eq 'AgentTaskNeedsReview').Count -eq 1) 'Health did not report failed task startup.'
  Assert-True (@($healthEvents | Where-Object status -eq 'StaleAgentProcessDetected').Count -eq 2) 'Health must match the cached bootstrap process too.'
  Assert-True (-not $sentinel.HasExited) 'Additional health checks must not terminate the process.'
} finally {
  if (-not $sentinel.HasExited) { Stop-Process -Id $sentinel.Id -Force }
}

[pscustomobject]@{
  Status = 'Passed'
  RunRoot = $runRoot
  MachineScenario = if ($machineSuccess.Skipped) { $machineSuccess.SkipReason } else { 'Passed' }
  Scenarios = @($machineSuccess.Name, $userSuccess.Name, $stillOffered.Name, $installer1605.Name, $installer1603.Name, $continuedInstallerSuccess.Name, $downloadFailure.Name, $noApplicable.Name, $manualOnly.Name, $majorVersion.Name, $quiescence.Name, 'health-stale-process', 'health-orphan-and-launch-failure', 'edge-technology-mismatch','copilot-no-upgrade','chrome-download-timeout','msi-later-failure','msi-restart','inventory-failure','verification-malformed','installed-package-not-found','native-restart','burn-restart','node-lts-manual-only','edge-native-updater','ReleaseHold','ReleaseHoldUnreadable','ReleaseHoldCorrected')
} | ConvertTo-Json -Depth 5
