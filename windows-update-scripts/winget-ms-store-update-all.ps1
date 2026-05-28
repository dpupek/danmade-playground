# =====================================================================
# Winget updater (interactive selection)
# - Lists upgrades with numbers so the user can pick specific items
# - Runs winget only for the chosen packages
# - Still re-launches elevated to catch machine-scope installs
# =====================================================================

param(
  [switch]$MachinePhase, # internal flag when relaunching elevated
  [switch]$NonSilentPhase, # internal flag for elevated reruns without --silent
  [string]$SelectedList,  # comma-separated package ids when elevated
  [switch]$TestModeAddPowerShell,  # inject Microsoft.PowerShell into the picker to test deferred handling
  [switch]$CleanupPortablePackages # select installed winget portable packages to uninstall with --purge
)

$script:InstallerFailures = @()
$script:FailureHistory = @()
$script:RestartRequiredPackages = @()
$script:CachedUninstallEntries = $null
$script:DeferredPowerShellPackageIds = @('Microsoft.PowerShell')

# Args for single-package upgrade runs. Do not include --all/--recurse (those are for bulk upgrades).
$WingetCommonArgs = @(
  "--include-unknown",
  "--silent",
  "--exact",
  "--accept-package-agreements",
  "--accept-source-agreements"
)

function Normalize-DisplayText {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $normalized = $Text.ToLowerInvariant()
  $normalized = $normalized -replace '\b\d+(?:\.\d+)+(?:[-_]\d+)?\b', ' '
  $normalized = $normalized -replace '[^a-z0-9]+', ' '
  $normalized = $normalized -replace '\s+', ' '
  return $normalized.Trim()
}

function Test-IsDeferredPowerShellPackage {
  param([object]$Package)

  if (-not $Package) { return $false }
  if ([string]::IsNullOrWhiteSpace([string]$Package.Id)) { return $false }

  return $script:DeferredPowerShellPackageIds -contains [string]$Package.Id
}

function Get-TestPowerShellPackage {
  return [pscustomobject]@{
    Name         = 'PowerShell 7 (test mode)'
    Id           = 'Microsoft.PowerShell'
    Installed    = 'test-mode'
    Available    = 'test-mode'
    Source       = 'winget'
    InstallScope = 'Machine'
    ScopeNote    = 'Test mode injected this entry so you can exercise the deferred PowerShell helper path even when no real upgrade is currently available.'
    Provider     = 'Winget'
  }
}

function Get-CachedUninstallEntries {
  if ($script:CachedUninstallEntries) { return $script:CachedUninstallEntries }

  function Get-OptionalPropertyValue {
    param(
      [object]$Object,
      [string]$Name
    )

    if (-not $Object) { return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if (-not $prop) { return $null }
    return [string]$prop.Value
  }

  $paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
  )

  $script:CachedUninstallEntries = @(
    Get-ItemProperty $paths -ErrorAction SilentlyContinue |
      Where-Object {
        $_.PSObject.Properties['DisplayName'] -and
        -not [string]::IsNullOrWhiteSpace([string]$_.PSObject.Properties['DisplayName'].Value)
      } |
      ForEach-Object {
        $productCode = $null
        if ($_.PSChildName) {
          $productCode = [string]$_.PSChildName
        } elseif ($_.PSPath) {
          $productCode = [string]$_.PSPath
          if ($productCode -match '\\([^\\]+)$') {
            $productCode = $Matches[1]
          }
        }

        [pscustomobject]@{
          DisplayName      = Get-OptionalPropertyValue -Object $_ -Name 'DisplayName'
          DisplayNameNorm  = Normalize-DisplayText -Text (Get-OptionalPropertyValue -Object $_ -Name 'DisplayName')
          DisplayVersion   = Get-OptionalPropertyValue -Object $_ -Name 'DisplayVersion'
          InstallLocation  = Get-OptionalPropertyValue -Object $_ -Name 'InstallLocation'
          UninstallString  = Get-OptionalPropertyValue -Object $_ -Name 'UninstallString'
          ProductCode      = $productCode
          PSPath           = [string]$_.PSPath
        }
      }
  )

  return $script:CachedUninstallEntries
}

function Get-ScopeFromEntry {
  param([object]$Entry)

  $psPath = [string]$Entry.PSPath
  $installLocation = ([string]$Entry.InstallLocation).Trim('"')
  $uninstallString = ([string]$Entry.UninstallString).Trim('"')
  $localAppData = [regex]::Escape($env:LOCALAPPDATA)
  $userProfile = [regex]::Escape($env:USERPROFILE)
  $programFiles = [regex]::Escape($env:ProgramFiles)
  $programFilesX86Path = ${env:ProgramFiles(x86)}
  $programFilesX86 = if ($programFilesX86Path) { [regex]::Escape($programFilesX86Path) } else { $null }

  if ($psPath -match 'HKEY_CURRENT_USER') { return 'User' }
  if ($psPath -match 'HKEY_LOCAL_MACHINE') { return 'Machine' }
  if ($installLocation -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($uninstallString -match "^(?:$localAppData|$userProfile)") { return 'User' }
  if ($installLocation -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $installLocation -match "^$programFilesX86") { return 'Machine' }
  if ($uninstallString -match "^$programFiles") { return 'Machine' }
  if ($programFilesX86 -and $uninstallString -match "^$programFilesX86") { return 'Machine' }

  return 'Unknown'
}

function Resolve-PackageInstallMetadata {
  param([object]$Package)

  if (-not $Package -or [string]::IsNullOrWhiteSpace($Package.Name)) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $packageName = [string]$Package.Name
  $packageNameNorm = Normalize-DisplayText -Text $packageName
  if (-not $packageNameNorm) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $bestMatch = $null
  $bestScore = -1
  foreach ($entry in @(Get-CachedUninstallEntries)) {
    if (-not $entry.DisplayNameNorm) { continue }

    $score = -1
    if ($entry.DisplayName -eq $packageName) {
      $score = 140
    } elseif ($entry.DisplayNameNorm -eq $packageNameNorm) {
      $score = 120
    } elseif ($entry.DisplayNameNorm.StartsWith($packageNameNorm)) {
      $score = 100
    } elseif ($entry.DisplayNameNorm.Contains($packageNameNorm)) {
      $score = 90
    } elseif ($packageNameNorm.Contains($entry.DisplayNameNorm)) {
      $score = 70
    }

    if ($score -lt 0) { continue }
    if ($score -gt $bestScore) {
      $bestScore = $score
      $bestMatch = $entry
    }
  }

  if (-not $bestMatch) {
    return [pscustomobject]@{
      Scope = 'Unknown'
      Note  = $null
    }
  }

  $scope = Get-ScopeFromEntry -Entry $bestMatch
  $note = switch ($scope) {
    'User' { 'Per-user install; keep this in the current session.' }
    'Machine' { 'Machine install; elevated retry can apply.' }
    default { $null }
  }

  if (Test-IsDeferredPowerShellPackage -Package $Package) {
    $deferredNote = 'PowerShell 7 updates are deferred into a separate helper after all pwsh.exe sessions close.'
    if ([string]::IsNullOrWhiteSpace($note)) {
      $note = $deferredNote
    } else {
      $note = "$note $deferredNote"
    }
  }

  return [pscustomobject]@{
    Scope = $scope
    Note  = $note
  }
}

function Compare-VersionText {
  param(
    [string]$Left,
    [string]$Right
  )

  if ([string]::IsNullOrWhiteSpace($Left) -and [string]::IsNullOrWhiteSpace($Right)) { return 0 }
  if ([string]::IsNullOrWhiteSpace($Left)) { return -1 }
  if ([string]::IsNullOrWhiteSpace($Right)) { return 1 }

  $leftVersion = $null
  $rightVersion = $null
  $leftParsed = [version]::TryParse($Left, [ref]$leftVersion)
  $rightParsed = [version]::TryParse($Right, [ref]$rightVersion)
  if ($leftParsed -and $rightParsed) {
    return $leftVersion.CompareTo($rightVersion)
  }

  return [string]::Compare($Left, $Right, [System.StringComparison]::OrdinalIgnoreCase)
}

function Ensure-WingetMsStoreSource {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Warning "winget not found."
    return
  }
  try {
    $sources = winget source list 2>$null
    if ($sources -notmatch 'msstore') {
      Write-Host "Adding winget msstore source..."
      winget source add --name msstore --arg https://storeedgefd.dsx.mp.microsoft.com/v9.0 | Out-Null
    }
    winget source update | Out-Null
  } catch {
    Write-Warning "Unable to verify/update winget sources: $($_.Exception.Message)"
  }
}

function Normalize-WingetPackageId {
  param([string]$RawId)

  if ([string]::IsNullOrWhiteSpace($RawId)) { return $null }

  $candidate = $RawId.Trim()

  # Strip leading non-identifier artifacts from mojibake/truncated table output.
  $candidate = $candidate -replace '^[^A-Za-z0-9]+', ''

  # Keep the first token that looks like a winget package identifier.
  # Support both dotted IDs (e.g., Microsoft.Teams) and non-dotted store IDs.
  $match = [regex]::Match($candidate, '[A-Za-z0-9][A-Za-z0-9._+-]*')
  if ($match.Success) { return $match.Value }

  return $candidate
}

function Get-WingetUpgradeList {
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return @() }

  function Get-FirstValue {
    param(
      [object]$Object,
      [string[]]$Names
    )

    if (-not $Object) { return $null }
    $propNames = $Object.PSObject.Properties.Name
    foreach ($name in $Names) {
      if ($propNames -contains $name) {
        $value = $Object.$name
        if ($null -eq $value) { continue }
        $text = [string]$value
        if (-not [string]::IsNullOrWhiteSpace($text)) { return $text.Trim() }
      }
    }
    return $null
  }

  function Add-WingetJsonPackages {
    param(
      [object]$Node,
      [System.Collections.Generic.List[object]]$Collector,
      [string]$InheritedSource
    )

    if ($null -eq $Node) { return }

    if ($Node -is [System.Collections.IEnumerable] -and -not ($Node -is [string])) {
      foreach ($child in $Node) {
        Add-WingetJsonPackages -Node $child -Collector $Collector -InheritedSource $InheritedSource
      }
      return
    }

    $nodeProps = $Node.PSObject.Properties.Name
    if (-not $nodeProps) { return }

    $sourceFromNode = $InheritedSource
    if ($nodeProps -contains 'Source') {
      $candidate = Get-FirstValue -Object $Node -Names @('Source')
      if ($candidate) { $sourceFromNode = $candidate }
    } elseif (($nodeProps -contains 'Name') -and ($nodeProps -contains 'Packages')) {
      # In many winget JSON payloads, each source entry is { Name, Packages[] }.
      $candidate = Get-FirstValue -Object $Node -Names @('Name')
      if ($candidate) { $sourceFromNode = $candidate }
    } elseif ($nodeProps -contains 'Details') {
      $candidate = Get-FirstValue -Object $Node.Details -Names @('Name')
      if ($candidate) { $sourceFromNode = $candidate }
    }

    $idRaw = Get-FirstValue -Object $Node -Names @('PackageIdentifier', 'Id')
    $id = Normalize-WingetPackageId -RawId $idRaw
    if ($id) {
      $name = Get-FirstValue -Object $Node -Names @('PackageName', 'Name')
      if (-not $name) { $name = $id }

      $installed = Get-FirstValue -Object $Node -Names @('InstalledVersion', 'Version', 'Installed')
      $available = Get-FirstValue -Object $Node -Names @('AvailableVersion', 'Available', 'LatestVersion')
      $source = Get-FirstValue -Object $Node -Names @('Source', 'Repository', 'Origin')
      if (-not $source) { $source = $sourceFromNode }

      if ($available -or $installed) {
        $Collector.Add([pscustomobject]@{
          Name       = $name
          Id         = $id
          Installed  = $installed
          Available  = $available
          Source     = $source
        })
      }
    }

    foreach ($prop in $Node.PSObject.Properties) {
      if ($null -eq $prop.Value) { continue }
      if ($prop.Value -is [string]) { continue }
      Add-WingetJsonPackages -Node $prop.Value -Collector $Collector -InheritedSource $sourceFromNode
    }
  }

  # Prefer JSON for reliable parsing; fall back to table parsing.
  try {
    $json = winget upgrade --include-unknown --output json 2>$null
    if ($json) {
      $jsonText = ($json -join "`n").Trim()
      if ($jsonText) {
        $firstBraceIndex = $jsonText.IndexOfAny(@('[', '{'))
        if ($firstBraceIndex -ge 0) {
          $jsonBody = $jsonText.Substring($firstBraceIndex)
          $data = $jsonBody | ConvertFrom-Json -ErrorAction Stop
          $collected = New-Object System.Collections.Generic.List[object]
          Add-WingetJsonPackages -Node $data -Collector $collected -InheritedSource $null

          if ($collected.Count -gt 0) {
            $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            $deduped = foreach ($pkg in $collected) {
              if (-not $pkg.Id) { continue }
              if ($seen.Add($pkg.Id)) { $pkg }
            }
            if ($deduped) {
              return @(
                foreach ($pkg in $deduped) {
                  $metadata = Resolve-PackageInstallMetadata -Package $pkg
                  $pkg | Add-Member -NotePropertyName InstallScope -NotePropertyValue $metadata.Scope -Force
                  $pkg | Add-Member -NotePropertyName ScopeNote -NotePropertyValue $metadata.Note -Force
                  $pkg | Add-Member -NotePropertyName Provider -NotePropertyValue 'Winget' -Force
                  $pkg
                }
              )
            }
          }
        }
      }
    }
  } catch {
    # winget versions prior to full JSON support often emit "Invalid JSON primitive"; fall back quietly.
    Write-Verbose "winget JSON parse failed; using text fallback. $_"
  }

  try {
    $output = winget upgrade --include-unknown
    $lines = $output -split "`r?`n"
    $header = $lines | Where-Object { $_ -match '^Name\s+Id\s+Version\s+Available\s+Source' } | Select-Object -First 1
    if (-not $header) { return @() }

    $headerIndex = [array]::IndexOf($lines, $header)
    $candidateLines = $lines | Select-Object -Skip ($headerIndex + 2)
    # Some winget rows collapse to single-space separators when columns are full width.
    # Accept both dotted and non-dotted package identifiers.
    $rowPattern = '^(?<Name>.+?)\s+(?<Id>[A-Za-z0-9][A-Za-z0-9._+-]*)\s+(?<Installed>\S+)\s+(?<Available>\S+)\s+(?<Source>\S+)\s*$'

    $parsed = foreach ($line in $candidateLines) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      if ($line -match '^[-\s]+$') { continue }
      if ($line -match '^\d+\s+upgrades available\.?$') { continue }
      if ($line -match '^The following packages have an upgrade available') { continue }

      if ($line -notmatch $rowPattern) { continue }
      $name = $matches['Name'].Trim()
      $id = Normalize-WingetPackageId -RawId $matches['Id']
      $installed = $matches['Installed'].Trim()
      $available = $matches['Available'].Trim()
      $source = $matches['Source'].Trim()

      if (-not $id) { continue }

      [pscustomobject]@{
        Name      = $name
        Id        = $id
        Installed = $installed
        Available = $available
        Source    = $source
      }
    }
    return @(
      foreach ($pkg in $parsed) {
        $metadata = Resolve-PackageInstallMetadata -Package $pkg
        $pkg | Add-Member -NotePropertyName InstallScope -NotePropertyValue $metadata.Scope -Force
        $pkg | Add-Member -NotePropertyName ScopeNote -NotePropertyValue $metadata.Note -Force
        $pkg | Add-Member -NotePropertyName Provider -NotePropertyValue 'Winget' -Force
        $pkg
      }
    )
  } catch {
    Write-Warning "winget upgrade listing failed: $($_.Exception.Message)"
    return @()
  }
}

function Get-PythonInstallManagerUpgradeList {
  $py = Get-Command py -ErrorAction SilentlyContinue
  if (-not $py) { return @() }

  $installedJson = $null
  try {
    $installedJson = & $py.Path list --only-managed -f json 2>$null
  } catch {
    return @()
  }
  if (-not $installedJson) { return @() }

  try {
    $installedData = ($installedJson -join "`n") | ConvertFrom-Json -ErrorAction Stop
  } catch {
    Write-Verbose "Python Install Manager JSON parse failed. $_"
    return @()
  }

  $installedVersions = @($installedData.versions)
  if (-not $installedVersions -or $installedVersions.Count -eq 0) { return @() }

  $results = New-Object System.Collections.Generic.List[object]
  foreach ($runtime in $installedVersions) {
    $tag = [string]$runtime.tag
    if ([string]::IsNullOrWhiteSpace($tag)) { continue }

    $installed = [string]$runtime.'sort-version'
    $displayName = [string]$runtime.'display-name'
    if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = "Python $tag" }

    $latest = $null
    try {
      $onlineJson = & $py.Path list --online -f json $tag 2>$null
      if ($onlineJson) {
        $onlineData = ($onlineJson -join "`n") | ConvertFrom-Json -ErrorAction Stop
        $onlineMatches = @($onlineData.versions | Where-Object { [string]$_.tag -eq $tag })
        if (-not $onlineMatches -or $onlineMatches.Count -eq 0) {
          $onlineMatches = @($onlineData.versions)
        }

        foreach ($candidate in $onlineMatches) {
          $candidateVersion = [string]$candidate.'sort-version'
          if ([string]::IsNullOrWhiteSpace($candidateVersion)) { continue }
          if (-not $latest -or (Compare-VersionText -Left $candidateVersion -Right $latest) -gt 0) {
            $latest = $candidateVersion
          }
        }
      }
    } catch {
      Write-Verbose "Python Install Manager online lookup failed for tag '$tag'. $_"
    }

    if ([string]::IsNullOrWhiteSpace($latest)) { continue }
    if ((Compare-VersionText -Left $latest -Right $installed) -le 0) { continue }

    $results.Add([pscustomobject]@{
      Name       = $displayName
      Id         = "PythonInstallManager::$tag"
      Installed  = $installed
      Available  = $latest
      Source     = 'python-install-manager'
      InstallScope = 'User'
      ScopeNote  = "Managed by Python Install Manager. Updates run via 'py install --update $tag' in this session."
      Provider   = 'PythonInstallManager'
      UpdateTag  = $tag
    })
  }

  return $results.ToArray()
}

function Get-UpgradeList {
  $all = New-Object System.Collections.Generic.List[object]
  foreach ($pkg in @(Get-WingetUpgradeList)) { $all.Add($pkg) }
  foreach ($pkg in @(Get-PythonInstallManagerUpgradeList)) { $all.Add($pkg) }

  if ($TestModeAddPowerShell) {
    $hasPowerShell = @($all | Where-Object { $_.Id -eq 'Microsoft.PowerShell' }).Count -gt 0
    if (-not $hasPowerShell) {
      $all.Add((Get-TestPowerShellPackage))
      Write-Host "Test mode: injected Microsoft.PowerShell into the upgrade picker." -ForegroundColor Yellow
    } else {
      Write-Host "Test mode: Microsoft.PowerShell is already present in the upgrade picker." -ForegroundColor Yellow
    }
  }

  return $all.ToArray()
}

function Test-IsWingetPortablePackageEntry {
  param([object]$Entry)

  if (-not $Entry) { return $false }
  if ([string]::IsNullOrWhiteSpace([string]$Entry.ProductCode)) { return $false }
  if ([string]::IsNullOrWhiteSpace([string]$Entry.UninstallString)) { return $false }
  if ([string]::IsNullOrWhiteSpace([string]$Entry.InstallLocation)) { return $false }

  $installLocation = ([string]$Entry.InstallLocation).Trim('"')
  $localPackagesRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
  $programFilesWindowsApps = if ($env:ProgramFiles) { Join-Path $env:ProgramFiles 'WindowsApps' } else { $null }
  $portableRoots = @($localPackagesRoot)
  if ($programFilesWindowsApps) {
    $portableRoots += $programFilesWindowsApps
  }

  $underPortableRoot = $false
  foreach ($root in @($portableRoots | Where-Object { $_ })) {
    if ($installLocation.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
      $underPortableRoot = $true
      break
    }
  }

  if (-not $underPortableRoot) { return $false }

  return ([string]$Entry.UninstallString) -match '(^|\s)winget(\.exe)?\s+uninstall\s+--product-code\s+'
}

function Get-WingetPortablePackageList {
  $packages = New-Object System.Collections.Generic.List[object]

  foreach ($entry in @(Get-CachedUninstallEntries)) {
    if (-not (Test-IsWingetPortablePackageEntry -Entry $entry)) { continue }

    $scope = Get-ScopeFromEntry -Entry $entry
    $note = 'Portable winget package. Cleanup mode removes the package registration and package directory with --purge.'
    if ($scope -eq 'Machine') {
      $note = 'Machine-scope portable winget package. Cleanup mode uses --purge and may require elevation.'
    }

    $packages.Add([pscustomobject]@{
      Name            = [string]$entry.DisplayName
      Id              = [string]$entry.ProductCode
      ProductCode     = [string]$entry.ProductCode
      Installed       = [string]$entry.DisplayVersion
      Available       = $null
      Source          = 'winget'
      InstallScope    = $scope
      ScopeNote       = $note
      Provider        = 'WingetPortableCleanup'
      InstallLocation = [string]$entry.InstallLocation
    })
  }

  return @(
    $packages |
      Sort-Object @{ Expression = { if ($_.InstallScope -eq 'Machine') { 0 } elseif ($_.InstallScope -eq 'User') { 1 } else { 2 } } }, Name, Installed
  )
}

function Show-NumberedUpgrades($packages) {
  Write-Host "`n=== Available upgrade candidates ===`n"
  $machineCount = @($packages | Where-Object InstallScope -eq 'Machine').Count
  $userCount = @($packages | Where-Object InstallScope -eq 'User').Count
  $unknownCount = @($packages | Where-Object { $_.InstallScope -ne 'Machine' -and $_.InstallScope -ne 'User' }).Count

  Write-Host ("{0} package(s): {1} machine, {2} user, {3} unknown`n" -f $packages.Count, $machineCount, $userCount, $unknownCount) -ForegroundColor DarkCyan
  Write-Host "Note: [User] items stay in the current session; elevated retries only help [Machine] items.`n" -ForegroundColor DarkYellow

  $groupedPackages = @(
    [pscustomobject]@{ Label = 'Machine-scope updates'; Items = @($packages | Where-Object InstallScope -eq 'Machine') }
    [pscustomobject]@{ Label = 'User-scope updates'; Items = @($packages | Where-Object InstallScope -eq 'User') }
    [pscustomobject]@{ Label = 'Other updates'; Items = @($packages | Where-Object { $_.InstallScope -ne 'Machine' -and $_.InstallScope -ne 'User' }) }
  )

  $indexWidth = [Math]::Max(2, [string]$packages.Count).Length
  $nameWidth = 34
  $versionWidth = 14
  $i = 1

  foreach ($group in $groupedPackages) {
    if (-not $group.Items.Count) { continue }

    Write-Host $group.Label -ForegroundColor Cyan
    foreach ($pkg in $group.Items) {
      $installed = if ($pkg.Installed) { [string]$pkg.Installed } else { 'unknown' }
      $available = if ($pkg.Available) { [string]$pkg.Available } else { 'unknown' }
      $displayName = [string]$pkg.Name
      if ($displayName.Length -gt $nameWidth) {
        $displayName = $displayName.Substring(0, $nameWidth - 3) + '...'
      }

      $badges = switch ($pkg.InstallScope) {
        'User' { @('[User]') }
        'Machine' { @('[Machine]') }
        default { @('[Unknown]') }
      }
      if ($pkg.Source -and $pkg.Source -ne 'winget') { $badges += "[{0}]" -f $pkg.Source }

      $primaryLine = "[{0}] {1} {2} -> {3} {4}" -f `
        $i.ToString().PadLeft($indexWidth),
        $displayName.PadRight($nameWidth),
        $installed.PadRight($versionWidth),
        $available.PadRight($versionWidth),
        ($badges -join ' ')

      Write-Host $primaryLine

      $provider = if ($pkg.Provider) { [string]$pkg.Provider } else { 'Winget' }
      $details = @("Provider: $provider", "Id: $($pkg.Id)")
      if ($pkg.UpdateTag) { $details += "Update tag: $($pkg.UpdateTag)" }
      if ($pkg.ScopeNote) { $details += "Details: $($pkg.ScopeNote)" }
      Write-Host ("     " + ($details -join ' | ')) -ForegroundColor DarkGray
      $i++
    }

    Write-Host ""
  }
}

function Test-OutGridViewAvailable {
  return [bool](Get-Command Out-GridView -ErrorAction SilentlyContinue)
}

function Select-UpgradesWithGrid([object[]]$Packages) {
  Write-Host ""
  Write-Host "Opening the upgrade picker..." -ForegroundColor Cyan
  Write-Host "Use Ctrl+Click for nonconsecutive rows, Shift+Click for ranges, and Ctrl+A for all rows. Choose OK to continue." -ForegroundColor DarkYellow

  $rows = @(
    $index = 1
    foreach ($pkg in @($Packages)) {
      [pscustomobject]@{
        Index      = $index
        Name       = [string]$pkg.Name
        Installed  = if ($pkg.Installed) { [string]$pkg.Installed } else { 'unknown' }
        Available  = if ($pkg.Available) { [string]$pkg.Available } else { 'unknown' }
        Scope      = if ($pkg.InstallScope) { [string]$pkg.InstallScope } else { 'Unknown' }
        Elevation  = if ($pkg.InstallScope -eq 'User') { 'Current session only' } else { 'Elevated retry supported' }
        Source     = if ($pkg.Source) { [string]$pkg.Source } else { '' }
        Id         = [string]$pkg.Id
        Note       = if ($pkg.ScopeNote) { [string]$pkg.ScopeNote } else { '' }
      }
      $index++
    }
  )

  $selectedRows = @(
    $rows |
      Out-GridView -Title "Select available upgrades | Ctrl+Click: multi-select | Shift+Click: range | Ctrl+A: all | OK: continue" -OutputMode Multiple
  )

  if (-not $selectedRows -or $selectedRows.Count -eq 0) { return @() }

  $selectedPackages = foreach ($row in $selectedRows) {
    $selectedIndex = [int]$row.Index
    if ($selectedIndex -ge 1 -and $selectedIndex -le $Packages.Count) {
      $Packages[$selectedIndex - 1]
    }
  }

  return @($selectedPackages | Where-Object { $_ })
}

function Prompt-SelectionText([int]$Count) {
  $input = Read-Host "Select packages: 1,3,5 / 2-4 / all / Enter to skip"
  if ([string]::IsNullOrWhiteSpace($input)) { return @() }

  $tokens = @()
  foreach ($t in ($input -split '[,\s]+')) { if ($t) { $tokens += [string]$t } }

  if ($tokens.Count -eq 1 -and ([string]$tokens[0]).ToLowerInvariant() -eq 'all') {
    return 1..$Count
  }

  $selected = New-Object System.Collections.Generic.List[int]

  foreach ($token in $tokens) {
    if ($token -match '^([0-9]+)-([0-9]+)$') {
      $start = [int]$matches[1]
      $end   = [int]$matches[2]
      if ($end -lt $start) { Write-Warning "Ignoring reversed range: $token"; continue }
      foreach ($n in $start..$end) {
        if ($n -lt 1 -or $n -gt $Count) { Write-Warning "Ignoring out-of-range selection: $n"; continue }
        if (-not $selected.Contains($n)) { $selected.Add($n) }
      }
      continue
    }

    $number = 0
    if (-not [int]::TryParse($token, [ref]$number)) {
      Write-Warning "Ignoring non-numeric entry: '$token'"
      continue
    }
    if ($number -lt 1 -or $number -gt $Count) {
      Write-Warning "Ignoring out-of-range selection: $number"
      continue
    }
    if (-not $selected.Contains($number)) { $selected.Add($number) }
  }

  return $selected
}

function Select-Upgrades([object[]]$Packages) {
  if (Test-OutGridViewAvailable) {
    return @(Select-UpgradesWithGrid -Packages $Packages)
  }

  Write-Warning "Out-GridView is not available in this session. Falling back to text selection."
  Show-NumberedUpgrades -packages $Packages
  $selectedIndexes = Prompt-SelectionText -Count $Packages.Count
  if (-not $selectedIndexes -or $selectedIndexes.Count -eq 0) { return @() }

  return @(
    foreach ($idx in $selectedIndexes) {
      $Packages[$idx - 1]
    }
  )
}

function Show-NumberedPortablePackages([object[]]$Packages) {
  Write-Host "`n=== Installed winget portable packages ===`n"
  $machineCount = @($Packages | Where-Object InstallScope -eq 'Machine').Count
  $userCount = @($Packages | Where-Object InstallScope -eq 'User').Count
  $unknownCount = @($Packages | Where-Object { $_.InstallScope -ne 'Machine' -and $_.InstallScope -ne 'User' }).Count

  Write-Host ("{0} package(s): {1} machine, {2} user, {3} unknown`n" -f $Packages.Count, $machineCount, $userCount, $unknownCount) -ForegroundColor DarkCyan
  Write-Host "Cleanup mode uses winget uninstall --purge, which removes the package directory for the selected portable package.`n" -ForegroundColor DarkYellow

  $groupedPackages = @(
    [pscustomobject]@{ Label = 'Machine-scope packages'; Items = @($Packages | Where-Object InstallScope -eq 'Machine') }
    [pscustomobject]@{ Label = 'User-scope packages'; Items = @($Packages | Where-Object InstallScope -eq 'User') }
    [pscustomobject]@{ Label = 'Other packages'; Items = @($Packages | Where-Object { $_.InstallScope -ne 'Machine' -and $_.InstallScope -ne 'User' }) }
  )

  $indexWidth = [Math]::Max(2, [string]$Packages.Count).Length
  $nameWidth = 34
  $versionWidth = 16
  $i = 1

  foreach ($group in $groupedPackages) {
    if (-not $group.Items.Count) { continue }

    Write-Host $group.Label -ForegroundColor Cyan
    foreach ($pkg in $group.Items) {
      $installed = if ($pkg.Installed) { [string]$pkg.Installed } else { 'unknown' }
      $displayName = [string]$pkg.Name
      if ($displayName.Length -gt $nameWidth) {
        $displayName = $displayName.Substring(0, $nameWidth - 3) + '...'
      }

      $badge = switch ($pkg.InstallScope) {
        'User' { '[User]' }
        'Machine' { '[Machine]' }
        default { '[Unknown]' }
      }

      $primaryLine = "[{0}] {1} {2} {3}" -f `
        $i.ToString().PadLeft($indexWidth),
        $displayName.PadRight($nameWidth),
        $installed.PadRight($versionWidth),
        $badge

      Write-Host $primaryLine
      $details = @("Product code: $($pkg.ProductCode)")
      if ($pkg.InstallLocation) { $details += "Install: $($pkg.InstallLocation)" }
      if ($pkg.ScopeNote) { $details += "Details: $($pkg.ScopeNote)" }
      Write-Host ("     " + ($details -join ' | ')) -ForegroundColor DarkGray
      $i++
    }

    Write-Host ""
  }
}

function Select-WingetPortablePackagesWithGrid([object[]]$Packages) {
  Write-Host ""
  Write-Host "Opening the portable package cleanup picker..." -ForegroundColor Cyan
  Write-Host "Use Ctrl+Click for nonconsecutive rows, Shift+Click for ranges, and Ctrl+A for all rows. Choose OK to continue." -ForegroundColor DarkYellow

  $rows = @(
    $index = 1
    foreach ($pkg in @($Packages)) {
      [pscustomobject]@{
        Index       = $index
        Name        = [string]$pkg.Name
        Version     = if ($pkg.Installed) { [string]$pkg.Installed } else { 'unknown' }
        Scope       = if ($pkg.InstallScope) { [string]$pkg.InstallScope } else { 'Unknown' }
        ProductCode = [string]$pkg.ProductCode
        InstallPath = if ($pkg.InstallLocation) { [string]$pkg.InstallLocation } else { '' }
        Note        = if ($pkg.ScopeNote) { [string]$pkg.ScopeNote } else { '' }
      }
      $index++
    }
  )

  $selectedRows = @(
    $rows |
      Out-GridView -Title "Select installed portable packages to remove | Ctrl+Click: multi-select | Shift+Click: range | Ctrl+A: all | OK: continue" -OutputMode Multiple
  )

  if (-not $selectedRows -or $selectedRows.Count -eq 0) { return @() }

  $selectedPackages = foreach ($row in $selectedRows) {
    $selectedIndex = [int]$row.Index
    if ($selectedIndex -ge 1 -and $selectedIndex -le $Packages.Count) {
      $Packages[$selectedIndex - 1]
    }
  }

  return @($selectedPackages | Where-Object { $_ })
}

function Select-WingetPortablePackages([object[]]$Packages) {
  if (Test-OutGridViewAvailable) {
    return @(Select-WingetPortablePackagesWithGrid -Packages $Packages)
  }

  Write-Warning "Out-GridView is not available in this session. Falling back to text selection."
  Show-NumberedPortablePackages -Packages $Packages
  $selectedIndexes = Prompt-SelectionText -Count $Packages.Count
  if (-not $selectedIndexes -or $selectedIndexes.Count -eq 0) { return @() }

  return @(
    foreach ($idx in $selectedIndexes) {
      $Packages[$idx - 1]
    }
  )
}

function Prompt-RunMode {
  param(
    [string]$ActionLabel = 'selected package changes'
  )

  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Current session (no elevation)", "Run $ActionLabel here without elevation"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Elevated window", "Launch a UAC prompt and run $ActionLabel as admin")
  )
  return $Host.UI.PromptForChoice("Run location", "Where should the $ActionLabel run?", $choices, 0)
}

function Convert-ToHexCode {
  param([int]$Code)
  $unsigned = [BitConverter]::ToUInt32([BitConverter]::GetBytes([int]$Code), 0)
  return ('0x{0:X8}' -f $unsigned)
}

function Get-InstallerExitHint {
  param(
    [Nullable[int]]$InstallerExitCode,
    [string[]]$OutputLines,
    [string]$DiagnosticText
  )

  $allText = @()
  if ($OutputLines) { $allText += @($OutputLines) }
  if ($DiagnosticText) { $allText += @($DiagnosticText) }
  $combinedText = ($allText -join "`n")

  if ($combinedText -match '(?i)application is currently running\. exit the application then try again\.?') {
    return "Installer reports the app is still running. Close it, then check Task Manager for a lingering background process before retrying."
  }

  if ($null -eq $InstallerExitCode) { return $null }

  switch ([int]$InstallerExitCode) {
    1602 { return "Installer canceled by user." }
    1603 {
      if ($combinedText -match '(?i)uninstall(?:er)? failed') {
        return "MSI 1603 during uninstall/upgrade. Usually caused by locked files, running app/processes, or insufficient privileges."
      }
      return "MSI 1603 (fatal install error). Usually caused by locked files, running app/processes, or insufficient privileges."
    }
    1618 { return "Another installation is already in progress. Wait for it to finish, then retry." }
    1638 { return "Another version is already installed. Uninstall conflicting version or retry with elevation." }
    3010 { return "Installer succeeded but requires a restart to finish." }
    default { return "Installer returned exit code $($InstallerExitCode.Value)." }
  }
}

function Get-FailureRetryHint {
  param([object]$Failure)

  if ($Failure.RebootRequired) {
    return "Restart Windows first, then rerun the updater to confirm whether the new version is already in place."
  }
  if ($Failure.WingetHint) { return $Failure.WingetHint }
  if ($Failure.Stage -like 'Machine-scope*' -and $Failure.InstallScope -eq 'User') {
    return "Retry this package in the current session instead of the elevated window."
  }
  if ($Failure.Stage -eq 'Current session') {
    return "Retry in Elevated window for this package."
  }
  return "Close related apps/processes and retry."
}

function Get-WingetFailureHint {
  param(
    [int]$WingetExitCode,
    [object]$Package,
    [string]$Stage,
    [string]$LogPath,
    [object]$DiagnosticContext,
    [object]$PreviousFailure
  )

  $diagText = $null
  if ($DiagnosticContext -and $DiagnosticContext.CliLogText) {
    $diagText = [string]$DiagnosticContext.CliLogText
  } elseif ($LogPath -and (Test-Path $LogPath)) {
    try { $diagText = Get-Content -Path $LogPath -Raw -ErrorAction Stop } catch { $diagText = $null }
  }

  switch ($WingetExitCode) {
    -1978335226 {
      if ($DiagnosticContext -and $DiagnosticContext.InstallerExitCode -eq 3010) {
        return "Installer completed and requested a restart. Reboot Windows, then rerun the updater to confirm the final version."
      }
      return "Winget launched the installer, but the installer did not report a normal success path back to winget."
    }
    -1978335184 {
      if ($DiagnosticContext -and $DiagnosticContext.InstallerExitCode -eq 1603) {
        return "The uninstall step failed with MSI 1603. Close the app and related processes, then retry non-silent or in an elevated window."
      }
      return "The uninstall step failed before winget could complete the upgrade."
    }
    -1978335189 {
      if ($diagText -and $diagText -match 'Installer scope does not match currently installed scope:\s*(?<available>\w+)\s*!=\s*(?<installed>\w+)') {
        $availableScope = $Matches['available']
        $installedScope = $Matches['installed']
        return "Installed as $($installedScope.ToLowerInvariant())-scope, but the available winget upgrade only supports $($availableScope.ToLowerInvariant())-scope installs. Winget cannot upgrade this install in place."
      }
      if ($PreviousFailure -and $PreviousFailure.RebootRequired) {
        return "The previous attempt completed into a restart-required state. Reboot Windows, then re-check whether the upgrade is already applied before retrying."
      }
      if ($PreviousFailure -and $PreviousFailure.ExitCode -eq -1978335184) {
        return "The previous attempt failed during uninstall, and winget no longer considers an in-place update applicable. Retry in an elevated window; if it still fails, uninstall and reinstall the package explicitly."
      }
      if ($Package -and $Package.InstallScope -eq 'User' -and $Stage -like 'Machine-scope*') {
        return "This package is installed per-user, so the machine-scope run does not apply. Run or retry it in the current session instead."
      }
      return "A newer version exists, but winget says it does not apply to this install or to this system's requirements."
    }
    default { return $null }
  }
}

function Get-WingetCliDiagDirectory {
  return Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir'
}

function Get-WingetCliDiagnosticLogPath {
  param(
    [string]$InstallerLogPath,
    [string]$PackageId
  )

  $diagDir = Get-WingetCliDiagDirectory
  if (-not (Test-Path $diagDir)) { return $null }

  $escapedInstallerLogPath = if ($InstallerLogPath) { [regex]::Escape($InstallerLogPath) } else { $null }
  $candidateFiles = @(Get-ChildItem -Path $diagDir -Filter 'WinGet-*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 24)

  if ($escapedInstallerLogPath) {
    foreach ($file in $candidateFiles) {
      try {
        $text = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
        if ($text -match $escapedInstallerLogPath) {
          return $file.FullName
        }
      } catch {
        continue
      }
    }

    return $null
  }

  if ($PackageId) {
    foreach ($file in $candidateFiles) {
      try {
        $text = Get-Content -Path $file.FullName -Raw -ErrorAction Stop
        if ($text -match [regex]::Escape($PackageId)) {
          return $file.FullName
        }
      } catch {
        continue
      }
    }
  }

  return $null
}

function Get-WingetDiagnosticContext {
  param(
    [string]$InstallerLogPath,
    [string]$PackageId
  )

  $cliLogPath = Get-WingetCliDiagnosticLogPath -InstallerLogPath $InstallerLogPath -PackageId $PackageId
  $cliLogText = $null
  if ($cliLogPath -and (Test-Path $cliLogPath)) {
    try { $cliLogText = Get-Content -Path $cliLogPath -Raw -ErrorAction Stop } catch { $cliLogText = $null }
  }

  $installerExitCode = $null
  if ($cliLogText) {
    $patterns = @(
      'ShellExecute installer failed:\s*(-?\d+)',
      'MsiExec uninstaller failed:\s*(-?\d+)',
      'MsiExec installer failed:\s*(-?\d+)'
    )

    foreach ($pattern in $patterns) {
      $match = [regex]::Match($cliLogText, $pattern)
      if (-not $match.Success) { continue }
      $parsed = 0
      if ([int]::TryParse($match.Groups[1].Value, [ref]$parsed)) {
        $installerExitCode = $parsed
        break
      }
    }
  }

  return [pscustomobject]@{
    CliLogPath        = $cliLogPath
    CliLogText        = $cliLogText
    InstallerExitCode = $installerExitCode
    RebootRequired    = ($installerExitCode -eq 3010)
  }
}

function Get-MostRecentFailureForPackage {
  param([string]$Id)

  if ([string]::IsNullOrWhiteSpace($Id)) { return $null }

  return @(
    $script:FailureHistory |
      Where-Object { $_.Id -eq $Id } |
      Select-Object -Last 1
  ) | Select-Object -First 1
}

function Register-InstallerFailure {
  param([object]$Failure)

  if (-not $Failure) { return }
  $script:InstallerFailures += $Failure
  $script:FailureHistory += $Failure
}

function Add-RestartRequiredPackage {
  param(
    [object]$Package,
    [string]$Stage,
    [string]$LogPath,
    [string]$CliLogPath
  )

  if (-not $Package -or [string]::IsNullOrWhiteSpace([string]$Package.Id)) { return }
  if (@($script:RestartRequiredPackages | Where-Object { $_.Id -eq $Package.Id }).Count -gt 0) { return }

  $script:RestartRequiredPackages += [pscustomobject]@{
    Id           = [string]$Package.Id
    Name         = [string]$Package.Name
    Stage        = $Stage
    LogPath      = $LogPath
    CliLogPath   = $CliLogPath
    InstallScope = $Package.InstallScope
  }
}

function Show-RestartRequiredSummary {
  if (-not $script:RestartRequiredPackages -or $script:RestartRequiredPackages.Count -eq 0) { return }

  Write-Host ""
  Write-Host "Restart required:" -ForegroundColor Yellow
  foreach ($pkg in @($script:RestartRequiredPackages)) {
    $logNotes = @()
    if ($pkg.LogPath) { $logNotes += "Installer log: $($pkg.LogPath)" }
    if ($pkg.CliLogPath) { $logNotes += "Winget log: $($pkg.CliLogPath)" }
    $detail = if ($logNotes.Count -gt 0) { " (" + ($logNotes -join ' | ') + ")" } else { "" }
    Write-Host "- $($pkg.Id): installer requested a restart to finish the upgrade. Reboot Windows, then rerun the updater to confirm the new version.$detail"
  }
}

function Get-InstallerExitCodeFromLog {
  param([string]$Path)
  if (-not $Path -or -not (Test-Path $Path)) { return $null }

  try {
    $text = Get-Content -Path $Path -Raw -ErrorAction Stop
    $patterns = @(
      '(?i)(?:install|uninstall)\s+failed\s+with\s+exit\s+code:\s*(-?\d+)',
      '(?i)installer\s+return\s+code\s*[:=]\s*(-?\d+)',
      '(?i)msi(?:\s+installer)?\s+(?:return|exit)\s+code\s*[:=]\s*(-?\d+)'
    )

    foreach ($pattern in $patterns) {
      $m = [regex]::Match($text, $pattern)
      if (-not $m.Success) { continue }
      $tmp = 0
      if ([int]::TryParse($m.Groups[1].Value, [ref]$tmp)) {
        return $tmp
      }
    }
  } catch {
    return $null
  }

  return $null
}

function Start-WingetSelected {
  param(
    [object[]]$Packages,
    [string]$Stage,
    [switch]$NonSilent
  )

  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }

  foreach ($pkg in @($Packages)) {
    if (-not $pkg.Id) { continue }
    $safeId = ($pkg.Id -replace '[^A-Za-z0-9._-]','_')
    $logDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winget-update-script-logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $wingetLogPath = Join-Path -Path $logDir -ChildPath ("{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    $args = @('upgrade','--id',$pkg.Id,'--log',$wingetLogPath) + $WingetCommonArgs
    if ($NonSilent) {
      $args = @($args | Where-Object { $_ -ne '--silent' })
    }
    if ($pkg.Source) { $args += @('--source', $pkg.Source) }
    Write-Host "`n[$Stage] Running: winget $($args -join ' ')"

    $wingetOutput = @()
    & winget @args
    $wingetExitCode = $LASTEXITCODE

    $diagnosticContext = Get-WingetDiagnosticContext -InstallerLogPath $wingetLogPath -PackageId $pkg.Id
    $installerExitCode = Get-InstallerExitCodeFromLog -Path $wingetLogPath
    if ($installerExitCode -eq $null -and $diagnosticContext -and $diagnosticContext.InstallerExitCode -ne $null) {
      $installerExitCode = $diagnosticContext.InstallerExitCode
    }

    if ($wingetExitCode -ne 0) {
      if ($diagnosticContext -and $diagnosticContext.RebootRequired) {
        Add-RestartRequiredPackage -Package $pkg -Stage $Stage -LogPath $wingetLogPath -CliLogPath $diagnosticContext.CliLogPath
        Write-Host "[$Stage] $($pkg.Id) completed into a restart-required state. Reboot Windows, then rerun the updater to confirm the final version." -ForegroundColor Yellow
        continue
      }

      $wingetHex = Convert-ToHexCode -Code $wingetExitCode
      $previousFailure = Get-MostRecentFailureForPackage -Id $pkg.Id
      $wingetHint = Get-WingetFailureHint -WingetExitCode $wingetExitCode -Package $pkg -Stage $Stage -LogPath $wingetLogPath -DiagnosticContext $diagnosticContext -PreviousFailure $previousFailure
      $installerHint = Get-InstallerExitHint -InstallerExitCode $installerExitCode -OutputLines $wingetOutput -DiagnosticText (if ($diagnosticContext) { $diagnosticContext.CliLogText } else { $null })
      $warningText = "winget failed for $($pkg.Id) in $Stage. winget code: $wingetExitCode ($wingetHex)"
      if ($installerExitCode -ne $null) {
        $warningText += "; installer code: $installerExitCode"
      }
      if ($wingetHint) {
        $warningText += ". $wingetHint"
      } elseif ($installerHint) {
        $warningText += ". $installerHint"
      }
      Write-Warning $warningText

      Register-InstallerFailure -Failure ([pscustomobject]@{
        Id                = $pkg.Id
        Provider          = 'Winget'
        Stage             = $Stage
        ExitCode          = $wingetExitCode
        ExitCodeHex       = $wingetHex
        InstallerExitCode = $installerExitCode
        InstallerHint     = $installerHint
        WingetHint        = $wingetHint
        RetryHint         = $null
        LogPath           = $wingetLogPath
        InstallScope      = $pkg.InstallScope
        CliLogPath        = if ($diagnosticContext) { $diagnosticContext.CliLogPath } else { $null }
        RebootRequired    = $false
      })
    }
  }
}

function Start-PythonManagerSelected {
  param(
    [object[]]$Packages,
    [string]$Stage
  )

  $py = Get-Command py -ErrorAction SilentlyContinue
  if (-not $py) {
    Write-Warning "Skipping Python Install Manager upgrades; 'py' command was not found."
    return
  }

  foreach ($pkg in @($Packages)) {
    $tag = [string]$pkg.UpdateTag
    if ([string]::IsNullOrWhiteSpace($tag)) {
      Write-Warning "Skipping Python Install Manager entry without update tag: $($pkg.Id)"
      continue
    }

    $args = @('install', '--update', '-y', $tag)
    Write-Host "`n[$Stage] Running: py $($args -join ' ')"
    & $py.Path @args
    $pyExitCode = $LASTEXITCODE

    if ($pyExitCode -ne 0) {
      $pyHex = Convert-ToHexCode -Code $pyExitCode
      Write-Warning "Python Install Manager failed for $tag in $Stage. exit code: $pyExitCode ($pyHex)"
      Register-InstallerFailure -Failure ([pscustomobject]@{
        Id                = $pkg.Id
        Provider          = 'PythonInstallManager'
        Stage             = $Stage
        ExitCode          = $pyExitCode
        ExitCodeHex       = $pyHex
        InstallerExitCode = $null
        InstallerHint     = "Python Install Manager update failed for tag '$tag'."
        WingetHint        = $null
        RetryHint         = "Retry this package in the current session."
        LogPath           = $null
        InstallScope      = 'User'
        CliLogPath        = $null
        RebootRequired    = $false
      })
    }
  }
}

function Start-SelectedUpgrades {
  param(
    [object[]]$Packages,
    [string]$Stage,
    [switch]$NonSilent
  )

  $wingetPackages = @($Packages | Where-Object { -not $_.Provider -or $_.Provider -eq 'Winget' })
  if ($wingetPackages.Count -gt 0) {
    Start-WingetSelected -Packages $wingetPackages -Stage $Stage -NonSilent:$NonSilent
  }

  $pythonPackages = @($Packages | Where-Object { $_.Provider -eq 'PythonInstallManager' })
  if ($pythonPackages.Count -gt 0) {
    Start-PythonManagerSelected -Packages $pythonPackages -Stage $Stage
  }
}

function Start-WingetPortableUninstallSelected {
  param(
    [object[]]$Packages,
    [string]$Stage
  )

  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return }

  foreach ($pkg in @($Packages)) {
    if (-not $pkg.ProductCode) { continue }
    $safeId = ($pkg.ProductCode -replace '[^A-Za-z0-9._-]','_')
    $logDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath 'winget-update-script-logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $wingetLogPath = Join-Path -Path $logDir -ChildPath ("cleanup-{0}-{1}.log" -f $safeId, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    $args = @(
      'uninstall',
      '--product-code', $pkg.ProductCode,
      '--purge',
      '--silent',
      '--disable-interactivity',
      '--accept-source-agreements',
      '--log', $wingetLogPath
    )
    if ($pkg.Source) { $args += @('--source', $pkg.Source) }
    Write-Host "`n[$Stage] Running: winget $($args -join ' ')"

    $wingetOutput = @()
    & winget @args
    $wingetExitCode = $LASTEXITCODE

    $installerExitCode = Get-InstallerExitCodeFromLog -Path $wingetLogPath

    if ($wingetExitCode -ne 0) {
      $wingetHex = Convert-ToHexCode -Code $wingetExitCode
      $installerHint = Get-InstallerExitHint -InstallerExitCode $installerExitCode -OutputLines $wingetOutput -DiagnosticText $null
      $warningText = "winget cleanup failed for $($pkg.ProductCode) in $Stage. winget code: $wingetExitCode ($wingetHex)"
      if ($installerExitCode -ne $null) {
        $warningText += "; installer code: $installerExitCode"
      }
      if ($installerHint) {
        $warningText += ". $installerHint"
      }
      Write-Warning $warningText

      Register-InstallerFailure -Failure ([pscustomobject]@{
        Id                = $pkg.ProductCode
        Provider          = 'WingetPortableCleanup'
        Stage             = $Stage
        ExitCode          = $wingetExitCode
        ExitCodeHex       = $wingetHex
        InstallerExitCode = $installerExitCode
        InstallerHint     = $installerHint
        WingetHint        = $null
        RetryHint         = $null
        LogPath           = $wingetLogPath
        InstallScope      = $pkg.InstallScope
        CliLogPath        = $null
        RebootRequired    = $false
      })
    }
  }
}

function Summarize-LogReason {
  param([string]$Path)
  if (-not $Path -or -not (Test-Path $Path)) { return $null }
  try {
    $text = Get-Content -Path $Path -Raw -ErrorAction Stop

    if ($text -match 'The following process\(es\) use Git for Windows:\s*(?<procs>[^\r\n]+)') {
      $procs = $Matches['procs'].Trim()
      return "Git installer blocked because running processes: $procs. Close them and retry."
    }
    if ($text -match 'Some applications could not be shut down') {
      return "Installer could not close running applications (Restart Manager). Close WinMerge or related processes and retry."
    }
    if ($text -match 'User canceled the installation process') {
      return "Installer was canceled (likely after a prompt). Rerun and allow it to continue."
    }
    return $null
  } catch { return $null }
}

function Summarize-Failures {
  param([System.Collections.IEnumerable]$Failures)
  if (-not $Failures -or $Failures.Count -eq 0) { return }

  Write-Host "" # blank line
  Write-Host "Failure summary:" -ForegroundColor Yellow
  foreach ($fail in $Failures) {
    if (-not $fail.RetryHint) { $fail.RetryHint = Get-FailureRetryHint -Failure $fail }
    $reason = Summarize-LogReason -Path $fail.LogPath
    $logParts = @()
    if ($fail.LogPath) { $logParts += "Installer log: $($fail.LogPath)" }
    if ($fail.CliLogPath) { $logParts += "Winget log: $($fail.CliLogPath)" }
    $logNote = if ($logParts.Count -gt 0) { $logParts -join " | " } else { "No log path provided" }
    $toolName = if ($fail.Provider) { [string]$fail.Provider } else { 'winget' }
    $codeNote = "$toolName=$($fail.ExitCode)"
    if ($fail.ExitCodeHex) { $codeNote += " ($($fail.ExitCodeHex))" }
    if ($fail.InstallerExitCode -ne $null) { $codeNote += ", installer=$($fail.InstallerExitCode)" }

    if ($reason) {
      Write-Host "- $($fail.Id): $reason [$codeNote]. $($fail.RetryHint) ($logNote)"
    } elseif ($fail.InstallerHint) {
      Write-Host "- $($fail.Id): $($fail.InstallerHint) [$codeNote]. $($fail.RetryHint) ($logNote)"
    } else {
      Write-Host "- $($fail.Id): Upgrade failed [$codeNote]. $($fail.RetryHint) ($logNote)"
    }
  }
}

function Analyze-FailuresWithCodex {
  param([System.Collections.IEnumerable]$Failures)

  if (-not $Failures -or $Failures.Count -eq 0) { return }

  $codex = Get-Command codex -ErrorAction SilentlyContinue
  if (-not $codex) {
    Write-Host "Codex executable not found; skipping auto-analysis. Logs are listed above if available."
    return
  }

  foreach ($fail in $Failures) {
    if (-not $fail.LogPath -or -not (Test-Path $fail.LogPath)) {
      Write-Host "No log file found for $($fail.Id); skipping Codex analysis."
      continue
    }
    try {
      $logText = Get-Content -Path $fail.LogPath -Raw -ErrorAction Stop
      $prompt = "Summarize in plain language why the installer failed and suggest the next step to fix it. Log follows:\n\n$logText"
      Write-Host "\n--- Codex analysis for $($fail.Id) ($($fail.Stage)) ---"
      $analysis = $prompt | & $codex.Path 2>&1
      if ($analysis) { $analysis | ForEach-Object { Write-Host $_ } }
    } catch {
      Write-Warning "Codex analysis failed for $($fail.Id): $($_.Exception.Message)"
    }
  }
}

function Start-MachinePhase {
  param(
    [string[]]$Ids,
    [switch]$NonSilent,
    [switch]$CleanupPortablePackages
  )

  if (-not $Ids -or $Ids.Count -eq 0) { return }

  $csv = ($Ids -join ',')
  $argList = @(
    '-NoProfile',
    '-ExecutionPolicy','Bypass',
    '-File',"`"$PSCommandPath`"",
    '-MachinePhase'
  )
  if ($NonSilent) {
    $argList += '-NonSilentPhase'
  }
  if ($CleanupPortablePackages) {
    $argList += '-CleanupPortablePackages'
  }
  $argList += @(
    '-SelectedList',"`"$csv`""
  )

  $currentShellPath = $null
  try { $currentShellPath = (Get-Process -Id $PID -ErrorAction Stop).Path } catch { $currentShellPath = $null }
  if (-not $currentShellPath) {
    $currentShellPath = if ($PSVersionTable.PSEdition -eq 'Core') { 'pwsh' } else { 'powershell' }
  }

  try { Start-Process -FilePath $currentShellPath -Verb runas -ArgumentList $argList | Out-Null }
  catch { Write-Warning "Elevation cancelled; machine-scope upgrades skipped." }
}

function Prompt-RetryFailedInElevated {
  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Retry failed packages in an elevated window"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not retry now")
  )
  $result = $Host.UI.PromptForChoice("Retry failed upgrades", "Retry failed package upgrades in an elevated window?", $choices, 0)
  return ($result -eq 0)
}

function Prompt-RetryFailedNonSilent {
  param([string]$ScopeDescription = "this window")

  $choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    (New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Retry failed packages without --silent in $ScopeDescription"),
    (New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Do not retry now")
  )
  $message = "Retry failed package upgrades without --silent? This can surface installer dialogs/prompts that were hidden during the silent run. Target: $ScopeDescription."
  $result = $Host.UI.PromptForChoice("Retry failed upgrades without --silent", $message, $choices, 0)
  return ($result -eq 0)
}

function Show-ActivePwshProcesses {
  $pwshProcesses = @(Get-Process pwsh -ErrorAction SilentlyContinue | Sort-Object Id)
  if ($pwshProcesses.Count -eq 0) {
    Write-Host "No active pwsh.exe processes were detected." -ForegroundColor DarkGray
    return
  }

  Write-Host "Active PowerShell 7 processes that can block the update:" -ForegroundColor Yellow
  foreach ($proc in $pwshProcesses) {
    $windowTitle = if ([string]::IsNullOrWhiteSpace($proc.MainWindowTitle)) { '(no window title)' } else { $proc.MainWindowTitle.Trim() }
    $path = $null
    try { $path = $proc.Path } catch { $path = $null }
    if ($path) {
      Write-Host ("- PID {0}: {1} | {2}" -f $proc.Id, $windowTitle, $path)
    } else {
      Write-Host ("- PID {0}: {1}" -f $proc.Id, $windowTitle)
    }
  }
}

function Start-DeferredPowerShellUpgradeHelper {
  param([object[]]$Packages)

  $deferredPackages = @($Packages | Where-Object { Test-IsDeferredPowerShellPackage -Package $_ })
  if ($deferredPackages.Count -eq 0) { return }

  $helperPath = Join-Path -Path $PSScriptRoot -ChildPath 'deferred-winget-package-upgrade.ps1'
  if (-not (Test-Path $helperPath)) {
    Write-Warning "Deferred PowerShell helper script was not found at $helperPath"
    return
  }

  $packageIds = @($deferredPackages | Select-Object -ExpandProperty Id -Unique)
  if ($packageIds.Count -eq 0) { return }

  Write-Host ""
  Write-Host "PowerShell 7 updates will run in a separate Windows PowerShell helper after all pwsh.exe sessions close." -ForegroundColor Yellow
  Write-Host "Close any remaining PowerShell 7 terminals and check Task Manager if a pwsh.exe process is lingering." -ForegroundColor Yellow
  Show-ActivePwshProcesses

  $argList = @(
    '-NoLogo',
    '-NoProfile',
    '-ExecutionPolicy','Bypass',
    '-File',"`"$helperPath`""
  )
  foreach ($packageId in $packageIds) {
    $argList += @('-PackageId',"`"$packageId`"")
  }

  $needsElevation = @($deferredPackages | Where-Object { $_.InstallScope -ne 'User' }).Count -gt 0
  try {
    if ($needsElevation) {
      Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $argList | Out-Null
    } else {
      Start-Process -FilePath 'powershell.exe' -ArgumentList $argList | Out-Null
    }
    Write-Host "Opened the deferred PowerShell update helper in a separate window." -ForegroundColor Cyan
  } catch {
    Write-Warning "Unable to start the deferred PowerShell update helper: $($_.Exception.Message)"
  }
}

function Exit-WithPause {
  param(
    [int]$Code = 0,
    [string]$Prompt = "Press Enter to close this window"
  )

  $script:RequestedExitCode = $Code
  $script:PausePrompt = $Prompt
  throw [System.OperationCanceledException]::new("__WINGET_SCRIPT_EXIT__")
}

# ------------------ CONTROL FLOW ------------------

$script:RequestedExitCode = 0
$script:PausePrompt = if ($MachinePhase) { "Press Enter to close this elevated window" } else { "Press Enter to close this window" }

if (-not $PSVersionTable.PSVersion -or $PSVersionTable.PSVersion.Major -lt 7) {
  $launcherPath = Join-Path -Path $PSScriptRoot -ChildPath 'run-winget-ms-store-update-all.cmd'
  Write-Host ""
  Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Yellow
  Write-Host "Current host: $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
  Write-Host ""
  if (Test-Path $launcherPath) {
    Write-Host "Run this launcher instead:" -ForegroundColor Yellow
    Write-Host "  $launcherPath" -ForegroundColor Cyan
    Write-Host "It can install the latest PowerShell and then run this script." -ForegroundColor Yellow
  } else {
    Write-Host "Install PowerShell 7+ first, then rerun this script with pwsh." -ForegroundColor Yellow
  }
  exit 1
}

try {
  Ensure-WingetMsStoreSource

  if (-not $MachinePhase) {
    if ($CleanupPortablePackages) {
      $portablePackages = @(Get-WingetPortablePackageList)
      if (-not $portablePackages -or $portablePackages.Count -eq 0) {
        Write-Host "No installed winget portable packages were found for cleanup."
        Exit-WithPause
      }

      $selectedPackages = @(Select-WingetPortablePackages -Packages $portablePackages)
      if (-not $selectedPackages -or $selectedPackages.Count -eq 0) {
        Write-Host "No selection made. Exiting without changes."
        Exit-WithPause
      }

      $runMode = Prompt-RunMode -ActionLabel 'selected package removals'
      if ($runMode -eq 0) {
        Start-WingetPortableUninstallSelected -Packages $selectedPackages -Stage 'Current session cleanup'

        if ($script:InstallerFailures.Count -gt 0) {
          Summarize-Failures -Failures $script:InstallerFailures

          $failedMachineIds = @(
            $script:InstallerFailures |
              Where-Object { $_.Provider -eq 'WingetPortableCleanup' -and $_.InstallScope -ne 'User' } |
              Select-Object -ExpandProperty Id -Unique
          )
          if ($failedMachineIds.Count -gt 0 -and (Prompt-RetryFailedInElevated)) {
            Start-MachinePhase -Ids $failedMachineIds -CleanupPortablePackages
            Write-Host "Opening an elevated window to retry failed package removals..."
          }
        } else {
          Write-Host "`nSelected package removals completed successfully."
        }

        Exit-WithPause
      }

      $userScopedPackages = @($selectedPackages | Where-Object { $_.InstallScope -eq 'User' })
      $machineScopedPackages = @($selectedPackages | Where-Object { $_.InstallScope -ne 'User' })

      if ($userScopedPackages.Count -gt 0) {
        $userScopedIds = ($userScopedPackages | Select-Object -ExpandProperty ProductCode) -join ', '
        Write-Host "Removing user-scope packages in the current session: $userScopedIds" -ForegroundColor Yellow
        Start-WingetPortableUninstallSelected -Packages $userScopedPackages -Stage 'Current session cleanup (user-scope)'
        if ($script:InstallerFailures.Count -gt 0) {
          Summarize-Failures -Failures $script:InstallerFailures
        }
      }

      if ($machineScopedPackages.Count -gt 0) {
        Start-MachinePhase -Ids ($machineScopedPackages | Select-Object -ExpandProperty ProductCode) -CleanupPortablePackages
        Write-Host "Opening an elevated window to run machine-scope package removals..."
      } else {
        Write-Host "No machine-scope packages remain after filtering out user-scope installs."
      }

      Exit-WithPause
    }

    $upgrades = Get-UpgradeList
    if (-not $upgrades -or $upgrades.Count -eq 0) {
      Write-Host "No upgrades found from winget or Python Install Manager."
      Exit-WithPause
    }

    $selectedPackages = @(Select-Upgrades -Packages $upgrades)
    if (-not $selectedPackages -or $selectedPackages.Count -eq 0) {
      Write-Host "No selection made. Exiting without changes."
      Exit-WithPause
    }
    $runMode = Prompt-RunMode -ActionLabel 'selected upgrades'
    $deferredPowerShellPackages = @($selectedPackages | Where-Object { Test-IsDeferredPowerShellPackage -Package $_ })
    $immediatePackages = @($selectedPackages | Where-Object { -not (Test-IsDeferredPowerShellPackage -Package $_) })
    if ($runMode -eq 0) {
      if ($immediatePackages.Count -gt 0) {
        Start-SelectedUpgrades -Packages $immediatePackages -Stage "Current session"
      }

      if ($script:InstallerFailures.Count -gt 0) {
        Summarize-Failures -Failures $script:InstallerFailures
        $failedPackageIds = @(
          $script:InstallerFailures |
            Where-Object { $_.Id } |
            Select-Object -ExpandProperty Id -Unique
        )
        $failedPackages = @(
          foreach ($failedId in $failedPackageIds) {
            $immediatePackages | Where-Object { $_.Id -eq $failedId } | Select-Object -First 1
          }
        ) | Where-Object { $_ }

        if ($failedPackages.Count -gt 0 -and (Prompt-RetryFailedNonSilent -ScopeDescription "the current session")) {
          $script:InstallerFailures = @()
          Start-SelectedUpgrades -Packages $failedPackages -Stage "Current session (non-silent retry)" -NonSilent
          if ($script:InstallerFailures.Count -gt 0) {
            Summarize-Failures -Failures $script:InstallerFailures
          } else {
            Write-Host "`nNon-silent retry completed successfully."
          }
        }

        $failedMachineIds = @(
          $script:InstallerFailures |
            Where-Object { $_.Provider -eq 'Winget' -and $_.InstallScope -ne 'User' } |
            Select-Object -ExpandProperty Id -Unique
        )
        if ($failedMachineIds.Count -gt 0 -and (Prompt-RetryFailedInElevated)) {
          Start-MachinePhase -Ids $failedMachineIds
          Write-Host "Opening an elevated window to retry failed upgrades..."
        }
      }

      if ($deferredPowerShellPackages.Count -gt 0) {
        Start-DeferredPowerShellUpgradeHelper -Packages $deferredPowerShellPackages
      }

      Show-RestartRequiredSummary

      if ($script:InstallerFailures.Count -eq 0 -and $immediatePackages.Count -gt 0) {
        Write-Host "`nAll immediate upgrades completed successfully."
      } elseif ($script:InstallerFailures.Count -eq 0 -and $deferredPowerShellPackages.Count -gt 0) {
        Write-Host "`nDeferred PowerShell update helper queued successfully."
      }
      Exit-WithPause
    }

    # Relaunch elevated to run the same selection as admin.
    $userScopedPackages = @($immediatePackages | Where-Object { $_.InstallScope -eq 'User' })
    $machineScopedPackages = @($immediatePackages | Where-Object { $_.InstallScope -ne 'User' })

    if ($userScopedPackages.Count -gt 0) {
      $userScopedIds = ($userScopedPackages | Select-Object -ExpandProperty Id) -join ', '
      Write-Host "Running user-scope packages in the current session: $userScopedIds" -ForegroundColor Yellow
      Start-SelectedUpgrades -Packages $userScopedPackages -Stage "Current session (user-scope)"
      if ($script:InstallerFailures.Count -gt 0) {
        Summarize-Failures -Failures $script:InstallerFailures
      }
    }

    if ($machineScopedPackages.Count -gt 0) {
      Start-MachinePhase -Ids ($machineScopedPackages | Select-Object -ExpandProperty Id)
      Write-Host "Opening an elevated window to run machine-scope upgrades..."
    } else {
      Write-Host "No machine-scope packages remain after filtering out user-scope installs."
    }

    if ($deferredPowerShellPackages.Count -gt 0) {
      Start-DeferredPowerShellUpgradeHelper -Packages $deferredPowerShellPackages
    }
    Show-RestartRequiredSummary
    Exit-WithPause
  }

  # Phase 2 (elevated): machine-scope upgrade for the same selection
  $machineIds = if ($SelectedList) { $SelectedList -split ',' | Where-Object { $_ } } else { @() }
  if (-not $machineIds -or $machineIds.Count -eq 0) {
    Write-Warning "No package ids provided to machine phase; nothing to do."
    Exit-WithPause
  }

  if ($CleanupPortablePackages) {
    $machinePackages = foreach ($id in $machineIds) {
      [pscustomobject]@{
        Name         = $id
        Id           = $id
        ProductCode  = $id
        Source       = 'winget'
        Provider     = 'WingetPortableCleanup'
        InstallScope = 'Machine'
      }
    }
    $machineStage = "Machine-scope cleanup"
    Start-WingetPortableUninstallSelected -Packages $machinePackages -Stage $machineStage
  } else {
    $machinePackages = foreach ($id in $machineIds) { [pscustomobject]@{ Name = $id; Id = $id; Source = $null; Provider = 'Winget' } }
    $machineStage = if ($NonSilentPhase) { "Machine-scope (non-silent retry)" } else { "Machine-scope" }
    Start-SelectedUpgrades -Packages $machinePackages -Stage $machineStage -NonSilent:$NonSilentPhase
  }
  if ($script:InstallerFailures.Count -gt 0) {
    Write-Host "`n$machineStage completed with failures."
    Analyze-FailuresWithCodex -Failures $script:InstallerFailures
    Summarize-Failures -Failures $script:InstallerFailures

    if (-not $CleanupPortablePackages -and -not $NonSilentPhase) {
      $failedIds = @($script:InstallerFailures | Select-Object -ExpandProperty Id -Unique)
      if ($failedIds.Count -gt 0 -and (Prompt-RetryFailedNonSilent -ScopeDescription "this elevated window")) {
        $script:InstallerFailures = @()
        $retryPackages = foreach ($id in $failedIds) { [pscustomobject]@{ Name = $id; Id = $id; Source = $null; Provider = 'Winget' } }
        Start-SelectedUpgrades -Packages $retryPackages -Stage "Machine-scope (non-silent retry)" -NonSilent

        if ($script:InstallerFailures.Count -gt 0) {
          Write-Host "`nMachine-scope non-silent retry still has failures."
          Analyze-FailuresWithCodex -Failures $script:InstallerFailures
          Summarize-Failures -Failures $script:InstallerFailures
        } else {
          Write-Host "Machine-scope non-silent retry completed successfully."
        }
      }
    }
  } else {
    Write-Host "$machineStage upgrades complete."
  }

  Show-RestartRequiredSummary

  Exit-WithPause -Prompt "Press Enter to close this elevated window"
} catch [System.OperationCanceledException] {
  if ($_.Exception.Message -ne "__WINGET_SCRIPT_EXIT__") {
    $script:RequestedExitCode = 1
    throw
  }
} catch {
  $script:RequestedExitCode = 1
  Write-Warning "Script failed unexpectedly: $($_.Exception.Message)"
} finally {
  Read-Host $script:PausePrompt | Out-Null
}

exit $script:RequestedExitCode
