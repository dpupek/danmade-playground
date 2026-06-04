param(
  [string]$TargetScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Assert-True {
  param(
    [bool]$Condition,
    [string]$Message
  )
  if (-not $Condition) {
    throw "Assertion failed: $Message"
  }
}

function Assert-Equal {
  param(
    [object]$Actual,
    [object]$Expected,
    [string]$Message
  )
  if ($Actual -ne $Expected) {
    throw "Assertion failed: $Message (expected '$Expected', got '$Actual')"
  }
}

function Contains-Id {
  param(
    [object[]]$Packages,
    [string]$Id
  )
  foreach ($pkg in @($Packages)) {
    if ($pkg -and $pkg.Id -eq $Id) { return $true }
  }
  return $false
}

if (-not $TargetScriptPath) {
  $repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..\..\..\..")
  $TargetScriptPath = Join-Path $repoRoot "windows-update-scripts\winget-ms-store-update-all.ps1"
}

if (-not (Test-Path -LiteralPath $TargetScriptPath)) {
  throw "Target script not found: $TargetScriptPath"
}

$rawScript = Get-Content -LiteralPath $TargetScriptPath -Raw
$controlFlowMarker = "# ------------------ CONTROL FLOW ------------------"
$markerIndex = $rawScript.IndexOf($controlFlowMarker)
Assert-True ($markerIndex -ge 0) "Could not find control flow marker in target script."

# Load only helper/function definitions, not interactive control flow.
$definitionsOnly = $rawScript.Substring(0, $markerIndex)
Invoke-Expression $definitionsOnly

Assert-True ([bool](Get-Command Normalize-WingetPackageId -ErrorAction SilentlyContinue)) "Normalize-WingetPackageId must exist."
Assert-True ([bool](Get-Command Get-WingetUpgradeList -ErrorAction SilentlyContinue)) "Get-WingetUpgradeList must exist."

$originalWinget = $null
$hadWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)
if ($hadWinget) {
  $originalWinget = (Get-Command winget).Path
}

try {
  function global:winget {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)

    $argLine = ($Args -join " ")
    if ($argLine -match "^upgrade\b" -and $argLine -match "--output json") {
      # Force JSON path failure to exercise fallback table parsing.
      return @("{ invalid json")
    }

    if ($argLine -match "^upgrade\b") {
      return @(
        "Name Id Version Available Source",
        "----------------------------------------------------------",
        "LibreOffice 25.8.4.2 TheDocumentFoundation.LibreOffice 25.8.4.2 26.2.0.3 winget",
        "Microsoft Teams (personal) Microsoft.Teams.Free 25255.501.3956.3603 26043.2101.4385.7548 winget",
        "Microsoft Store App 9WZDNCRFJ3Q8 1.0.0 1.1.0 msstore",
        "ImageMagick 7.1.2-13 Q16-HDRI (64-bit) ImageMagick.ImageMagick 7.1.2.13 7.1.2.15 winget",
        "Zoom Workplace XP99J3KP4XZ4VV 6.7.8 (32670) 7.0.2 (34412) msstore",
        "Python 3.14.5 Python.Python.3.12 > 3.12.10 3.14.6 winget",
        "6 upgrades available."
      )
    }

    return @()
  }

  $packages = @(Get-WingetUpgradeList)
  Assert-Equal -Actual $packages.Count -Expected 6 -Message "Fallback parser should return all sample packages."

  Assert-True (Contains-Id -Packages $packages -Id "TheDocumentFoundation.LibreOffice") "Should include dotted ID."
  Assert-True (Contains-Id -Packages $packages -Id "Microsoft.Teams.Free") "Should include dotted ID with single-space separators."
  Assert-True (Contains-Id -Packages $packages -Id "9WZDNCRFJ3Q8") "Should include non-dotted Store ID."
  Assert-True (Contains-Id -Packages $packages -Id "ImageMagick.ImageMagick") "Should include ImageMagick ID from mojibake name row."
  Assert-True (Contains-Id -Packages $packages -Id "XP99J3KP4XZ4VV") "Should include Store ID when version columns include parenthetical build numbers."
  Assert-True (Contains-Id -Packages $packages -Id "Python.Python.3.12") "Should include Python row when installed version includes a comparator and space."

  $zoom = $packages | Where-Object { $_.Id -eq "XP99J3KP4XZ4VV" } | Select-Object -First 1
  Assert-Equal -Actual $zoom.Installed -Expected "6.7.8 (32670)" -Message "Should parse parenthetical installed version."
  Assert-Equal -Actual $zoom.Available -Expected "7.0.2 (34412)" -Message "Should parse parenthetical available version."

  $python = $packages | Where-Object { $_.Id -eq "Python.Python.3.12" } | Select-Object -First 1
  Assert-Equal -Actual $python.Installed -Expected "> 3.12.10" -Message "Should parse comparator installed version."
  Assert-Equal -Actual $python.Available -Expected "3.14.6" -Message "Should parse Python available version."

  $normalizedMojibake = Normalize-WingetPackageId -RawId (([char]0x00AA) + " ImageMagick.ImageMagick")
  Assert-Equal -Actual $normalizedMojibake -Expected "ImageMagick.ImageMagick" -Message "Normalizer should strip leading mojibake."

  $normalizedStoreId = Normalize-WingetPackageId -RawId "  XP9KHM4BK9FZ7Q "
  Assert-Equal -Actual $normalizedStoreId -Expected "XP9KHM4BK9FZ7Q" -Message "Normalizer should preserve non-dotted Store IDs."

  Write-Host "PASS: winget fallback parser regression checks succeeded."
  exit 0
}
finally {
  if (Get-Command winget -ErrorAction SilentlyContinue) {
    Remove-Item function:\winget -ErrorAction SilentlyContinue
  }
  if ($hadWinget -and $originalWinget) {
    # No restore needed for external executable; removing mock function reveals it again.
    $null = $originalWinget
  }
}
