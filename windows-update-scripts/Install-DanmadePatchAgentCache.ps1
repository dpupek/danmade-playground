<#
.SYNOPSIS
Stages a signed immutable local patch-agent release and optionally switches its tasks.
.DESCRIPTION
Run elevated after signing the release. Package is protected; -SecureParent also
removes broad write grants from its fixed parent while retaining read access.
Log/state contents, task triggers and principals are preserved. Publication is
atomic through current.json. Prior releases remain available for rollback.
Configure GPO actions to the stable local bootstrap or preference refresh will
restore UNC execution. Never accepts unsigned or differently signed PS1 files.
.EXAMPLE
Install-DanmadePatchAgentCache.ps1 -SourcePath '\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent' -ExpectedSignerThumbprint DB13837BA676E68359DC408F476C7DD1F57AFF41 -UpdateTasks -WhatIf
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
  [Parameter(Mandatory = $true)][string]$SourcePath,
  [Parameter(Mandatory = $true)][ValidatePattern('^[0-9a-fA-F]{40}$')][string]$ExpectedSignerThumbprint,
  [switch]$UpdateTasks,
  # Explicit opt-in for AllSigned -File task actions; PowerShell 5.1 cannot bind -Confirm:$false there.
  [switch]$Unattended,
  [switch]$SecureParent
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($Unattended) { $ConfirmPreference = 'None' }
$root = 'C:\ProgramData\DanmadePatchAgent\Package'
$files = @('danmade-patch-agent.ps1', 'Test-DanmadePatchAgentHealth.ps1', 'Show-DanmadePatchAgentNotification.ps1', 'Invoke-DanmadePatchAgentCached.ps1', 'Start-DanmadePatchAgentHidden.vbs', 'danmade-patch-agent.policy.json')
$plans = @()
$hashes = @{}

function Get-ReleaseFileHash {
  param([string]$Path)
  # Stream input avoids PowerShell 5.1's WhatIf-sensitive ProviderPath reader.
  $stream = [IO.File]::OpenRead($Path)
  try { return (Get-FileHash -InputStream $stream -Algorithm SHA256).Hash }
  finally { $stream.Dispose() }
}

function Get-SecuredAgentParentAcl {
  param([Security.AccessControl.DirectorySecurity]$Acl)
  # Preserve raw generic ACEs (notably CREATOR OWNER) that FileSystemAccessRule cannot reconstruct.
  $descriptor = New-Object Security.AccessControl.RawSecurityDescriptor($Acl.GetSecurityDescriptorSddlForm([Security.AccessControl.AccessControlSections]::All))
  $rawAcl = New-Object Security.AccessControl.RawAcl($descriptor.DiscretionaryAcl.Revision, $descriptor.DiscretionaryAcl.Count)
  $writeMask = [Security.AccessControl.FileSystemRights]::Write -bor [Security.AccessControl.FileSystemRights]::Delete -bor [Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor [Security.AccessControl.FileSystemRights]::ChangePermissions -bor [Security.AccessControl.FileSystemRights]::TakeOwnership
  foreach ($ace in $descriptor.DiscretionaryAcl) {
    if ($ace -isnot [Security.AccessControl.CommonAce]) { throw 'Unsupported parent ACE; manual ACL review required.' }
    $sid = $ace.SecurityIdentifier.Value
    $ace.AceFlags = [Security.AccessControl.AceFlags]([int]$ace.AceFlags -band (-bnot [int][Security.AccessControl.AceFlags]::Inherited))
    if ($ace.AceQualifier -eq [Security.AccessControl.AceQualifier]::AccessAllowed -and ($sid -in @('S-1-1-0','S-1-5-11','S-1-5-32-545') -or $sid -match '-513$')) {
      $rights = $ace.AccessMask
      if ($rights -band 0x10000000) { $rights = ($rights -band (-bnot 0x10000000)) -bor [int][Security.AccessControl.FileSystemRights]::FullControl }
      $ace.AccessMask = $rights -band (-bnot ([int]$writeMask -bor 0x40000000))
    }
    if ($ace.AccessMask -ne 0) { $rawAcl.InsertAce($rawAcl.Count, $ace) }
  }
  $descriptor.DiscretionaryAcl = $rawAcl
  $descriptor.SetFlags($descriptor.ControlFlags -bor [Security.AccessControl.ControlFlags]::DiscretionaryAclProtected)
  $secured = New-Object Security.AccessControl.DirectorySecurity
  $secured.SetSecurityDescriptorSddlForm($descriptor.GetSddlForm([Security.AccessControl.AccessControlSections]::All))
  if (-not $secured.AreAccessRulesCanonical) { throw 'Parent ACL requires manual canonicalization.' }
  return $secured
}

foreach ($file in $files) {
  $path = Join-Path $SourcePath $file
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing release file: $file" }
  if ($file -like '*.ps1') {
    $signature = Get-AuthenticodeSignature -FilePath $path
    if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $ExpectedSignerThumbprint) { throw "Invalid signature or unexpected signer: $file" }
  }
  $hashes[$file] = Get-ReleaseFileHash -Path $path
}
$policy = Get-Content -LiteralPath (Join-Path $SourcePath 'danmade-patch-agent.policy.json') -Raw | ConvertFrom-Json
if ($policy.enabled -isnot [bool] -or $policy.blockedPackageIds -isnot [array] -or $policy.packageRules -isnot [array]) { throw 'Release policy is malformed.' }

if ($UpdateTasks) {
  foreach ($name in @('Danmade Patch Agent - Machine', 'Danmade Patch Agent - Health', 'Danmade Patch Agent - User', 'Danmade Patch Agent - User Notification')) {
    $task = Get-ScheduledTask -TaskName $name -ErrorAction Stop
    if ($task.State -eq 'Running') { throw "Stop and ask for help: $name is running. No release was installed." }
    $mode = switch ($name) { 'Danmade Patch Agent - Machine' { 'Machine' }; 'Danmade Patch Agent - Health' { 'Health' }; 'Danmade Patch Agent - User' { 'user' }; default { 'notification' } }
    $action = if ($mode -in @('user', 'notification')) {
      New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\wscript.exe" -Argument "//B //NoLogo `"$root\Start-DanmadePatchAgentHidden.vbs`" $mode"
    } else {
      New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-NoProfile -NonInteractive -ExecutionPolicy AllSigned -File `"$root\Invoke-DanmadePatchAgentCached.ps1`" -Mode $mode"
    }
    $plans += [pscustomobject]@{ Task = $task; Action = $action }
  }
}

if (-not $PSCmdlet.ShouldProcess($root, 'Stage verified release, publish local bootstrap, and update requested patch-agent actions')) { return }
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Run this installer from an elevated operator terminal.' }
foreach ($ancestor in @('C:\ProgramData\DanmadePatchAgent', $root)) {
  if ((Test-Path -LiteralPath $ancestor) -and ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Refusing reparse point: $ancestor" }
}
$parent = Split-Path -Parent $root
if (Test-Path -LiteralPath $parent) {
  $parentAcl = Get-Acl -LiteralPath $parent
  $broadWriteRules = @(
  foreach ($rule in $parentAcl.Access) {
    $sid = $rule.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value
    $writeMask = [Security.AccessControl.FileSystemRights]::Write -bor [Security.AccessControl.FileSystemRights]::Delete -bor [Security.AccessControl.FileSystemRights]::DeleteSubdirectoriesAndFiles -bor [Security.AccessControl.FileSystemRights]::ChangePermissions -bor [Security.AccessControl.FileSystemRights]::TakeOwnership
    if ($rule.AccessControlType -eq 'Allow' -and ($sid -in @('S-1-1-0','S-1-5-11','S-1-5-32-545') -or $sid -match '-513$') -and ([int]$rule.FileSystemRights -band ([int]$writeMask -bor 0x50000000))) { $rule }
  }
  )
  if ($broadWriteRules.Count -gt 0) {
    if (-not $SecureParent) { throw 'Patch-agent parent directory allows non-admin writes. Review its ACL and explicitly use -SecureParent before staging.' }
    # Protect only the fixed agent parent. Retain read/execute and all non-broad principals.
    $beforeSddl = $parentAcl.Sddl
    $parentAcl = Get-SecuredAgentParentAcl -Acl $parentAcl
    Set-Acl -LiteralPath $parent -AclObject $parentAcl
    foreach ($rule in (Get-Acl $parent).Access) {
      $sid = $rule.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value
      if ($rule.AccessControlType -eq 'Allow' -and ($sid -in @('S-1-1-0','S-1-5-11','S-1-5-32-545') -or $sid -match '-513$') -and ([int]$rule.FileSystemRights -band ([int]$writeMask -bor 0x50000000))) { throw 'Parent ACL write protection did not take effect; no release published.' }
    }
    [ordered]@{ timestamp=(Get-Date).ToString('o'); path=$parent; beforeSddl=$beforeSddl; afterSddl=(Get-Acl $parent).Sddl } | ConvertTo-Json | Set-Content (Join-Path $parent ('cache-parent-acl-' + [guid]::NewGuid().ToString('N') + '.json')) -Encoding UTF8
  }
}
$acl = New-Object Security.AccessControl.DirectorySecurity
$acl.SetAccessRuleProtection($true, $false)
$acl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
foreach ($entry in @(@('S-1-5-18', 'FullControl'), @('S-1-5-32-544', 'FullControl'), @('S-1-5-32-545', 'ReadAndExecute'))) {
  $sid = New-Object Security.Principal.SecurityIdentifier($entry[0])
  $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sid, $entry[1], 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
}
if (-not (Test-Path -LiteralPath $parent)) {
  New-Item -ItemType Directory -Path $parent | Out-Null
  Set-Acl -LiteralPath $parent -AclObject $acl
}
New-Item -ItemType Directory -Path $root -Force | Out-Null
Set-Acl -LiteralPath $root -AclObject $acl
$fileAcl = New-Object Security.AccessControl.FileSecurity
$fileAcl.SetAccessRuleProtection($true, $false)
$fileAcl.SetOwner((New-Object Security.Principal.SecurityIdentifier('S-1-5-32-544')))
foreach ($entry in @(@('S-1-5-18', 'FullControl'), @('S-1-5-32-544', 'FullControl'), @('S-1-5-32-545', 'ReadAndExecute'))) {
  $fileAcl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule((New-Object Security.Principal.SecurityIdentifier($entry[0])), $entry[1], 'Allow')))
}
$pointer = Join-Path $root 'current.json'
$current = $null
if (Test-Path -LiteralPath $pointer) {
  if ((Get-Item -LiteralPath $pointer -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Refusing current pointer reparse point.' }
  try { $current = Get-Content -LiteralPath $pointer -Raw | ConvertFrom-Json } catch { }
}
$unchanged = $false
if ($current -and [string]$current.id -match '^[a-f0-9]{32}$' -and $current.signerThumbprint -eq $ExpectedSignerThumbprint) {
  $existing = Join-Path $root $current.id
  $unchanged = Test-Path -LiteralPath $existing -PathType Container
  if ($unchanged -and ((Get-Item -LiteralPath $existing -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Refusing release directory reparse point.' }
  foreach ($file in $files) {
    $path = Join-Path $existing $file
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $unchanged = $false; break }
    if ((Get-Item -LiteralPath $path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Refusing release file reparse point.' }
    if ((Get-ReleaseFileHash -Path $path) -ne $hashes[$file]) { $unchanged = $false; break }
  }
}
$id = if ($unchanged) { [string]$current.id } else { [guid]::NewGuid().ToString('N') }
$directory = Join-Path $root $id
if (-not $unchanged) {
  New-Item -ItemType Directory -Path $directory | Out-Null
  foreach ($file in $files) { Copy-Item -LiteralPath (Join-Path $SourcePath $file) -Destination (Join-Path $directory $file) }
}
foreach ($file in $files) {
  if ((Get-ReleaseFileHash -Path (Join-Path $directory $file)) -ne $hashes[$file]) { throw "Source changed during staging: $file" }
}
# Revalidate copied bytes rather than relying on preflight against a changing share.
foreach ($file in ($files | Where-Object { $_ -like '*.ps1' })) {
  $signature = Get-AuthenticodeSignature -FilePath (Join-Path $directory $file)
  if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Thumbprint -ne $ExpectedSignerThumbprint) { throw "Staged signature check failed: $file" }
}
foreach ($file in @('Invoke-DanmadePatchAgentCached.ps1', 'Start-DanmadePatchAgentHidden.vbs')) {
  $destination = Join-Path $root $file
  if ((Test-Path -LiteralPath $destination) -and ((Get-Item -LiteralPath $destination -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Refusing reparse point: $destination" }
  if (Test-Path -LiteralPath $destination) { Set-Acl -LiteralPath $destination -AclObject $fileAcl }
  $temporary = "$destination.$id.tmp"
  Copy-Item -LiteralPath (Join-Path $directory $file) -Destination $temporary
  if ([IO.File]::Exists($destination)) { [IO.File]::Replace($temporary, $destination, [System.Management.Automation.Language.NullString]::Value) } else { [IO.File]::Move($temporary, $destination) }
}
if ((Test-Path -LiteralPath $pointer) -and ((Get-Item -LiteralPath $pointer -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw "Refusing reparse point: $pointer" }
$previousPointer = $null
if (Test-Path -LiteralPath $pointer) {
  Set-Acl -LiteralPath $pointer -AclObject $fileAcl
  if (-not $unchanged) {
    $previousPointer = Join-Path $root ("previous-$id.json")
    Copy-Item -LiteralPath $pointer -Destination $previousPointer
    Set-Acl -LiteralPath $previousPointer -AclObject $fileAcl
  }
}
$release = [ordered]@{ id = $id; signerThumbprint = $ExpectedSignerThumbprint; policySha256 = (Get-ReleaseFileHash -Path (Join-Path $directory 'danmade-patch-agent.policy.json')); installedAt = $(if ($unchanged) { $current.installedAt } else { (Get-Date).ToString('o') }); verifiedAt = (Get-Date).ToString('o') }
$temporary = "$pointer.$id.tmp"
$release | ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding UTF8
if ([IO.File]::Exists($pointer)) { [IO.File]::Replace($temporary, $pointer, [System.Management.Automation.Language.NullString]::Value) } else { [IO.File]::Move($temporary, $pointer) }
foreach ($plan in $plans) { Set-ScheduledTask -TaskName $plan.Task.TaskName -TaskPath $plan.Task.TaskPath -Action $plan.Action | Out-Null }
[pscustomobject]@{ Status = 'LocalReleasePublished'; Directory = $directory; PreviousPointer = $previousPointer; TasksUpdated = $plans.Count; PolicySha256 = $release.policySha256 }

# SIG # Begin signature block
# MIIgEQYJKoZIhvcNAQcCoIIgAjCCH/4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDbAVEdUUGzZDHI
# t3cAMcN8TQzF63WsDkGgSh4qEKNwhqCCGiwwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwgga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphB
# cr48RsAcrHXbo0ZodLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6p
# vF4uGjwjqNjfEvUi6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHe
# HYNnQxqXmRinvuNgxVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEd
# gkFiDNYiOTx4OtiFcMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjU
# jsZvkgFkriK9tUKJm/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bR
# VFLeGkuAhHiGPMvSGmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeS
# LsJygoLPp66bkDX1ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIV
# NSaz7BX8VtYGqLt9MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL
# 6s36czwzsucuoKs7Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2Zd
# SoQbU2rMkpLiQ6bGRinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFU
# eEY0qVjPKOWug/G6X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/
# BAgwBgEB/wIBADAdBgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0j
# BBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8E
# PDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVz
# dGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEw
# DQYJKoZIhvcNAQELBQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/
# T8ObXAZz8OjuhUxjaaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQ
# E7jU/kXjjytJgnn0hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9r
# EVKChHyfpzee5kH0F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y
# 1IsA0QF8dTXqvcnTmpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gx
# dEkMx1NKU4uHQcKfZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3t
# y9qIijanrUR3anzEwlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcy
# tL5TTLL4ZaoBdqbhOhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEB
# YTptMSbhdhGQDpOXgpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud
# /v4+7RWsWCiKi9EOLLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiS
# uEtQvLsNz3Qbp7wGWqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZP
# ubdcMIIG7TCCBNWgAwIBAgIQCE/cM09+RU7bww+P+ZIYNTANBgkqhkiG9w0BAQsF
# ADBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNV
# BAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hB
# MjU2IDIwMjUgQ0ExMB4XDTI2MDgwNTAwMDAwMFoXDTM3MTEwNDIzNTk1OVowYzEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMTswOQYDVQQDEzJE
# aWdpQ2VydCBTSEEyNTYgUlNBNDA5NiBUaW1lc3RhbXAgUmVzcG9uZGVyIDIwMjYg
# MTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBALZ7pvLJ/s1K+NSbTGWz
# /TjGMPh8CQ6RucZCLv5anHzWJjF/NWJrFIhy24fcpKXlgRiky4WAawDfU3YP0BMx
# t9l3Dm5oCG5Z69AqEN1kgHg2epx+l+lZBcmJCcN0ASURML5uFIS80sZsDwO3BSkU
# xDjLJhBI+qiZP3aixAC/qEGLjsBNlLol9VZ7pfGEXiMlneJIC5/YKuizVzNFKZZE
# eoy/0B8Zm+nzKBgSWG52lCO1w+nCg6XpCtklTJXeIg283hw7TmmsZXR+SMbjbrEO
# vZ3fP2VxIgeR28Y90ZStd3F9VuA5RVynb/whITPAo9b75Zr4Ta6Mj3URm26QZYMn
# /FnbuTegcoRcFEZ9FOqM5T6MTdtr/n74lIT/ug0eeOzmZ6QTFg33otX+bFRsIolv
# ykE1jive4PuESaT8zzVeFWDAMDtozNgLctkGD1ZjkEyZtJrLl5ya0m5doH/ScpaZ
# CZVl6pNUOCybMc/kxC6EAmSJY24L0yYKD1Nkddsnb/ItVKi/2nXpQNMu1PT5prW8
# 3vV8d67WowuUs0HdY4H8AMLGvdL/WHEj3ZnqMqAQQP9u3Ai9t+5eQ02GDwy0ODjd
# zi0xlp70W+ow63/0++YDEX1M0iwgUHwbrJvfpklkZQvw3+kv3vUPItdwroczk9ic
# flf55W1zOEKAcJVAIXpcMCU9AgMBAAGjggGVMIIBkTAMBgNVHRMBAf8EAjAAMB0G
# A1UdDgQWBBQUyWOKMC7USvtulPPm40B+9ezN4jAfBgNVHSMEGDAWgBTvb1NK6eQG
# fHrK4pBW9i/USezLTjAOBgNVHQ8BAf8EBAMCB4AwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwgwgZUGCCsGAQUFBwEBBIGIMIGFMCQGCCsGAQUFBzABhhhodHRwOi8vb2Nz
# cC5kaWdpY2VydC5jb20wXQYIKwYBBQUHMAKGUWh0dHA6Ly9jYWNlcnRzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEy
# NTYyMDI1Q0ExLmNydDBfBgNVHR8EWDBWMFSgUqBQhk5odHRwOi8vY3JsMy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hB
# MjU2MjAyNUNBMS5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcB
# MA0GCSqGSIb3DQEBCwUAA4ICAQCNxTphHp1SCt+ZrAmAfn0oQLFr0mLywSLaDXQI
# ENoyKqxrFbJblzCVP/pkXmwXOdrOpWygLzlT12os5ipDCy35RBCg2UMeApEtrfGh
# z45F4Wt4WGdNdIbRWt3YTYJmpR+b7lr4d7Uwn+H600u4D7RnOGf8Wj4UNgAdZkfH
# hHv1mx9EVh71SJelcEN/oORSjXzdjfw1iZH9d8Nh/thn6hH23d+VsPAr6GAYyzSA
# 02nXD1nYLI7Ijmiv+xLCiYC41DSFYL3GhTiy0PxpawPtGRyaBVGzq+UiTfM8pD7K
# VyF5aQyWP4KhVGUUTnmm/RlYJoW3TiXA/+t0YcT2oRVBm3JETjajHug2AL+v5jht
# KVnd3D0rbHXEu27o+Q8p4sEWPMqKDB+qbceb6T/6WcwTwXmQ9lOCLLYcsQeSWmvK
# qzpAec9etE14jOQAzLKWdE3w/TCaKtLRaRT7LCkRYVnhA2D73FLje1O5b3HR5eHs
# 0NzU/+xX7NbEdcofy0W3Wdwd1XOqtlpg/JgwtKfZM5dqO94lbUveOiJBI+xZEbGR
# sMNbXmMREUTgu+Oca7Y73MPWcslIx2VhkSKSXjDbD6rgg39H5Mh7QfieAIjWagkJ
# Nt68Yfim6cjEzVSiLSeZfdkr5dtFPTW6jATlWJdYeeDRGCyatf8R1hSjzSvdN8yW
# QPT9gzCCBu4wggXWoAMCAQICE1YAAAD7uI/G5fXiwIQAAQAAAPswDQYJKoZIhvcN
# AQELBQAwSDEVMBMGCgmSJomT8ixkARkWBWxvY2FsMRcwFQYKCZImiZPyLGQBGRYH
# bmV4cG9ydDEWMBQGA1UEAxMNbmV4cG9ydC1sb2NhbDAeFw0yNjA3MDYxNjEzMjla
# Fw0yNzA3MDYxNjEzMjlaMHQxFTATBgoJkiaJk/IsZAEZFgVsb2NhbDEXMBUGCgmS
# JomT8ixkARkWB25leHBvcnQxEjAQBgNVBAsTCURpdmlzaW9uczEaMBgGA1UECxMR
# TmV4cG9ydCBTb2x1dGlvbnMxEjAQBgNVBAMTCURhbiBQdXBlazCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAMZQ4zJZc+eugIuUhFyk+eVYC9gkueExn2Sx
# JvuJ+K9v9ykS2jTErJWxVzACQokpRuu7X7t7ZTkMIjQvl7ovzxDsRHOZ93gvBR61
# D7rwHUll1lMqOUzoA/5pHRMmlrxzLk0TK5I7XW0aCU94RHfktU0DIb0MpiGUyKcO
# 5LCVer4D0lXzqTTIDiizY+lJPRYKNfxnK3tGdROO/NObmWeFOlKnDlNxmCBsSMv8
# e6gBs7pk70TTUAf6yq3sCQ45Oo+q58Z0dMw5+5pir2K3pY/En/BWsULa5Qn9Xy66
# lcr19YpqVjq83NoOtgmaw4hGz8zjGdZLOj9YA3bKQNxDDZqcCqkCAwEAAaOCA6Mw
# ggOfMD4GCSsGAQQBgjcVBwQxMC8GJysGAQQBgjcVCIWJh2SF+5lKhLmbLYXr+FeH
# /JYDgRWE9e0NhtrrNgIBZAIBADATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNVHQ8B
# Af8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU
# egRE7DA1bfxWX5Xw8IpDWAn/P3wwHwYDVR0jBBgwFoAUp/6qLWthmwXFyv6qX9Gn
# /xl4WqQwgc0GA1UdHwSBxTCBwjCBv6CBvKCBuYaBtmxkYXA6Ly8vQ049bmV4cG9y
# dC1sb2NhbCxDTj1PS0NEQzAyLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2
# aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPW5leHBvcnQsREM9
# bG9jYWw/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNz
# PWNSTERpc3RyaWJ1dGlvblBvaW50MIIBhQYIKwYBBQUHAQEEggF3MIIBczCBrgYI
# KwYBBQUHMAKGgaFsZGFwOi8vL0NOPW5leHBvcnQtbG9jYWwsQ049QUlBLENOPVB1
# YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRp
# b24sREM9bmV4cG9ydCxEQz1sb2NhbD9jQUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0
# Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhvcml0eTBeBggrBgEFBQcwAYZSaHR0cDov
# L09LQ0RDMDIubmV4cG9ydC5sb2NhbC9DZXJ0RW5yb2xsL09LQ0RDMDIubmV4cG9y
# dC5sb2NhbF9uZXhwb3J0LWxvY2FsKDEpLmNydDBgBggrBgEFBQcwAoZUZmlsZTov
# Ly8vT0tDREMwMi5uZXhwb3J0LmxvY2FsL0NlcnRFbnJvbGwvT0tDREMwMi5uZXhw
# b3J0LmxvY2FsX25leHBvcnQtbG9jYWwoMSkuY3J0MDIGA1UdEQQrMCmgJwYKKwYB
# BAGCNxQCA6AZDBdkYW4ucHVwZWtAbmV4cG9ydC5sb2NhbDBOBgkrBgEEAYI3GQIE
# QTA/oD0GCisGAQQBgjcZAgGgLwQtUy0xLTUtMjEtMTI1NDMxNjM0MC0yODY0MzMx
# OTY4LTQ3NzkwNTQ0MS0xMTA1MA0GCSqGSIb3DQEBCwUAA4IBAQCbq973SpW2d0oM
# JAyWvJiZWDyuPCmzSubthWPRxVhbi6vdh40+rt3fZ6uTJ2Z8MdGLqtmNPK1HXBuj
# 0X0/2m/5MUzyYmZjfqeRyBuSzXAdfWuZ27J5jhOOj51+KDAPEOxjRea6MZBr1ezC
# K7asWYX1mcK19E2ZlnZTic0fx2cCVJNRTTNaQmYvxnJq0XNbbbaZcOSIW8TTERsb
# p+NCKwm0mzfrIvA3QaX0DrpYTdNLgwm0u7kD7d/f9S/aGbXZp6L/VUqNPWc9Ld0x
# OAu8qXJ/zFjIbs4MeClYeLsWTn0f6h3+J++egix4lZZTBfmn4tcPJL2QT7WzVKPG
# +lmwH+uqMYIFOzCCBTcCAQEwXzBIMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxFzAV
# BgoJkiaJk/IsZAEZFgduZXhwb3J0MRYwFAYDVQQDEw1uZXhwb3J0LWxvY2FsAhNW
# AAAA+7iPxuX14sCEAAEAAAD7MA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcC
# AQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYB
# BAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINFVTfhpi43I
# sh7bH3SM0TsVfC879x2p/ii0BX0jMnx/MA0GCSqGSIb3DQEBAQUABIIBAH3Q0reE
# UmsDaY+Ed0zcsjRDvxjddhlWQvOsDv68BFYJSa9gqcRh3gi+ImVNBC3pEHcC4eXt
# 0kCaaExKF7grEbSm0P2Auz6zj/OrStaqTgBLEU+kRyIkKC8PxYZI5FISJnOqIZ7R
# J76Inr+YwG5DFQzAfY9neRbwzpXhtYOP8xahUou3PCnVG2dpFDpfjZz9QjYI2Gpn
# a8jYl3/sOmWVVPXHBJWIylttiknzCoVDIjS9OgNHlFRUE5GMD0qibw+PtKLkT2BW
# 7r1VWELgP2tevNRz5Sb1aoS5964mDvr1a2GCGPcwNj4E5nYOJpikN9Wy9/olR3+t
# /+L2V5Sqhoq4JVyhggMmMIIDIgYJKoZIhvcNAQkGMYIDEzCCAw8CAQEwfTBpMQsw
# CQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERp
# Z2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIw
# MjUgQ0ExAhAIT9wzT35FTtvDD4/5khg1MA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjYxMDE2MTc0ODM4
# WjAvBgkqhkiG9w0BCQQxIgQgU432WYnT1lJHPJlKUdu+CQsAWOQHVu7mIkdVxUkI
# fn0wDQYJKoZIhvcNAQEBBQAEggIAKN44Pt65xru8bB3D30Emh637R9yChw5Vwx5s
# 1RxDRdIrAEJt/lrT44hnQRuB6NCvxt7lDYTrnO9MiK9lFyXjuQB9dek+Y1ja0CkS
# GoBUAMcvOqjEPq+CCRnna2AQxSG7iQFm8/ciQcBDcXT2RNI2NYBBndlso03+u0J/
# DFvOspcr0Le5Z1DTO/hqcWjdsbsoo5kwCXRJJgwZ57Jb60jga4Mby4/tAGwR2Gem
# uOxN858AdfHN5fjRA9DIkl6bIEVf5rTM02TdVYsUA6UDM/s9jhJs8D8BRo5hFnZG
# abWOIKCHWyP/439HId684br4w66sjgtu+uoX1tU7rF/6KR/M9gkmvPUYKzTeeAGz
# E2tRazldauT1kfEmiaKtBRcS7Zkv4AFbYA7e+okv4etAEAJVefzmQwTpNiWMuRAU
# doDx6eGOUKCZWAP1aC+cpYmGj+C6bgMHoxSYMEBW7bDiNawuHhoGepGWaK6zlibn
# VFberX37V1MQnyY8G/NJVsTLccayk7NJb4cRFQfiqDLkrvyqmGadOnccACWvOF8O
# aYJSWChKdPhcDTRe+UtmzH5Fhoq5AAOKRkmn3+eHWVB39nPgFbSC/dQPGVry6VLa
# TGbEZ8Z6Nxw/qnKhVO2ExbbymEEzM0kON1PpHmj2dUCRuj+GYzftVM6RKVtPneEO
# 1s9+F8U=
# SIG # End signature block
