<# Tests cache dispatch and fail-closed validation in .scratch; signer API is mocked. #>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$root=Join-Path (Join-Path $repoRoot '.scratch') ('cache-tests-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$original=Get-Content (Join-Path $repoRoot 'windows-update-scripts\Invoke-DanmadePatchAgentCached.ps1') -Raw
$original=($original -split '# SIG # Begin signature block',2)[0]
$count=0
foreach($case in @(
  @{Name='valid';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=0},
  @{Name='valid-whatif';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=0},
  @{Name='bad-signature';Signature='HashMismatch';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=2},
  @{Name='wrong-signer';Signature='Valid';Signer=('B'*40);HashValid=$true;IdValid=$true;Exit=2},
  @{Name='bad-policy';Signature='Valid';Signer=('A'*40);HashValid=$false;IdValid=$true;Exit=2},
  @{Name='path-traversal';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$false;Exit=2},
  @{Name='stale-policy';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=2},
  @{Name='missing-age';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=2},
  @{Name='future-age';Signature='Valid';Signer=('A'*40);HashValid=$true;IdValid=$true;Exit=2}
)) {
  # Arrange: immutable fixture release with a stub that cannot install packages.
  $fixture=Join-Path $root $case.Name
  $id=[guid]::NewGuid().ToString('N')
  $releasePath=Join-Path $fixture $id
  New-Item -ItemType Directory -Path $releasePath -Force | Out-Null
  '{"enabled":false}' | Set-Content (Join-Path $releasePath 'danmade-patch-agent.policy.json')
  '[CmdletBinding(SupportsShouldProcess=$true)]param($PolicyPath,$Mode); if($PSCmdlet.ShouldProcess($env:CACHE_TEST_MARKER,"Dispatch")){Set-Content -LiteralPath $env:CACHE_TEST_MARKER -Value "Dispatched"}; exit 0' | Set-Content (Join-Path $releasePath 'danmade-patch-agent.ps1')
  $hash=(Get-FileHash (Join-Path $releasePath 'danmade-patch-agent.policy.json') -Algorithm SHA256).Hash
  if(-not $case.HashValid){$hash='0'*64}
  $pointerId=if($case.IdValid){$id}else{'..\escape'}
  $verified = if ($case.Name -eq 'stale-policy') { (Get-Date).AddDays(-8) } elseif ($case.Name -eq 'future-age') { (Get-Date).AddDays(1) } else { Get-Date }
  $pointer = [pscustomobject]@{id=$pointerId;signerThumbprint=('A'*40);policySha256=$hash;verifiedAt=$verified.ToString('o')}
  if ($case.Name -eq 'missing-age') { $pointer.PSObject.Properties.Remove('verifiedAt') }
  $pointer | ConvertTo-Json | Set-Content (Join-Path $fixture 'current.json')
  $mock="function Get-AuthenticodeSignature { param(`$FilePath) [pscustomobject]@{Status='$($case.Signature)';SignerCertificate=[pscustomobject]@{Thumbprint='$($case.Signer)'}} }"
  $bootstrap=Join-Path $fixture 'Invoke-DanmadePatchAgentCached.ps1'
  $original.Replace('Set-StrictMode -Version Latest', $mock + [Environment]::NewLine + 'Set-StrictMode -Version Latest') | Set-Content $bootstrap
  $marker=Join-Path $fixture 'dispatched.txt'
  if(Test-Path $marker){throw 'Initial assert: fixture already dispatched.'}
  $oldLocal=$env:LOCALAPPDATA;$oldMarker=$env:CACHE_TEST_MARKER
  try {
    $env:LOCALAPPDATA=$fixture;$env:CACHE_TEST_MARKER=$marker
    # Act: execute bootstrap without modifying production storage or tasks.
    if ($case.Name -eq 'valid-whatif') {
      & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $bootstrap -Mode User -WhatIf | Out-Null
    } else {
      & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $bootstrap -Mode User | Out-Null
    }
    $exit=$LASTEXITCODE
  } finally {$env:LOCALAPPDATA=$oldLocal;$env:CACHE_TEST_MARKER=$oldMarker}
  # Assert: failures never dispatch; evidence stays inside the fixture.
  if($exit -ne $case.Exit){throw "$($case.Name): wrong exit $exit"}
  if((Test-Path $marker) -ne ($case.Exit -eq 0 -and $case.Name -ne 'valid-whatif')){throw "$($case.Name): unexpected dispatch"}
  if($case.Exit -ne 0){
    $events=@(Get-Content (Join-Path $fixture 'DanmadePatchAgent\Events\patch-agent.jsonl') | ForEach-Object {$_|ConvertFrom-Json})
    if(@($events | Where-Object status -eq 'LocalBootstrapFailed').Count -ne 1){throw "$($case.Name): missing local failure evidence"}
  }
  $count++
}
[pscustomobject]@{Status='Passed';Cases=$count;ScratchRoot=$root;Signing='Mocked API; actual Authenticode validation is a separate release gate'} | ConvertTo-Json
