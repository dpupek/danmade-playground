<# Pure ACL transformation tests. No filesystem permissions are changed. #>
[CmdletBinding()]param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$path=Join-Path (Split-Path $PSScriptRoot) 'Install-DanmadePatchAgentCache.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Installer parse failed.'}
$node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-SecuredAgentParentAcl'},$true)
Invoke-Expression $node.Extent.Text
$count=0
foreach($grant in @('0x116','GA','GW')) {
foreach($inherited in @($true,$false)){
  # Arrange: default ProgramData grants, including creator-owner and broad inherited Write.
  $flags=if($inherited){'OICI ID'.Replace(' ','')}else{'OICI'}
  $creatorFlags=if($inherited){'OICIIOID'}else{'OICIIO'}
  $acl=New-Object Security.AccessControl.DirectorySecurity
  $acl.SetSecurityDescriptorSddlForm("O:SYG:SYD:(A;$flags;FA;;;SY)(A;$flags;FA;;;BA)(A;$flags;0x1200a9;;;BU)(A;$flags;$grant;;;BU)(A;$creatorFlags;GA;;;CO)")
  $before=$acl.Sddl
  if(@($acl.Access|Where-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value -eq 'S-1-5-32-545' -and ([int]$_.FileSystemRights -band (278 -bor 0x50000000)) }).Count -eq 0){throw 'Initial assert: fixture lacks user write grant.'}
  # Act: use actual production transform, not a reimplementation.
  $secured=Get-SecuredAgentParentAcl -Acl $acl
  # Assert: read and privileged access remain; no user writes; transformation is idempotent.
  if($acl.Sddl -ne $before){throw 'Input ACL was mutated.'}
  if(-not $secured.AreAccessRulesProtected){throw 'Parent still inherits unsafe grants.'}
  $users=@($secured.Access|Where-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value -eq 'S-1-5-32-545' })
  if(@($users|Where-Object { [int]$_.FileSystemRights -band (278 -bor 0x50000000) }).Count){throw 'Broad writes survived normalization.'}
  if(@($users|Where-Object { $_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::ReadAndExecute }).Count -eq 0){throw 'User read access was removed.'}
  foreach($sid in @('S-1-5-18','S-1-5-32-544')){
    if(@($secured.Access|Where-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value -eq $sid -and ($_.FileSystemRights -band [Security.AccessControl.FileSystemRights]::FullControl) -eq [Security.AccessControl.FileSystemRights]::FullControl }).Count -ne 1){throw 'Privileged access was removed.'}
  }
  if((Get-SecuredAgentParentAcl $secured).Sddl -ne $secured.Sddl){throw 'ACL normalization is not idempotent.'}
  $count++
}
}
[pscustomobject]@{Status='Passed';Cases=$count;Scope='In-memory ACLs only'}|ConvertTo-Json
