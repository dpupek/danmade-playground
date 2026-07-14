Option Explicit

' Launches only the fixed, signed Danmade Patch Agent user commands without a console host.
Dim shell
Dim mode
Dim command
Dim quote
Dim exitCode
Dim powerShellPath

If WScript.Arguments.Count <> 1 Then
  WScript.Quit 87
End If

mode = LCase(WScript.Arguments(0))
quote = Chr(34)
Set shell = CreateObject("WScript.Shell")
powerShellPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe")

Select Case mode
  Case "user"
    command = quote & powerShellPath & quote & _
      " -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy AllSigned -File " & _
      quote & "\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent\danmade-patch-agent.ps1" & quote & _
      " -PolicyPath " & quote & "\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent\danmade-patch-agent.policy.json" & quote & _
      " -Mode User"
  Case "notification"
    command = quote & powerShellPath & quote & _
      " -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy AllSigned -File " & _
      quote & "\\nexport.local\SYSVOL\nexport.local\scripts\danmade-patch-agent\Show-DanmadePatchAgentNotification.ps1" & quote & _
      " -MachineStatePath " & quote & "C:\ProgramData\DanmadePatchAgent\State\pending-notification.json" & quote & _
      " -RepeatHours 4 -MaxShowCount 12 -PopupTimeoutSeconds 90"
  Case Else
    WScript.Quit 87
End Select

exitCode = shell.Run(command, 0, True)
WScript.Quit exitCode
