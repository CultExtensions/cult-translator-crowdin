Option Explicit
Dim shell, fso, base, support, ps1, psCmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
base = fso.GetParentFolderName(WScript.ScriptFullName)
support = fso.BuildPath(base, "SupportFiles")
ps1 = fso.BuildPath(support, "Install.ps1")
If Not fso.FileExists(ps1) Then
  MsgBox "Keep Install.vbs and the SupportFiles folder in the same place, then try again.", vbCritical, "Cult Connector"
  WScript.Quit 1
End If
shell.CurrentDirectory = support
psCmd = "powershell.exe -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & ps1 & """"
shell.Run psCmd, 0, False
Set shell = Nothing
Set fso = Nothing
