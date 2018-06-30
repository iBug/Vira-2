' Vira Installer

Dim ShellApp, Fso
Set ShellApp = CreateObject("Shell.Application")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim WD
WorkDir = Fso.GetParentFolderName(WScript.ScriptFullName)

If Not Fso.FileExists(WorkDir & "\Vira_2.vbs") Then
  MsgBox "Vira 2 main script not found." & vbCrLf & "Installation aborted.", 16, "Vira 2 Installer"
  WScript.Quit 1
End If

ShellApp.ShellExecute "wscript.exe", """" & WorkDir & "\Vira_2.vbs"" install", , "runas", 1