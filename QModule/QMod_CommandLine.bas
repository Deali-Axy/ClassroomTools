Attribute VB_Name = "QMod_CommandLine"
Option Explicit

Public Function CommandProc(ParamString As String) As Boolean
    On Error GoTo Err
    Dim CommandString As String, CmdParam As String
    If InStr(1, ParamString, " ") <> 0 Then
        CommandString = Mid(ParamString, 1, InStr(1, ParamString, " ") - 1)
    Else
        CommandString = ParamString
    End If
    Select Case CommandString
    Case "project"
        Shell "explorer H:\_code\vb\Tools\DATOOLS\DATOOLSproj.vbp", vbNormalFocus
    Case "exit"
        fExit = True
        QApp.ExitQApp
        Unload QFrm_Main
    Case "about"
        QFrm_About.Show
    Case Else
        CommandProc = False
        Exit Function
    End Select
    CommandProc = True
    Exit Function
Err:

End Function

Function IsChinese(paramStr As String) As Boolean
    On Error GoTo Err
    If Asc(paramStr) > 128 Or Asc(paramStr) < 0 Then
        IsChinese = True
    Else
        IsChinese = False
    End If
    Exit Function
Err:
    IsChinese = False
End Function

