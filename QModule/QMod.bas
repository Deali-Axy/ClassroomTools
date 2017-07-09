Attribute VB_Name = "QMod_Main"
Option Explicit
#Const QApp_Activate_Server = False
Public fExit As Boolean

Sub Main()
    fExit = False
    #If QApp_Activate_Server Then
        If Not HasActive Then
            QFrm_QApp_Activate.Show
        Else
            QApp_Main
        End If
    #Else
        QApp_Main
    #End If
End Sub

Function QMsg(paramQMsg As String)
    On Error GoTo Err
    Select Case paramQMsg
    Case "load qfrm_main"
        QApp_Main
    Case "exitqapp"
        Mod_HookSkinner.Detach QFrm_Main.Hwnd
        Dim F As Form
        For Each F In Forms
            Unload F
        Next
    End Select
    QDB.Log "[QMsg]" & paramQMsg
    Exit Function
Err:
    QDB.Runtime_Error "QMod_Main_QMsg", Err.Description, Err.Number
    Resume Next
End Function

Function QMsgHanding(paramQMsg As QMsg_Struct)
    On Error GoTo Err
    
    Exit Function
Err:
    QDB.Runtime_Error "QMod_Main_QMsgHanding", Err.Description, Err.Number
    Resume Next
End Function

Function QMsgHandingStr(paramQMsgStr As String)
    On Error GoTo Err
    Select Case paramQMsgStr
        Case "load qfrm_main"
            QApp_Main
        Case "exitqapp"
            Dim F As Form
            For Each F In Forms
                Unload F
            Next
    End Select
    QDB.Log "[QMsgStr]" & paramQMsgStr
    Exit Function
Err:
    QDB.Runtime_Error "QMod_Main_QMsg", Err.Description, Err.Number
    Resume Next
End Function

Function HasActive() As Boolean
    If GetSetting(QApp_Name, "Info", "active") = "true" Then
        HasActive = True
    Else
        HasActive = False
    End If
End Function

Function IsNewVer() As Boolean
    If GetSetting(QApp_Name, "Info", "build") <> "" Then
        If Val(GetSetting(QApp_Name, "Info", "build")) < App.Revision Then
            IsNewVer = True
        Else
            IsNewVer = False
        End If
    Else
        IsNewVer = True
    End If
End Function
