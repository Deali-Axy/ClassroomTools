Attribute VB_Name = "modTray"

Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
                                      (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
                                        (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const TRAY_CALLBACK = (&H400 + 1001&)
Public Const GWL_WNDPROC = -4

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206

Public Enum WIN_STATUS
    STA_MIN
    STA_NORMAL
End Enum

Public glWinRet As Long
Public OrgWinRet As Long
Public Status As WIN_STATUS    '保存窗体状态

Public Function CallbackMsgs(ByVal wHwnd As Long, ByVal wMsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
    On Error Resume Next
    If wMsg = TRAY_CALLBACK Then
        With Frm_Menu
            Select Case CLng(lp_id)
            Case WM_RBUTTONUP    '右键
                .PopupMenu .Menu_托盘菜单, , , , .Menu_Tray_打开教室管理
            Case WM_LBUTTONUP    '左键
                If QFrm_Main.Visible Then
                    QFrm_Main.Visible = False
                Else
                    QFrm_Main.Visible = True
                End If
            End Select
        End With
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)
End Function

