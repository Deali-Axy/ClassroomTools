VERSION 5.00
Begin VB.Form QFrm_Tray 
   BorderStyle     =   0  'None
   ClientHeight    =   4605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "QFrm_Tray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   3720
   End
   Begin VB.Label lbl_Text 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3240
      TabIndex        =   2
      Top             =   840
      Width           =   645
   End
   Begin VB.Image Img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lbl_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   720
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lbl_Close 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "关闭"
      Height          =   465
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape_Main 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      Height          =   1455
      Left            =   120
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "QFrm_Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'实现鼠标拖动窗口的API
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim AnimeFlag As String, ExitFlag As Boolean, h As Integer, w As Integer, t As Integer
Attribute ExitFlag.VB_VarUserMemId = 1073938432
Attribute h.VB_VarUserMemId = 1073938432
Attribute w.VB_VarUserMemId = 1073938432
Attribute t.VB_VarUserMemId = 1073938432

Private Sub Form_Load()
    Call InitUI
    ExitFlag = False
    Tmr.Enabled = True
    QApp.SetFormTransparency Me, 200
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'实现鼠标拖动窗口
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    '或把上面注释掉用这行，效果相同 SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ExitFlag Then
    Else
        AnimeFlag = "unload"
        Tmr.Enabled = True
        h = 4600
        w = 4410
        t = 0
        Cancel = 1
    End If
End Sub

Private Sub lbl_Close_Click()
    Unload Me
End Sub

Sub InitUI()
    Dim hwnd As Long
    Dim WinRECT As RECT
    hwnd = FindWindow("Shell_TrayWnd", vbNullString)
    GetWindowRect hwnd, WinRECT
    t = Screen.Height - (WinRECT.Bottom - WinRECT.Top) * 13
    With Me
        .Left = Screen.Width - Me.Width
        '.Top = Screen.Height - Me.Height - (WinRECT.Bottom - WinRECT.Top) * 15
        .Top = Screen.Height - (WinRECT.Bottom - WinRECT.Top) * 13
    End With
    With lbl_Close
        .Left = Me.Width - .Width - 80
        .Top = 5
    End With
    With Lbl_Title
        .Left = 50
        .Top = (lbl_Close.Top + lbl_Close.Height / 2) / 4
    End With
    With lbl_Text
        .Left = 50
        .Top = lbl_Close.Top + lbl_Close.Height + 50
        .Width = Me.Width - 80
        .Height = Me.Height - .Top - 50
    End With
    With Line1
        .X1 = 0
        .X2 = Me.Width
        .Y1 = lbl_Close.Top + lbl_Close.Height + 5
        .Y2 = lbl_Close.Top + lbl_Close.Height + 5
    End With
    With Shape_Main
        .Left = 10
        .Top = 10
        .Width = Me.Width - 20
        .Height = Me.Height - 38
    End With
End Sub

Private Sub lbl_Text_Click()
    MsgBox lbl_Text.Caption
End Sub

Private Sub Tmr_Timer()
    Select Case AnimeFlag
        Case "load"
            If h >= 4600 Then
                h = 0
                Me.Height = 4600
                Tmr.Enabled = False
                Exit Sub
            End If

            If h <= 4600 Then
                Me.Height = h
                Me.Top = Me.Top - 100
                h = h + 100
            End If

        Case "unload"
            If Me.Top >= Screen.Height Then
                Tmr.Enabled = False
                ExitFlag = True
                t = 0
                h = 0
                w = 0
                Unload Me
                Exit Sub
            End If

            If h >= 0 Then
                Me.Top = Me.Top + t
                t = t + 1
            End If

        Case "refresh1"
            If Me.Left < 0 And Me.Width >= 100 Then
                Me.Width = Me.Width - 100
                If Me.Width <= 100 Then
                    Me.Width = 0
                End If
            Else
                Me.Left = Me.Left - 1000
            End If

            If Me.Width = 0 Then
                Me.Left = Screen.Width
                Me.Width = 4410
                AnimeFlag = "refresh2"
            End If

        Case "refresh2"
            If Me.Left >= (Screen.Width - 4600) Then
                Me.Left = Me.Left - 100
            Else
                Me.Left = (Screen.Width - 4410)
                Tmr.Enabled = False
            End If
    End Select

    Debug.Print "Me.Left=" & Me.Left & vbCrLf & "Me.Width=" & Me.Width
End Sub

Sub ShowTray(strTitle As String, strText As String, _
             Optional BackColor As Long = &H8000000F, _
             Optional ForeColor As Long = &H80000012)
'UI处理
    With Lbl_Title
        .Caption = strTitle
        .ForeColor = ForeColor
    End With
    With lbl_Text
        .Caption = strText
        .ToolTipText = strText
        .Visible = True
        .ForeColor = ForeColor
        .Height = Me.Height - .Top - 50
    End With
    Me.BackColor = BackColor
    Me.ForeColor = ForeColor
    lbl_Close.ForeColor = ForeColor
    Shape_Main.BorderColor = ForeColor
    Img.Visible = False

    If Me.Visible Then
        AnimeFlag = "refresh1"
    Else
        AnimeFlag = "load"
        Me.Visible = True
    End If
    Tmr.Enabled = True
End Sub

Sub ShowTrayGraphics(strTitle As String, strText As String, _
                     ByRef Pic As StdPicture, _
                     Optional ShowText As Boolean = False, _
                     Optional ShowImgBorder As Boolean = False, _
                     Optional BackColor As Long = &H8000000F, _
                     Optional ForeColor As Long = &H80000012)

    Me.BackColor = BackColor
    Me.ForeColor = ForeColor
    lbl_Close.ForeColor = ForeColor
    Shape_Main.BorderColor = ForeColor

    With lbl_Text
        If ShowText Then
            .Visible = True
            .Height = 420
        Else
            .Visible = False
        End If
        .Caption = strText
        .ToolTipText = strText
        .ForeColor = ForeColor
    End With

    With Lbl_Title
        .Caption = strTitle
        .ForeColor = ForeColor
    End With

    With Img
        If ShowImgBorder Then
            Img.BorderStyle = 1
        Else
            Img.BorderStyle = 0
        End If
        If ShowText Then
            .Top = lbl_Text.Top + lbl_Text.Height + 50
            .Width = 3500
            .Height = 3500
        Else
            .Top = Lbl_Title.Top + Lbl_Title.Height + 50
            .Width = 4000
            .Height = 4000
        End If
        .Left = (Me.Width - .Width) / 2
        .Visible = True
        Set .Picture = Pic
    End With

    If Me.Visible Then
        AnimeFlag = "refresh1"
    Else
        AnimeFlag = "load"
        Me.Visible = True
    End If
    Tmr.Enabled = True
End Sub

Sub SetForeColor(ForeColor As Long)
    On Error GoTo Err
    Me.ForeColor = ForeColor
    Shape_Main.BorderColor = ForeColor
    Line1.BorderColor = ForeColor
    lbl_Close.ForeColor = ForeColor
    Lbl_Title.ForeColor = ForeColor
    lbl_Text.ForeColor = ForeColor
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->SetForeColor()", Err.Description, Err.Number
End Sub

Sub SetBackColor(BackColor As Long)
    On Error GoTo Err
    Me.BackColor = BackColor
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "->SetBackColor()", Err.Description, Err.Number
End Sub

Sub RandomColor()    '随机色彩
    Dim ForeColor As Long, BackColor As Long
    Randomize
    Dim R, G, b As Integer
    R = Int(Rnd * 255)
    G = Int(Rnd * 255)
    b = Int(Rnd * 255)
    ForeColor = RGB(R, G, b)
    BackColor = RGB(R Xor 255, G Xor 255, b Xor 255)
    Me.SetBackColor BackColor
    Me.SetForeColor ForeColor
End Sub

Sub RandomForeColor()    '随机前景色
    Dim ForeColor As Long
    Randomize
    Dim R, G, b As Integer
    R = Int(Rnd * 255)
    G = Int(Rnd * 255)
    b = Int(Rnd * 255)
    ForeColor = RGB(R, G, b)
    Me.SetForeColor ForeColor
End Sub

Sub RandomBackColor()    '随机背景色
    Dim BackColor As Long
    Randomize
    Dim R, G, b As Integer
    R = Int(Rnd * 255)
    G = Int(Rnd * 255)
    b = Int(Rnd * 255)
    BackColor = RGB(R, G, b)
    Me.SetBackColor BackColor
End Sub
