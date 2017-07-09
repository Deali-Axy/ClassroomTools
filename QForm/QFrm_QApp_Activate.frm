VERSION 5.00
Begin VB.Form QFrm_QApp_Activate 
   BackColor       =   &H00FFCC00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2040
   End
   Begin VB.Label lbl_Trip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "按任意键继续"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   1710
   End
   Begin VB.Label lbl_info 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "应用正在激活"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   2520
      TabIndex        =   1
      Top             =   1800
      Width           =   2430
   End
   Begin VB.Label lbl_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   45
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   2805
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "QFrm_QApp_Activate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private ExitFlag As Boolean

Private Sub Form_Click()
    If ExitFlag Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ExitFlag Then
        Unload Me
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0
    '或把上面注释掉用这行，效果相同 SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Load()
    ExitFlag = False
    With Shape
        .Top = 1
        .Left = 1
        .Width = Me.Width
        .Height = Me.Height
    End With
    With Lbl_Title
        .Caption = QApp_Title
        .Left = (Me.Width - .Width) / 2
    End With
    With lbl_Trip
        .Left = (Me.Width - .Width) / 2
        .Visible = False
    End With
    ShowLblInfo "应用正在激活"
    Tmr.Enabled = True
End Sub

Sub ActivateProc()
    If Not ExitFlag Then
        Dim QCM As QApp_CommonMsg
        QCM = QMod_QApp_Activate.QApp_Acitvate(QApp_Activate_Email)
        If QCM.TrueOrFalse Then
            ShowLblInfo "应用激活成功"
            ExitFlag = True
        Else
            ShowLblInfo "应用激活失败"
            ExitFlag = True
        End If
        Me.SetFocus
        lbl_Trip.Visible = True
    End If
End Sub

Sub ShowLblInfo(Info As String)
    With lbl_info
        .Caption = Info
        .Left = (Me.Width - .Width) / 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QMod_Main.QMsg "load qfrm_main"
End Sub

Private Sub lbl_info_Click()
    If ExitFlag Then
        Unload Me
    End If
End Sub

Private Sub lbl_Title_Click()
    If ExitFlag Then
        Unload Me
    End If
End Sub

Private Sub Tmr_Timer()
    ActivateProc
    Tmr.Enabled = False
End Sub
