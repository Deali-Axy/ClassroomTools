VERSION 5.00
Begin VB.Form QFrm_Load 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   7560
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Tmr_Load 
      Interval        =   10
      Left            =   4920
      Top             =   840
   End
   Begin ClassroomTools.进度条 pb 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "pb"
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00FFCC00&
      BorderWidth     =   3
      Height          =   855
      Left            =   5580
      Top             =   300
      Width           =   1695
   End
   Begin VB.Image Img_AppIcon 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   120
      Picture         =   "QFrm_Load.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lbl_AppName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "App"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   930
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1440
   End
End
Attribute VB_Name = "QFrm_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    QDB.Log Me.Name & " Activate"
End Sub

Private Sub Form_Load()
    On Error GoTo load_err
    lbl_AppName.Caption = QApp_Title
    With lbl_AppName
        .Left = (Me.Width + (Img_AppIcon.Width / 2) - .Width) / 2
    End With
    With Shape
        .Top = 1
        .Left = 1
        .Width = Me.Width
        .Height = Me.Height
    End With
    If Dir(App_Icon_Gif) = "" Then
        'Img_AppIcon.Picture = QFrm_Main.Icon
    Else
        Img_AppIcon.Picture = LoadPicture(App_Icon_Gif)
    End If
    QDB.Log Me.Name & " Load hWnd=" & Me.hwnd
    Exit Sub
load_err:
    QDB.Runtime_Error Me.Name & "_Load", Err.Description, Err.Number
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QDB.Log Me.Name & " Unload"
End Sub

Private Sub Tmr_Load_Timer()
    Static i As Integer
    i = i + 1
    pb.Value = i
    If i >= 100 Then
        Tmr_Load.Enabled = False
        Load QFrm_Main
        QFrm_Main.Show
        Unload Me
    End If
End Sub
