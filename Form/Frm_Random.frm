VERSION 5.00
Begin VB.Form Frm_Random 
   BackColor       =   &H00FFCC00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖĞĞÄ
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   6720
   End
   Begin ClassroomTools.QTextButton QTBtn_Main 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   6960
      Width           =   2595
      _extentx        =   4577
      _extenty        =   1720
      autosize        =   -1
      backcolor       =   16763904
      color_down      =   16711935
      color_normal    =   16777215
      color_over      =   65280
      font            =   "Frm_Random.frx":0000
      text            =   "Ëæ»ú³éºÅ"
      text            =   "Ëæ»ú³éºÅ"
   End
   Begin ClassroomTools.QTextButton QTBtn_Exit 
      Height          =   975
      Left            =   3720
      TabIndex        =   3
      Top             =   6960
      Width           =   2595
      _extentx        =   4577
      _extenty        =   1720
      autosize        =   -1
      backcolor       =   16763904
      color_down      =   16711935
      color_normal    =   16777215
      color_over      =   65280
      font            =   "Frm_Random.frx":0028
      text            =   "ÍË³ö³éºÅ"
      text            =   "ÍË³ö³éºÅ"
   End
   Begin VB.Shape Shape 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   6495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   0
      X2              =   6480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Lbl_Title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ëæ»ú³éºÅ 1.0"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6240
   End
   Begin VB.Label lbl_Main 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   300
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5160
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   6240
   End
End
Attribute VB_Name = "Frm_Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MaxNum As Integer = 37    '°à¼¶ÈËÊı
Dim fState As fStateEnum
Dim R As Integer
Private Enum fStateEnum
    Running = 0
    Pause = 1
End Enum

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        QTBtn_Main_Click
    Case 27
        Unload Me
    End Select
End Sub

Private Sub Form_Load()
    fState = Pause
    With Shape
        .Top = 10
        .Left = 10
    End With
    With Me
        .Width = Shape.Left + Shape.Width
        .Height = Shape.Top + Shape.Height
    End With
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    QTBtn_Main.Reset
    QTBtn_Exit.Reset
End Sub

Private Sub lbl_Main_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    QTBtn_Main.Reset
    QTBtn_Exit.Reset
End Sub

Private Sub Lbl_Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    QTBtn_Main.Reset
    QTBtn_Exit.Reset
End Sub

Private Sub QTBtn_Exit_Click()
    Unload Me
End Sub

Private Sub QTBtn_Main_Click()
    If fState = Pause Then
        Tmr.Enabled = True
        lbl_Main = "--"
        QTBtn_Main.Text = "Í£Ö¹³éºÅ"
        fState = Running
    ElseIf fState = Running Then
        Tmr.Enabled = False
        lbl_Main = R
        QTBtn_Main.Text = "¿ªÊ¼³éºÅ"
        fState = Pause
    End If
End Sub

Private Sub Tmr_Timer()
    Dim tmp As Integer
    Static bS As Boolean
    Randomize
    If bS Then
        lbl_Main = ">-"
        bS = False
    Else
        lbl_Main = "-<"
        bS = True
    End If
    tmp = Int((MaxNum + 1) * Rnd) + 1
    R = tmp
    Debug.Print Str(R)
End Sub
