VERSION 5.00
Begin VB.Form Frm_Tray 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ"
      Size            =   18
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Label lbl_Text 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
      Width           =   645
   End
   Begin VB.Label lbl_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Caption         =   "¹Ø±Õ"
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
Attribute VB_Name = "Frm_Tray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Me
        .Left = Screen.Width - Me.Width
        .Top = Screen.Height - Me.Height - 500
    End With
    With lbl_Close
        .Left = Me.Width - .Width - 5
        .Top = 5
    End With
    With lbl_Title
        .Left = 10
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
        .Width = Me.Width - 10
        .Height = Me.Height - 10
    End With
End Sub

Private Sub lbl_Close_Click()
    Unload Me
End Sub

Sub ShowTray(strTitle As String, strText As String, _
             Optional BackColor As Long = &H8000000F, _
             Optional ForeColor As Long = &H80000012)
    lbl_Title.Caption = strTitle
    lbl_Text.Caption = strText
    Me.BackColor = BackColor
    Me.ForeColor = ForeColor
    lbl_Close.ForeColor = ForeColor
    lbl_Title.ForeColor = ForeColor
    lbl_Text.ForeColor = ForeColor
    Shape_Main.BorderColor = ForeColor
    Me.Visible = True
End Sub
