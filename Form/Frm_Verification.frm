VERSION 5.00
Begin VB.Form Frm_Verification 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Verification"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5820
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
   ScaleHeight     =   1200
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Txt_Verification 
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "-"
      TabIndex        =   1
      Top             =   555
      Width           =   5775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please enter your Verification Info"
      Height          =   465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5820
   End
End
Attribute VB_Name = "Frm_Verification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varForm As Form

Private Sub Txt_Verification_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        QFrm_Main.mVerification Txt_Verification, varForm
        Unload Me
    End Select
End Sub

Sub mInfo(paramForm As Form)
    Set varForm = paramForm
End Sub
