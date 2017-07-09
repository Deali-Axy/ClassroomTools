VERSION 5.00
Begin VB.UserControl QTextButton 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipBehavior    =   0  '无
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HitBehavior     =   2  '使用画图
   ScaleHeight     =   1110
   ScaleWidth      =   4800
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTextButton"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1740
   End
End
Attribute VB_Name = "QTextButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'///////////////////////////事件声明///////////////////////////
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event DbClick()
'///////////////////////////常量声明///////////////////////////
Const QName As String = "QTextButton"
'///////////////////////////变量声明///////////////////////////
Dim cProp As QTextButtonProperties
Dim fLblMouseState As MouseState
'///////////////////////////枚举定义///////////////////////////
Private Enum MouseState
    m_None = 0
    m_Over = 1
    m_Down = 2
End Enum
'///////////////////////////结构定义///////////////////////////
Private Type QTextButtonProperties    'QTextButton属性
    AutoSize As Boolean
    BackColor As OLE_COLOR
    Color_Down As OLE_COLOR
    Color_Normal As OLE_COLOR
    Color_Over As OLE_COLOR
    Font As StdFont
    Text As String
End Type
'///////////////////////////属性实现///////////////////////////
Public Property Let AutoSize(param As Boolean)
    cProp.AutoSize = param
    Lbl.AutoSize = param
    Lbl.Top = 100
    Lbl.Left = 100
    If param Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
    PropertyChanged "AutoSize"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = cProp.AutoSize
End Property

Public Property Let BackColor(param As OLE_COLOR)
    cProp.BackColor = param
    UserControl.BackColor = param
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = cProp.BackColor
End Property


Public Property Let Color_Down(param As OLE_COLOR)
    cProp.Color_Down = param
    PropertyChanged "Color_Down"
End Property

Public Property Get Color_Down() As OLE_COLOR
    Color_Down = cProp.Color_Down
End Property

Public Property Let Color_Normal(param As OLE_COLOR)
    cProp.Color_Normal = param
    Lbl.ForeColor = param
    PropertyChanged "Color_Normal"
End Property

Public Property Get Color_Normal() As OLE_COLOR
    Color_Normal = cProp.Color_Normal
End Property

Public Property Let Color_Over(param As OLE_COLOR)
    cProp.Color_Over = param
    PropertyChanged "Color_Over"
End Property

Public Property Get Color_Over() As OLE_COLOR
    Color_Over = cProp.Color_Over
End Property

Public Property Set Font(param As StdFont)
    On Error GoTo Err
    Set cProp.Font = param
    Set UserControl.Font = param
    Set Lbl.Font = param
    Lbl.Top = 100
    Lbl.Left = 100
    If cProp.AutoSize Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
    PropertyChanged "Font"
    Exit Property
Err:
    Debug.Print "[" & QName & ".Property.Set.Font.Err]" & Err.Number & "," & Err.Description
End Property

Public Property Get Font() As StdFont
    On Error GoTo Err
    Set Font = cProp.Font
    Exit Property
Err:
    Debug.Print "[" & QName & ".Property.Get.Font.Err]" & Err.Number & "," & Err.Description
End Property

Public Property Let Text(param As String)
    cProp.Text = param
    Lbl.Caption = param
    Lbl.Top = 100
    Lbl.Left = 100
    If cProp.AutoSize Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
    PropertyChanged "Text"
End Property

Public Property Get Text() As String
    Text = cProp.Text
End Property
'///////////////////////////内部事件实现///////////////////////////
Private Sub Lbl_Change()
    Lbl.Top = 100
    Lbl.Left = 100
    If cProp.AutoSize Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
End Sub

Private Sub Lbl_Click()
    RaiseEvent Click
End Sub

Private Sub Lbl_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub Lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.ForeColor = cProp.Color_Down
    fLblMouseState = m_Down
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fLblMouseState = m_Down Then
    ElseIf fLblMouseState = m_None Then
        fLblMouseState = m_Over
        Lbl.ForeColor = cProp.Color_Over
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Private Sub Lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fLblMouseState = m_Down Then
        Lbl.ForeColor = cProp.Color_Over
        fLblMouseState = m_Over
    ElseIf fLblMouseState = m_None Then
        Lbl.ForeColor = cProp.Color_Normal
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
    If Not (TypeOf UserControl.Parent Is MDIForm) Then
        Set cProp.Font = Parent.Font
        Set UserControl.Font = Parent.Font
        Set Lbl.Font = Parent.Font
    End If
    With cProp
        .AutoSize = True
        Lbl.AutoSize = True
    End With
    With Lbl
        .Top = 100
        .Left = 100
    End With
    If cProp.AutoSize Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
    fLblMouseState = m_None
    PropertyChanged "AutoSize"
    'PropertyChanged "Font"
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.ForeColor = cProp.Color_Normal
    fLblMouseState = m_None
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.ForeColor = cProp.Color_Normal
    fLblMouseState = m_None
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl.ForeColor = cProp.Color_Normal
    fLblMouseState = m_None
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        cProp.AutoSize = .ReadProperty("AutoSize")
        Lbl.AutoSize = cProp.AutoSize
        cProp.BackColor = .ReadProperty("BackColor")
        UserControl.BackColor = cProp.BackColor
        cProp.Color_Down = .ReadProperty("Color_Down")
        cProp.Color_Normal = .ReadProperty("Color_Normal")
        Lbl.ForeColor = cProp.Color_Normal
        cProp.Color_Over = .ReadProperty("Color_Over")
        Set cProp.Font = .ReadProperty("Font")
        Set Lbl.Font = cProp.Font
        cProp.Text = .ReadProperty("Text")
        Lbl.Caption = .ReadProperty("Text")
    End With
End Sub

Private Sub UserControl_Resize()
    Lbl.Top = 100
    Lbl.Left = 100
    If cProp.AutoSize Then
        UserControl.Width = Lbl.Left + Lbl.Width + 100
        UserControl.Height = Lbl.Top + Lbl.Height + 100
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "AutoSize", cProp.AutoSize
        .WriteProperty "BackColor", cProp.BackColor
        .WriteProperty "Color_Down", cProp.Color_Down
        .WriteProperty "Color_Normal", cProp.Color_Normal
        .WriteProperty "Color_Over", cProp.Color_Over
        .WriteProperty "Font", cProp.Font
        .WriteProperty "Text", cProp.Text
    End With
End Sub

Public Sub Reset()
    Lbl.ForeColor = cProp.Color_Normal
    fLblMouseState = m_None
End Sub
