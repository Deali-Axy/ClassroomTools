VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QClass_QApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'保持属性值的局部变量
Private mvarQSettings As QClass_QAppSettings    '局部复制
Private mvarQDebug As QClass_QDebug    '局部复制
Private mvarQEverydayTips As QClass_QEverydayTips    '局部复制

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Const GWL_EXSTYLE = -20
Private Const WS_EX_LAYERED = &H80000


Public Property Set QEverydayTips(ByVal vData As QClass_QEverydayTips)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.QEverydayTips = Form1
    Set mvarQEverydayTips = vData
End Property


Public Property Get QEverydayTips() As QClass_QEverydayTips
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.QEverydayTips
    Set QEverydayTips = mvarQEverydayTips
End Property




Public Property Set QDebug(ByVal vData As QClass_QDebug)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.QDebug = Form1
    Set mvarQDebug = vData
End Property


Public Property Get QDebug() As QClass_QDebug
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.QDebug
    Set QDebug = mvarQDebug
End Property


Public Property Set QSettings(ByVal vData As QClass_QAppSettings)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.QSettings = Form1
    Set mvarQSettings = vData
End Property


Public Property Get QSettings() As QClass_QAppSettings
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.QSettings
    Set QSettings = mvarQSettings
End Property

Public Property Get Name() As String
    Name = QApp_Name
End Property

Public Property Get Author() As String
    Author = QApp_Author
End Property

Public Property Get Author_Website() As String
    Author_Website = QApp_Author_Website
End Property

Public Property Get Version() As String
    Version = QApp_Version
End Property

Public Property Get MajorVersion() As Integer
    MajorVersion = QApp_MajorVersion
End Property

Public Property Get MinorVersion() As Integer
    MinorVersion = QApp_MinorVersion
End Property

Public Property Get ReleaseVersion() As Integer
    ReleaseVersion = QApp_ReleaseVersion
End Property

Public Property Get Comments() As String
    Comments = QApp_Comments
End Property

Public Property Get FileDescription() As String
    FileDescription = QApp_FileDescription
End Property

Public Property Get Website() As String
    Website = QApp_Website
End Property

Public Property Get LegalCopyright() As String
    LegalCopyright = QApp_LegalCopyright
End Property

Public Property Get LegalTrademarks() As String
    LegalTrademarks = QApp_LegalTrademarks
End Property

Public Property Get SubTitle() As String
    SubTitle = QApp_SubTitle
End Property

Public Property Get Title() As String
    Title = QApp_Title
End Property

Public Property Get ThreadID()
    ThreadID = App.ThreadID
End Property

Public Property Get hInstance()
    hInstance = App.hInstance
End Property

Public Property Get Path() As String
    Path = App.Path
End Property

Public Property Let DefaultConfigFile(vData As String)
    Dim strTmp As String
    If Len(Dir(vData)) > 0 Then
        strTmp = vData
    Else
        strTmp = QApp_Default_ConfigFile
    End If
    QApp_Default_ConfigFile = strTmp
End Property

Public Property Get DefaultConfigFile() As String
    DefaultConfigFile = QApp_Default_ConfigFile
End Property


Public Function SubmitStrMsg(Msg As String)
    On Error GoTo Err
    Dim ReturnVal As String
    ReturnVal = QMod_Main.QMsgHandingStr(Msg)
    SubmitStrMsg = ReturnVal
    Exit Function
Err:
    QDB.Runtime_Error Me.Name & "_SubmitStrMsg", Err.Description, Err.Number
    Resume Next
End Function

Private Function SubmitMsg(ByRef QMsg As QMsg_Struct)
    On Error GoTo Err
    Dim ReturnVal As String
    ReturnVal = QMod_Main.QMsgHanding(QMsg)
    SubmitMsg = ReturnVal
    Exit Function
Err:
    QDB.Runtime_Error Me.Name & "_SubmitMsg", Err.Description
    Resume Next
End Function

Public Sub ExitQApp()
    QMod_Main.QMsgHandingStr "exitqapp"
End Sub

Public Sub Tray_ShowText(strTitle As String, strText As String, _
                         Optional BackColor As Long = &H8000000B, _
                         Optional ForeColor As Long = &H80000012)
    QFrm_Tray.ShowTray strTitle, strText, BackColor, ForeColor
End Sub

Public Sub SetGlobalFormsTransparency(ByVal limpid As Long)    ' 设置全局窗体透明度
    Dim tmpForm As Form
    For Each tmpForm In Forms
        Call SetWindowLong(tmpForm.hwnd, GWL_EXSTYLE, GetWindowLong(tmpForm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(tmpForm.hwnd, 0, limpid, LWA_ALPHA)     'limpid在0--255之间
    Next
End Sub

Public Sub SetFormTransparency(ByRef TargetForm As Form, ByVal limpid As Long)    ' 设置窗体透明度
    Call SetWindowLong(TargetForm.hwnd, GWL_EXSTYLE, GetWindowLong(TargetForm.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(TargetForm.hwnd, 0, limpid, LWA_ALPHA)     'limpid在0--255之间
End Sub

Private Sub Class_Initialize()
    Set mvarQSettings = New QClass_QAppSettings
    Set mvarQDebug = New QClass_QDebug
    Set mvarQEverydayTips = New QClass_QEverydayTips
End Sub

Private Sub Class_Terminate()
    Set mvarQSettings = Nothing
    Set mvarQDebug = Nothing
    Set mvarQEverydayTips = Nothing
End Sub


