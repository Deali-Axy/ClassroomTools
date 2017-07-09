VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------- 系统热键 ----------------
'需要 mdlSysHotKeySupport 标准模块的支持
'需要 clsHashLK、clsStack、clsSubClass 类模块 和 mdlSubClass 标准模块的支持

'#提示信息字符串常量已国际化，但都是静态常量，不需替换占位符号；如不 _
 是多国语言版本，可不添加 mdlRes 模块

'#国际化字符串常量：
#If Not con_MulLang Then
    Private Const syhkResStr1 As String = "系统热键定义成功！"
    Private Const syhkResStr2 As String = "(无)"
    Private Const syhkResStr3 As String = "系统热键定义失败。请重新指定一个其他的热键。"
#End If


'================================================================================
'一个本对象只管理一个系统热键
'关于系统热键 ID 的说明：
'本对象一旦建立，就有一个热键ID了：lIDSysHotKey （有效范围1～32767）。 _
 '该 ID 将伴随整个对象的生存期且保持不变，而不论当前对象中是否设置了一个系统热键
'设置热键时，将使用该ID；取消设置热键时，仍保留该ID留作下次重新再设置热键时使用
'该ID是在 Class_Initialize 时，向 mdlSysHotKeySupport 申请的
'在对象 Class_Terminate 时，向 mdlSysHotKeySupport 释放ID
'如果在本对象运行期间，lIDSysHotKey=0，则本对象不能用于管理任何系统热键，即 _
 '本对象不能使用
'================================================================================


Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const MOD_ALT = &H1
Private Const MOD_WIN = &H8


Public Event SysKeyPressed()

Private lHwnd As Long    '目标窗体句柄
Private lKeyCode As Long    '本对象管理的热键的虚拟键码
Private lShift As Long  '本对象管理的热键的组合键
Private lIDSysHotKey As Long  '系统热键ID
Private lIsSysHotKeySet As Boolean    '是否已通过本对象设置了一个系统热键

Private mHashVKSpe As clsHashLK    '特殊按键的虚拟键码与键名对应关系：Key=虚拟键码，DataString=键名

Public Sub RaiseSysKeyPressedEvent()
    RaiseEvent SysKeyPressed
End Sub

Public Function SetASysHotKey(ByVal hwnd As Long, _
                              ByVal kKeyCode As Long, _
                              ByVal kShift As Long, _
                              Optional ByVal bShowMsg As Boolean = True) As Boolean

'设置系统热键
'hwnd：窗体句柄
'KeyCode：按键虚拟键码
'Shift：为Alt(=1),Control(=2),Shift(=4),WinKey(=8)的组合
'bShowMsg：是否提示成功或失败

    If lIDSysHotKey = 0 Then GoTo errH    '本对象没有热键ID，不能使用
    If KeyCode < 0 Or Shift < 0 Then GoTo errH

    '======== 设置系统热键，ID 为 lIDSysHotKey ========
    Dim ret As Long
    ret = RegisterHotKey(hwnd, lIDSysHotKey, kShift, kKeyCode)
    If ret = 0 Then GoTo errH    '定义失败

    '======== 子类化窗口 hwnd，自定义窗口程序为 _
     'mdlSysHotKeySupport 模块的 SHKWndProc 函数  ========
    '向 mdlSysHotKeySupport 请求子类化，标准模块会防止重复子类化
    If Not SHKSubClassHwnd(hwnd) Then GoTo errH

    '======== 系统热键设置成功 ========
    '相关参数赋值到模块级属性变量
    lHwnd = hwnd
    lKeyCode = kKeyCode
    lShift = kShift
    lIsSysHotKeySet = True

    '显示提示信息，定义的热键是什么
    If bShowMsg Then
        Dim strHotKey As String
        strHotKey = ShiftStr()
        If strHotKey = syhkResStr2 Then strHotKey = "" Else strHotKey = strHotKey & " + "
        strHotKey = strHotKey & KeyCodeStr() & " (KeyCode=" & CStr(KeyCode) & ")"
        MsgBox syhkResStr1 & strHotKey, vbInformation
    End If

    '返回成功
    SetASysHotKey = True
    Exit Function
errH:
    SetASysHotKey = False
    If bShowMsg Then MsgBox syhkResStr3, 16
End Function


Public Sub UnSetSysHotKey()
'取消通过本对象设置过的一个系统热键， _
  '但保留本对象的 lIDSysHotKey，以备仍能用本对象设置新的系统热键
'只有在 Class_Terminate 时，才释放 lIDSysHotKey

    If lIsSysHotKeySet Then
        '取消系统热键
        UnregisterHotKey lHwnd, lIDSysHotKey

        '取消子类化 lhwnd
        '向 mdlSysHotKeySupport 请求取消子类化，标准模块会防止取消尚在使用的子类化
        SHKUnSubClassHwnd lHwnd

        '变量清0
        lHwnd = 0
        lKeyCode = 0
        lShift = 0
        '不能清除 lIDSysHotKey，只有在 Class_Terminate 时才能清除
        lIsSysHotKeySet = False
    End If
End Sub









Public Sub ToComboShifts(tComboBox As ComboBox)
'把 Shift,Alt,Ctrl,WinKey 键的各种组合添加到 tComboBox 列表
'并且把 tComboBox 的 ItemData 设置为组合键常量值，便于设计设置热键的程序界面
    Dim i As Integer
    With tComboBox
        tComboBox.Clear
        For i = 0 To 15
            .AddItem ShiftStr(i)
            .itemData(.ListCount - 1) = i
        Next i
        If .ListCount Then .ListIndex = 0
    End With
End Sub

Public Sub ToComboKeys(tComboBox As ComboBox)
'把设置热键的各种键添加到 tComboBox 列表
'并且把 tComboBox 的 ItemData 设置为对应键的虚拟键码，便于设计设置热键的程序界面

    Dim arrKeys() As Long
    Dim i As Long
    With tComboBox
        .Clear
        'F1-F12 功能键
        For i = 1 To 12
            .AddItem "F" & CStr(i)
            .itemData(.ListCount - 1) = 111 + i
        Next i
        '字母键
        For i = 1 To 26
            .AddItem Chr(64 + i)
            .itemData(.ListCount - 1) = 64 + i
        Next i
        '数字键
        For i = 0 To 9
            .AddItem Chr(48 + i)
            .itemData(.ListCount - 1) = 48 + i
        Next i

        '特殊键
        For i = 1 To mHashVKSpe.GetKeyArray(arrKeys)
            .AddItem mHashVKSpe.DataString(arrKeys(i))
            .itemData(.ListCount - 1) = arrKeys(i)
        Next i

        '小键盘数字键
        For i = 0 To 9
            .AddItem "NumPAD:" & Chr(48 + i)
            .itemData(.ListCount - 1) = 96 + i
        Next i

        '默认选中第0项
        If .ListCount Then .ListIndex = 0
    End With
End Sub


Public Function ShiftStr(Optional ByVal kShift As Long = -1) As String
'把系统热键的组合键 kShift 转换为字符串说明
'缺省转换本对象保存的系统热键的组合键 lShift 属性
    Dim strResu As String

    If kShift < 0 Then kShift = lShift
    If kShift < 0 Then ShiftStr = "": Exit Function

    If kShift And MOD_CONTROL Then strResu = strResu & "Ctrl + "
    If kShift And MOD_ALT Then strResu = strResu & "Alt + "
    If kShift And MOD_SHIFT Then strResu = strResu & "Shift + "
    If kShift And MOD_WIN Then strResu = strResu & "WinKey + "

    If Len(strResu) Then
        strResu = Left(strResu, Len(strResu) - 3)    '除去最后的" + "
    Else
        strResu = syhkResStr2
    End If
    ShiftStr = strResu
End Function

Public Function KeyCodeStr(Optional ByVal kKeyCode As Long = -1) As String
'把虚拟键码 kKeyCode 转换为字符串说明
'缺省转换本对象保存的系统热键的按键属性 lKeyCode

    If kKeyCode < 0 Then kKeyCode = lKeyCode
    If kKeyCode < 0 Then KeyCodeStr = "": Exit Function

    Select Case kKeyCode
    Case 112 To 123
        'F1-F12功能键
        KeyCodeStr = "F" & CStr((kKeyCode - 111))
    Case 96 To 105
        '小键盘数字键
        KeyCodeStr = "NumPAD:" & Chr(kKeyCode - 48)
    Case Else
        If mHashVKSpe.IsKeyExist(kKeyCode) Then
            KeyCodeStr = mHashVKSpe.DataString(kKeyCode, False)
        Else
            KeyCodeStr = Chr(kKeyCode)
        End If
    End Select
End Function





Private Sub Class_Initialize()
'在本对象建立初期，就要向 mdlSysHotKeySupport 申请一个热键 ID， _
  '保存在 lIDSysHotKey 中，同时让 mdlSysHotKeySupport 登记本对象 _
  '地址与此ID的对应关系
'在本对象 Class_Terminate 时，向 mdlSysHotKeySupport 取消 _
  '此登记并释放热键 ID

'函数可能返回0表示失败，这时 lIDSysHotKey 为 0，则不能继续使用本对象建立系统热键
    lIDSysHotKey = SHKRegOneObject(ObjPtr(Me))


    '------ 初始化特殊键虚拟键码 ------
    Dim s() As String, i As Integer
    Set mHashVKSpe = New clsHashLK
    mHashVKSpe.AlloMem 255

    mHashVKSpe.Add 0, 8, 0, "Backspace", False
    mHashVKSpe.Add 0, 9, 0, "Tab", False
    mHashVKSpe.Add 0, 13, 0, "Return", False
    mHashVKSpe.Add 0, 19, 0, "Pause", False
    mHashVKSpe.Add 0, 44, 0, "PrintScreen", False
    mHashVKSpe.Add 0, 45, 0, "Insert", False
    mHashVKSpe.Add 0, 46, 0, "Delete", False
    mHashVKSpe.Add 0, 91, 0, "LeftWin", False
    mHashVKSpe.Add 0, 92, 0, "RightWin", False
    mHashVKSpe.Add 0, 93, 0, "App", False
    mHashVKSpe.Add 0, 20, 0, "CapsLock", False
    mHashVKSpe.Add 0, 144, 0, "NumLock", False
    mHashVKSpe.Add 0, 145, 0, "ScrollLock", False
    '32-40
    s = Split("SPACE,PageUp,PageDown,END,HOME,LEFT,UP,RIGNT,DOWN", ",")
    For i = 0 To UBound(s)
        mHashVKSpe.Add 0, i + 32, 0, s(i), False
    Next i
    '106-111
    s = Split("NumPAD:*,NumPAD:+,NumPAD:Sep,NumPAD:-,NumPAD:.,NumPAD:/", ",")
    For i = 0 To UBound(s)
        mHashVKSpe.Add 0, i + 106, 0, s(i), False
    Next i
End Sub

Private Sub Class_Terminate()
'取消通过本对象设置过的一个系统热键（如果有的话）
    UnSetSysHotKey

    '向 mdlSysHotKeySupport 取消登记并释放热键 ID：lIDSysHotKey
    SHKUnRegOneObject lIDSysHotKey

    '释放 mHashVKSpe 引用
    Set mHashVKSpe = Nothing
End Sub

















'目标窗口句柄（只读属性）
Public Property Get hWndNotify() As Long
    hWndNotify = lHwnd
End Property

'热键虚拟键码（只读属性）
Public Property Get KeyCode() As Long
    KeyCode = lKeyCode
End Property

'热键按键组合键（只读属性）
Public Property Get Shift() As Long
    Shift = lShift
End Property

'本对象现在是否已经管理着一个系统热键（只读属性）
Public Property Get IsSysHotKeySet() As Boolean
    IsSysHotKeySet = lIsSysHotKeySet
End Property


'系统热键 ID （在整个本对象的生存期内，其值不变）（只读属性）
Public Property Get IDSysHotKey() As Long
    IDSysHotKey = lIDSysHotKey
End Property


