VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------系统托盘 类------------------------
'需要 mdlSysTraySupport 标准模块的支持
'需要 clsHashLK、clsSubClass 类模块 和 mdlSubClass 标准模块的支持

'# 国际化：无提示信息字符串常量

'================================================================================
'一个窗体最多只能使用一个本对象设置托盘
'================================================================================

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Type NOTIFYICONDATA    'Shell_NotifyIcon 函数使用的类型
    cbSize As Long  '本类型数据的字节大小
    hwnd As Long    '要接收托盘消息的窗口的句柄
    uID As Long    '应用程序自定义的托盘图标ID（可设为0）
    uFlags As Long  '一个掩码，表示类型中哪些成员的值有效
    uCallbackMessage As Long    '应用程序自定义的消息编号值
    hIcon As Long   '要添加、编辑或删除的图标句柄
    szTip As String * 64    '图标的工具提示文本
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Const NOTIFYICON_VERSION = 3       'V5 style taskbar
Private Const NOTIFYICON_OLDVERSION = 0    'Win95 style taskbar

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const SW_HIDE = 0
Private Const SW_MINIMIZE = 6
Private Const SW_MAXIMIZE = 3
Private Const SW_RESTORE = 9
Private Const SW_SHOW = 5

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206

Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4

Public Event MouseMove()
Public Event MouseUp(Button As Integer)    'Button=1,2,4 分别表示鼠标左键、右键、中间键
Public Event MouseDown(Button As Integer)    'Button=1,2,4 分别表示鼠标左键、右键、中间键
Public Event DblClick(Button As Integer)    'Button=1,2,4 分别表示鼠标左键、右键、中间键

Private lHwnd As Long    '目标窗口句柄
Private lIconHandle As Long, lToolTipText As String    '图标句柄和托盘 ToolTip
Attribute lToolTipText.VB_VarUserMemId = 1073938433
Private lIsSysTraySet As Boolean    '本对象现在是否已经管理着一个系统托盘图标
Attribute lIsSysTraySet.VB_VarUserMemId = 1073938435

Private mNoti As NOTIFYICONDATA    'Shell_NotifyIcon 参数使用的数据
Attribute mNoti.VB_VarUserMemId = 1073938436


Function AddSysTray(ByVal hWndForm As Long, _
                    Optional ByVal hIcon As Long = 0, _
                    Optional ByRef sToolTip As String = vbNullChar, _
                    Optional ByVal sTrayTitle As String = "CQAppFramework - CI", _
                    Optional ByVal sTrayText As String = "Welcome!", _
                    Optional ByVal bThenHideForm As Boolean = True) As Boolean
'添加一个系统托盘
'hWndForm 为目标窗体句柄
'hIcon、sToolTip 可缺省，缺省时使用之前通过 IconHandle 属性 _
  '或 SetIcon 方法设置的图标、使用 ToolTipText 属性设置的 ToolTip _
  '如不缺省自动改变相应属性的值为参数的值
'bThenHideForm 表示设置系统托盘后，是否将 hWndForm 窗体最小化和隐藏

'一个类模块对象只能设置一个窗体为系统托盘
'如果现在已经管理着一个窗体的系统托盘，先将它删除
    If lHwnd Then RemoveSysTray

    '默认返回值
    AddSysTray = False

    '若未设置图标，无法建立托盘
    If lIconHandle = 0 And hIcon = 0 Then Exit Function

    '======== 添加新托盘，向 mdlSysTraySupport 登记 ========
    If Not STRegOneObject(hWndForm, ObjPtr(Me)) Then Exit Function

    '======== 子类化窗口 hWndForm，自定义窗口程序为 _
     'mdlSysTraySupport 模块的 STWndProc 函数  ========
    If Not SCCreateSubClass(hWndForm, AddressOf STWndProc) Then Exit Function


    '======== 数据准备 ========
    lHwnd = hWndForm
    If hIcon Then lIconHandle = hIcon
    If sToolTip <> vbNullChar Then lToolTipText = sToolTip

    '======== 创建托盘 ========
    With mNoti
        .cbSize = LenB(mNoti)
        .uID = 0
        .hIcon = lIconHandle
        .hwnd = lHwnd
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = ST_NOTI_MSG  '为 mdlSysTraySupport 中定义的公有常量
        .szTip = lToolTipText & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfoTitle = sTrayTitle
        .szInfo = sTrayText
        .dwInfoFlags = NIIF_INFO    '气泡的图标
        .uTimeout = 10000    '气泡消失时间
    End With

    If Shell_NotifyIcon(NIM_ADD, mNoti) Then
        If bThenHideForm Then
            '最小化和隐藏窗口
            ShowWindow lHwnd, SW_MINIMIZE
            DoEvents
            ShowWindow lHwnd, SW_HIDE
        End If
        lIsSysTraySet = True
        '返回值
        AddSysTray = True
    End If
End Function

Public Function ShowText(ByVal paramTitle As String, ByVal paramText As String)      '显示托盘气泡
    With mNoti
        .szInfoTitle = paramTitle
        .szInfo = paramText
        .dwInfoFlags = NIIF_GUID
    End With
    Shell_NotifyIcon NIM_MODIFY, mNoti
End Function

Public Function RemoveSysTray(Optional ByVal bThenShowForm As Boolean = True) As Boolean
'删除本对象维护的系统托盘
'bThenShowForm：是否之后恢复窗体显示

    If lIsSysTraySet Then
        '取消在 mdlSysTraySupport 中的登记
        STUnRegOneObject lHwnd
        '恢复子类处理
        SCRestoreSubClassOne lHwnd, AddressOf STWndProc
        '删除托盘
        With mNoti
            .uFlags = 0
        End With
        Shell_NotifyIcon NIM_DELETE, mNoti

        '恢复窗口显示
        If bThenShowForm Then ShowWindow lHwnd, SW_SHOW

        '属性值变量清0
        lHwnd = 0
        lIsSysTraySet = False
    End If
End Function

Public Function EventsGen(ByVal wParam As Long, ByVal lParam As Long) As Long
'事件生成器
'由 mdlSysTraySupport 模块的自定义窗口程序截获消息后调用，用于本对象生成事件
'lParam 参数表示鼠标消息
    Select Case lParam
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(1)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(2)
    Case WM_MBUTTONUP
        RaiseEvent MouseUp(4)

    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(1)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(2)
    Case WM_MBUTTONDOWN
        RaiseEvent MouseDown(4)

    Case WM_LBUTTONDBLCLK
        RaiseEvent DblClick(1)
    Case WM_RBUTTONDBLCLK
        RaiseEvent DblClick(2)
    Case WM_MBUTTONDBLCLK
        RaiseEvent DblClick(4)

    Case WM_MOUSEMOVE
        RaiseEvent MouseMove
    End Select

    EventsGen = gc_lngEventsGenDefautRet
End Function

Public Sub SetIcon(picIcon As Picture)
'以图片方式更新图标（图片必须是图标格式的）
    If picIcon.Type <> vbPicTypeIcon Then Exit Sub

    If lIsSysTraySet Then
        '////////// 已经设置了系统托盘，更新图标 //////////
        With mNoti
            .hIcon = picIcon.Handle
            .uFlags = NIF_ICON
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '设置新图标成功时再改变 lIconHandle 的值
            lIconHandle = picIcon.Handle
        End If
    Else
        '////////// 尚未设置系统托盘仅设置 lIconHandle 的值 //////////
        lIconHandle = picIcon.Handle
    End If
End Sub


Public Sub ShowForm()
'显示窗体并将窗体置于前台
    ShowWindow lHwnd, SW_SHOW
    ShowWindow lHwnd, SW_RESTORE
    SetForegroundWindow lHwnd
End Sub

Public Sub ShowMenuBeforePatch()
'在弹出快捷菜单之前，需先调用本过程，再 Form.PopupMenu
'以便弹出菜单后单击其他位置菜单能自动消失
    SetForegroundWindow lHwnd
End Sub


Private Sub Class_Initialize()
    lIsSysTraySet = False
End Sub

Private Sub Class_Terminate()
    RemoveSysTray False
End Sub






'返回目标窗体句柄（只读属性）
Public Property Get hwnd() As Long
    hwnd = lHwnd
End Property


'以图标句柄方式返回或更新托盘图标，也可用 SetIcon 以图片方式更新图标
Public Property Get IconHandle() As Long
    IconHandle = lIconHandle
End Property

Public Property Let IconHandle(ByVal vNewValue As Long)
    If lIsSysTraySet Then
        '////////// 已经设置了系统托盘，更新图标 //////////
        With mNoti
            .hIcon = vNewValue
            .uFlags = NIF_ICON
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '设置新图标成功时再改变 lIconHandle 的值
            lIconHandle = vNewValue
        End If
    Else
        '////////// 尚未设置系统托盘仅设置 lIconHandle 的值 //////////
        lIconHandle = vNewValue
    End If
End Property

'返回或设置托盘工具提示文本 ToolTip
Public Property Get ToolTipText() As String
    ToolTipText = lToolTipText
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
    If lIsSysTraySet Then
        '////////// 已经设置了系统托盘，更新 ToolTip //////////
        With mNoti
            .szTip = vNewValue & vbNullChar
            .uFlags = NIF_TIP
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '设置新 ToolTip 成功时再改变 lToolTipText 的值
            lToolTipText = vNewValue
        End If
    Else
        '////////// 尚未设置系统托盘仅设置 lToolTipText 的值 //////////
        lToolTipText = vNewValue
    End If
End Property



'本对象现在是否已经管理着一个系统托盘（只读属性）
Public Property Get IsSysTraySet() As Boolean
    IsSysTraySet = lIsSysTraySet
End Property
