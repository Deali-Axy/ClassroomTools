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

'--------------------ϵͳ���� ��------------------------
'��Ҫ mdlSysTraySupport ��׼ģ���֧��
'��Ҫ clsHashLK��clsSubClass ��ģ�� �� mdlSubClass ��׼ģ���֧��

'# ���ʻ�������ʾ��Ϣ�ַ�������

'================================================================================
'һ���������ֻ��ʹ��һ����������������
'================================================================================

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Type NOTIFYICONDATA    'Shell_NotifyIcon ����ʹ�õ�����
    cbSize As Long  '���������ݵ��ֽڴ�С
    hwnd As Long    'Ҫ����������Ϣ�Ĵ��ڵľ��
    uID As Long    'Ӧ�ó����Զ��������ͼ��ID������Ϊ0��
    uFlags As Long  'һ�����룬��ʾ��������Щ��Ա��ֵ��Ч
    uCallbackMessage As Long    'Ӧ�ó����Զ������Ϣ���ֵ
    hIcon As Long   'Ҫ���ӡ��༭��ɾ����ͼ����
    szTip As String * 64    'ͼ��Ĺ�����ʾ�ı�
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
Public Event MouseUp(Button As Integer)    'Button=1,2,4 �ֱ��ʾ���������Ҽ����м��
Public Event MouseDown(Button As Integer)    'Button=1,2,4 �ֱ��ʾ���������Ҽ����м��
Public Event DblClick(Button As Integer)    'Button=1,2,4 �ֱ��ʾ���������Ҽ����м��

Private lHwnd As Long    'Ŀ�괰�ھ��
Private lIconHandle As Long, lToolTipText As String    'ͼ���������� ToolTip
Attribute lToolTipText.VB_VarUserMemId = 1073938433
Private lIsSysTraySet As Boolean    '�����������Ƿ��Ѿ�������һ��ϵͳ����ͼ��
Attribute lIsSysTraySet.VB_VarUserMemId = 1073938435

Private mNoti As NOTIFYICONDATA    'Shell_NotifyIcon ����ʹ�õ�����
Attribute mNoti.VB_VarUserMemId = 1073938436


Function AddSysTray(ByVal hWndForm As Long, _
                    Optional ByVal hIcon As Long = 0, _
                    Optional ByRef sToolTip As String = vbNullChar, _
                    Optional ByVal sTrayTitle As String = "CQAppFramework - CI", _
                    Optional ByVal sTrayText As String = "Welcome!", _
                    Optional ByVal bThenHideForm As Boolean = True) As Boolean
'����һ��ϵͳ����
'hWndForm ΪĿ�괰����
'hIcon��sToolTip ��ȱʡ��ȱʡʱʹ��֮ǰͨ�� IconHandle ���� _
  '�� SetIcon �������õ�ͼ�ꡢʹ�� ToolTipText �������õ� ToolTip _
  '�粻ȱʡ�Զ��ı���Ӧ���Ե�ֵΪ������ֵ
'bThenHideForm ��ʾ����ϵͳ���̺��Ƿ� hWndForm ������С��������

'һ����ģ�����ֻ������һ������Ϊϵͳ����
'��������Ѿ�������һ�������ϵͳ���̣��Ƚ���ɾ��
    If lHwnd Then RemoveSysTray

    'Ĭ�Ϸ���ֵ
    AddSysTray = False

    '��δ����ͼ�꣬�޷���������
    If lIconHandle = 0 And hIcon = 0 Then Exit Function

    '======== ���������̣��� mdlSysTraySupport �Ǽ� ========
    If Not STRegOneObject(hWndForm, ObjPtr(Me)) Then Exit Function

    '======== ���໯���� hWndForm���Զ��崰�ڳ���Ϊ _
     'mdlSysTraySupport ģ��� STWndProc ����  ========
    If Not SCCreateSubClass(hWndForm, AddressOf STWndProc) Then Exit Function


    '======== ����׼�� ========
    lHwnd = hWndForm
    If hIcon Then lIconHandle = hIcon
    If sToolTip <> vbNullChar Then lToolTipText = sToolTip

    '======== �������� ========
    With mNoti
        .cbSize = LenB(mNoti)
        .uID = 0
        .hIcon = lIconHandle
        .hwnd = lHwnd
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = ST_NOTI_MSG  'Ϊ mdlSysTraySupport �ж���Ĺ��г���
        .szTip = lToolTipText & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfoTitle = sTrayTitle
        .szInfo = sTrayText
        .dwInfoFlags = NIIF_INFO    '���ݵ�ͼ��
        .uTimeout = 10000    '������ʧʱ��
    End With

    If Shell_NotifyIcon(NIM_ADD, mNoti) Then
        If bThenHideForm Then
            '��С�������ش���
            ShowWindow lHwnd, SW_MINIMIZE
            DoEvents
            ShowWindow lHwnd, SW_HIDE
        End If
        lIsSysTraySet = True
        '����ֵ
        AddSysTray = True
    End If
End Function

Public Function ShowText(ByVal paramTitle As String, ByVal paramText As String)      '��ʾ��������
    With mNoti
        .szInfoTitle = paramTitle
        .szInfo = paramText
        .dwInfoFlags = NIIF_GUID
    End With
    Shell_NotifyIcon NIM_MODIFY, mNoti
End Function

Public Function RemoveSysTray(Optional ByVal bThenShowForm As Boolean = True) As Boolean
'ɾ��������ά����ϵͳ����
'bThenShowForm���Ƿ�֮��ָ�������ʾ

    If lIsSysTraySet Then
        'ȡ���� mdlSysTraySupport �еĵǼ�
        STUnRegOneObject lHwnd
        '�ָ����ദ��
        SCRestoreSubClassOne lHwnd, AddressOf STWndProc
        'ɾ������
        With mNoti
            .uFlags = 0
        End With
        Shell_NotifyIcon NIM_DELETE, mNoti

        '�ָ�������ʾ
        If bThenShowForm Then ShowWindow lHwnd, SW_SHOW

        '����ֵ������0
        lHwnd = 0
        lIsSysTraySet = False
    End If
End Function

Public Function EventsGen(ByVal wParam As Long, ByVal lParam As Long) As Long
'�¼�������
'�� mdlSysTraySupport ģ����Զ��崰�ڳ���ػ���Ϣ����ã����ڱ����������¼�
'lParam ������ʾ�����Ϣ
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
'��ͼƬ��ʽ����ͼ�꣨ͼƬ������ͼ���ʽ�ģ�
    If picIcon.Type <> vbPicTypeIcon Then Exit Sub

    If lIsSysTraySet Then
        '////////// �Ѿ�������ϵͳ���̣�����ͼ�� //////////
        With mNoti
            .hIcon = picIcon.Handle
            .uFlags = NIF_ICON
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '������ͼ��ɹ�ʱ�ٸı� lIconHandle ��ֵ
            lIconHandle = picIcon.Handle
        End If
    Else
        '////////// ��δ����ϵͳ���̽����� lIconHandle ��ֵ //////////
        lIconHandle = picIcon.Handle
    End If
End Sub


Public Sub ShowForm()
'��ʾ���岢����������ǰ̨
    ShowWindow lHwnd, SW_SHOW
    ShowWindow lHwnd, SW_RESTORE
    SetForegroundWindow lHwnd
End Sub

Public Sub ShowMenuBeforePatch()
'�ڵ�����ݲ˵�֮ǰ�����ȵ��ñ����̣��� Form.PopupMenu
'�Ա㵯���˵��󵥻�����λ�ò˵����Զ���ʧ
    SetForegroundWindow lHwnd
End Sub


Private Sub Class_Initialize()
    lIsSysTraySet = False
End Sub

Private Sub Class_Terminate()
    RemoveSysTray False
End Sub






'����Ŀ�괰������ֻ�����ԣ�
Public Property Get hwnd() As Long
    hwnd = lHwnd
End Property


'��ͼ������ʽ���ػ��������ͼ�꣬Ҳ���� SetIcon ��ͼƬ��ʽ����ͼ��
Public Property Get IconHandle() As Long
    IconHandle = lIconHandle
End Property

Public Property Let IconHandle(ByVal vNewValue As Long)
    If lIsSysTraySet Then
        '////////// �Ѿ�������ϵͳ���̣�����ͼ�� //////////
        With mNoti
            .hIcon = vNewValue
            .uFlags = NIF_ICON
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '������ͼ��ɹ�ʱ�ٸı� lIconHandle ��ֵ
            lIconHandle = vNewValue
        End If
    Else
        '////////// ��δ����ϵͳ���̽����� lIconHandle ��ֵ //////////
        lIconHandle = vNewValue
    End If
End Property

'���ػ��������̹�����ʾ�ı� ToolTip
Public Property Get ToolTipText() As String
    ToolTipText = lToolTipText
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
    If lIsSysTraySet Then
        '////////// �Ѿ�������ϵͳ���̣����� ToolTip //////////
        With mNoti
            .szTip = vNewValue & vbNullChar
            .uFlags = NIF_TIP
        End With
        If Shell_NotifyIcon(NIM_MODIFY, mNoti) Then
            '������ ToolTip �ɹ�ʱ�ٸı� lToolTipText ��ֵ
            lToolTipText = vNewValue
        End If
    Else
        '////////// ��δ����ϵͳ���̽����� lToolTipText ��ֵ //////////
        lToolTipText = vNewValue
    End If
End Property



'�����������Ƿ��Ѿ�������һ��ϵͳ���̣�ֻ�����ԣ�
Public Property Get IsSysTraySet() As Boolean
    IsSysTraySet = lIsSysTraySet
End Property