Attribute VB_Name = "QMod_SysTray"
Option Explicit

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
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

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4

Private myData As NOTIFYICONDATA    '��������ͼ������


Function AddTray(ByVal hwnd As Long)
    OrgWinRet = GetWindowLong(hwnd, GWL_WNDPROC)
    With myData
        .cbSize = Len(myData)
        .hwnd = hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = TRAY_CALLBACK    '����ͼ�귢���¼�ʱ����������Ϣ��
        .hIcon = Me.Icon    'ͼ�ꡣ����ΪStdPicture�����Կ�������Ϊpicturebox�е�ͼƬ
        .szTip = "��ӭʹ��" & vbNullChar    'tooltip����
        .dwState = 0
        .dwStateMask = 0
        .szInfoTitle = "��ӭʹ��" & vbNullChar    '������ʾ����
        .szInfo = "������ͼ�꽫��ʾ/����������" & vbNullChar    '������ʾ����
        .dwInfoFlags = NIIF_INFO    '���ݵ�ͼ��
        .uTimeout = 10000    '������ʧʱ��
    End With
    Shell_NotifyIcon NIM_ADD, myData
    glWinRet = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Function

Function ShowText()
    With myData
        .szInfoTitle = "���ԡ�" & vbNullChar
        .szInfo = "��ť������ԡ�" & vbNullChar
        .dwInfoFlags = NIIF_GUID
    End With
    Shell_NotifyIcon NIM_MODIFY, myData
End Function

Function DeleteTray()
    Shell_NotifyIcon NIM_DELETE, myData
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OrgWinRet)
End Function
