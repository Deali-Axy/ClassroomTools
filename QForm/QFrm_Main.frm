VERSION 5.00
Begin VB.Form QFrm_Main 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "QFrm_Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame Frame_Other 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   7080
      TabIndex        =   42
      Top             =   720
      Width           =   2175
      Begin VB.CommandButton Btn_Other_�����ʾ 
         Caption         =   "�����ޱ���"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   960
         Width           =   1755
      End
      Begin VB.CommandButton Btn_Other_�����޶��� 
         Caption         =   "�����޶���"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Timer Tmr_Announcement 
      Interval        =   1
      Left            =   8880
      Top             =   120
   End
   Begin VB.Frame Frame_Announcement 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�༶����"
      ForeColor       =   &H00000000&
      Height          =   4095
      Left            =   2640
      TabIndex        =   38
      Top             =   2280
      Width           =   6615
      Begin VB.CommandButton Btn_Announcement_OK 
         Caption         =   "ȷ��"
         Height          =   375
         Left            =   5760
         TabIndex        =   41
         Top             =   3600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Txt_Announcement 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3840
         MultiLine       =   -1  'True
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl_Announcement 
         BackStyle       =   0  'Transparent
         Caption         =   "���������ӹ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3495
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   6375
      End
   End
   Begin VB.Frame Frame_Tools 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ù���"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   60
      TabIndex        =   27
      Top             =   3240
      Width           =   2535
      Begin VB.CommandButton Btn_CUseTool_�������� 
         Caption         =   "dlt��������"
         Height          =   435
         Left            =   240
         TabIndex        =   33
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CUseTool_dltPaint��ͼ 
         Caption         =   "dltPaint��ͼ"
         Height          =   435
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CUseTool_Fonroid��ȫ�� 
         Caption         =   "Fonroid��ȫ��"
         Height          =   435
         Left            =   240
         TabIndex        =   31
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CUseTool_HideTaskbar 
         Caption         =   "ϵͳ������"
         Height          =   435
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl_AppCenter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFCC00&
         Caption         =   "[App Center]"
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   2640
         Width           =   2055
      End
   End
   Begin VB.Frame Frame_Game 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��Ϸ"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   9360
      TabIndex        =   20
      Top             =   660
      Width           =   2115
      Begin VB.CommandButton Btn_Game_������ 
         Caption         =   "������"
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         TabIndex        =   30
         Top             =   2400
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton Btn_Game_�������� 
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         TabIndex        =   29
         Top             =   1920
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton Btn_Game_������ 
         Caption         =   "������"
         Height          =   375
         Left            =   180
         TabIndex        =   26
         Top             =   960
         Width           =   1755
      End
      Begin VB.CommandButton Btn_Game_2048 
         Caption         =   "2048"
         Height          =   375
         Left            =   180
         TabIndex        =   22
         Top             =   480
         Width           =   1755
      End
      Begin VB.CommandButton Btn_Game_FlappyBird 
         Caption         =   "FlappyBird"
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lbl_GameCenter 
         Alignment       =   2  'Center
         BackColor       =   &H00FFCC00&
         Caption         =   "GameCenter"
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   180
         TabIndex        =   35
         Top             =   1440
         Width           =   1755
      End
   End
   Begin VB.TextBox Txt_CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6360
      TabIndex        =   1
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���׽��̹���"
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   2640
      TabIndex        =   11
      Top             =   720
      Width           =   4335
      Begin VB.CommandButton Btn_Create 
         Caption         =   "����"
         Height          =   435
         Left            =   300
         TabIndex        =   18
         Top             =   900
         Width           =   735
      End
      Begin VB.CommandButton Btn_TerminateProcess 
         Caption         =   "����"
         Height          =   435
         Left            =   3480
         TabIndex        =   17
         Top             =   900
         Width           =   735
      End
      Begin VB.CommandButton Btn_ResumeProcess 
         Caption         =   "�ָ�"
         Height          =   435
         Left            =   2640
         TabIndex        =   16
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton Btn_SuspendProcess 
         Caption         =   "����"
         Height          =   435
         Left            =   1860
         TabIndex        =   15
         Top             =   900
         Width           =   735
      End
      Begin VB.CommandButton Btn_ProcessTest 
         Caption         =   "���"
         Height          =   435
         Left            =   1080
         TabIndex        =   14
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox Txt_ProcessName 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   1320
         TabIndex        =   12
         Top             =   420
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1140
      End
   End
   Begin VB.Frame Frame_System 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ϵͳ"
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   9360
      TabIndex        =   7
      Top             =   3840
      Width           =   2115
      Begin VB.CommandButton Btn_Exit 
         Caption         =   "�˳�"
         Height          =   375
         Left            =   300
         TabIndex        =   45
         Top             =   2520
         Width           =   1515
      End
      Begin VB.CommandButton Btn_Sys_Cmd 
         Caption         =   "������"
         Height          =   375
         Left            =   300
         TabIndex        =   25
         Top             =   1260
         Width           =   1515
      End
      Begin VB.CommandButton Btn_Sys_Gpedit 
         Caption         =   "�����"
         Height          =   375
         Left            =   300
         TabIndex        =   24
         Top             =   840
         Width           =   1515
      End
      Begin VB.CommandButton Btn_Sys_Regedit 
         Caption         =   "ע���"
         Height          =   375
         Left            =   300
         TabIndex        =   23
         Top             =   420
         Width           =   1515
      End
      Begin VB.CommandButton Btn_Reboot 
         Caption         =   "����"
         Height          =   375
         Left            =   300
         TabIndex        =   9
         Top             =   2100
         Width           =   1515
      End
      Begin VB.CommandButton Btn_Shutdown 
         Caption         =   "�ػ�"
         Height          =   375
         Left            =   300
         TabIndex        =   8
         Top             =   1680
         Width           =   1515
      End
   End
   Begin VB.Frame Frame_InterTools 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ù���"
      ForeColor       =   &H00000000&
      Height          =   2475
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   2535
      Begin VB.CommandButton Btn_CommonTools 
         Caption         =   "װ�ƹ���"
         Height          =   435
         Index           =   4
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CommonTools 
         Caption         =   "������"
         Height          =   435
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CommonTools 
         Caption         =   "��������"
         Height          =   435
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CommonTools 
         Caption         =   "��������"
         Height          =   435
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton Btn_CommonTools 
         BackColor       =   &H00404040&
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
   End
   Begin VB.Timer Tmr_ShowTime 
      Interval        =   1
      Left            =   9360
      Top             =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DA������"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   6480
      Width           =   1260
   End
   Begin VB.Label lbl_Trip 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   45
      TabIndex        =   10
      Top             =   6480
      Width           =   4935
   End
   Begin VB.Label lbl_Time 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   9840
      TabIndex        =   2
      Top             =   60
      Width           =   1650
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lbl_DesTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˫��������Ը�������Ŷ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "QFrm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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

Private Const NOTIFYICON_VERSION = 3        'V5 style taskbar
Private Const NOTIFYICON_OLDVERSION = 0        'Win95 style taskbar

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

Const MAINFORMKEY = vbKeyF6

Private myData As NOTIFYICONDATA        '��������ͼ������
Dim WithEvents cSHK As clsSysHotKey
Attribute cSHK.VB_VarHelpID = -1
'Dim WithEvents cSysTray As clsSysTray

Dim VerificationPassString As String
Dim strAnnouncement As String

Private Sub Btn_About_Click()
    QFrm_About.Show
End Sub

Private Sub Btn_Announcement_OK_Click()
    strAnnouncement = Txt_Announcement.Text
    lbl_Announcement.Caption = strAnnouncement
    mSaveConfig
    Txt_Announcement.Visible = False
    Btn_Announcement_OK.Visible = False
    ShowTrayText "", "�༶�����Ѹ��£�"
    QFrm_Tray.ShowTray "�༶����", strAnnouncement
End Sub

Private Sub Btn_CommonTools_Click(index As Integer)
    Select Case index
        Case 1
            Shell "tskill explorer"
        Case 2
            SystemClean
        Case 3
            Frm_Random.Show
        Case 4
            PreventB
    End Select
End Sub

Private Sub Btn_Create_Click()
    If Len(Txt_ProcessName) > 0 Then
        Dim TmpPid As Long
        TmpPid = Shell(Txt_ProcessName)
        If TmpPid <> 0 Then
            mShowMsg "�����ɹ���PID=" & TmpPid
        Else
            mShowMsg "����ʧ�ܣ�"
        End If
    Else
        mShowMsg "�㻹û����- -"
    End If
End Sub

Private Sub Btn_CUseTool_dltPaint��ͼ_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\dltPaint.exe", "", "", 5
End Sub

Private Sub Btn_CUseTool_Fonroid��ȫ��_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\FonroidOS Lock.exe", "", "", 5
End Sub

Private Sub Btn_CUseTool_HideTaskbar_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\HideTaskbar.exe", "", "", 5
End Sub

Private Sub Btn_CUseTool_��������_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\��������.exe", "", "", 5
End Sub

Private Sub Btn_Exit_Click()
    Mod_HookSkinner.Detach Me.hwnd
    End
End Sub

Private Sub Btn_Game_2048_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\Game\2048\2048 V3.exe", "", "", 5
End Sub

Private Sub Btn_Game_FlappyBird_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\Game\FlappyBird\FlappyBird.exe", "", "", 5
End Sub

Private Sub Btn_Game_������_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\Game\������\������.exe", "", "", 5
End Sub

Private Sub Btn_Game_������_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\Game\������\������2.exe", "", "", 5
End Sub

Private Sub Btn_Game_��������_Click()
    ShellExecute Me.hwnd, "open", App.Path & "\Bin\Game\��������\��������.exe", "", "", 5
End Sub

Private Sub Btn_Other_�����ʾ_Click()
    Call Frm_Menu.MenuFunc_ͼƬ����Ƽ�
End Sub

Private Sub Btn_Other_�����޶���_Click()
    Call Frm_Menu.MenuFunc_��ʾ��������¼
End Sub

Private Sub Btn_ProcessTest_Click()
    If GetPID(Txt_ProcessName) = 0 Then
        mShowMsg "���̲����ڣ�"
    Else
        mShowMsg "����������..."
    End If
End Sub

Private Sub Btn_Reboot_Click()
    Shell "shutdown -r -t 0"
End Sub

Private Sub Btn_ResumeProcess_Click()
    If GetPID(Txt_ProcessName) = 0 Then
        mShowMsg "�Ҳ����ý��̣�"
    Else
        ResumeProcess Txt_ProcessName
        mShowMsg "�ѻָ��ý���"
    End If
End Sub

Private Sub Btn_Shutdown_Click()
    Shell "shutdown -s -t 0"
End Sub

Private Sub Btn_SuspendProcess_Click()
    If GetPID(Txt_ProcessName) = 0 Then
        mShowMsg "�Ҳ����ý��̣�"
    Else
        SuspendProcess Txt_ProcessName
        mShowMsg "�Ѷ���ý��̣�������ָ�����ť�ָ��ý���"
    End If
End Sub

Private Sub Btn_Sys_Cmd_Click()
    ShellExecute Me.hwnd, "open", "cmd", "", "", 5
End Sub

Private Sub Btn_Sys_Gpedit_Click()
    ShellExecute Me.hwnd, "open", "gpedit.msc", "", "", 5
End Sub

Private Sub Btn_Sys_Regedit_Click()
    ShellExecute Me.hwnd, "open", "regedit", "", "", 5
End Sub

Private Sub Btn_TerminateProcess_Click()
    If GetPID(Txt_ProcessName) = 0 Then
        mShowMsg "�Ҳ����ý��̣�"
    Else
        TerminateProcessEx Txt_ProcessName
        mShowMsg "�ѽ����ý���"
    End If
End Sub

Private Sub cSHK_SysKeyPressed()
    If Me.Visible Then
        Me.Hide
    Else
        Me.Show
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo Err
    QDB.Log Me.Name & " Activate"
    Exit Sub
Err:
    QDB.Runtime_Error Me.Name & "_Activate", Err.Description, Err.Number
    Resume Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            QFrm_About.Show
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo load_err
    Me.Caption = QApp_Title
    Set cSHK = New clsSysHotKey        '����SysHookKey����
    'Set cSysTray = New clsSysTray '����SysTray����
    AddSysTray        '�������
    cSHK.SetASysHotKey Me.hwnd, MAINFORMKEY, 0, False        '������ʾ�������ȼ�
    'cSysTray.AddSysTray Me.hwnd, Me.Icon, QApp.Title, "", "��ӭʹ��ClassTools���ҹ�����", False '�������
    QDB.Log Me.Name & " Load hWnd=" & Me.hwnd
    lbl_DesTitle = lbl_DesTitle & " " & QApp_Version
    mLoadConfig
    ShowTrayText "", "��ӭʹ��DA���ҹ���ϵͳ��"
    'mAutoStart True
    'QApp.QEverydayTips.GraphicsInit    '��ʼ��ͼƬ���� ����
    VerificationPassString = Chr(100) & Chr(101) & Chr(97) & Chr(108) & Chr(105) & Chr(97) & Chr(120) & Chr(121)
    mShowMsg "˫�����Ͻǵı�������޸ı�������Ŷ."
    'Frm_Tray.ShowTray "Classroom Tools!", "��ӭʹ��DA���ҹ���ϵͳ��" & vbCrLf & "���չ��棺" & strAnnouncement
    'QApp.QEverydayTips.ShowText 5    'Ψ����ʫ��
    'QApp.QEverydayTips.InitFile App.Path & "\_Output\��̾�����¼.txt", App.Path & "\fujiao.qdata"
    Exit Sub
load_err:
    QDB.Runtime_Error Me.Name & "_Load", Err.Description, Err.Number
    'Debug.Print Err.Description
    Resume Next
End Sub

Private Sub Form_Terminate()
    mSaveConfig
    Set cSHK = Nothing
    DeleteTray
    Shell "cmd.exe /c del " & App.Path & "\*.bat"
    'Set cSysTray = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not fExit Then
        QDB.Log Me.Name & " Unload"
        Me.Hide
        Cancel = 1
    Else
        QDB.Log Me.Name & " Unload"
        Set cSHK = Nothing
    End If
    mSaveConfig
End Sub

Private Sub lbl_Announcement_Click()
    With Txt_Announcement
        .Top = lbl_Announcement.Top
        .Left = lbl_Announcement.Left
        .Width = lbl_Announcement.Width
        .Height = lbl_Announcement.Height
        .Visible = True
        .Text = strAnnouncement
    End With
    Btn_Announcement_OK.Visible = True

End Sub

Private Sub lbl_AppCenter_Click()
    Frm_Verification.Show
    Frm_Verification.mInfo Frm_AppCenter
End Sub

Private Sub lbl_DesTitle_DblClick()
    Dim TmpStr As String
    TmpStr = InputBox("�������ʺ�����", "���ҹ����� By DA")
    If Len(TmpStr) > 0 Then
        lbl_DesTitle = TmpStr
    End If
End Sub

Private Sub lbl_DesTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 2
            PopupMenu Frm_Menu.Menu_FrmMain_������˵�
    End Select
End Sub

Private Sub lbl_GameCenter_Click()
    Frm_Verification.Show
    Frm_Verification.mInfo Frm_GameCenter
End Sub

Private Sub Tmr_Announcement_Timer()
    If Minute(Time) Mod 5 = 0 And Second(Time) = 0 Then
        QFrm_Tray.ShowTray "�༶����", strAnnouncement
    End If
End Sub

Private Sub Tmr_ShowTime_Timer()
    lbl_Time = Str(Time)
    'If Minute(Time) = 0 And Second(Time) = 0 Then
        'ShowTrayText "ClassroomTools ���㱨ʱ~", "������" & Date & " " & Time & vbCrLf & "By Deali-Axy"
    'End If

    If Minute(Time) Mod 5 = 0 And Second(Time) = 1 Then
        'Call Frm_Menu.MenuFunc_��Լ����Ƽ�
    End If
End Sub

Sub mLoadConfig()
    Dim TmpStr As String
    TmpStr = GetSetting(QApp_Name, Me.Name, "DesTitle")
    If Len(TmpStr) > 0 Then
        lbl_DesTitle.Caption = ""
        lbl_DesTitle.Caption = TmpStr
    End If
    '��ȡ�༶���棡
    strAnnouncement = GetSetting(QApp.Name, "Class", "Announcement", "�������������ù��棡" & vbCrLf & "By Deali-Axy")
    lbl_Announcement.Caption = strAnnouncement
End Sub

Sub mSaveConfig()
    SaveSetting QApp_Name, Me.Name, "DesTitle", lbl_DesTitle.Caption
    SaveSetting QApp.Name, "Class", "Announcement", strAnnouncement
End Sub

Sub SystemClean()
    Dim a As String
    a = """"
    Open App.Path & "\Tmp_SystemClean.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "title " & QApp_Title
    Print #1, "echo exit|%ComSpec% /k prompt e 100 B4 00 B0 12 CD 10 B0 03 CD 10 CD 20 $_g$_q$_|debug>nul"
    Print #1, "chcp 437 >nul"
    Print #1, "graftabl 936 >nul"
    Print #1, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #1, "echo ���ڰ������ϵͳ�����ļ������Ե�......"
    Print #1, "del /f /s /q %systemdrive%\*.tmp"
    Print #1, "del /f /s /q %systemdrive%\*._mp"
    Print #1, "del /f /s /q %systemdrive%\*.log"
    Print #1, "del /f /s /q %systemdrive%\*.gid"
    Print #1, "del /f /s /q %systemdrive%\*.chk"
    Print #1, "del /f /s /q %systemdrive%\*.old"
    Print #1, "del /f /s /q %systemdrive%\recycled\*.*"
    Print #1, "del /f /s /q %windir%\*.bak"
    Print #1, "del /f /s /q %windir%\prefetch\*.*"
    Print #1, "rd /s /q %windir%\temp & md %windir%\temp"
    Print #1, "del /f /q %userprofile%\cookies\*.*"
    Print #1, "del /f /q %userprofile%\recent\*.*"
    Print #1, "del /f /s /q " + a + "%userprofile%\Local Settings\Temporary Internet Files\*.*" + a
    Print #1, "del /f /s /q " + a + "%userprofile%\Local Settings\Temp\*.*" + a
    Print #1, "del /f /s /q %userprofile%\recent\*.*"
    'Print #1, "echo ϵͳ������� ��������˳�����"
    'Print #1, "pause>nul"
    Print #1, "del %0"
    Print #1, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Close #1
    Shell App.Path & "\Tmp_SystemClean.bat"
End Sub

Sub PreventB()
    Dim a As String
    a = """"
    Open App.Path & "\Tmp_PreventB.bat" For Output As #1
    Print #1, "@echo off"
    Print #1, "title " & QApp_Title
    Print #1, "echo exit|%ComSpec% /k prompt e 100 B4 00 B0 12 CD 10 B0 03 CD 10 CD 20 $_g$_q$_|debug>nul"
    Print #1, "chcp 437 >nul"
    Print #1, "graftabl 936 >nul"
    Print #1, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Print #1, "color 0a"
    Print #1, ":start"
    Print #1, "set num=0"
    Print #1, "set ""echos= "" "
    Print #1, ":num"
    Print #1, "set /a a1=%random%%%3"
    Print #1, "if ""%a1%"" == ""1"" set ""a1= "" "
    Print #1, "if ""%a1%"" == ""2"" set ""a1= "" "
    Print #1, "if ""%a1%"" == ""0"" set /a a1=%random%%%2 "
    Print #1, "set echos=%echos%%a1% "
    Print #1, "set /a num=%num%+1"
    Print #1, "if ""%num%"" == ""75"" echo %echos%&&goto :start "
    Print #1, "goto :num"
    'Print #1, "echo ϵͳ������� ��������˳�����"
    'Print #1, "pause>nul"
    Print #1, "del %0"
    Print #1, ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::"
    Close #1

    Open App.Path & "\Tmp_Launch.bat" For Output As #1
    Print #1, "call %1"
    Print #1, "del %1"
    Print #1, "del %0"
    Close #1
    'Shell App.Path & "\Tmp_Launch.bat " & App.Path & "\Tmp_PreventB.bat"
    Shell App.Path & "\Tmp_PreventB.bat"
End Sub

Sub TmpClear(FileName As String)
    Open App.Path & "\Tmp_Clear.bat" For Output As #1
    Print #1, ":loop"
    Print #1, "if not exist %1 goto exit"
    Print #1, "del %1"
    Print #1, "goto loop"
    Print #1, ":exit"
    Print #1, "del %0"
    Close #1
    Shell App.Path & "\Tmp_Clear.bat " & FileName
End Sub

Sub mAutoStart(pAutoStart As Boolean)
    On Error GoTo Err
    Dim tmp As Object
    Set tmp = CreateObject("WScript.Shell")
    tmp.regDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\DAClassroom"
    If pAutoStart Then
        pAutoStart = False
        tmp.regDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\DAClassroom"
    Else
        pAutoStart = True
        tmp.regWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\DAClassroom", App.Path & "\" & App.EXEName & ".exe boot", "REG_SZ"
    End If
    Exit Sub
Err:
    Resume Next
End Sub

Sub mVerification(paramVerificationInfo As String, Form As Form)
    If paramVerificationInfo = VerificationPassString Then
        Form.Show
    Else
        mShowMsg "Verification Failed."
    End If
End Sub

Sub mShowMsg(pMsg As String)
    lbl_Trip = pMsg
End Sub

Private Sub Txt_CMD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            Dim RetVal As Boolean
            If Len(Txt_CMD) > 0 Then RetVal = CommandProc(Txt_CMD)
            If RetVal Then
                Txt_CMD = ""
                lbl_Trip = "����ִ����� ^ ^"
            Else
                lbl_Trip = "����ִ��ʧ�� T T.."
            End If
            If IsChinese(Txt_CMD.Text) Then
                ShellExecute Me.hwnd, "open", "http://www.baidu.com/s?wd=" & Txt_CMD, "", "", 5
                Txt_CMD = ""
                lbl_Trip = "����ִ����� ^ ^"
            End If
    End Select
End Sub

Private Sub AddSysTray()
    OrgWinRet = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    With myData
        .cbSize = Len(myData)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = TRAY_CALLBACK        '����ͼ�귢���¼�ʱ����������Ϣ��
        .hIcon = Me.Icon        'ͼ�ꡣ����ΪStdPicture�����Կ�������Ϊpicturebox�е�ͼƬ
        .szTip = "��ӭʹ��" & QApp.Title & vbNullChar        'tooltip����
        .dwState = 0
        .dwStateMask = 0
        .szInfoTitle = "��ӭʹ��" & QApp.Title & vbNullChar        '������ʾ����
        .szInfo = "������" & Date & vbCrLf & "������ͼ��(���߰�F10)����ʾ������" & vbNullChar        '������ʾ����
        .dwInfoFlags = NIIF_INFO        '���ݵ�ͼ��
        .uTimeout = 1        '������ʧʱ��
    End With
    Shell_NotifyIcon NIM_ADD, myData
    glWinRet = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Sub

Private Sub ShowTrayText(sTitle As String, sText As String)
    With myData
        .szInfoTitle = sTitle & vbNullChar
        .szInfo = sText & vbNullChar
        .dwInfoFlags = NIIF_GUID
        .uTimeout = 1
    End With
    Shell_NotifyIcon NIM_MODIFY, myData
End Sub

Private Sub DeleteTray()
    Shell_NotifyIcon NIM_DELETE, myData
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OrgWinRet)
End Sub
