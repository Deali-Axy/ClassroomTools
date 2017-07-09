Attribute VB_Name = "QMod_AppConfig"
'[Mod Name]Code Institute Common QApp Framework Config Module
'[Mod Author]Deali-Axy
'----------------------------QApp Public Object-------------------------------
Public QDB As New QClass_QDebug    'ȫ�ֵ�QDebug������
Public QApp As New QClass_QApp
'----------------------------QApp Config-------------------------------
Public Const QApp_Name As String = "ClassroomTools"    '��������
Public Const QApp_Author As String = "DealiAxy"    '��������
Public Const QApp_Author_Website As String = "http://weibo.com/dealiaxy"    '������վ
Public Const QApp_Version As String = "0.5.60 Beta5"    '����汾(�ַ�������)
Public Const QApp_MajorVersion As Integer = 0    '�������汾
Public Const QApp_MinorVersion As Integer = 5    '����ΰ汾
Public Const QApp_ReleaseVersion As Integer = 60    '���������汾
Public Const QApp_Comments As String = "[DA_Dev]���ҹ�����"    '����ע��
Public Const QApp_FileDescription As String = "[DA_Dev]���ҹ�����"    '�ļ�˵��
Public Const QApp_Website As String = "http://weibo.com/dealiaxy"    '�����ҳ
Public Const QApp_LegalCopyright As String = "Copyright @ Deali-Axy"    '���ɰ�Ȩ
Public Const QApp_LegalTrademarks As String = "DealiAxy"    '�����̱�
Public Const QApp_SubTitle = "���ҹ����� " & QApp_Version    '�����ӱ���
Public Const QApp_Title = "DA���ҹ�����-�����"    '����������
Public QApp_Default_ConfigFile As String  'Ĭ�������ļ�·��
Public QApp_Icon_Gif As String    '����ͼ��(����Sub Main������)
'----------------------------CQAF Config-------------------------------
Public Const CQAF_Version = "0.6.0"    'CQAF�汾
'----------------------------QApp Standard Error Config-------------------------------
Public Const ErrNum_SubMain = 1
Public Const ErrNum_FormLoad = 2
Public Const ErrNum_Form = 3
Public Const ErrNum_Other = 1024
'----------------------------QApp Pretreatment-------------------------------
#Const App_Load_Interface = False    'QFrm_Load ����
#Const MLC_HookSkin = True    'ʹ��QHookSkinƤ������

Private Type QApp_Info
    App_Name As String
    App_Authuor As String
    App_Version As String
    App_MajorVersion As Integer
    App_MinorVersion As Integer
    App_ReleaseVersion As Integer
    App_Comments As String
    App_FileDescription As String
    App_Website As String
    App_LegalCopyright As String
    App_LegalTrademarks As String
End Type

Public Type QMsg_Struct
    hwnd As Long
    Date As Date
    Time As Date
    msgType As String
    Head As String
    Body As String
End Type

Public Sub QApp_Main()
    On Error GoTo Main_Err
    QDB.Log "QApp Run! Name=" & QApp.Name
    QDB.Log "QApp ThreadID=" & QApp.ThreadID
    QDB.Log "QApp hInstance=" & QApp.hInstance
    App_Icon_Gif = QApp.Path & "\CI_Icon.gif"

    '����Ĭ�������ļ���·��
    QApp_Default_ConfigFile = QApp.Path & "\Config\default.config"
    
    Load QFrm_Main
    QDB.Log "Load QFrm_main"
    With QFrm_Main
        .Caption = QApp.Title & Space(1) & QApp.Version
        #If MLC_HookSkin Then
            Mod_HookSkinner.Attach .hwnd
            QDB.Log "Load QHookSkin"
        #End If
    End With

    #If App_Load_Interface Then
        Load QFrm_Load
        QDB.Log "Load QFrm_Load"
        With QFrm_Load
            .Caption = QApp.Title & "  ���ڼ���..."
            .Show
            QDB.Log "QFrm_Load.Show"
        End With
    #Else
        If Command = "boot" Then
            QFrm_Main.Hide
        Else
            QFrm_Main.Show
            QDB.Log "QFrm_Main.Show"
        End If
    #End If

    Exit Sub
Main_Err:
    QDB.Runtime_Error "Sub Main", Err.Description, Err.Number
End Sub

