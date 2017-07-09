VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QClass_QDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"QDB"
'[Mod Name]Code Institute Common QApp Framework QDebug Class
'[Mod Author]Deali-Axy
'Private Constant
Private Const cLogFile As Integer = 201
Private Const cRtErrFile As Integer = 202
'��������ֵ�ľֲ�����
Private mvarLogFile As String    '�ֲ�����
Private LogFileName As String
Private RtErrFileName As String
Private DateTime As String
'----------------------------CQAF Config-------------------------------
Private Const cQDB_Version As String = "0.2.0"
'----------------------------QApp Pretreatment-------------------------------
#Const IfWriteToLogFile = False
#Const IfRecordRunTimeError = False    '��������IfWriteToLogFileΪTrue


Private Sub Class_Initialize()
    On Error GoTo Err:
    RtErrFileName = App.Path & "\QDB\qdb_RtErr.log"
    LogFileName = App.Path & "\QDB\qdb.log"    '����Ĭ����־�ļ�·��
    DateTime = Date & "|" & Time
    mvarLogFile = LogFileName
    #If IfWriteToLogFile Then    '���������ļ�
        If Dir(App.Path & "\QDB", vbDirectory) = "" Then
            MkDir App.Path & "\QDB"
        End If
        Open RtErrFileName For Append As cRtErrFile
        Open LogFileName For Append As cLogFile
    #End If
    Exit Sub
Err:
    Resume Next
End Sub

Private Sub Class_Terminate()
    Close
End Sub

Public Property Get QDB_Version() As String
    QDB_Version = cQDB_Version
End Property

Public Property Let LogFile(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.LogFile = 5
    Close #cLogFile
    LogFileName = vData
    Open LogFileName For Append As cLogFile
    mvarLogFile = vData
End Property


Public Property Get LogFile() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.LogFile
    mvarLogFile = LogFileName
    LogFile = mvarLogFile
End Property

Public Sub Log(DebugMsg As String)
    Debug.Print DebugMsg
    #If IfWriteToLogFile Then
        Print #cLogFile, DateTime & "|" & "Debug_Msg" & "|" & DebugMsg
    #End If
End Sub

Public Sub Runtime_Error(Position As String, ErrMsg As String, Optional ErrNum As Integer = 0)
    Debug.Print "Runtime_Err: " & ErrMsg & " ErrNum: " & ErrNum
    #If IfRecordRunTimeError Then
        Print #cRtErrFile, DateTime & "|" & Position & "|" & "Runtime_Err" & "|" & ErrMsg & "|" & ErrNum
    #End If
End Sub