VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QClass_QAppSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary
'@@require
'@@reference
'@@license
'@@author
'@@create
'@@modify
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       接口继承
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有常量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       事件声明
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       私有常量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------
Dim Rtn As String
Dim Success As String

'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
Private Declare Function GetPrivateProfileStringByKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringKeys& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function GetPrivateProfileStringSections& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName&, ByVal lpszKey&, ByVal lpszDefault$, ByVal lpszReturnBuffer$, ByVal cchReturnBuffer&, ByVal lpszFile$)
Private Declare Function WritePrivateProfileStringByKeyName& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileStringToDeleteKey& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Long, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileStringToDeleteSection& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Long, ByVal lpString As Long, ByVal lplFileName As String)
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long


'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
'
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
'
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      事件处理
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////

Private Function GetAllSections(File$) As Collection    'Read All Sections
    Dim strTemp As String, strSet As New Collection

    On Error Resume Next
    Rtn = String$(128, 0)
    Success = GetPrivateProfileStringSections(0, 0, "", Rtn, 127, File$)
    Dim NullOffset%
    Do
        NullOffset% = InStr(Rtn, Chr$(0))
        If NullOffset% > 1 Then
            strTemp = Mid$(Rtn, 1, NullOffset% - 1)
            strSet.Add strTemp
            Rtn = Mid$(Rtn, NullOffset% + 1)
        End If
    Loop While NullOffset% > 1
    
    Set GetAllSections = strSet
    Set strSet = Nothing
End Function

Private Function GetAllKeys(Section$, File$) As Collection
    Dim strTemp As String, strSet As New Collection

    Rtn = String$(256, 0)
    Success = GetPrivateProfileStringKeys(Section$, 0, "", Rtn, 255, File$)    'get the entries from the section
    Dim NullOffset%
    Do
        NullOffset% = InStr(Rtn, Chr$(0))
        If NullOffset% > 1 Then
            strTemp = Mid$(Rtn, 1, NullOffset% - 1)
            strSet.Add Mid$(Rtn, 1, NullOffset% - 1)
            Rtn = Mid$(Rtn, NullOffset% + 1)
        End If
    Loop While NullOffset% > 1

    Set GetAllKeys = strSet
    Set strSet = Nothing
End Function

Private Function GetPrivateStringValue(Section$, Key$, File$) As String    'Read Value
    Dim KeyValue$
    Dim characters As Long
    KeyValue$ = String$(256, 0)
    characters = GetPrivateProfileStringByKeyName(Section$, Key$, "", KeyValue$, 255, File$)
    If characters > 1 Then
        KeyValue$ = Left$(KeyValue$, characters)
    End If
    GetPrivateStringValue = KeyValue$
End Function

Private Sub AddSection(Section$, File$)
    WritePrivateProfileStringByKeyName Section$, "", "", File$
End Sub

Private Sub AddKey(Section$, Key$, Value$, File$)
    WritePrivateProfileStringByKeyName Section$, Key$, Value$, File$
End Sub

Private Sub DeleteSection(Section$, File$)
    WritePrivateProfileStringToDeleteSection Section$, 0&, 0&, File$
End Sub

Private Sub DeleteKey(Section$, Key$, File$)
    WritePrivateProfileStringToDeleteKey Section$, Key$, 0&, File$
End Sub

'//////////////////////////////////////////////////////////////////////////////
'//
'//      继承实现
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////

Function SaveToConfig(Section As String, Key As String, Value As String) As Long
    Dim X As Long
    X = WritePrivateProfileString(Section, Key, Value, QApp.DefaultConfigFile)
    SaveToConfig = X
End Function

Function SaveToFile(IniFile As String, appName As String, keyName As String, valueNew As String) As Long
    Dim X As Long
    X = WritePrivateProfileString(appName, keyName, valueNew, IniFile)
    SaveToFile = X
End Function

Function ReadFromConfig(Section As String, Key As String) As String
    If Len(Dir(QApp.DefaultConfigFile)) = 0 Then '检测默认配置文件是否存在
        ReadFromConfig = ""
        Exit Function
    End If
    
    Dim strDefault As String
    Dim lngBuffLen As Long
    Dim strResu As String
    Dim X As Long

    strResu = String$(1025, vbNullChar): lngBuffLen = 1025
    strDefault = ""
    X = GetPrivateProfileString(Section, Key, _
                                strDefault, strResu, lngBuffLen, QApp.DefaultConfigFile)
    ReadFromConfig = Left$(strResu, X)
End Function

Function ReadFromFile(IniFile As String, appName As String, keyName As String) As String
    Dim strDefault As String
    Dim lngBuffLen As Long
    Dim strResu As String
    Dim X As Long

    strResu = String$(1025, vbNullChar): lngBuffLen = 1025
    strDefault = ""
    X = GetPrivateProfileString(appName, keyName, _
                                strDefault, strResu, lngBuffLen, IniFile)
    ReadFromFile = Left$(strResu, X)
End Function

Function ReadAllSectionsFromFile(IniFile As String) As Collection
    Set ReadAllSectionsFromFile = GetAllSections(IniFile)
End Function

Function ReadAllKeysFromFile(IniFile As String, Section As String) As Collection
    Set ReadAllKeysFromFile = GetAllKeys(Section, IniFile)
End Function

Sub Save(Section As String, Key As String, Setting As String)
    SaveSetting QApp.Name, Section, Key, Setting
End Sub

Function Read(Section As String, Key As String) As String
    Read = GetSetting(QApp.Name, Section, Key)
End Function

