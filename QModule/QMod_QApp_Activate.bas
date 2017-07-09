Attribute VB_Name = "QMod_QApp_Activate"
Option Explicit
'Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Type QApp_CommonMsg
    TrueOrFalse As Boolean
    Description As String
End Type
Public Enum QApp_Activate_Way_Enum
    QApp_Activate_Email = 0
    QApp_Activate_FTP = 1
    QApp_Activate_Server = 2
End Enum
Const QApp_Activate_Version = "1.2.2"
Const QApp_Mail_Post_1 = "CI_QApp_MsgPost_t1@163.com"
Const QApp_Mail_Post_1P = "MjI5MjcwNTc3A"
Const QApp_Mail_Post_1S = "smtp.163.com"
Const QApp_Activate_Mail_1 = "CI_QApp_Active_t1@163.com"
Dim QApp_Mail_Title As String
Dim QApp_Mail_Body As String

Public Function QApp_Acitvate(ActivateWay As QApp_Activate_Way_Enum) As QApp_CommonMsg
    Select Case ActivateWay
    Case 0    'Email
        Dim SendEmailState As QApp_CommonMsg, dLabel As String, sLabel As String
        QApp_Mail_Title = Chr(91) & Chr(81) & Chr(65) & Chr(112) & Chr(112) & Chr(95) & Chr(65) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(97) & Chr(116) & Chr(101) & Chr(93) _
                        & QApp_Name
        dLabel = "==================================="
        sLabel = "--------------------------------------------------------------"
        QApp_Mail_Body = dLabel & vbCrLf _
                       & "QApp_Activate Mail " & Date & " " & Time & vbCrLf _
                       & "[QApp_Name]" & QApp_Name & vbCrLf _
                       & "[QApp_Author]" & QApp_Author & vbCrLf _
                       & "[QApp_Version]" & QApp_Version & vbCrLf _
                       & "[QApp_Version_1]" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & sLabel & vbCrLf _
                       & "[WAN_IP]" & GetWANIP & vbCrLf & sLabel & vbCrLf _
                       & "[System Info]" & vbCrLf & GetSystemInfo_Str & vbCrLf _
                       & dLabel & vbCrLf _
                       & "(By SkyUI QApp_Activate Module)Ver " & QApp_Activate_Version
        DoEvents
        SendEmailState = SendEmail(QApp_Mail_Post_1, QApp_Activate_Mail_1, QApp_Mail_Post_1P, _
                                   QApp_Mail_Title, QApp_Mail_Body, QApp_Mail_Post_1S)
        If SendEmailState.TrueOrFalse Then
            SaveSetting QApp_Name, "Info", "active", "true"
        End If
        QApp_Acitvate = SendEmailState
    Case 1    'FTP
    Case 2    'Server
    End Select
End Function

Private Function SendEmail(StrFrom As String, StrTo As String, StrPassword As String, StrTitle As String, StrBody As String, StrServer As String, Optional StrAttachment As String) As QApp_CommonMsg
    On Error GoTo Err
    Dim strName As String, ReturnStruct As QApp_CommonMsg
    strName = "http://schemas.microsoft.com/cdo/configuration/"
    Dim objEmail As Object
    Set objEmail = CreateObject("CDO.Message")
    ReturnStruct.TrueOrFalse = False
    SendEmail = ReturnStruct
    objEmail.From = StrFrom    '发送邮件地址
    objEmail.To = StrTo    '接受邮件地址
    objEmail.Subject = StrTitle    '邮件标题
    objEmail.Textbody = StrBody   '邮件内容
    objEmail.Configuration.Fields.Item(strName & "sendusing") = 2
    objEmail.Configuration.Fields.Item(strName & "smtpserver") = StrServer    '发送邮箱的服务器(如：smtp.163.com)
    objEmail.Configuration.Fields.Item(strName & "smtpserverport") = 25
    objEmail.Configuration.Fields.Item(strName & "smtpconnectiontimeout") = 10
    objEmail.Configuration.Fields.Item(strName & "smtpauthenticate") = 1
    objEmail.Configuration.Fields.Item(strName & "sendusername") = Left(StrFrom, InStr(StrFrom, "@") - 1)
    objEmail.Configuration.Fields.Item(strName & "sendpassword") = StrPassword    '发送邮件邮箱密码
    objEmail.Configuration.Fields.Item(strName & "languagecode") = "0x0804"
    objEmail.Configuration.Fields.Update
    objEmail.AddAttachment "" & StrAttachment
    objEmail.BodyPart.Charset = "GB2312"    '设置字符集，防止乱码
    DoEvents
    objEmail.Send
    Set objEmail = Nothing
    ReturnStruct.TrueOrFalse = True
    SendEmail = ReturnStruct
    Exit Function
Err:
    ReturnStruct.TrueOrFalse = False
    ReturnStruct.Description = Err.Description
    SendEmail = ReturnStruct
    'MsgBox "发送失败：" & Err.Description
End Function

Function Base64Encode(Str() As Byte) As String                                  'Base64 编码
    On Error GoTo over                                                          '排错
    Dim buf() As Byte, Length As Long, mods As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    mods = (UBound(Str) + 1) Mod 3
    Length = UBound(Str) + 1 - mods
    ReDim buf(Length / 3 * 4 + IIf(mods <> 0, 3, 0))
    Dim i As Long
    For i = 0 To Length - 1 Step 3
        buf(i / 3 * 4) = (Str(i) And &HFC) / &H4
        buf(i / 3 * 4 + 1) = (Str(i) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(i / 3 * 4 + 2) = (Str(i + 1) And &HF) * &H4 + (Str(i + 2) And &HC0) / &H40
        buf(i / 3 * 4 + 3) = Str(i + 2) And &H3F
    Next
    If mods = 1 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10
        buf(Length / 3 * 4 + 2) = 64
        buf(Length / 3 * 4 + 3) = 64
    ElseIf mods = 2 Then
        buf(Length / 3 * 4) = (Str(Length) And &HFC) / &H4
        buf(Length / 3 * 4 + 1) = (Str(Length) And &H3) * &H10 + (Str(i + 1) And &HF0) / &H10
        buf(Length / 3 * 4 + 2) = (Str(Length) And &HF) * &H4
        buf(Length / 3 * 4 + 3) = 64
    End If
    For i = 0 To UBound(buf)
        Base64Encode = Base64Encode + Mid(B64_CHAR_DICT, buf(i) + 1, 1)
    Next
over:
End Function

Function Base64Uncode(B64 As String) As Byte()                                  'Base64 解码
    On Error GoTo over                                                          '排错
    Dim OutStr() As Byte, i As Long, j As Long
    Const B64_CHAR_DICT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    If InStr(1, B64, "=") <> 0 Then B64 = Left(B64, InStr(1, B64, "=") - 1)     '判断Base64真实长度,除去补位
    Dim Length As Long, mods As Long
    mods = Len(B64) Mod 4
    Length = Len(B64) - mods
    ReDim OutStr(Length / 4 * 3 - 1 + Switch(mods = 2, 1, mods = 3, 2))
    For i = 1 To Length Step 4
        Dim buf(3) As Byte
        For j = 0 To 3
            buf(j) = InStr(1, B64_CHAR_DICT, Mid(B64, i + j, 1)) - 1
        Next
        OutStr((i - 1) / 4 * 3) = buf(0) * &H4 + (buf(1) And &H30) / &H10
        OutStr((i - 1) / 4 * 3 + 1) = (buf(1) And &HF) * &H10 + (buf(2) And &H3C) / &H4
        OutStr((i - 1) / 4 * 3 + 2) = (buf(2) And &H3) * &H40 + buf(3)
    Next
    If mods = 2 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
    ElseIf mods = 3 Then
        OutStr(Length / 4 * 3) = (InStr(1, B64_CHAR_DICT, Mid(B64, Length + 1, 1)) - 1) * &H4 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &H30) / 16
        OutStr(Length / 4 * 3 + 1) = ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 2, 1)) - 1) And &HF) * &H10 + ((InStr(1, B64_CHAR_DICT, Mid(B64, Length + 3, 1)) - 1) And &H3C) / &H4
    End If
    Base64Uncode = OutStr                                                       '读取解码结果
over:
End Function

Private Function GetWANIP() As String
    Dim h As Object, s As String
    Set h = CreateObject("Microsoft.XMLHTTP")
    h.Open "GET", "http://city.ip138.com/ip2city.asp", False
    h.Send
    If h.ReadyState = 4 Then s = StrConv(h.Responsebody, vbUnicode)
    GetWANIP = Mid(s, InStr(1, s, "[") + 1, InStrRev(s, "]") - InStr(1, s, "[") - 1)
End Function

Private Function GetSystemInfo_Str() As String
    Dim System As Object, Item As Object, i As Integer
    Dim ret As String
    Set System = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
    For Each Item In System
        ret = ret & "Computer Name=" & Item.Name & vbCrLf
        ret = ret & ("Status=" & Item.Status) & vbCrLf
        ret = ret & ("SystemType=" & Item.SystemType) & vbCrLf
        ret = ret & ("Manufacturer=" & Item.Manufacturer) & vbCrLf
        ret = ret & ("Model=" & Item.Model) & vbCrLf
        ret = ret & ("totalPhysicalMemory=" & Item.totalPhysicalMemory \ 1024 \ 1024 & "MB") & vbCrLf
        ret = ret & ("domain=" & Item.domain) & vbCrLf
        ret = ret & ("Workgroup=" & Item.Workgroup) & vbCrLf
        ret = ret & ("username=" & Item.username) & vbCrLf
        ret = ret & ("BootupState=" & Item.BootupState) & vbCrLf
        ret = ret & ("Primary Owner Name=" & Item.PrimaryOwnerName) & vbCrLf
        ret = ret & ("CreationClassName" & Item.CreationClassName) & vbCrLf    '系统类型
        ret = ret & ("Description=" & Item.Description)    '电脑类型
    Next
    GetSystemInfo_Str = ret
End Function

Private Function GetWindowsVersion_CMD() As String
    Shell "cmd /C ver> " & App.Path & "\winver"
    Sleep 200
    Dim WinVer As String
    Open App.Path & "\winver" For Input As #1
    Line Input #1, WinVer
    Line Input #1, WinVer
    Close #1
    Kill App.Path & "\winver"
    GetWindowsVersion_CMD = WinVer
End Function
