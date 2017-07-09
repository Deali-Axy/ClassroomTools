Attribute VB_Name = "mdlSysTraySupport"
Option Explicit

'-------------------------- ϵͳ���� clsSysTray ��ģ���֧��ģ�� ----------------------------

'��Ҫ clsSysTray ��ģ���֧��
'��Ҫ clsHashLK��clsSubClass ��ģ�� �� mdlSubClass ��׼ģ���֧��

'#���ʻ�������ʾ�ַ�������

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Const WM_USER = &H400
Public Const ST_NOTI_MSG = WM_USER + 1001&    '�Զ���ϵͳ���̵���Ϣ

'���Ǽǡ����������е�ϵͳ���̡�Key=��Ӧ�����hWnd��Data=һ��clsSysTray����ĵ�ַ
'һ���������ֻ��ʹ��һ�� clsSysTray ������������
Private mHashSysTrays As New clsHashLK


Public Function STRegOneObject(ByVal hwnd As Long, _
                               ByVal addrClsSysTray As Long) As Boolean
'���Ǽǡ�һ�� clsSysTray ����
'һ���������ֻ��ʹ��һ�� clsSysTray ������������
    If mHashSysTrays.IsKeyExist(hwnd) Then
        '�ô����Ѿ����Ǽǹ���������ϵͳ���̡��������ظ�����ϵͳ����
        STRegOneObject = False
    Else
        STRegOneObject = mHashSysTrays.Add(addrClsSysTray, hwnd, 0, "", False)
    End If
End Function

Public Function STUnRegOneObject(ByVal hwnd As Long) As Boolean
'ȡ��"�Ǽ�"һ������� clsSysTray ���󣬼�ɾ�� mHashSysTrays �е�һ����¼
    mHashSysTrays.Remove hwnd
End Function


Public Function STWndProc(ByVal hwnd As Long, _
                          ByVal Msg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

'������ϵͳ���̵Ĵ��ڱ����໯���Զ��崰�ڳ���
    If Msg = ST_NOTI_MSG Then
        '////////// ���ϵͳ������Ϣ�����ݸ���Ӧ�� clsSysTray ���� //////////
        Dim lngAddrObj As Long    '��Ӧ clsSysTray �����ַ

        '�ӹ�ϣ���� clsSysTray �����ַ
        lngAddrObj = mHashSysTrays.Item(hwnd, False)

        If lngAddrObj Then    '��� lngAddrObj Ϊ0��ʾʧ�ܻ�VB�����и���û�д˴��ڶ�Ӧ������
            'ͨ�������õ��� clsSysTray ����� EventsGen ���������¼�
            Dim objCls As clsSysTray
            CopyMemory objCls, lngAddrObj, 4
            objCls.EventsGen wParam, lParam
            CopyMemory objCls, 0&, 4
        End If
    End If

    '����ֵ�������ദ��ͨ��ģ�顱���Զ�����Ĭ�ϴ��ڳ�����
    STWndProc = gc_lngEventsGenDefautRet
End Function



