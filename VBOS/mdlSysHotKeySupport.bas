Attribute VB_Name = "mdlSysHotKeySupport"
Option Explicit

'-------------------------- ϵͳ�ȼ� clsSysHotKey ��ģ���֧��ģ�� ----------------------------
'��Ҫ clsSysHotKey ��ģ���֧��
'��Ҫ clsHashLK��clsStack��clsSubClass ��ģ�� �� mdlSubClass ��׼ģ���֧��

'#���ʻ�������ʾ�ַ�������

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_HOTKEY = &H312

'���Ǽǡ����������е�ϵͳ�ȼ���Key=�ȼ�ID��Data=һ��clsSysHotKey����ĵ�ַ
Private mHashSysHotkeys As New clsHashLK

'����ϵͳ�ȼ�ID�ã�����ֵ���������������� mc_MaxHotKeyID��
Private mIDInc As Long
Private Const mc_MaxHotKeyID As Long = &H7FFF
'����ϵͳ�ȼ�ID��ʹ�ö�ջ��
Private mStackRecycleIDs As New clsStack

'��� clsSysHotKey ����Ϊͬһ������ hWnd �����ȼ�ʱ��Ҫ�ظ��Ըô������໯ _
 '��ȡ�����໯�����Զ��崰�ڳ�����ͬһ�� SHKWndProc
'Ϊ�����ظ����໯����ȡ������ʹ�õ����໯�����໯/ȡ�����໯Ҫ�ɱ�ģ��ͳһ����
'"�Ǽ�"����������Ϊ����ϵͳ�ȼ������໯�Ĵ��� hWnd
Private mHashSubClsedHwnds As New clsHashLK    'Key=hWnd��Data=�ô������໯�����������ظ������˼��Σ�Ϊ0ʱ�Զ�ȡ�����໯��

Public Function SHKRegOneObject(ByVal addrClsSysHotKey As Long) As Long
'���Ǽǡ�һ�� clsSysHotKey ���󣬷���һ������ʹ�õ�ϵͳ�ȼ�ID����¼�ö� _
  '�����ID��Ӧ���� clsSysHotKey �� Class_Initialize ʱ���ã�

'�������ظ�ϵͳ�ȼ� ID��ʧ�ܷ��� 0
'addrClsSysHotKey Ϊһ��clsSysHotKey����ĵ�ַ
'API����Ҫ����ȼ� id ��Χ�� 0x0000��0xBFFF�� _
  '��������涨��Χ��0x0001��0x7FFF����ʹ��0�͸�����

    Dim lngNewID As Long
    '�� mIDInc ������1 �ķ�ʽ����һ���µ��ȼ� ID��lngNewID
    mIDInc = mIDInc + 1
    If mIDInc > mc_MaxHotKeyID Then
        '//////////  mIDInc ����������ͷ //////////
        '�ָ� mIDInc ������ֵ
        mIDInc = mIDInc - 1
        '����л���ID��ʹ�û���ID
        If mStackRecycleIDs.IsEmpty Then
            'û�л���ID������ʧ��
            SHKRegOneObject = 0
            Exit Function
        Else
            'ʹ�û��յ�ջ��ID
            lngNewID = mStackRecycleIDs.PopLong
        End If
    Else
        '////////// �� mIDInc ������ֵ��Ϊ���ȼ� ID //////////
        lngNewID = mIDInc
    End If

    '�ڹ�ϣ�� mHashSysHotkeys �м�¼ addrClsSysHotKey ��ַ�� lngNewID ��Ӧ
    If mHashSysHotkeys.Add(addrClsSysHotKey, lngNewID, 0, "", False) Then
        SHKRegOneObject = lngNewID
    Else
        SHKRegOneObject = 0
    End If
End Function

Public Function SHKUnRegOneObject(ByVal IDSysHotKey As Long) As Boolean
'ȡ�����Ǽǡ�һ�� clsSysHotKey ���󣬼�ɾ�� mHashSysHotkeys �е�һ�� _
  '��¼���� clsSysHotKey �� Class_Terminate ʱ���ã�

    If mHashSysHotkeys.Remove(IDSysHotKey, False) Then
        '�����ͷŵ��ȼ� ID
        mStackRecycleIDs.PushLong IDSysHotKey
        '���سɹ�
        SHKUnRegOneObject = True
    Else
        '����ʧ��
        SHKUnRegOneObject = False
    End If
End Function

Public Function SHKSubClassHwnd(ByVal hwnd As Long) As Boolean
'���󽫴��� hwnd ���ദ��Ϊ SHKWndProc �Զ��崰�ڳ���
'�� clsSysHotKey ��ģ�����

    If mHashSubClsedHwnds.IsKeyExist(hwnd) Then
        '�ô����ѱ����໯Ϊ SHKWndProc �Զ��崰�ڳ��򣬲����ظ����໯
        '��Ҫ�������໯�������������е� Item ֵΪ����ֵ +1
        Dim lngCt As Long
        lngCt = mHashSubClsedHwnds.Item(hwnd, False)
        lngCt = lngCt + 1
        mHashSubClsedHwnds.Remove hwnd, False
        mHashSubClsedHwnds.Add lngCt, hwnd, 0, "", False
        SHKSubClassHwnd = True
    Else
        '�ô��ڻ�û�����໯Ϊ SHKWndProc �Զ��崰�ڳ���
        '�������໯
        SHKSubClassHwnd = SCCreateSubClass(hwnd, AddressOf SHKWndProc)
        'mHashSubClsedHwnds �м�¼1��
        mHashSubClsedHwnds.Add 1, hwnd, 0, "", False
    End If
End Function

Public Sub SHKUnSubClassHwnd(ByVal hwnd As Long)
'���󽫴��� hwnd ȡ�����ദ��Ϊ SHKWndProc �Զ��崰�ڳ���
'�� clsSysHotKey ��ģ�����

    If mHashSubClsedHwnds.IsKeyExist(hwnd) Then
        '�ô����ѱ����໯Ϊ SHKWndProc �Զ��崰�ڳ���
        '���໯������1�����໯����Ϊ0ʱ������ȡ�����໯
        Dim lngCt As Long
        lngCt = mHashSubClsedHwnds.Item(hwnd, False)
        lngCt = lngCt - 1
        mHashSubClsedHwnds.Remove hwnd, False
        If lngCt <= 0 Then
            'ȡ�����໯
            SCRestoreSubClassOne hwnd, AddressOf SHKWndProc
        Else
            '������ȡ�����໯���Ѽ� 1 �� lngCt �ٴ��� mHashSubClsedHwnds
            mHashSubClsedHwnds.Add lngCt, hwnd, 0, "", False
        End If
    Else
        'mHashSubClsedHwnds ��û��¼�ô������໯����Ϊ�ݴ�ֱ��ȡ���������໯
        SCRestoreSubClassOne hwnd, AddressOf SHKWndProc
    End If
End Sub

Public Function SHKWndProc(ByVal hwnd As Long, _
                           ByVal Msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
'������ϵͳ�ȼ��Ĵ��ڱ����໯���Զ��崰�ڳ���


    If Msg = WM_HOTKEY Then
        '////////// ϵͳ�ȼ������£�wParam ����Ϊ�ȼ� ID //////////

        Dim lngAddrObj As Long    '��ӦclsSysHotKey�����ַ

        '�ӹ�ϣ���� clsSysHotKey �����ַ
        lngAddrObj = mHashSysHotkeys.Item(wParam, False)

        If lngAddrObj Then    '��� lngAddrObj Ϊ0��ʾʧ�ܻ�VB�����и���û�д��ȼ�
            'ͨ�������õ��� clsSysHotKey ����� RaiseSysKeyPressedEvent ���������¼�
            Dim objCls As clsSysHotKey
            CopyMemory objCls, lngAddrObj, 4
            objCls.RaiseSysKeyPressedEvent
            CopyMemory objCls, 0&, 4
        End If
    End If

    '����ֵ�������ദ��ͨ��ģ�顱���Զ�����Ĭ�ϴ��ڳ�����
    SHKWndProc = gc_lngEventsGenDefautRet
End Function



