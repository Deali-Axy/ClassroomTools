Attribute VB_Name = "mdlSubClass"
Option Explicit

'------------------ ���ദ��ģ�� --------------------------
'-----Support: ��Ҫ clsSubClass, clsHashLK ��ģ���֧�� ------------------
'# ���ʻ�������ʾ�ַ�������
'
'//////////////////////////////////////////////////////////////
'�ô�ģ��ĺ�����ʵ�ֲ�����һ�������е����������������֧�ֶ�ͬһ������ _
 '������໯����ͬһ�������ظ����� SCCreateSubClass() ����ʱֻ�������� _
 '���ô��ڳ����ַһ�Σ�ÿ�ζ���¼�û��������Զ��崰�ڳ����ַ���˵�ַ�γ� _
 'һ���б��Ժ���������б��е����Զ��崰�ڳ���
'��֧�ַ�ֹ ������໯ʱ ͬһ���Զ��崰�ڳ����ַ��������ӵ��б�
'ֻΪ �Զ��崰�ڳ����ַ ά��һ���б�Ĭ�ϴ��ڳ�����Զ _
 '������ mCSC(i).PrevWndProcAddr �в��䡣������������˵����Զ SetWindowLong _
 '�Ѵ��ڳ���ĵ�ַ��Ϊ��ģ��� SCMyWndProc �������� SCMyWndProc �ٵ��� _
 '�б��е������Զ��崰�ڳ���
'�� clsSubClass ����ȷ���������ʱ�ָ�����
'//////////////////////////////////////////////////////////////

'�÷���
'1. ��Ҫ��һ���������໯ʱ���� SCCreateSubClass() �������ú���ָ�� _
 '���໯����Զ���Ĵ��ڳ������磺SCCreateSubClass(Me.hwnd, AddressOf MyProc2) _
 '����ͬһ�����ڶ�����࣬�ظ����ñ��������ɣ��� hwnd ����ÿ�α���һ�£�
'ע�⣺�����Զ��崰�ڳ��򶼲����С�����Ĭ�ϴ��ڳ��򡱵���䣬������ _
 '�з���ֵ��䣬return gc_lngEventsGenDefautRet�����������Ƿ��ظ�ֵ _
 '��Ĭ�ϵĴ��ڳ����ܱ�����
'2. ʹ�� SCGetPreWndProcAddr() ����Ѿ����໯�Ĵ��ڵ� ԭʼĬ�ϴ��ڳ���ĵ�ַ _
 '�ɾݴˣ��������������Ĭ�ϴ��ڳ���
'3. �ָ�һ���������໯ʱ���� SCRestoreSubClassWhole() �������÷���ֵ�ж��Ƿ� _
 �ָ��ɹ�����û�лָ��ɹ���������Ҳ����Ϊ������ mCSC() �ռ䣨Ϊ���� _
 �����ͷŶ��󣩣��ɱ�ģ��� cHashUnRestPreprocs ��¼û�ָ��ɹ�����Ϣ _
 'ע�⣺���øú����Դ��� hwnd �����е����໯����������
'4. Ҫ�ָ����������д��ڵ����໯���� SCRestoreSubClassAllWnds() ������δ�ָ� _
 �ɹ��Ĵ�����Ϣ���� SCGetUnSucRestCount() �� SCGetUnSucRestWins() �������
'
'
'�����������໯�����д�����Ϣ�ñ�ģ��� mCSC() ��¼����������ûʹ�õ� _
 '����ռ��±��¼�� mCSCIdxUnused() �У��ɱ�ģ���Զ�����ִ�� _
 'SCRestoreSubClassWhole �� mCSC() �Ŀ���Ԫ�ر���¼�� mCSCIdxUnused() �У� _
 '�� mCSC() Ԫ��ֻռһ���±꣬������Զ� Set=Nothing ��ռ�ڴ档 _
 '�´��½�������Ҫ�ռ�ʱ����ʹ�� mCSCIdxUnused() �м�¼���±�ռ䡣
'���������һ�� SCRestoreSubClassAllWnds() �������� mCSC() �ִ�0��ʼ����, _
 ' mCSCIdxUnused() Ҳ�����
'---------------------------------------------------------------
Public Const gc_lngEventsGenDefautRet As Long = 1147483647    '��ֵ=API�����ؼ�ʹ�õ� gc_APICEventsGenDefautRet

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const GWL_WNDPROC = (-4)

Private mCSC() As clsSubClass, mCSCCount As Long
Attribute mCSCCount.VB_VarUserMemId = 1073741824
Private mCSCIdxUnused() As Long, mCSCIdxUnusedCount As Long    '��¼ mCSC() ��ûʹ�õĿռ��±꼴���ָ����໯��Ŀռ��±꣬�Ա��Ժ����½�����ʱʹ�ø��±�Ԫ��
Attribute mCSCIdxUnused.VB_VarUserMemId = 1073741826
Attribute mCSCIdxUnusedCount.VB_VarUserMemId = 1073741826
Private cHashCSCIdxes As New clsHashLK    'Data=mCSC() ���±ꣻkey= hwnd��ͨ����Ҳ��ģ���м�¼����Щ���ڱ����໯�ˣ�û������Ĵ�����Ϊû�����໯��Ҳ��Ҫ�� cHashUnRestPreprocs ���ǲ������ڱ������ָ�ʧ�ܵ������
Attribute cHashCSCIdxes.VB_VarUserMemId = 1073741828
Private cHashUnRestPreprocs As New clsHashLK    'û�лָ��ɹ��ı�����Ĵ�����Ϣ��key=û�лָ��ɹ��� hwnd��Data=Ĭ�ϴ��ڳ����ַ
Attribute cHashUnRestPreprocs.VB_VarUserMemId = 1073741829

Dim mIdx As Long, mR As Long, mRet As Long, mAddr As Long, mI As Long    '����������ÿ��ִ�к��������¶���
Attribute mIdx.VB_VarUserMemId = 1073741830
Attribute mR.VB_VarUserMemId = 1073741830
Attribute mRet.VB_VarUserMemId = 1073741830
Attribute mAddr.VB_VarUserMemId = 1073741830
Attribute mI.VB_VarUserMemId = 1073741830

Public Function SCMyWndProc(ByVal hwnd As Long, _
                            ByVal Msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
'�������໯��������Ϊ�˴��ڳ����������໯�����ô˴��ڳ���
'��������һ���� �� mCSC() �еĶ����м�¼�����С��Զ��崰�ڳ��򡱣� _
  '��֧�ֶ�ͬһ���ڵĶ�����໯
'�����Զ��崰�ڳ����У���һ�������� gc_lngEventsGenDefautRet �������Ͳ� _
  '�����Ĭ�ϵĴ��ڳ��򣬵�������λ�����������Զ��崰�ڳ��򣻱������� _
  '��ֵΪ���һ�������� gc_lngEventsGenDefautRet ���Զ��崰�ڳ���ķ���ֵ


'-------- �ҵ����໯�� mCSC() ���±� ���룺 mIdx --------
    mIdx = cHashCSCIdxes.Item(hwnd, False)
    If mIdx = 0 Then SCMyWndProc = 0: Exit Function

    '-------- �� mCSC(mIdx) ������� �Զ��崰�ڳ��򣬲���һ���� --------
    '������ mRet = gc_lngEventsGenDefautRet����һ���ô��ڳ���ʱ������ _
     '���ڳ��򷵻�ֵ <> gc_lngEventsGenDefautRet���ͻ����� mRet=�˷��� _
     'ֵ�����Ժ��� ����ֵ <> gc_lngEventsGenDefautRet �ģ��ͻ���� _
     'mRet����������� mRet(�� mRet <> gc_lngEventsGenDefautRet ��)
    mRet = gc_lngEventsGenDefautRet
    With mCSC(mIdx)
        For mI = 1 To .UserProcAddrsCount
            mAddr = .UserProcAddr(mI)
            If mAddr Then  'If mAddr <> 0 Then
                mR = CallWindowProc(mAddr, hwnd, Msg, wParam, ByVal lParam)
                If mR <> gc_lngEventsGenDefautRet Then mRet = mR
            End If
        Next mI
    End With

    '-------- ����ֵ --------
    If mRet = gc_lngEventsGenDefautRet And (Not mCSC(mIdx) Is Nothing) Then
        'mRet��gc_lngEventsGenDefautRet ˵�������Զ��崰�ڳ��򶼷��ص� _
         '�� gc_lngEventsGenDefautRet���͵���Ĭ�ϴ��ڳ���
        SCMyWndProc = CallWindowProc(mCSC(mIdx).PrevWndProcAddr, _
                                     hwnd, Msg, wParam, ByVal lParam)
    Else
        'mRet<>gc_lngEventsGenDefautRet ˵�����Զ��崰�ڳ��� ���� _
         '<>gc_lngEventsGenDefautRet �ģ��������һ�� ����ֵ _
         '��Ϊ gc_lngEventsGenDefautRet ���Զ��崰�ڳ���ķ���ֵ
        SCMyWndProc = mRet
    End If
End Function


Public Function SCCreateSubClass(ByVal hwnd As Long, _
                                 ByVal addrMyWndProc As Long) As Boolean
'���໯���� hwnd������¼������໯
'addrMyWndProc ָ�����໯����Զ���Ĵ��ڳ���=0��ʹ�ñ�ģ�� _
  �� SCMyWndProc�����磺SCCreateSubClass(Me.hwnd, AddressOf MyProc2)
'����ͬһ�����ڶ�����࣬�ظ����ñ��������ɣ��� hwnd ����ÿ�α���һ��

    Dim blCannotRest As Boolean
    'ͨ�� cHashCSCIdxes ���Ƿ��Ѿ������໯�����Ƿ����ظ����໯��
    If SCIsHwndSubClassed(hwnd, blCannotRest) Then

        '//////////// �ô����Ѿ������໯���ˡ������������汻���໯���ˣ� _
         '�����������໯���ָֻ�����ʱ�ָ�ʧ�ܵ������ ////////////
        If blCannotRest Then
            '//// �����໯���ָֻ�����ʱ�ָ�ʧ�ܵ����,�޷��ٽ��˴������໯,���� false

            SCCreateSubClass = False
            Exit Function

        Else
            '//// �������汻���໯����

            '�ҵ���ǰ���໯ʱʹ�õ���ģ������ҵ� mCSC() ����Ԫ�ص��±꼴�ɣ�
            Dim lngIdx As Long
            lngIdx = cHashCSCIdxes.Item(hwnd, False)

            '��ǰ�����໯����Ӧ���ܹ��ҵ� mCSC() ���±꣬���Ҳ������ͳ����˳�
            If lngIdx <= 0 Then SCCreateSubClass = False: Exit Function

            '��������佨���µ����࣬�����ѱ����໯����CreateSubClass ����ֻ�� _
             '���һ��"�Զ��崰�ڳ���ĵ�ַ" addrMyWndProc ���ѣ������ظ����໯
            SCCreateSubClass = mCSC(lngIdx).CreateSubClass(hwnd, addrMyWndProc)

            Exit Function

        End If

    Else
        '//////////// �ô���û�����໯�����½����໯ ////////////
        SCCreateSubClass = CreateNewFirstOneSubClass(hwnd, addrMyWndProc)
    End If

End Function

Public Function SCRestoreSubClassOne(ByVal hwnd As Long, _
                                     ByVal addrUserProc As Long) As Boolean
'�ָ����໯,���ָ� addrUserProc ���ڳ�������࣬���� �Զ��崰�ڳ��� �б���ɾ������
'��ɾ���� �Զ��崰�ڳ��� �б�Ϊ�գ��򳹵׻ָ��ô��ڵ�����
'�����׻ָ��ô��ڵ����࣬�����Ƿ�ָ��ɹ��������� cHashCSCIdxes.Remove _
  '��¼������δ�ָ��ɹ�����Ϣ���ٱ���¼�� cHashUnRestPreprocs ��

    Dim lngIdx As Long, blRet As Boolean

    lngIdx = cHashCSCIdxes.Item(hwnd, False)
    If lngIdx = 0 Then SCRestoreSubClassOne = False: Exit Function
    If mCSC(lngIdx) Is Nothing Then SCRestoreSubClassOne = False: Exit Function

    '�ָ�����
    '���溯�����غ�if �б�Ϊ�գ���ʲô������������ʹ�ã������б�Ϊ�գ� _
     '�ж����溯������ֵ������ֵ��ʾ�����溯���Զ����� UnSubclassWhole �ĳɹ����
    blRet = mCSC(lngIdx).UnSubclassOne(addrUserProc)

    '�ж��б��Ƿ�Ϊ��
    If mCSC(lngIdx).UserProcAddrsCount = 0 Then
        If blRet Then
            SCRestoreSubClassOne = True
        Else
            '������� UnSubclassOne ���ɹ��� _
             �Ѵ˲��ɹ���Ϣ��ӵ� cHashUnRestPreprocs
            cHashUnRestPreprocs.Add _
                    mCSC(lngIdx).PrevWndProcAddr, mCSC(lngIdx).hwnd
            SCRestoreSubClassOne = False
        End If

        '�ͷ� mCSC(lngIdx) �ռ䣬�����˿���ռ��¼��"����վ" mCSCIdxUnused()
        Set mCSC(lngIdx) = Nothing
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(1 To mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = lngIdx

        'ɾ�� cHashCSCIdxes �Ķ�Ӧ��
        cHashCSCIdxes.Remove hwnd, False
    Else
        'ʲô���������������÷���ֵ��ʾ�Ƿ�ɾ�������ɹ�
        SCRestoreSubClassOne = blRet
    End If
End Function

Public Function SCRestoreSubClassWhole(ByVal hwnd As Long) As Boolean
'�ָ����໯���Դ��� hwnd �����е����໯����������
'�����Ƿ�ָ��ɹ��������� cHashCSCIdxes.Remove ��¼
'����δ�ָ��ɹ�����Ϣ���ٱ������� cHashUnRestPreprocs ��
    Dim lngIdx As Long, blRet As Boolean
    lngIdx = cHashCSCIdxes.Item(hwnd, False)
    If lngIdx Then    'If lngIdx <> 0 Then
        blRet = UnSubclassOneCSCWhole(lngIdx)
        cHashCSCIdxes.Remove hwnd, False
        SCRestoreSubClassWhole = blRet
    Else
        SCRestoreSubClassWhole = False
    End If
End Function

Public Function SCGetPreWndProcAddr(hwnd As Long) As Long
'���� hwnd �ҵ�����Ĭ�ϴ��ڳ���ĵ�ַ�������� 0
'�� hwnd δ�����ദ��Ҳ���� 0

    Dim lngCSCIdx As Long

    '-------- �� hwnd �� cHashCSCIdxes �����޼�¼���ֱ��� --------
    If cHashCSCIdxes.IsKeyExist(hwnd) Then
        '//////// hwnd �� cHashCSCIdxes ���м�¼��ֱ�ӻ�ȡ������ ////////
        lngCSCIdx = cHashCSCIdxes.Item(hwnd)
        If mCSC(lngCSCIdx) Is Nothing Then
            'mCSC(lngCSCIdx) �����ѱ�����
            SCGetPreWndProcAddr = 0
        Else
            If mCSC(lngCSCIdx).hwnd = hwnd Then    '����֤һ�� hwnd �Ƿ����
                SCGetPreWndProcAddr = mCSC(lngCSCIdx).PrevWndProcAddr
            Else
                SCGetPreWndProcAddr = 0
            End If
        End If
        Exit Function
    Else
        '//////// hwnd �� cHashCSCIdxes ��û�м�¼������ cHashUnRestPreprocs _
         '����û�м�¼�����б�������û�лָ��ɹ��� ////////
        If cHashUnRestPreprocs.IsKeyExist(hwnd) Then
            '//////// ����û�лָ��ɹ��� ////////
            SCGetPreWndProcAddr = cHashUnRestPreprocs.Item(hwnd)
        Else
            '//////// ��������û�лָ��ɹ��ģ���� hwnd �����໯û�б���¼�� _
             '�޷�ȷ����Ĭ�ϴ��ڳ���ĵ�ַ������0 ////////
            SCGetPreWndProcAddr = 0
        End If
        Exit Function
    End If
End Function

Public Function SCIsHwndSubClassed(ByVal hwnd As Long, _
                                   Optional ByRef blRetCannotRestore As Boolean) As Boolean
'�жϾ��Ϊ hwnd �Ĵ����Ƿ����ദ��������� True/False ��ʾ�Ƿ����ദ���
'������� True������Ҫ�鿴 blRetCannotRestore �����ķ���ֵ _
  'blRetCannotRestore �����Ƿ�ô��� hwnd Ϊ"���ɻָ�����"�Ĵ��ڣ�����ô��� _
  '��ǰ�����ദ��������ָ�����ʱ�ָ�ʧ�ܣ��Ͳ����ٽ��ô������µ����ദ�� _
  '����������ѩ�ϼ�˪�����ϼӴ���ʱ blRetCannotRestore ���� True
'���������� false������� blRetCannotRestore �ķ���ֵ������
    If cHashCSCIdxes.IsKeyExist(hwnd) Then
        '//////// hwnd �� cHashCSCIdxes ���м�¼����ô�ض��Ѿ����ദ���ˣ� _
         �������� true��blRetCannotRestore ���� false ////////
        SCIsHwndSubClassed = True
        blRetCannotRestore = False
    Else
        '//////// hwnd �� cHashCSCIdxes ��û�м�¼�����Ƿ����ڻָ� _
         'ʧ�ܵ���� ////////
        If cHashUnRestPreprocs.IsKeyExist(hwnd) Then
            '//////// hwnd ���ڻָ�ʧ�ܵ��������ôҲ���Ѿ������ദ���� _
             '�������� true��blRetCannotRestore ���� true ////////
            SCIsHwndSubClassed = True
            blRetCannotRestore = True
        Else
            '//////// hwnd �����ڻָ�ʧ�ܵ������ȷʵ�ô���Ŀǰ��û�б����ദ�� ////////
            SCIsHwndSubClassed = False
            'blRetCannotRestore ����ֵ������
        End If
    End If
End Function

Public Function SCRestoreSubClassAllWnds() As Boolean
'��������е����С����໯��
'���� false ����Щ�����໯��û�лָ��ɹ�����������ν���� _
  mCSCCount��mCSC()��cHashCSCIdxes
'�����̽�û�лָ��ɹ��Ĵ�����Ϣ���� cHashUnRestPreprocs
'�����̻�������ͼ�ָ�һ�� cHashUnRestPreprocs �м�¼��δ�ָ��ɹ��Ĵ���
'������ mSCS() �еĻ��� cHashUnRestPreprocs �еģ�ֻҪ��û�ָ��ɹ� _
  �ģ������ͷ��� false

    Dim i As Long
    Dim blRet As Boolean

    '------------ ������Ĭ�Ϸ���ֵ�� true�����������ʧ�ܵ� _
     '�ٸ�Ϊ false ------------
    blRet = True

    '------------ �ָ����� mCSC() �е�Ԫ�� ------------
    If mCSCCount > 0 Then
        For i = 1 To mCSCCount
            If Not UnSubclassOneCSCWhole(i) Then blRet = False
        Next i
        Erase mCSC
        mCSCCount = 0
        cHashCSCIdxes.Clear
    End If

    '------------ �ָ����� cHashUnRestPreprocs �е�Ԫ�� ------------
    Dim retKeys() As Long, ret As Long
    For i = 1 To cHashUnRestPreprocs.GetKeyArray(retKeys)
        ret = 0
        SetLastError 0&
        On Error Resume Next
        ret = SetWindowLong(retKeys(i), GWL_WNDPROC, _
                            cHashUnRestPreprocs.Item(retKeys(i), False))
        If ret = 0 And GetLastError <> 0 Then
            blRet = False
        Else
            cHashUnRestPreprocs.Remove retKeys(i), False
        End If
    Next i

    '------------ ��ա�����վ��mCSCIdxUnused() ------------
    Erase mCSCIdxUnused
    mCSCIdxUnusedCount = 0

    '------------ ���� ------------
    SCRestoreSubClassAllWnds = blRet
End Function

Public Function SCGetUnSucRestCount() As Long
'����û�лָ�����ɹ��Ĵ��ڵĸ���
'û�лָ��ɹ��Ĵ������� SCRestoreSubClassAllWnds ���̼�¼��
    SCGetUnSucRestCount = cHashUnRestPreprocs.Count
End Function

Public Function SCGetUnSucRestWins(retHwnds() As Long, retPrevWnds() As Long) As Long
'����û�лָ�����ɹ��Ĵ��ڵ���Ϣ(�Ӳ������� hwnd �� PrevWndProc)
'�������� û������ָ��ɹ��Ĵ��ڵĸ���

    Dim k As Long, b As Boolean, ret As Long
    Dim i As Long
    If cHashUnRestPreprocs.Count > 0 Then
        ReDim retHwnds(1 To cHashUnRestPreprocs.Count)
        ReDim retPrevWnds(1 To cHashUnRestPreprocs.Count)
    End If

    '������ϣ�� cHashUnRestPreprocs
    cHashUnRestPreprocs.StartTraversal
    i = 1
    Do
        ret = cHashUnRestPreprocs.NextItem(0, "", k, b)
        If b Then Exit Do
        retPrevWnds(i) = ret
        retHwnds(i) = k
        i = i + 1
    Loop
    SCGetUnSucRestWins = cHashUnRestPreprocs.Count
End Function

Private Function CreateNewFirstOneSubClass(ByVal hwnd As Long, _
                                           ByVal addrUserProc As Long) As Long
'�½�����¼һ�����ࣨ���·����һ�� mCSC() ��������¼���� hwnd ������Ĭ�� _
  ���ڳ���ĵ�ַ��Ĭ�ϴ��ڳ���ĵ�ַ�Ļ���ɶ�������ɣ�
'�������ǵ�һ������ʱ���õģ�����Ƕ�ͬһ�������ظ����࣬��Ҫ���ñ��� _
  '������Ҫ�����ǰ����� mCSC() ���󲢵������� CreateSubClass ����
'�������������ж� cHashCSCIdxes ������ hwnd����ȷ���Ƿ��ǵ�һ������
'cHashCSCIdxes ��û�У�����һ�����ࣩʱ�ſɵ��ñ�����
'����ʹ�õ� mCSC() ���±꣬������0

    Dim lngAddrCsc As Long
    Dim idxCSC As Long

    '-------- �·���һ�� mCSC() �����±���룺 idxCSC --------
    idxCSC = 0
    If mCSCIdxUnusedCount > 0 Then
        '//// ʹ��"����վ"�п���ռ��±꣨ʹ��"����վ"�����һ��Ԫ�أ� ////
        idxCSC = mCSCIdxUnused(mCSCIdxUnusedCount)
        '����ռ�ɾ���˼�¼
        mCSCIdxUnusedCount = mCSCIdxUnusedCount - 1
        If mCSCIdxUnusedCount > 0 _
           Then ReDim Preserve mCSCIdxUnused(mCSCIdxUnusedCount)
    Else
        '//// "����վ"��û�п���ռ䣬���� mCSC() ////
        mCSCCount = mCSCCount + 1
        ReDim Preserve mCSC(1 To mCSCCount)
        idxCSC = mCSCCount
    End If

    '-------- �������� --------
    Set mCSC(idxCSC) = New clsSubClass

    '-------- �ö��� mCSC(idxCSC) �������ಢ _
     '��¼ hwnd��Ĭ�ϴ��ں����ĵ�ַ �� �Զ��崰�ڳ���ĵ�ַ --------
    '��ʱ���� mCSC(idxCSC) ������˵��CreateSubClass �����Ǳ���һ�ε���
    lngAddrCsc = mCSC(idxCSC).CreateSubClass(hwnd, addrUserProc)
    If lngAddrCsc = 0 Then GoTo errH

    '-------- ��¼"���ദ��"��ģ�鼶��ϣ��cHashCSCIdxes --------
    cHashCSCIdxes.Add idxCSC, hwnd

    '-------- ����ʹ�õ� mCSC() ���±꣨>=1�� --------
    CreateNewFirstOneSubClass = idxCSC
    Exit Function
errH:
    If idxCSC Then  'If idxCSC <> 0 Then
        Set mCSC(idxCSC) = Nothing

        '����ռ��Ѿ��������˻��Ѿ�ʹ����"����վ"�е�һ���ռ䣬 _
         ������ΰ�����¼��"����վ" mCSCIdxUnused()
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = idxCSC
    End If
    '���� 0 ��ʧ��
    CreateNewFirstOneSubClass = 0
End Function

Private Function UnSubclassOneCSCWhole(ByVal idxCSC As Long) As Boolean
'ȡ�����໯ mCSC(idxCSC)���Դ��ڵ����е����໯����������
'�����Ƿ�ȡ���ɹ������ռ�����Ч(mCSC(idxCSC)=nothing) Ҳ���� true
'�粻�ɹ��򽫴����໯����Ϣ��ӵ� cHashUnRestPreprocs
'��������ζ��ͷ� mCSC(idxCSC) �ռ�
'����������Ӱ�� cHashCSCIdxes����Ҫ����������� cHashCSCIdxes.Remove
    If mCSC(idxCSC) Is Nothing Then
        UnSubclassOneCSCWhole = True
    Else
        If mCSC(idxCSC).UnSubclassWhole Then
            UnSubclassOneCSCWhole = True
        Else
            '�ָ����໯���ɹ���Ҳ�ͷ� mCSC(idxCSC) �ռ䣬���� _
             '���ɹ���Ϣ��ӵ� cHashUnRestPreprocs
            cHashUnRestPreprocs.Add _
                    mCSC(idxCSC).PrevWndProcAddr, mCSC(idxCSC).hwnd
            UnSubclassOneCSCWhole = False
        End If

        '�ͷ� mCSC(idxCSC) �ռ䣬�����˿���ռ��¼��"����վ" mCSCIdxUnused()
        Set mCSC(idxCSC) = Nothing
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(1 To mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = idxCSC
    End If
End Function



