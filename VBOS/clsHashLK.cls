VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHashLK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------- ��ϣ����ģ�� -----------------------------
'# ���ʻ�������ʾ�ַ�������

'============================= �÷���=================================
'�� collection ���ƣ��� Key ������ Long ���͵�
'Data ֻ�ܴ�� Long ���͵����ݣ������� DataLong, DataString �������ݣ�Ҳ�� _
 '���2�� Long ���͵����ݣ�Ҳ��ͬʱ���һ���ַ���������
'Item ������ Remove �������˲������ô˲����ɾ������û���ҵ�Ԫ���Ƿ񱨴�
'������ IsKeyExist ����ֱ���ж�һ�� Long ���͵� Key �Ƿ����
'������ Clear ���������൱����ɾ������Ԫ�أ��������µĶ�������������

'----- ʹ�ñ�������һ������������ ------
'�ȵ��� StartTraversal
'Ȼ���� NextItem��NextData��NextDataLong��NextDataString �� NextKey �����һ�� _
 Ԫ�ص����ݣ������Щ�����Ĳ������� true ���ѱ�����ɣ����� true ʱ _
 �ĺ�������ֵ����Ч��
'ע�⣺�������Ӱ�����ݵķ����� Add��Remove��Clear �ȣ��������µ��� _
 'StartTraversal ���±���
'=====================================================================

'============================= ������=================================
'-------lMem() ���󷽷���
'�������ݸ��� Count Ϊ lMemUsedCount+lMemUsedCount2 �� lMemCount �Ľϴ�ֵ
'1. �� Count δ���� mcExpandCountThres��������ԭ����С��2���������� _
 mcExpandCountThres ��δ������ mcExpandCountThresMax������ԭ����С�� 3/2��
'2. �� Count ���� mcExpandCountThresMax����ÿ��ֻ���� mcExpandBigPer��

'-------lMem2() ���󷽷���
'lMem2() ˳��ʹ�ã�lMemCount2 Ϊ�ѿ��ٿռ��С��lMemUsedCount2 Ϊ��ʹ�ÿռ�
'ÿ�������ռ�ֻ���� mcExpandMem2Per ��Ԫ��
'=====================================================================

Private Type Hs_DataType    '��ϣ����ÿ�����ݵ�����
    Key As Long    '������
    Data As Long    '����ֵ
    DataLong As Long    '���ӳ���������
    DataString As String    '�����ַ���������
    Used As Byte    '���÷�ı�־��=1��ʾ�ÿռ��ѱ�ʹ�ã�=0��ʾδ��ʹ�ã�
End Type



Private lMem() As Hs_DataType, lMemCount As Long, lMemUsedCount As Long    '�±��1��ʼ
Attribute lMemCount.VB_VarUserMemId = 1073938432
Attribute lMemUsedCount.VB_VarUserMemId = 1073938432
Private lMem2() As Hs_DataType, lMemCount2 As Long, lMemUsedCount2 As Long    'û���� lmem() ���ҵ����ʴ洢λ�õ�Ԫ�ش��ڴˣ��±��1��ʼ��˳��ʹ�ã�lMemCount2 Ϊ�����С��lMemUsedCount2 Ϊ��ʹ�ÿռ�
Attribute lMem2.VB_VarUserMemId = 1073938435
Attribute lMemCount2.VB_VarUserMemId = 1073938435
Attribute lMemUsedCount2.VB_VarUserMemId = 1073938435
Private mTravIdxCurr As Long    '�� NextXXX ���������ĵ�ǰ index����ֵ��ʾ lMem() �е��±꣬��ֵ��ʾ lMem2() �е��±�
Attribute mTravIdxCurr.VB_VarUserMemId = 1073938438

Private Const mcIniMemSize As Long = 10    '��ʼ lmem() ��С
Private Const mcMaxItemCount As Long = 214748364    '���Ԫ�ظ������������ֵ�� long ��ʾ�ķ�Χ֮�ڣ�
Private Const mcExpandMaxPort As Single = 0.75    '����Ԫ�ظ������� 0.75*lmemCount ʱ������ lmem() �Ŀռ�
Private Const mcExpandCountThres As Long = 10000    '���� lmem() �ռ�ʱ���� Count С�ڴ�ֵ��ÿ������ lmemCount*2���� Count ���ڴ�ֵ��ÿ������ Count+Count/2
Private Const mcExpandCountThresMax As Long = 10000000    '���� lmem() �ռ�ʱ���� Count �Ѵ��ڴ�ֵ����ÿ�β������󵽵� Count+Count/2����ֻ���� Count+mcExpandBigPer
Private Const mcExpandBigPer As Long = 1000000    '���� lmem() �ռ�ʱ���� Count �Ѵ��� mcExpandCountThresMax����ÿ�β������󵽵� Count+Count/2����ֻ���� Count+mcExpandBigPer
Private Const mcExpandMem2Per As Long = 10    'ÿ������ lMem2() �Ĵ�С
Private Const mcSeqMax As Long = 5    '˳��������ֵ

Public Function Add(ByVal Data As Long, _
                    ByVal Key As Long, _
                    Optional ByVal DataLong As Long, _
                    Optional DataString As String, _
                    Optional ByVal RaiseErrorIfNotHas As Boolean = True) As Boolean

    Dim lngIdx As Long

    '��ϣ���е����ݸ�����಻�ܳ��� mcMaxItemCount
    If lMemUsedCount + lMemUsedCount2 >= mcMaxItemCount Then
        If RaiseErrorIfNotHas Then Err.Raise 7    '�����ڴ�
        Add = False
        Exit Function
    End If

    '��ǰ��ϣ���в�������ͬ�ġ���������
    If IsKeyExist(Key) Then
        If RaiseErrorIfNotHas Then Err.Raise 5    '��Ч�Ĺ��̵��û����
        Add = False
        Exit Function
    End If

    'ͨ������ AlloMemIndex �������һ�����ÿռ���±꣺lngIdx
    lngIdx = AlloMemIndex(Key)

    If lngIdx > 0 Then
        '��õ��±�ֵΪ����ʱ��ʹ�� lMem() ����Ŀռ�
        With lMem(lngIdx)
            .Data = Data
            .DataLong = DataLong
            .DataString = DataString
            .Key = Key
            .Used = 1
        End With
        lMemUsedCount = lMemUsedCount + 1
    Else
        '��õ��±�ֵΪ����ʱ��ʹ�� lMem2() ����Ŀռ䣬_
        '�ռ��±�Ϊ lngIdx �ľ���ֵ
        With lMem2(-lngIdx)
            .Data = Data
            .DataLong = DataLong
            .DataString = DataString
            .Key = Key
            .Used = 1
        End With
        lMemUsedCount2 = lMemUsedCount2 + 1
    End If

    '��ϣ��������ָ������
    mTravIdxCurr = 0

    '�������سɹ�
    Add = True
End Function

Public Function Item(ByVal Key As Long, _
                     Optional ByVal RaiseErrorIfNotHas As Boolean = True) As Long
    Dim lngIdx As Long
    lngIdx = FindMemIndex(Key)
    If lngIdx = 0 Then
        If RaiseErrorIfNotHas Then Err.Raise 5
        Item = 0
        Exit Function
    ElseIf lngIdx > 0 Then
        Item = lMem(lngIdx).Data
    Else
        'lngIdx<0
        Item = lMem2(-lngIdx).Data
    End If
End Function

Public Function DataLong(ByVal Key As Long, _
                         Optional ByVal RaiseErrorIfNotHas As Boolean = True) As Long
    Dim lngIdx As Long
    lngIdx = FindMemIndex(Key)
    If lngIdx = 0 Then
        If RaiseErrorIfNotHas Then Err.Raise 5
        DataLong = 0
        Exit Function
    ElseIf lngIdx > 0 Then
        DataLong = lMem(lngIdx).DataLong
    Else
        'lngIdx<0
        DataLong = lMem2(-lngIdx).DataLong
    End If
End Function


Public Function DataString(ByVal Key As Long, _
                           Optional ByVal RaiseErrorIfNotHas As Boolean = True) As String
    Dim lngIdx As Long
    lngIdx = FindMemIndex(Key)
    If lngIdx = 0 Then
        If RaiseErrorIfNotHas Then Err.Raise 5
        DataString = ""
        Exit Function
    ElseIf lngIdx > 0 Then
        DataString = lMem(lngIdx).DataString
    Else
        'lngIdx<0
        DataString = lMem2(-lngIdx).DataString
    End If
End Function

Public Function Remove(ByVal Key As Long, _
                       Optional ByVal RaiseErrorIfNotHas As Boolean = True) As Boolean

    Dim lngIdx As Long

    '���� FindMemIndex ������á�����Ϊ Key ���������ڿռ���±�
    lngIdx = FindMemIndex(Key)
    If lngIdx = 0 Then
        '��ϣ����û�С�����Ϊ Key �����ݴ���
        If RaiseErrorIfNotHas Then Err.Raise 5   '��Ч�Ĺ��̵��û����
        Remove = False
        Exit Function
    ElseIf lngIdx > 0 Then
        '��ϣ���С�����Ϊ Key �������� lMem() �����У��±�Ϊ lngIdx
        With lMem(lngIdx)
            .Used = 0
            .Key = 0
        End With
        lMemUsedCount = lMemUsedCount - 1
    Else
        'lngIdx<0 ��ʾ����ϣ���С�����Ϊ Key �������� lMem2() ������ _
         '�±�Ϊ lngIdx �ľ���ֵ
        'ɾ���±�Ϊ��-lngIdx����Ԫ��
        Dim i As Long
        For i = -lngIdx To lMemUsedCount2 - 1
            lMem2(i) = lMem2(i + 1)
        Next i
        lMemUsedCount2 = lMemUsedCount2 - 1
    End If

    '��ϣ��������ָ������
    mTravIdxCurr = 0

    '�������سɹ�
    Remove = True
End Function

Public Function IsKeyExist(ByVal Key As Long) As Boolean
'�ж�һ�� Key �Ƿ��ڵ�ǰ�����д���
    Dim lngIdx As Long
    lngIdx = FindMemIndex(Key)
    IsKeyExist = (lngIdx <> 0)
End Function

Public Sub StartTraversal()
'��ʼ�� NextXXX ... ��������
    mTravIdxCurr = 1
End Sub


Public Function NextItem(Optional ByRef rDataLong As Long, _
                         Optional ByRef rDataString As String, _
                         Optional ByRef rKey As Long, _
                         Optional ByRef bRetNotValid As Boolean = False) As Long

'���� StartTraversal ���ô˹��̱��� Data (�˲�ͬ�� _
  NextData��NextDataLong��NextDataString �� NextKey�� _
  ���ɷ���������Ŀ�����ݣ�������ֻ������Ŀ��һ������)
'��������ֵΪһ����Ŀ�� Data
'bRetNotValid ���� True�����˴α����ѽ�������ʱ��������ֵҲ��Ч��
    Dim lngIdx As Long
    lngIdx = TraversalGetNextIdx
    If lngIdx > 0 Then
        With lMem(lngIdx)
            NextItem = .Data
            rDataLong = .DataLong
            rDataString = .DataString
            rKey = .Key
        End With
    ElseIf lngIdx < 0 Then
        With lMem2(-lngIdx)
            NextItem = .Data
            rDataLong = .DataLong
            rDataString = .DataString
            rKey = .Key
        End With
    Else
        bRetNotValid = True
        Exit Function
    End If
End Function

Public Function NextData(Optional ByRef bRetNotValid As Boolean = False) As Long
'���� StartTraversal ���ô˹��̱��� Data
'bRetNotValid ���� True�����˴α����ѽ�������ʱ��������ֵҲ��Ч��
    Dim lngIdx As Long
    lngIdx = TraversalGetNextIdx
    If lngIdx > 0 Then
        NextData = lMem(lngIdx).Data
    ElseIf lngIdx < 0 Then
        NextData = lMem2(-lngIdx).Data
    Else
        bRetNotValid = True
        Exit Function
    End If
End Function

Public Function NextDataLong(Optional ByRef bRetNotValid As Boolean = False) As Long
'���� StartTraversal ���ô˹��̱��� DataLong
'bRetEndNotValid ���� True�����˴α����ѽ�������ʱ��������ֵҲ��Ч��
    Dim lngIdx As Long
    lngIdx = TraversalGetNextIdx
    If lngIdx > 0 Then
        NextDataLong = lMem(lngIdx).DataLong
    ElseIf lngIdx < 0 Then
        NextDataLong = lMem2(-lngIdx).DataLong
    Else
        bRetNotValid = True
        Exit Function
    End If
End Function

Public Function NextDataString(Optional ByRef bRetNotValid As Boolean = False) As String
'���� StartTraversal ���ô˹��̱��� String
'bRetEndNotValid ���� True�����˴α����ѽ�������ʱ��������ֵҲ��Ч��
    Dim lngIdx As Long
    lngIdx = TraversalGetNextIdx
    If lngIdx > 0 Then
        NextDataString = lMem(lngIdx).DataString
    ElseIf lngIdx < 0 Then
        NextDataString = lMem2(-lngIdx).DataString
    Else
        bRetNotValid = True
        Exit Function
    End If
End Function

Public Function NextKey(Optional ByRef bRetNotValid As Boolean = False) As Long
'���� StartTraversal ���ô˹��̱��� String
'bRetEndNotValid ���� True�����˴α����ѽ�������ʱ��������ֵҲ��Ч��
    Dim lngIdx As Long
    lngIdx = TraversalGetNextIdx
    If lngIdx > 0 Then
        NextKey = lMem(lngIdx).Key
    ElseIf lngIdx < 0 Then
        NextKey = lMem2(-lngIdx).Key
    Else
        bRetNotValid = True
        Exit Function
    End If
End Function

Public Function GetDataArray(retData() As Long) As Long
'�����������ݵ� Data�����������������ݵĸ���
    Dim lngCount As Long
    Dim i As Long, j As Long
    lngCount = lMemUsedCount + lMemUsedCount2
    If lngCount <= 0 Then GetDataArray = 0: Exit Function
    ReDim retData(1 To lngCount)
    j = 1
    For i = 1 To lMemCount
        If lMem(i).Used Then
            retData(j) = lMem(i).Data
            j = j + 1
        End If
    Next i
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used Then
            retData(j) = lMem2(i).Data
            j = j + 1
        End If
    Next i
    GetDataArray = lngCount
End Function

Public Function GetDataLongArray(retDataLong() As Long) As Long
'�����������ݵ� DataLong�����������������ݵĸ���
    Dim lngCount As Long
    Dim i As Long, j As Long
    lngCount = lMemUsedCount + lMemUsedCount2
    If lngCount <= 0 Then GetDataLongArray = 0: Exit Function
    ReDim retDataLong(1 To lngCount)
    j = 1
    For i = 1 To lMemCount
        If lMem(i).Used Then
            retDataLong(j) = lMem(i).DataLong
            j = j + 1
        End If
    Next i
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used Then
            retDataLong(j) = lMem2(i).DataLong
            j = j + 1
        End If
    Next i
    GetDataLongArray = lngCount
End Function


Public Function GetDataStringArray(retDataString() As String) As Long
'�����������ݵ� DataString�����������������ݵĸ���
    Dim lngCount As Long
    Dim i As Long, j As Long
    lngCount = lMemUsedCount + lMemUsedCount2
    If lngCount <= 0 Then GetDataStringArray = 0: Exit Function
    ReDim retDataString(1 To lngCount)
    j = 1
    For i = 1 To lMemCount
        If lMem(i).Used Then
            retDataString(j) = lMem(i).DataString
            j = j + 1
        End If
    Next i
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used Then
            retDataString(j) = lMem2(i).DataString
            j = j + 1
        End If
    Next i
    GetDataStringArray = lngCount
End Function

Public Function GetKeyArray(retKeys() As Long) As Long
'�����������ݵ� Key�����������������ݵĸ���
    Dim lngCount As Long
    Dim i As Long, j As Long
    lngCount = lMemUsedCount + lMemUsedCount2
    If lngCount <= 0 Then GetKeyArray = 0: Exit Function
    ReDim retKeys(1 To lngCount)
    j = 1
    For i = 1 To lMemCount
        If lMem(i).Used Then
            retKeys(j) = lMem(i).Key
            j = j + 1
        End If
    Next i
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used Then
            retKeys(j) = lMem2(i).Key
            j = j + 1
        End If
    Next i
    GetKeyArray = lngCount
End Function

Public Sub Clear()
    Erase lMem
    Erase lMem2
    lMemCount = 0: lMemUsedCount = 0
    lMemCount2 = 0: lMemUsedCount2 = 0

    lMemCount = mcIniMemSize
    ReDim lMem(1 To lMemCount)
    lMemUsedCount = 0
    lMemCount2 = 0
    lMemUsedCount2 = 0

    mTravIdxCurr = 0
End Sub

Public Sub AlloMem(ByVal memSize As Long)
'�����ʼ��ʱֻ������ mcIniMemSize ��С�� lMem()���Ժ���ʹ�����Զ��� _
  �󣬵�������֪���ж�󣬿������ñ����������㹻�������Ժ󲻶� _
  �Զ������ʱ��ע����ʱҪ��Ԥ�õ�Ԫ�ظ����ඨ��һЩ���������ռ� _
  ʱ���ռ��ͻ���໹���Զ�����
'�˺���Ҳ�����ڽض� lMem() ����û��ʹ�õĿռ�
'ע��memSize <= lMemUsedCount ʱ���ܾ����¶��壬��ȷ�����ݲ��ᶪʧ

    If memSize <= lMemUsedCount Or memSize > mcMaxItemCount Then Exit Sub

    Dim lngPreMemCount As Long
    lngPreMemCount = lMemCount
    lMemCount = memSize
    ReDim Preserve lMem(1 To lMemCount)

    ReLocaMem lngPreMemCount
End Sub

Private Sub ExpandMem()
    Dim lngCount As Long, lngPreMemCount As Long

    '�����ϣ���й�����������
    lngCount = lMemUsedCount + lMemUsedCount2

    'ȡ�����������������͡���ǰ lMem() �Ŀռ����������ߵĽϴ�ֵ
    If lngCount < lMemCount Then lngCount = lMemCount

    '���������ռ�֮ǰ�ġ�ԭ���� lMem() �Ŀռ�����
    lngPreMemCount = lMemCount

    If lngCount < mcExpandCountThres Then
        '��������������Ƚ��١����������ռ�Ϊԭ����2��
        lngCount = lngCount * 2
    ElseIf lngCount < mcExpandCountThresMax Then
        '������������Ѿ����е�ࡱ���������ռ�Ϊԭ����1.5��
        lngCount = lngCount * 3 / 2
    Else
        '��������������ܶࡱ�������� mcExpandBigPer ���ռ�
        lngCount = lngCount + mcExpandBigPer
    End If

    '�ض��������С
    lMemCount = lngCount
    ReDim Preserve lMem(1 To lMemCount)

    '���������С�����°��������������ݵ���λ�ã�������Ҫ����
    '�����ռ�֮ǰ�ġ�ԭ���� lMem() �Ŀռ�����
    ReLocaMem lngPreMemCount
End Sub

Private Sub ReLocaMem(ByVal preMemCountTo As Long)
'���·��� lMem(), lMem2() �ĸ�Ԫ�صĵ�ַ��lMem2() ��ĳЩԪ�ؿ��ܱ� _
  �����ƶ��� lMem
'���޸� lMemUsedCount,lMemUsedCount2, lMemCount2, lMem2() ��ֵ
'preMemCountTo ֻ���� lMem(1 to preMemCountTo)��preMemCountTo �Ժ��Ԫ�ر���Ϊ _
  δ�ã������ǣ���������ζ����� lMem2() �е�����Ԫ��

    Dim memUsed() As Hs_DataType, lngUsedCount As Long
    Dim i As Long

    '�� lMem() �е���ʹ��Ԫ�غ� lMem2() �е�����Ԫ���ȷ��� memUsed() �У� _
     '�� memUsed() ����Ϊ�㹻��ʵ�� memUsed() ֻʹ���� lngUsedCount ��Ԫ��
    ReDim memUsed(1 To preMemCountTo + lMemUsedCount2)
    lngUsedCount = 0

    'lMemUsedCount ��0���������ƶ���+1
    lMemUsedCount = 0

    '�� lMem() ����ʹ�õ�Ԫ�ش��� memUsed()
    For i = 1 To preMemCountTo
        If lMem(i).Used Then
            lngUsedCount = lngUsedCount + 1
            memUsed(lngUsedCount) = lMem(i)    '�Զ������ͱ�����ĸ�ֵ
        End If
    Next i
    '�� lMem2() �е�����Ԫ�ش��� memUsed()
    For i = 1 To lMemUsedCount2
        lngUsedCount = lngUsedCount + 1
        memUsed(lngUsedCount) = lMem2(i)    '�Զ������ͱ�����ĸ�ֵ
    Next i

    '//// ��ʱ memUsed(1 To lngUsedCount) ��Ϊ���� lMem() _
     �е���ʹ��Ԫ�� �� lMem2() �е�����Ԫ�� ////
    '��� lMem��Ҳ��������� Used ��
    ReDim lMem(1 To lMemCount)
    '��� lMem2��Ҳ��������� Used ��
    Erase lMem2
    lMemCount2 = 0
    lMemUsedCount2 = 0
    lMemUsedCount = 0
    '����� memUsed(1 To lngUsedCount) �е�Ԫ�ذ��������Сӳ���±���� lMem()
    Dim lngIdx As Long
    For i = 1 To lngUsedCount
        lngIdx = AlloMemIndex(memUsed(i).Key, False)
        If lngIdx > 0 Then
            lMem(lngIdx) = memUsed(i)
            lMem(lngIdx).Used = 1
            lMemUsedCount = lMemUsedCount + 1
        Else
            lMem2(-lngIdx) = memUsed(i)
            lMem2(-lngIdx).Used = 1
            lMemUsedCount2 = lMemUsedCount2 + 1
        End If
    Next i

    '��ϣ��������ָ������
    mTravIdxCurr = 0
End Sub


Private Function AlloMemIndex(ByVal Key As Long, _
                              Optional ByVal CanExpandMem As Boolean = True) As Long

'���� Key ����һ�� lMem() �е�δ�ô洢�ռ䣬���� lMem() �����±�
'��� Key �Ǹ�ֵ����ת��Ϊ�����������Ĵ洢�ռ�
'���ظ�ֵ�������� lMem() ���ҵ��ռ䣺����ֵ�ľ���ֵΪ lMem2() �� _
  ��һ�������±�ռ䣨lMem2()�Զ�Redim�����Դ��� lMem2()
'������ȷ������һ����ʹ�õĿռ䣬���������� lMem2() �еĿռ�
'�������������޸� lMemUsedCount2 ��ֵ���� redim lMem2()
'CanExpandMem=true ʱ�������������Զ����� lMem()�����򲻻��Զ�����

'������
'1. ���� Key Mod lMemCount + 1���� Index -> idxMod
'2. �������Ԫ���Ѿ�ʹ�ã���Key�Ƿ� < cMaxNumForSquare (sqr(2^31)=46340) _
  '�� <����ƽ�� Key��Ȼ�� mod lMemCount + 1�� _
  '�� >=���� ���� Key��Ȼ�� mod lMemCount + 1
'��������������˲� Index -> idxSq
'3. �� lMemCount-idxMod+1 -> idxModRev
'4. �� lMemCount-idxSq+1 -> idxSqRev
'5. �������ҵ��� Index ����ʹ���ˣ��� Count �Ƿ� > _
  mcExpandMaxPort*Count�����ǣ��� CanExpandMem=true�� _
  ������ lMem() �Ĵ洢�ռ䣬Ȼ��ݹ鱾���̣��ظ� 1-4 ��
'6. �� idxMod+1,+2,...,+mcSeqMax���� idxMod-1,-2,...,-mcSeqMax
'7. �� idxSq+1,+2,...,+mcSeqMax���� idxSq-1,-2,...,-mcSeqMax
'8. ��û�У����ظ�ֵ������ֵΪ lMem2() ����һ�����ÿռ䣬�Դ��� lMem2()

    Const cMaxNumForSquare As Long = 46340    'sqr(2^31)=46340

    Dim idxMod As Long, idxSq As Long
    Dim idxModRev As Long, idxSqRev As Long
    Dim lngCount As Long
    Dim keyToCalc As Long    '������ Key����ԶΪ>0����
    keyToCalc = Key
    If keyToCalc < 0 Then keyToCalc = 0 - keyToCalc
    lngCount = lMemUsedCount + lMemUsedCount2

    '1: ���� Key Mod lMemCount + 1���� Index -> idxMod
    idxMod = keyToCalc Mod lMemCount + 1
    If lMem(idxMod).Used = 0 Then AlloMemIndex = idxMod: Exit Function

    '2: �� ƽ��Key�򿪷�Key ���ٳ���ȡ�࣬�� Index -> idxSq
    If keyToCalc <= cMaxNumForSquare Then
        idxSq = (keyToCalc * keyToCalc) Mod lMemCount + 1
    Else
        idxSq = Sqr(keyToCalc) Mod lMemCount + 1
    End If
    If lMem(idxSq).Used = 0 Then AlloMemIndex = idxSq: Exit Function

    '3: ���Ե����� idxMod ���ռ� -> idxModRev
    idxModRev = lMemCount - idxMod + 1
    If lMem(idxModRev).Used = 0 Then AlloMemIndex = idxModRev: Exit Function

    '4: ���Ե����� idxSq ���ռ� -> idxSqRev
    idxSqRev = lMemCount - idxSq + 1
    If lMem(idxSqRev).Used = 0 Then AlloMemIndex = idxSqRev: Exit Function

    '5: ����ռ�ʹ�ðٷֱȳ�����ֵ�������� lMem() �� �ռ�
    If CanExpandMem And lngCount > mcExpandMaxPort * lMemCount Then
        ExpandMem    '���� lMem() �Ŀռ�
        AlloMemIndex = AlloMemIndex(Key, CanExpandMem)    '�ݹ飬�ظ�1-4��
        Exit Function
    End If

    Dim lngRetIdx As Long

    '6: �� idxMod ��ʼ��ǰ������������� mcSeqMax ���ռ�
    Dim idxMdSta As Long, idxMdEnd As Long
    idxMdSta = idxMod - mcSeqMax: idxMdEnd = idxMod + mcSeqMax
    lngRetIdx = AlloSeqIdx(idxMdSta, idxMod - 1)
    If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
    lngRetIdx = AlloSeqIdx(idxMod + 1, idxMdEnd)
    If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function

    '7: �� idxSq ��ʼ��ǰ������������� mcSeqMax ���ռ�
    Dim lngSqSta As Long, lngSqEnd As Long
    lngSqSta = idxSq - mcSeqMax: lngSqEnd = idxSq + mcSeqMax
    If lngSqSta < 1 Then lngSqSta = 1
    If lngSqEnd > lMemCount Then lngSqEnd = lMemCount

    '�� lngSqSta �� lngSqEnd �����������Ѵ� idxMdSta �� idxMdEnd �������ľͲ�������
    If lngSqEnd < idxMdSta Then
        '//// �ڶ��ε�ĩβ���ڵ�һ�ε���ߣ����ֵڶ��β��� ////
        lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
        If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
    ElseIf lngSqEnd <= idxMdEnd Then
        '=��if lngSqEnd >= idxMdSta And lngSqEnd <= idxMdEnd��
        '//// �ڶ��ε�ĩβ���ڵ�һ�ε��ڲ��������߽磩����ʱ�ٿ��ڶ��ε� _
         ��ʼ�����������������ڵ�һ�ε���ߡ�Ҳ�ڵ�һ�ε��ڲ��� ////
        If lngSqSta < idxMdSta Then
            '�ڶ��ε���ʼ�ڵ�һ�ε���ߣ��޸ĵڶ��ε�ĩβΪ��һ�ε���ʼ-1
            lngSqEnd = idxMdSta - 1
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        Else
            '�ڶ��ε���ʼҲ�ڵ�һ�ε��ڲ������ڶ��ΰ����ڵ�һ���У�����
            lngSqSta = 0: lngSqEnd = 0
        End If
    Else    '=��if lngSqEnd > idxMdEnd��
        '//// �ڶ��ε�ĩβ�ڵ�һ�ε��ұ�
        '���ݵڶ��ε���ʼ��λ�ã�������������ڶ��ε���ʼ��λ�ڵ�һ�ε� _
         �ұߣ���һ�ε��ڲ��������߽磩����һ�ε���� ////
        If lngSqSta > idxMdEnd Then
            '�ڶ��ε���ʼ��λ�ڵ�һ�ε��ұߣ����ֵڶ��β���
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        ElseIf lngSqSta >= idxMdSta Then
            '=��if lngSqSta >= idxMdSta and lngSqSta<=idxMdEnd��
            '�ڶ��ε���ʼ��λ�ڵ�һ��֮�ڣ��޸ĵڶ�����ʼ��Ϊ��һ��ĩβ+1
            lngSqSta = idxMdEnd + 1
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        Else
            '=��if lngSqSta < idxMdSta��
            '�ڶ��ε���ʼ��λ�ڵ�һ�ε���ߣ����ڶ��ΰ����˵�һ�Σ���ʱ�ڶ��η� _
             '�����Σ�lngSqSta �� ��һ����ʼ-1 �� ��һ��ĩβ+1 �� lngSqEnd

            lngRetIdx = AlloSeqIdx(lngSqSta, idxMdSta - 1)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function

            lngRetIdx = AlloSeqIdx(idxMdEnd + 1, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        End If
    End If

    '8: ���ظ�ֵ������ֵΪ lMem2() ����һ��Ԫ�أ��Դ��� lMem2()
    If lMemUsedCount2 + 1 > lMemCount2 Then
        lMemCount2 = lMemCount2 + mcExpandMem2Per
        ReDim Preserve lMem2(1 To lMemCount2)
    End If
    AlloMemIndex = -(lMemUsedCount2 + 1)
End Function

Private Function AlloSeqIdx(ByVal fromIndex As Long, ByVal toIndex As Long) As Long
'�� lMem() ��һ��ûʹ�õĿռ䣬�� fromIndex ��ʼ�� _
   �� toIndex ����
'���� lMem() ��һ��ûʹ��Ԫ�ص��±꣬û�ҵ����� 0
    Dim i As Long, fCt As Long
    If fromIndex <= 0 Then fromIndex = 1
    If toIndex > lMemCount Then toIndex = lMemCount
    For i = fromIndex To toIndex
        If lMem(i).Used = 0 Then AlloSeqIdx = i: Exit Function
    Next i
    AlloSeqIdx = 0
End Function


Private Function FindMemIndex(ByVal Key As Long) As Long

'���� Key ���� lMem() ��Ԫ�أ����� lMem() �����±� _
  '���ظ�ֵ�������� lMem() ���ҵ�λ�ã����صľ���ֵΪ lMem2() �� _
  'Ԫ���±꣬���� lMem2() ��Ҳû�ҵ����� 0
'��� Key �Ǹ�ֵ����ת��Ϊ�����������Ĵ洢�ռ�

    Const cMaxNumForSquare As Long = 46340    'sqr(2^31)=46340

    Dim idxMod As Long, idxSq As Long
    Dim idxModRev As Long, idxSqRev As Long
    Dim i As Long
    Dim keyToCalc As Long    '������ Key����ԶΪ>=0����
    keyToCalc = Key
    If keyToCalc < 0 Then keyToCalc = 0 - keyToCalc

    '1: ���� Key Mod lMemCount + 1���� Index -> idxMod
    idxMod = keyToCalc Mod lMemCount + 1
    If lMem(idxMod).Used And lMem(idxMod).Key = Key _
       Then FindMemIndex = idxMod: Exit Function

    '2: �� ƽ��Key�򿪷�Key���ٳ���ȡ�࣬�� Index -> idxSq
    If keyToCalc <= cMaxNumForSquare Then
        idxSq = (keyToCalc * keyToCalc) Mod lMemCount + 1
    Else
        idxSq = Sqr(keyToCalc) Mod lMemCount + 1
    End If
    If lMem(idxSq).Used And lMem(idxSq).Key = Key _
       Then FindMemIndex = idxSq: Exit Function

    '3: ���Ե����� idxMod ���ռ� -> idxModRev
    idxModRev = lMemCount - idxMod + 1
    If lMem(idxModRev).Used And lMem(idxModRev).Key = Key _
       Then FindMemIndex = idxModRev: Exit Function

    '4: ���Ե����� idxSq ���ռ� -> idxSqRev
    idxSqRev = lMemCount - idxSq + 1
    If lMem(idxSqRev).Used And lMem(idxSqRev).Key = Key _
       Then FindMemIndex = idxSqRev: Exit Function

    Dim lngRetIdx As Long

    '6: �� idxMod ��ʼ��ǰ������������� mcSeqMax ���ռ�
    Dim idxMdSta As Long, idxMdEnd As Long
    idxMdSta = idxMod - mcSeqMax: idxMdEnd = idxMod + mcSeqMax
    lngRetIdx = FindSeqIdx(Key, idxMdSta, idxMod - 1)
    If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
    lngRetIdx = FindSeqIdx(Key, idxMod + 1, idxMdEnd)
    If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function

    '7: �� idxSq ��ʼ��ǰ������������� mcSeqMax ���ռ�
    Dim lngSqSta As Long, lngSqEnd As Long
    lngSqSta = idxSq - mcSeqMax: lngSqEnd = idxSq + mcSeqMax
    If lngSqSta < 1 Then lngSqSta = 1
    If lngSqEnd > lMemCount Then lngSqEnd = lMemCount

    '�� lngSqSta �� lngSqEnd �����������Ѵ� idxMdSta �� _
     'idxMdEnd �������ľͲ�������
    If lngSqEnd < idxMdSta Then
        '//// �ڶ��ε�ĩβ���ڵ�һ�ε���ߣ����ֵڶ��β��� /////
        lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
        If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
    ElseIf lngSqEnd <= idxMdEnd Then
        '=��if lngSqEnd >= idxMdSta And lngSqEnd <= idxMdEnd��
        '//// �ڶ��ε�ĩβ���ڵ�һ�ε��ڲ��������߽磩����ʱ�ٿ��ڶ��ε� _
         ��ʼ�����������������ڵ�һ�ε���ߡ�Ҳ�ڵ�һ�ε��ڲ��� ////
        If lngSqSta < idxMdSta Then
            '�ڶ��ε���ʼ�ڵ�һ�ε���ߣ��޸ĵڶ��ε�ĩβΪ��һ�ε���ʼ-1
            lngSqEnd = idxMdSta - 1
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        Else
            '�ڶ��ε���ʼҲ�ڵ�һ�ε��ڲ������ڶ��ΰ����ڵ�һ���У�����
            lngSqSta = 0: lngSqEnd = 0
        End If
    Else    '=��if lngSqEnd > idxMdEnd��
        '//// �ڶ��ε�ĩβ�ڵ�һ�ε��ұ�
        '���ݵڶ��ε���ʼ��λ�ã�������������ڶ��ε���ʼ��λ�ڵ�һ�ε� _
         �ұߣ���һ�ε��ڲ��������߽磩����һ�ε���� ////
        If lngSqSta > idxMdEnd Then
            '�ڶ��ε���ʼ��λ�ڵ�һ�ε��ұߣ����ֵڶ��β���
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        ElseIf lngSqSta >= idxMdSta Then
            '=��if lngSqSta >= idxMdSta and lngSqSta<=idxMdEnd��
            '�ڶ��ε���ʼ��λ�ڵ�һ��֮�ڣ��޸ĵڶ�����ʼ��Ϊ��һ��ĩβ+1
            lngSqSta = idxMdEnd + 1
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        Else
            '=��if lngSqSta < idxMdSta��
            '�ڶ��ε���ʼ��λ�ڵ�һ�ε���ߣ����ڶ��ΰ����˵�һ�Σ���ʱ�ڶ��η� _
             '�����Σ�lngSqSta �� ��һ����ʼ-1 �� ��һ��ĩβ+1 �� lngSqEnd

            lngRetIdx = FindSeqIdx(Key, lngSqSta, idxMdSta - 1)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function

            lngRetIdx = FindSeqIdx(Key, idxMdEnd + 1, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        End If
    End If

    '8: �ٲ鿴 lMem2() �е�Ԫ����û��
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used And lMem2(i).Key = Key Then FindMemIndex = -i: Exit Function
    Next i
    FindMemIndex = 0
End Function

Private Function FindSeqIdx(ByVal Key As Long, _
                            ByVal fromIndex As Long, _
                            ByVal toIndex As Long) As Long
'�� lMem() �м�ΪKey��Ԫ���±꣬�� fromIndex ��ʼ�� _
   �� toIndex ����
'���� lMem() ���ҵ������±꣨>0����û�ҵ����� 0
    Dim i As Long, fCt As Long
    If fromIndex < 1 Then fromIndex = 1
    If toIndex > lMemCount Then toIndex = lMemCount
    For i = fromIndex To toIndex
        If lMem(i).Used And lMem(i).Key = Key Then FindSeqIdx = i: Exit Function
    Next i
    FindSeqIdx = 0
End Function

Private Function TraversalGetNextIdx() As Long
'�� NextXXX ��������ʱ��������һ����Next���� lMem()�±꣨����ֵ>0���� _
   '�� lMem2() ���±꣨����ֵ<0�������ѱ�������������ֵ=0��

    Dim lngRetIdx As Long

    If mTravIdxCurr > lMemCount Or _
       -mTravIdxCurr > lMemCount2 Or _
       mTravIdxCurr = 0 Then lngRetIdx = 0: Exit Function

    If mTravIdxCurr > 0 Then
        '//////////// �� lMem() ���� ////////////
        Do Until lMem(mTravIdxCurr).Used
            mTravIdxCurr = mTravIdxCurr + 1
            If mTravIdxCurr > lMemCount Then Exit Do
        Loop
        If mTravIdxCurr > lMemCount Then
            '//// �ѱ������������� lMem2() �л������ݼ������� lMem2() ////
            If lMemCount2 > 0 Then
                '�����´α��� lMem2() �����ݵ��±�ĸ���
                mTravIdxCurr = -1
                '\\\\ ִ������� if mTravIdxCurr < 0 Then \\\\
            Else
                '���ؽ���
                lngRetIdx = 0
                TraversalGetNextIdx = lngRetIdx
                Exit Function
            End If
        Else
            '//// ���� mTravIdxCurr ////
            lngRetIdx = mTravIdxCurr
            '�����´α���ָ�� ָ����һ��λ�ã����� lMem() ����һ���� _
             ���� lMem2() ����ʼ��
            mTravIdxCurr = mTravIdxCurr + 1
            If mTravIdxCurr > lMemCount Then If lMemCount2 > 0 Then mTravIdxCurr = -1
            TraversalGetNextIdx = lngRetIdx
            Exit Function
        End If
    End If

    If mTravIdxCurr < 0 Then
        '//////////// �� lMem2() ���� ////////////
        Do Until lMem2(-mTravIdxCurr).Used
            mTravIdxCurr = mTravIdxCurr - 1
            If -mTravIdxCurr > lMemCount2 Then Exit Do
        Loop
        If -mTravIdxCurr > lMemCount2 Then
            '//// �ѱ������� ////
            '���ؽ���
            lngRetIdx = 0
        Else
            '���ظ�ֵ�� mTravIdxCurr
            lngRetIdx = mTravIdxCurr
            '���� mTravIdxCurr ��ָ��
            mTravIdxCurr = mTravIdxCurr - 1
        End If
        TraversalGetNextIdx = lngRetIdx
    End If
End Function






Private Sub Class_Initialize()
    lMemCount = mcIniMemSize
    ReDim lMem(1 To lMemCount)
    lMemUsedCount = 0
    lMemCount2 = 0
    lMemUsedCount2 = 0
End Sub

Private Sub Class_Terminate()
    Erase lMem
    Erase lMem2
    lMemCount = 0: lMemUsedCount = 0
    lMemCount2 = 0: lMemUsedCount2 = 0
End Sub





Public Property Get Count() As Long
    Count = lMemUsedCount + lMemUsedCount2
End Property

