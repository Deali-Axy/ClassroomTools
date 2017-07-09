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

'-------------------------- 哈希表类模块 -----------------------------
'# 国际化：无提示字符串常量

'============================= 用法：=================================
'与 collection 相似，但 Key 必须是 Long 类型的
'Data 只能存放 Long 类型的数据，但多了 DataLong, DataString 两项内容，也能 _
 '存放2个 Long 类型的数据，也能同时存放一个字符串的数据
'Item 方法和 Remove 方法多了参数，用此参数可决定如果没有找到元素是否报错
'还多了 IsKeyExist 方法直接判断一个 Long 类型的 Key 是否存在
'还多了 Clear 方法，此相当于是删除所有元素，可用作新的对象处理其他数据

'----- 使用本对象还有一个遍历方法： ------
'先调用 StartTraversal
'然后用 NextItem、NextData、NextDataLong、NextDataString 和 NextKey 获得下一个 _
 元素的内容，如果这些函数的参数返回 true 表已遍历完成（返回 true 时 _
 的函数返回值亦无效）
'注意：如调用了影响数据的方法如 Add，Remove，Clear 等，必须重新调用 _
 'StartTraversal 重新遍历
'=====================================================================

'============================= 技术：=================================
'-------lMem() 扩大方法：
'共有数据个数 Count 为 lMemUsedCount+lMemUsedCount2 和 lMemCount 的较大值
'1. 若 Count 未超过 mcExpandCountThres，则扩大到原来大小的2倍；若超过 _
 mcExpandCountThres 但未来超过 mcExpandCountThresMax，扩大到原来大小的 3/2；
'2. 若 Count 超过 mcExpandCountThresMax，则每次只扩大 mcExpandBigPer。

'-------lMem2() 扩大方法：
'lMem2() 顺序使用，lMemCount2 为已开辟空间大小，lMemUsedCount2 为已使用空间
'每次扩增空间只扩增 mcExpandMem2Per 个元素
'=====================================================================

Private Type Hs_DataType    '哈希表中每个数据的类型
    Key As Long    '“键”
    Data As Long    '数据值
    DataLong As Long    '附加长整型数据
    DataString As String    '附加字符串型数据
    Used As Byte    '是用否的标志（=1表示该空间已被使用，=0表示未被使用）
End Type



Private lMem() As Hs_DataType, lMemCount As Long, lMemUsedCount As Long    '下标从1开始
Attribute lMemCount.VB_VarUserMemId = 1073938432
Attribute lMemUsedCount.VB_VarUserMemId = 1073938432
Private lMem2() As Hs_DataType, lMemCount2 As Long, lMemUsedCount2 As Long    '没有在 lmem() 中找到合适存储位置的元素存于此，下标从1开始，顺序使用，lMemCount2 为缓存大小，lMemUsedCount2 为已使用空间
Attribute lMem2.VB_VarUserMemId = 1073938435
Attribute lMemCount2.VB_VarUserMemId = 1073938435
Attribute lMemUsedCount2.VB_VarUserMemId = 1073938435
Private mTravIdxCurr As Long    '用 NextXXX 方法遍历的当前 index，正值表示 lMem() 中的下标，负值表示 lMem2() 中的下标
Attribute mTravIdxCurr.VB_VarUserMemId = 1073938438

Private Const mcIniMemSize As Long = 10    '初始 lmem() 大小
Private Const mcMaxItemCount As Long = 214748364    '最多元素个数（可扩大此值到 long 表示的范围之内）
Private Const mcExpandMaxPort As Single = 0.75    '已有元素个数大于 0.75*lmemCount 时就扩大 lmem() 的空间
Private Const mcExpandCountThres As Long = 10000    '扩大 lmem() 空间时，若 Count 小于此值则每次扩大到 lmemCount*2；若 Count 大于此值则每次扩大到 Count+Count/2
Private Const mcExpandCountThresMax As Long = 10000000    '扩大 lmem() 空间时，若 Count 已大于此值，则每次不再扩大到到 Count+Count/2，而只扩大到 Count+mcExpandBigPer
Private Const mcExpandBigPer As Long = 1000000    '扩大 lmem() 空间时，若 Count 已大于 mcExpandCountThresMax，则每次不再扩大到到 Count+Count/2，而只扩大到 Count+mcExpandBigPer
Private Const mcExpandMem2Per As Long = 10    '每次扩大 lMem2() 的大小
Private Const mcSeqMax As Long = 5    '顺序检索最大值

Public Function Add(ByVal Data As Long, _
                    ByVal Key As Long, _
                    Optional ByVal DataLong As Long, _
                    Optional DataString As String, _
                    Optional ByVal RaiseErrorIfNotHas As Boolean = True) As Boolean

    Dim lngIdx As Long

    '哈希表中的数据个数最多不能超过 mcMaxItemCount
    If lMemUsedCount + lMemUsedCount2 >= mcMaxItemCount Then
        If RaiseErrorIfNotHas Then Err.Raise 7    '超出内存
        Add = False
        Exit Function
    End If

    '当前哈希表中不能有相同的“键”存在
    If IsKeyExist(Key) Then
        If RaiseErrorIfNotHas Then Err.Raise 5    '无效的过程调用或参数
        Add = False
        Exit Function
    End If

    '通过调用 AlloMemIndex 函数获得一个可用空间的下标：lngIdx
    lngIdx = AlloMemIndex(Key)

    If lngIdx > 0 Then
        '获得的下标值为正数时，使用 lMem() 数组的空间
        With lMem(lngIdx)
            .Data = Data
            .DataLong = DataLong
            .DataString = DataString
            .Key = Key
            .Used = 1
        End With
        lMemUsedCount = lMemUsedCount + 1
    Else
        '获得的下标值为负数时，使用 lMem2() 数组的空间，_
        '空间下标为 lngIdx 的绝对值
        With lMem2(-lngIdx)
            .Data = Data
            .DataLong = DataLong
            .DataString = DataString
            .Key = Key
            .Used = 1
        End With
        lMemUsedCount2 = lMemUsedCount2 + 1
    End If

    '哈希表遍历的指针重置
    mTravIdxCurr = 0

    '函数返回成功
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

    '调用 FindMemIndex 函数获得“键”为 Key 的数据所在空间的下标
    lngIdx = FindMemIndex(Key)
    If lngIdx = 0 Then
        '哈希表中没有“键”为 Key 的数据存在
        If RaiseErrorIfNotHas Then Err.Raise 5   '无效的过程调用或参数
        Remove = False
        Exit Function
    ElseIf lngIdx > 0 Then
        '哈希表中“键”为 Key 的数据在 lMem() 数组中，下标为 lngIdx
        With lMem(lngIdx)
            .Used = 0
            .Key = 0
        End With
        lMemUsedCount = lMemUsedCount - 1
    Else
        'lngIdx<0 表示：哈希表中“键”为 Key 的数据在 lMem2() 数组中 _
         '下标为 lngIdx 的绝对值
        '删除下标为“-lngIdx”的元素
        Dim i As Long
        For i = -lngIdx To lMemUsedCount2 - 1
            lMem2(i) = lMem2(i + 1)
        Next i
        lMemUsedCount2 = lMemUsedCount2 - 1
    End If

    '哈希表遍历的指针重置
    mTravIdxCurr = 0

    '函数返回成功
    Remove = True
End Function

Public Function IsKeyExist(ByVal Key As Long) As Boolean
'判断一个 Key 是否在当前集合中存在
    Dim lngIdx As Long
    lngIdx = FindMemIndex(Key)
    IsKeyExist = (lngIdx <> 0)
End Function

Public Sub StartTraversal()
'开始用 NextXXX ... 方法遍历
    mTravIdxCurr = 1
End Sub


Public Function NextItem(Optional ByRef rDataLong As Long, _
                         Optional ByRef rDataString As String, _
                         Optional ByRef rKey As Long, _
                         Optional ByRef bRetNotValid As Boolean = False) As Long

'调用 StartTraversal 后，用此过程遍历 Data (此不同于 _
  NextData、NextDataLong、NextDataString 和 NextKey， _
  它可返回所有项目的内容，而不是只返回项目的一个内容)
'函数返回值为一个项目的 Data
'bRetNotValid 返回 True，表此次遍历已结束（此时函数返回值也无效）
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
'调用 StartTraversal 后，用此过程遍历 Data
'bRetNotValid 返回 True，表此次遍历已结束（此时函数返回值也无效）
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
'调用 StartTraversal 后，用此过程遍历 DataLong
'bRetEndNotValid 返回 True，表此次遍历已结束（此时函数返回值也无效）
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
'调用 StartTraversal 后，用此过程遍历 String
'bRetEndNotValid 返回 True，表此次遍历已结束（此时函数返回值也无效）
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
'调用 StartTraversal 后，用此过程遍历 String
'bRetEndNotValid 返回 True，表此次遍历已结束（此时函数返回值也无效）
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
'返回所有数据的 Data，函数返回所有数据的个数
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
'返回所有数据的 DataLong，函数返回所有数据的个数
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
'返回所有数据的 DataString，函数返回所有数据的个数
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
'返回所有数据的 Key，函数返回所有数据的个数
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
'程序初始化时只定义了 mcIniMemSize 大小的 lMem()，以后随使用随自动扩 _
  大，但若事先知道有多大，可以先用本函数定义足够大以免以后不断 _
  自动扩大费时；注意这时要比预用的元素个数多定义一些，否则分配空间 _
  时若空间冲突本类还会自动扩大
'此函数也可用于截断 lMem() 后面没有使用的空间
'注：memSize <= lMemUsedCount 时，拒绝重新定义，以确保数据不会丢失

    If memSize <= lMemUsedCount Or memSize > mcMaxItemCount Then Exit Sub

    Dim lngPreMemCount As Long
    lngPreMemCount = lMemCount
    lMemCount = memSize
    ReDim Preserve lMem(1 To lMemCount)

    ReLocaMem lngPreMemCount
End Sub

Private Sub ExpandMem()
    Dim lngCount As Long, lngPreMemCount As Long

    '计算哈希表中共有数据总数
    lngCount = lMemUsedCount + lMemUsedCount2

    '取“共有数据总数”和“当前 lMem() 的空间总数”两者的较大值
    If lngCount < lMemCount Then lngCount = lMemCount

    '保存扩增空间之前的、原来的 lMem() 的空间总数
    lngPreMemCount = lMemCount

    If lngCount < mcExpandCountThres Then
        '如果数据总数“比较少”，就扩增空间为原来的2倍
        lngCount = lngCount * 2
    ElseIf lngCount < mcExpandCountThresMax Then
        '如果数据总数已经“有点多”，就扩增空间为原来的1.5倍
        lngCount = lngCount * 3 / 2
    Else
        '如果数据总数“很多”，就扩增 mcExpandBigPer 个空间
        lngCount = lngCount + mcExpandBigPer
    End If

    '重定义数组大小
    lMemCount = lngCount
    ReDim Preserve lMem(1 To lMemCount)

    '按新数组大小，重新安排其中所有数据的新位置，参数中要传递
    '扩增空间之前的、原来的 lMem() 的空间总数
    ReLocaMem lngPreMemCount
End Sub

Private Sub ReLocaMem(ByVal preMemCountTo As Long)
'重新分配 lMem(), lMem2() 的各元素的地址，lMem2() 的某些元素可能被 _
  重新移动到 lMem
'将修改 lMemUsedCount,lMemUsedCount2, lMemCount2, lMem2() 的值
'preMemCountTo 只考虑 lMem(1 to preMemCountTo)，preMemCountTo 以后的元素被认为 _
  未用，不考虑；但无论如何都考虑 lMem2() 中的所有元素

    Dim memUsed() As Hs_DataType, lngUsedCount As Long
    Dim i As Long

    '将 lMem() 中的已使用元素和 lMem2() 中的所有元素先放入 memUsed() 中， _
     '把 memUsed() 定义为足够大，实际 memUsed() 只使用了 lngUsedCount 个元素
    ReDim memUsed(1 To preMemCountTo + lMemUsedCount2)
    lngUsedCount = 0

    'lMemUsedCount 置0，后面随移动随+1
    lMemUsedCount = 0

    '将 lMem() 中已使用的元素存入 memUsed()
    For i = 1 To preMemCountTo
        If lMem(i).Used Then
            lngUsedCount = lngUsedCount + 1
            memUsed(lngUsedCount) = lMem(i)    '自定义类型变量间的赋值
        End If
    Next i
    '将 lMem2() 中的所有元素存入 memUsed()
    For i = 1 To lMemUsedCount2
        lngUsedCount = lngUsedCount + 1
        memUsed(lngUsedCount) = lMem2(i)    '自定义类型变量间的赋值
    Next i

    '//// 此时 memUsed(1 To lngUsedCount) 中为所有 lMem() _
     中的已使用元素 和 lMem2() 中的所有元素 ////
    '清空 lMem，也清空了所有 Used 域
    ReDim lMem(1 To lMemCount)
    '清空 lMem2，也清空了所有 Used 域
    Erase lMem2
    lMemCount2 = 0
    lMemUsedCount2 = 0
    lMemUsedCount = 0
    '逐个把 memUsed(1 To lngUsedCount) 中的元素按新数组大小映射下标存入 lMem()
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

    '哈希表遍历的指针重置
    mTravIdxCurr = 0
End Sub


Private Function AlloMemIndex(ByVal Key As Long, _
                              Optional ByVal CanExpandMem As Boolean = True) As Long

'根据 Key 分配一个 lMem() 中的未用存储空间，返回 lMem() 数组下标
'如果 Key 是负值，则转换为正数计算它的存储空间
'返回负值表不能在 lMem() 中找到空间：返回值的绝对值为 lMem2() 的 _
  下一个可用下标空间（lMem2()自动Redim），以存入 lMem2()
'本函数确保返回一个可使用的空间，最差情况返回 lMem2() 中的空间
'另：本函数不修改 lMemUsedCount2 的值，但 redim lMem2()
'CanExpandMem=true 时，允许本函数自动扩大 lMem()，否则不会自动扩大

'方法：
'1. 先用 Key Mod lMemCount + 1，此 Index -> idxMod
'2. 若上面的元素已经使用，则看Key是否 < cMaxNumForSquare (sqr(2^31)=46340) _
  '若 <，则平方 Key，然后 mod lMemCount + 1； _
  '若 >=，则 开根 Key，然后 mod lMemCount + 1
'无论哪种情况，此步 Index -> idxSq
'3. 用 lMemCount-idxMod+1 -> idxModRev
'4. 用 lMemCount-idxSq+1 -> idxSqRev
'5. 若上面找到的 Index 都被使用了，则看 Count 是否 > _
  mcExpandMaxPort*Count，若是，若 CanExpandMem=true， _
  则扩大 lMem() 的存储空间，然后递归本过程，重复 1-4 步
'6. 用 idxMod+1,+2,...,+mcSeqMax；用 idxMod-1,-2,...,-mcSeqMax
'7. 用 idxSq+1,+2,...,+mcSeqMax；用 idxSq-1,-2,...,-mcSeqMax
'8. 再没有，返回负值，绝对值为 lMem2() 的下一个可用空间，以存入 lMem2()

    Const cMaxNumForSquare As Long = 46340    'sqr(2^31)=46340

    Dim idxMod As Long, idxSq As Long
    Dim idxModRev As Long, idxSqRev As Long
    Dim lngCount As Long
    Dim keyToCalc As Long    '计算用 Key，永远为>0的数
    keyToCalc = Key
    If keyToCalc < 0 Then keyToCalc = 0 - keyToCalc
    lngCount = lMemUsedCount + lMemUsedCount2

    '1: 先用 Key Mod lMemCount + 1，此 Index -> idxMod
    idxMod = keyToCalc Mod lMemCount + 1
    If lMem(idxMod).Used = 0 Then AlloMemIndex = idxMod: Exit Function

    '2: 用 平方Key或开方Key 后再除法取余，此 Index -> idxSq
    If keyToCalc <= cMaxNumForSquare Then
        idxSq = (keyToCalc * keyToCalc) Mod lMemCount + 1
    Else
        idxSq = Sqr(keyToCalc) Mod lMemCount + 1
    End If
    If lMem(idxSq).Used = 0 Then AlloMemIndex = idxSq: Exit Function

    '3: 尝试倒数第 idxMod 个空间 -> idxModRev
    idxModRev = lMemCount - idxMod + 1
    If lMem(idxModRev).Used = 0 Then AlloMemIndex = idxModRev: Exit Function

    '4: 尝试倒数第 idxSq 个空间 -> idxSqRev
    idxSqRev = lMemCount - idxSq + 1
    If lMem(idxSqRev).Used = 0 Then AlloMemIndex = idxSqRev: Exit Function

    '5: 如果空间使用百分比超过阈值，就扩大 lMem() 的 空间
    If CanExpandMem And lngCount > mcExpandMaxPort * lMemCount Then
        ExpandMem    '扩大 lMem() 的空间
        AlloMemIndex = AlloMemIndex(Key, CanExpandMem)    '递归，重复1-4步
        Exit Function
    End If

    Dim lngRetIdx As Long

    '6: 从 idxMod 开始向前、向后线性搜索 mcSeqMax 个空间
    Dim idxMdSta As Long, idxMdEnd As Long
    idxMdSta = idxMod - mcSeqMax: idxMdEnd = idxMod + mcSeqMax
    lngRetIdx = AlloSeqIdx(idxMdSta, idxMod - 1)
    If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
    lngRetIdx = AlloSeqIdx(idxMod + 1, idxMdEnd)
    If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function

    '7: 从 idxSq 开始向前、向后线性搜索 mcSeqMax 个空间
    Dim lngSqSta As Long, lngSqEnd As Long
    lngSqSta = idxSq - mcSeqMax: lngSqEnd = idxSq + mcSeqMax
    If lngSqSta < 1 Then lngSqSta = 1
    If lngSqEnd > lMemCount Then lngSqEnd = lMemCount

    '从 lngSqSta 到 lngSqEnd 搜索，上面已从 idxMdSta 到 idxMdEnd 搜索过的就不搜索了
    If lngSqEnd < idxMdSta Then
        '//// 第二段的末尾落在第一段的左边，保持第二段不变 ////
        lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
        If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
    ElseIf lngSqEnd <= idxMdEnd Then
        '=“if lngSqEnd >= idxMdSta And lngSqEnd <= idxMdEnd”
        '//// 第二段的末尾落在第一段的内部（包含边界），此时再看第二段的 _
         起始落在哪里（两种情况：在第一段的左边、也在第一段的内部） ////
        If lngSqSta < idxMdSta Then
            '第二段的起始在第一段的左边，修改第二段的末尾为第一段的起始-1
            lngSqEnd = idxMdSta - 1
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        Else
            '第二段的起始也在第一段的内部，即第二段包含在第一段中，不做
            lngSqSta = 0: lngSqEnd = 0
        End If
    Else    '=“if lngSqEnd > idxMdEnd”
        '//// 第二段的末尾在第一段的右边
        '根据第二段的起始点位置，分三种情况：第二段的起始点位于第一段的 _
         右边，第一段的内部（包含边界），第一段的左边 ////
        If lngSqSta > idxMdEnd Then
            '第二段的起始点位于第一段的右边，保持第二段不变
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        ElseIf lngSqSta >= idxMdSta Then
            '=“if lngSqSta >= idxMdSta and lngSqSta<=idxMdEnd”
            '第二段的起始点位于第一段之内，修改第二段起始点为第一段末尾+1
            lngSqSta = idxMdEnd + 1
            lngRetIdx = AlloSeqIdx(lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        Else
            '=“if lngSqSta < idxMdSta”
            '第二段的起始点位于第一段的左边，即第二段包含了第一段，此时第二段分 _
             '成两段：lngSqSta 到 第一段起始-1 和 第一段末尾+1 到 lngSqEnd

            lngRetIdx = AlloSeqIdx(lngSqSta, idxMdSta - 1)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function

            lngRetIdx = AlloSeqIdx(idxMdEnd + 1, lngSqEnd)
            If lngRetIdx > 0 Then AlloMemIndex = lngRetIdx: Exit Function
        End If
    End If

    '8: 返回负值，绝对值为 lMem2() 的下一个元素，以存入 lMem2()
    If lMemUsedCount2 + 1 > lMemCount2 Then
        lMemCount2 = lMemCount2 + mcExpandMem2Per
        ReDim Preserve lMem2(1 To lMemCount2)
    End If
    AlloMemIndex = -(lMemUsedCount2 + 1)
End Function

Private Function AlloSeqIdx(ByVal fromIndex As Long, ByVal toIndex As Long) As Long
'找 lMem() 中一个没使用的空间，从 fromIndex 开始， _
   到 toIndex 结束
'返回 lMem() 的一个没使用元素的下标，没找到返回 0
    Dim i As Long, fCt As Long
    If fromIndex <= 0 Then fromIndex = 1
    If toIndex > lMemCount Then toIndex = lMemCount
    For i = fromIndex To toIndex
        If lMem(i).Used = 0 Then AlloSeqIdx = i: Exit Function
    Next i
    AlloSeqIdx = 0
End Function


Private Function FindMemIndex(ByVal Key As Long) As Long

'根据 Key 查找 lMem() 中元素，返回 lMem() 数组下标 _
  '返回负值表不能在 lMem() 中找到位置：返回的绝对值为 lMem2() 的 _
  '元素下标，若在 lMem2() 中也没找到返回 0
'如果 Key 是负值，则转换为正数计算它的存储空间

    Const cMaxNumForSquare As Long = 46340    'sqr(2^31)=46340

    Dim idxMod As Long, idxSq As Long
    Dim idxModRev As Long, idxSqRev As Long
    Dim i As Long
    Dim keyToCalc As Long    '计算用 Key，永远为>=0的数
    keyToCalc = Key
    If keyToCalc < 0 Then keyToCalc = 0 - keyToCalc

    '1: 先用 Key Mod lMemCount + 1，此 Index -> idxMod
    idxMod = keyToCalc Mod lMemCount + 1
    If lMem(idxMod).Used And lMem(idxMod).Key = Key _
       Then FindMemIndex = idxMod: Exit Function

    '2: 用 平方Key或开方Key后再除法取余，此 Index -> idxSq
    If keyToCalc <= cMaxNumForSquare Then
        idxSq = (keyToCalc * keyToCalc) Mod lMemCount + 1
    Else
        idxSq = Sqr(keyToCalc) Mod lMemCount + 1
    End If
    If lMem(idxSq).Used And lMem(idxSq).Key = Key _
       Then FindMemIndex = idxSq: Exit Function

    '3: 尝试倒数第 idxMod 个空间 -> idxModRev
    idxModRev = lMemCount - idxMod + 1
    If lMem(idxModRev).Used And lMem(idxModRev).Key = Key _
       Then FindMemIndex = idxModRev: Exit Function

    '4: 尝试倒数第 idxSq 个空间 -> idxSqRev
    idxSqRev = lMemCount - idxSq + 1
    If lMem(idxSqRev).Used And lMem(idxSqRev).Key = Key _
       Then FindMemIndex = idxSqRev: Exit Function

    Dim lngRetIdx As Long

    '6: 从 idxMod 开始向前、向后线性搜索 mcSeqMax 个空间
    Dim idxMdSta As Long, idxMdEnd As Long
    idxMdSta = idxMod - mcSeqMax: idxMdEnd = idxMod + mcSeqMax
    lngRetIdx = FindSeqIdx(Key, idxMdSta, idxMod - 1)
    If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
    lngRetIdx = FindSeqIdx(Key, idxMod + 1, idxMdEnd)
    If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function

    '7: 从 idxSq 开始向前、向后线性搜索 mcSeqMax 个空间
    Dim lngSqSta As Long, lngSqEnd As Long
    lngSqSta = idxSq - mcSeqMax: lngSqEnd = idxSq + mcSeqMax
    If lngSqSta < 1 Then lngSqSta = 1
    If lngSqEnd > lMemCount Then lngSqEnd = lMemCount

    '从 lngSqSta 到 lngSqEnd 搜索，上面已从 idxMdSta 到 _
     'idxMdEnd 搜索过的就不搜索了
    If lngSqEnd < idxMdSta Then
        '//// 第二段的末尾落在第一段的左边，保持第二段不变 /////
        lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
        If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
    ElseIf lngSqEnd <= idxMdEnd Then
        '=“if lngSqEnd >= idxMdSta And lngSqEnd <= idxMdEnd”
        '//// 第二段的末尾落在第一段的内部（包含边界），此时再看第二段的 _
         起始落在哪里（两种情况：在第一段的左边、也在第一段的内部） ////
        If lngSqSta < idxMdSta Then
            '第二段的起始在第一段的左边，修改第二段的末尾为第一段的起始-1
            lngSqEnd = idxMdSta - 1
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        Else
            '第二段的起始也在第一段的内部，即第二段包含在第一段中，不做
            lngSqSta = 0: lngSqEnd = 0
        End If
    Else    '=“if lngSqEnd > idxMdEnd”
        '//// 第二段的末尾在第一段的右边
        '根据第二段的起始点位置，分三种情况：第二段的起始点位于第一段的 _
         右边，第一段的内部（包含边界），第一段的左边 ////
        If lngSqSta > idxMdEnd Then
            '第二段的起始点位于第一段的右边，保持第二段不变
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        ElseIf lngSqSta >= idxMdSta Then
            '=“if lngSqSta >= idxMdSta and lngSqSta<=idxMdEnd”
            '第二段的起始点位于第一段之内，修改第二段起始点为第一段末尾+1
            lngSqSta = idxMdEnd + 1
            lngRetIdx = FindSeqIdx(Key, lngSqSta, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        Else
            '=“if lngSqSta < idxMdSta”
            '第二段的起始点位于第一段的左边，即第二段包含了第一段，此时第二段分 _
             '成两段：lngSqSta 到 第一段起始-1 和 第一段末尾+1 到 lngSqEnd

            lngRetIdx = FindSeqIdx(Key, lngSqSta, idxMdSta - 1)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function

            lngRetIdx = FindSeqIdx(Key, idxMdEnd + 1, lngSqEnd)
            If lngRetIdx > 0 Then FindMemIndex = lngRetIdx: Exit Function
        End If
    End If

    '8: 再查看 lMem2() 中的元素有没有
    For i = 1 To lMemUsedCount2
        If lMem2(i).Used And lMem2(i).Key = Key Then FindMemIndex = -i: Exit Function
    Next i
    FindMemIndex = 0
End Function

Private Function FindSeqIdx(ByVal Key As Long, _
                            ByVal fromIndex As Long, _
                            ByVal toIndex As Long) As Long
'找 lMem() 中键为Key的元素下标，从 fromIndex 开始， _
   到 toIndex 结束
'返回 lMem() 的找到键的下标（>0），没找到返回 0
    Dim i As Long, fCt As Long
    If fromIndex < 1 Then fromIndex = 1
    If toIndex > lMemCount Then toIndex = lMemCount
    For i = fromIndex To toIndex
        If lMem(i).Used And lMem(i).Key = Key Then FindSeqIdx = i: Exit Function
    Next i
    FindSeqIdx = 0
End Function

Private Function TraversalGetNextIdx() As Long
'用 NextXXX 方法遍历时，返回下一个（Next）的 lMem()下标（返回值>0）， _
   '或 lMem2() 的下标（返回值<0），或已遍历结束（返回值=0）

    Dim lngRetIdx As Long

    If mTravIdxCurr > lMemCount Or _
       -mTravIdxCurr > lMemCount2 Or _
       mTravIdxCurr = 0 Then lngRetIdx = 0: Exit Function

    If mTravIdxCurr > 0 Then
        '//////////// 在 lMem() 中找 ////////////
        Do Until lMem(mTravIdxCurr).Used
            mTravIdxCurr = mTravIdxCurr + 1
            If mTravIdxCurr > lMemCount Then Exit Do
        Loop
        If mTravIdxCurr > lMemCount Then
            '//// 已遍历结束，看若 lMem2() 中还有数据继续遍历 lMem2() ////
            If lMemCount2 > 0 Then
                '设置下次遍历 lMem2() 中数据的下标的负数
                mTravIdxCurr = -1
                '\\\\ 执行下面的 if mTravIdxCurr < 0 Then \\\\
            Else
                '返回结束
                lngRetIdx = 0
                TraversalGetNextIdx = lngRetIdx
                Exit Function
            End If
        Else
            '//// 返回 mTravIdxCurr ////
            lngRetIdx = mTravIdxCurr
            '调整下次遍历指针 指向下一个位置（或是 lMem() 的下一个， _
             或是 lMem2() 的起始）
            mTravIdxCurr = mTravIdxCurr + 1
            If mTravIdxCurr > lMemCount Then If lMemCount2 > 0 Then mTravIdxCurr = -1
            TraversalGetNextIdx = lngRetIdx
            Exit Function
        End If
    End If

    If mTravIdxCurr < 0 Then
        '//////////// 在 lMem2() 中找 ////////////
        Do Until lMem2(-mTravIdxCurr).Used
            mTravIdxCurr = mTravIdxCurr - 1
            If -mTravIdxCurr > lMemCount2 Then Exit Do
        Loop
        If -mTravIdxCurr > lMemCount2 Then
            '//// 已遍历结束 ////
            '返回结束
            lngRetIdx = 0
        Else
            '返回负值的 mTravIdxCurr
            lngRetIdx = mTravIdxCurr
            '调整 mTravIdxCurr 的指针
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


