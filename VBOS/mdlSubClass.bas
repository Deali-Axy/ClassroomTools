Attribute VB_Name = "mdlSubClass"
Option Explicit

'------------------ 子类处理模块 --------------------------
'-----Support: 需要 clsSubClass, clsHashLK 类模块的支持 ------------------
'# 国际化：无提示字符串常量
'
'//////////////////////////////////////////////////////////////
'用此模块的函数来实现并管理一个工程中的所有子类操作：可支持对同一个窗口 _
 '多次子类化。对同一个窗口重复调用 SCCreateSubClass() 函数时只对它真正 _
 '设置窗口程序地址一次，每次都记录用户给定的自定义窗口程序地址，此地址形成 _
 '一个列表。以后将逐个调用列表中的这自定义窗口程序
'还支持防止 多次子类化时 同一个自定义窗口程序地址被两次添加到列表
'只为 自定义窗口程序地址 维护一个列表，默认窗口程序永远 _
 '保存在 mCSC(i).PrevWndProcAddr 中不变。对所有子类来说，永远 SetWindowLong _
 '把窗口程序的地址改为本模块的 SCMyWndProc 函数，由 SCMyWndProc 再调用 _
 '列表中的所有自定义窗口程序
'用 clsSubClass 对象确保程序结束时恢复子类
'//////////////////////////////////////////////////////////////

'用法：
'1. 需要把一个窗口子类化时，用 SCCreateSubClass() 函数，该函数指定 _
 '子类化后的自定义的窗口程序，例如：SCCreateSubClass(Me.hwnd, AddressOf MyProc2) _
 '若对同一个窗口多次子类，重复调用本函数即可，但 hwnd 参数每次必须一致；
'注意：所有自定义窗口程序都不必有“调用默认窗口程序”的语句，但必须 _
 '有返回值语句，return gc_lngEventsGenDefautRet，否则若不是返回该值 _
 '则默认的窗口程序不能被调用
'2. 使用 SCGetPreWndProcAddr() 获得已经子类化的窗口的 原始默认窗口程序的地址 _
 '可据此，由主调程序调用默认窗口程序
'3. 恢复一个窗口子类化时，用 SCRestoreSubClassWhole() 函数，用返回值判断是否 _
 恢复成功。若没有恢复成功，本程序也不再为它保留 mCSC() 空间（为的是 _
 可以释放对象），由本模块的 cHashUnRestPreprocs 记录没恢复成功的信息 _
 '注意：调用该函数对窗口 hwnd 的所有的子类化都将被消除
'4. 要恢复工程中所有窗口的子类化，用 SCRestoreSubClassAllWnds() 函数，未恢复 _
 成功的窗口信息可用 SCGetUnSucRestCount() 和 SCGetUnSucRestWins() 函数获得
'
'
'程序中已子类化的所有窗口信息用本模块的 mCSC() 记录，该数组中没使用的 _
 '空余空间下标记录在 mCSCIdxUnused() 中，由本模块自动管理。执行 _
 'SCRestoreSubClassWhole 后 mCSC() 的空余元素被记录在 mCSCIdxUnused() 中， _
 '该 mCSC() 元素只占一个下标，其对象自动 Set=Nothing 不占内存。 _
 '下次新建子类需要空间时首先使用 mCSCIdxUnused() 中记录的下标空间。
'如果调用了一次 SCRestoreSubClassAllWnds() 函数，则 mCSC() 又从0开始计数, _
 ' mCSCIdxUnused() 也被清空
'---------------------------------------------------------------
Public Const gc_lngEventsGenDefautRet As Long = 1147483647    '此值=API建立控件使用的 gc_APICEventsGenDefautRet

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const GWL_WNDPROC = (-4)

Private mCSC() As clsSubClass, mCSCCount As Long
Attribute mCSCCount.VB_VarUserMemId = 1073741824
Private mCSCIdxUnused() As Long, mCSCIdxUnusedCount As Long    '记录 mCSC() 中没使用的空间下标即被恢复子类化后的空间下标，以便以后再新建子类时使用该下标元素
Attribute mCSCIdxUnused.VB_VarUserMemId = 1073741826
Attribute mCSCIdxUnusedCount.VB_VarUserMemId = 1073741826
Private cHashCSCIdxes As New clsHashLK    'Data=mCSC() 的下标；key= hwnd。通过此也在模块中记录了哪些窗口被子类化了，没在这里的窗口认为没被子类化（也还要看 cHashUnRestPreprocs 中是不是属于被子类后恢复失败的情况）
Attribute cHashCSCIdxes.VB_VarUserMemId = 1073741828
Private cHashUnRestPreprocs As New clsHashLK    '没有恢复成功的被子类的窗口信息，key=没有恢复成功的 hwnd；Data=默认窗口程序地址
Attribute cHashUnRestPreprocs.VB_VarUserMemId = 1073741829

Dim mIdx As Long, mR As Long, mRet As Long, mAddr As Long, mI As Long    '变量，避免每次执行函数都重新定义
Attribute mIdx.VB_VarUserMemId = 1073741830
Attribute mR.VB_VarUserMemId = 1073741830
Attribute mRet.VB_VarUserMemId = 1073741830
Attribute mAddr.VB_VarUserMemId = 1073741830
Attribute mI.VB_VarUserMemId = 1073741830

Public Function SCMyWndProc(ByVal hwnd As Long, _
                            ByVal Msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long
'所有子类化都被设置为此窗口程序，所有子类化都调用此窗口程序
'本程序将逐一调用 在 mCSC() 中的对象中记录的所有“自定义窗口程序”， _
  '以支持对同一窗口的多次子类化
'所有自定义窗口程序中，有一个不返回 gc_lngEventsGenDefautRet 本函数就不 _
  '会调用默认的窗口程序，但无论如何会调用其他的自定义窗口程序；本函数返 _
  '回值为最后一个不返回 gc_lngEventsGenDefautRet 的自定义窗口程序的返回值


'-------- 找到子类化的 mCSC() 的下标 存入： mIdx --------
    mIdx = cHashCSCIdxes.Item(hwnd, False)
    If mIdx = 0 Then SCMyWndProc = 0: Exit Function

    '-------- 从 mCSC(mIdx) 获得所有 自定义窗口程序，并逐一调用 --------
    '先设置 mRet = gc_lngEventsGenDefautRet，逐一调用窗口程序时，若有 _
     '窗口程序返回值 <> gc_lngEventsGenDefautRet，就会设置 mRet=此返回 _
     '值；若以后还有 返回值 <> gc_lngEventsGenDefautRet 的，就会更新 _
     'mRet。最后函数返回 mRet(若 mRet <> gc_lngEventsGenDefautRet 了)
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

    '-------- 返回值 --------
    If mRet = gc_lngEventsGenDefautRet And (Not mCSC(mIdx) Is Nothing) Then
        'mRet＝gc_lngEventsGenDefautRet 说明所有自定义窗口程序都返回的 _
         '是 gc_lngEventsGenDefautRet，就调用默认窗口程序
        SCMyWndProc = CallWindowProc(mCSC(mIdx).PrevWndProcAddr, _
                                     hwnd, Msg, wParam, ByVal lParam)
    Else
        'mRet<>gc_lngEventsGenDefautRet 说明有自定义窗口程序 返回 _
         '<>gc_lngEventsGenDefautRet 的，返回最后一个 返回值 _
         '不为 gc_lngEventsGenDefautRet 的自定义窗口程序的返回值
        SCMyWndProc = mRet
    End If
End Function


Public Function SCCreateSubClass(ByVal hwnd As Long, _
                                 ByVal addrMyWndProc As Long) As Boolean
'子类化窗口 hwnd，并记录这个子类化
'addrMyWndProc 指定子类化后的自定义的窗口程序，=0则使用本模块 _
  的 SCMyWndProc。例如：SCCreateSubClass(Me.hwnd, AddressOf MyProc2)
'若对同一个窗口多次子类，重复调用本函数即可，但 hwnd 参数每次必须一致

    Dim blCannotRest As Boolean
    '通过 cHashCSCIdxes 看是否已经被子类化过（是否是重复子类化）
    If SCIsHwndSubClassed(hwnd, blCannotRest) Then

        '//////////// 该窗口已经被子类化过了。但是是属于真被子类化过了， _
         '还是由于子类化后又恢复子类时恢复失败的情况？ ////////////
        If blCannotRest Then
            '//// 是子类化后又恢复子类时恢复失败的情况,无法再将此窗口子类化,返回 false

            SCCreateSubClass = False
            Exit Function

        Else
            '//// 是属于真被子类化过了

            '找到以前子类化时使用的类模块对象（找到 mCSC() 数组元素的下标即可）
            Dim lngIdx As Long
            lngIdx = cHashCSCIdxes.Item(hwnd, False)

            '以前被子类化过，应该能够找到 mCSC() 的下标，可找不到，就出错退出
            If lngIdx <= 0 Then SCCreateSubClass = False: Exit Function

            '用下面语句建立新的子类，由于已被子类化过，CreateSubClass 方法只是 _
             '添加一个"自定义窗口程序的地址" addrMyWndProc 而已，不会重复子类化
            SCCreateSubClass = mCSC(lngIdx).CreateSubClass(hwnd, addrMyWndProc)

            Exit Function

        End If

    Else
        '//////////// 该窗口没被子类化过，新建子类化 ////////////
        SCCreateSubClass = CreateNewFirstOneSubClass(hwnd, addrMyWndProc)
    End If

End Function

Public Function SCRestoreSubClassOne(ByVal hwnd As Long, _
                                     ByVal addrUserProc As Long) As Boolean
'恢复子类化,仅恢复 addrUserProc 窗口程序的子类，即从 自定义窗口程序 列表中删除该项
'若删除后 自定义窗口程序 列表为空，则彻底恢复该窗口的子类
'若彻底恢复该窗口的子类，无论是否恢复成功都会消除 cHashCSCIdxes.Remove _
  '记录；但若未恢复成功，信息会再被记录到 cHashUnRestPreprocs 中

    Dim lngIdx As Long, blRet As Boolean

    lngIdx = cHashCSCIdxes.Item(hwnd, False)
    If lngIdx = 0 Then SCRestoreSubClassOne = False: Exit Function
    If mCSC(lngIdx) Is Nothing Then SCRestoreSubClassOne = False: Exit Function

    '恢复单个
    '下面函数返回后，if 列表不为空，就什么都不做，继续使用；否则，列表为空， _
     '判断下面函数返回值，返回值表示了下面函数自动调用 UnSubclassWhole 的成功与否
    blRet = mCSC(lngIdx).UnSubclassOne(addrUserProc)

    '判断列表是否为空
    If mCSC(lngIdx).UserProcAddrsCount = 0 Then
        If blRet Then
            SCRestoreSubClassOne = True
        Else
            '类调用了 UnSubclassOne 不成功， _
             把此不成功信息添加到 cHashUnRestPreprocs
            cHashUnRestPreprocs.Add _
                    mCSC(lngIdx).PrevWndProcAddr, mCSC(lngIdx).hwnd
            SCRestoreSubClassOne = False
        End If

        '释放 mCSC(lngIdx) 空间，并将此空余空间记录到"回收站" mCSCIdxUnused()
        Set mCSC(lngIdx) = Nothing
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(1 To mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = lngIdx

        '删除 cHashCSCIdxes 的对应项
        cHashCSCIdxes.Remove hwnd, False
    Else
        '什么都不必做，但仍用返回值表示是否删除单个成功
        SCRestoreSubClassOne = blRet
    End If
End Function

Public Function SCRestoreSubClassWhole(ByVal hwnd As Long) As Boolean
'恢复子类化，对窗口 hwnd 的所有的子类化都将被消除
'无论是否恢复成功都会消除 cHashCSCIdxes.Remove 记录
'但若未恢复成功，信息会再被保存在 cHashUnRestPreprocs 中
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
'根据 hwnd 找到它的默认窗口程序的地址，出错返回 0
'若 hwnd 未被子类处理，也返回 0

    Dim lngCSCIdx As Long

    '-------- 看 hwnd 在 cHashCSCIdxes 中有无记录，分别处理 --------
    If cHashCSCIdxes.IsKeyExist(hwnd) Then
        '//////// hwnd 在 cHashCSCIdxes 中有记录，直接获取并返回 ////////
        lngCSCIdx = cHashCSCIdxes.Item(hwnd)
        If mCSC(lngCSCIdx) Is Nothing Then
            'mCSC(lngCSCIdx) 对象已被销毁
            SCGetPreWndProcAddr = 0
        Else
            If mCSC(lngCSCIdx).hwnd = hwnd Then    '再验证一下 hwnd 是否相符
                SCGetPreWndProcAddr = mCSC(lngCSCIdx).PrevWndProcAddr
            Else
                SCGetPreWndProcAddr = 0
            End If
        End If
        Exit Function
    Else
        '//////// hwnd 在 cHashCSCIdxes 中没有记录，看在 cHashUnRestPreprocs _
         '中有没有记录，若有便是属于没有恢复成功的 ////////
        If cHashUnRestPreprocs.IsKeyExist(hwnd) Then
            '//////// 属于没有恢复成功的 ////////
            SCGetPreWndProcAddr = cHashUnRestPreprocs.Item(hwnd)
        Else
            '//////// 不是属于没有恢复成功的，这个 hwnd 的子类化没有被记录， _
             '无法确定它默认窗口程序的地址，返回0 ////////
            SCGetPreWndProcAddr = 0
        End If
        Exit Function
    End If
End Function

Public Function SCIsHwndSubClassed(ByVal hwnd As Long, _
                                   Optional ByRef blRetCannotRestore As Boolean) As Boolean
'判断句柄为 hwnd 的窗口是否被子类处理过：返回 True/False 表示是否被子类处理过
'如果返回 True，还需要查看 blRetCannotRestore 参数的返回值 _
  'blRetCannotRestore 返回是否该窗口 hwnd 为"不可恢复子类"的窗口：如果该窗口 _
  '以前被子类处理过，但恢复子类时恢复失败，就不能再将该窗口做新的子类处理， _
  '否则无异于雪上加霜、错上加错。此时 blRetCannotRestore 返回 True
'若函数返回 false，则参数 blRetCannotRestore 的返回值无意义
    If cHashCSCIdxes.IsKeyExist(hwnd) Then
        '//////// hwnd 在 cHashCSCIdxes 中有记录，那么必定已经子类处理了， _
         函数返回 true，blRetCannotRestore 返回 false ////////
        SCIsHwndSubClassed = True
        blRetCannotRestore = False
    Else
        '//////// hwnd 在 cHashCSCIdxes 中没有记录，看是否属于恢复 _
         '失败的情况 ////////
        If cHashUnRestPreprocs.IsKeyExist(hwnd) Then
            '//////// hwnd 属于恢复失败的情况，那么也是已经被子类处理了 _
             '函数返回 true，blRetCannotRestore 返回 true ////////
            SCIsHwndSubClassed = True
            blRetCannotRestore = True
        Else
            '//////// hwnd 不属于恢复失败的情况，确实该窗口目前还没有被子类处理 ////////
            SCIsHwndSubClassed = False
            'blRetCannotRestore 返回值无意义
        End If
    End If
End Function

Public Function SCRestoreSubClassAllWnds() As Boolean
'清除工程中的所有“子类化”
'返回 false 表有些“子类化”没有恢复成功；但无论如何将清除 _
  mCSCCount、mCSC()、cHashCSCIdxes
'本过程将没有恢复成功的窗口信息存入 cHashUnRestPreprocs
'本过程还将再试图恢复一次 cHashUnRestPreprocs 中记录的未恢复成功的窗口
'无论是 mSCS() 中的还是 cHashUnRestPreprocs 中的，只要有没恢复成功 _
  的，函数就返回 false

    Dim i As Long
    Dim blRet As Boolean

    '------------ 先设置默认返回值是 true，如果下面有失败的 _
     '再改为 false ------------
    blRet = True

    '------------ 恢复所有 mCSC() 中的元素 ------------
    If mCSCCount > 0 Then
        For i = 1 To mCSCCount
            If Not UnSubclassOneCSCWhole(i) Then blRet = False
        Next i
        Erase mCSC
        mCSCCount = 0
        cHashCSCIdxes.Clear
    End If

    '------------ 恢复所有 cHashUnRestPreprocs 中的元素 ------------
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

    '------------ 清空“回收站”mCSCIdxUnused() ------------
    Erase mCSCIdxUnused
    mCSCIdxUnusedCount = 0

    '------------ 返回 ------------
    SCRestoreSubClassAllWnds = blRet
End Function

Public Function SCGetUnSucRestCount() As Long
'返回没有恢复子类成功的窗口的个数
'没有恢复成功的窗口是由 SCRestoreSubClassAllWnds 过程记录的
    SCGetUnSucRestCount = cHashUnRestPreprocs.Count
End Function

Public Function SCGetUnSucRestWins(retHwnds() As Long, retPrevWnds() As Long) As Long
'返回没有恢复子类成功的窗口的信息(从参数返回 hwnd 和 PrevWndProc)
'函数返回 没有子类恢复成功的窗口的个数

    Dim k As Long, b As Boolean, ret As Long
    Dim i As Long
    If cHashUnRestPreprocs.Count > 0 Then
        ReDim retHwnds(1 To cHashUnRestPreprocs.Count)
        ReDim retPrevWnds(1 To cHashUnRestPreprocs.Count)
    End If

    '遍历哈希表 cHashUnRestPreprocs
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
'新建并记录一个子类（用新分配的一个 mCSC() 对象来记录，记 hwnd 和它的默认 _
  窗口程序的地址，默认窗口程序的地址的获得由对象来完成）
'本函数是第一次子类时调用的，如果是对同一个窗口重复子类，不要调用本函 _
  '数，而要获得以前子类的 mCSC() 对象并调用它的 CreateSubClass 方法
'主调程序需先判断 cHashCSCIdxes 中有无 hwnd，来确定是否是第一次子类
'cHashCSCIdxes 中没有（即第一次子类）时才可调用本函数
'返回使用的 mCSC() 的下标，出错返回0

    Dim lngAddrCsc As Long
    Dim idxCSC As Long

    '-------- 新分配一个 mCSC() 对象，下标存入： idxCSC --------
    idxCSC = 0
    If mCSCIdxUnusedCount > 0 Then
        '//// 使用"回收站"中空余空间下标（使用"回收站"的最后一个元素） ////
        idxCSC = mCSCIdxUnused(mCSCIdxUnusedCount)
        '空余空间删除此记录
        mCSCIdxUnusedCount = mCSCIdxUnusedCount - 1
        If mCSCIdxUnusedCount > 0 _
           Then ReDim Preserve mCSCIdxUnused(mCSCIdxUnusedCount)
    Else
        '//// "回收站"中没有空余空间，扩增 mCSC() ////
        mCSCCount = mCSCCount + 1
        ReDim Preserve mCSC(1 To mCSCCount)
        idxCSC = mCSCCount
    End If

    '-------- 建立对象 --------
    Set mCSC(idxCSC) = New clsSubClass

    '-------- 用对象 mCSC(idxCSC) 建立子类并 _
     '记录 hwnd、默认窗口函数的地址 和 自定义窗口程序的地址 --------
    '此时对于 mCSC(idxCSC) 对象来说，CreateSubClass 方法是被第一次调用
    lngAddrCsc = mCSC(idxCSC).CreateSubClass(hwnd, addrUserProc)
    If lngAddrCsc = 0 Then GoTo errH

    '-------- 记录"子类处理"到模块级哈希表：cHashCSCIdxes --------
    cHashCSCIdxes.Add idxCSC, hwnd

    '-------- 返回使用的 mCSC() 的下标（>=1） --------
    CreateNewFirstOneSubClass = idxCSC
    Exit Function
errH:
    If idxCSC Then  'If idxCSC <> 0 Then
        Set mCSC(idxCSC) = Nothing

        '这个空间已经被开辟了或已经使用了"回收站"中的一个空间， _
         无论如何把它记录到"回收站" mCSCIdxUnused()
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = idxCSC
    End If
    '返回 0 表失败
    CreateNewFirstOneSubClass = 0
End Function

Private Function UnSubclassOneCSCWhole(ByVal idxCSC As Long) As Boolean
'取消子类化 mCSC(idxCSC)，对窗口的所有的子类化都将被消除
'返回是否取消成功，若空间已无效(mCSC(idxCSC)=nothing) 也返回 true
'如不成功则将此子类化的信息添加到 cHashUnRestPreprocs
'但无论如何都释放 mCSC(idxCSC) 空间
'本函数并不影响 cHashCSCIdxes，需要主调程序调用 cHashCSCIdxes.Remove
    If mCSC(idxCSC) Is Nothing Then
        UnSubclassOneCSCWhole = True
    Else
        If mCSC(idxCSC).UnSubclassWhole Then
            UnSubclassOneCSCWhole = True
        Else
            '恢复子类化不成功，也释放 mCSC(idxCSC) 空间，但把 _
             '不成功信息添加到 cHashUnRestPreprocs
            cHashUnRestPreprocs.Add _
                    mCSC(idxCSC).PrevWndProcAddr, mCSC(idxCSC).hwnd
            UnSubclassOneCSCWhole = False
        End If

        '释放 mCSC(idxCSC) 空间，并将此空余空间记录到"回收站" mCSCIdxUnused()
        Set mCSC(idxCSC) = Nothing
        mCSCIdxUnusedCount = mCSCIdxUnusedCount + 1
        ReDim Preserve mCSCIdxUnused(1 To mCSCIdxUnusedCount)
        mCSCIdxUnused(mCSCIdxUnusedCount) = idxCSC
    End If
End Function



