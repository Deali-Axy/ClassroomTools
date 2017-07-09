Attribute VB_Name = "mdlSysHotKeySupport"
Option Explicit

'-------------------------- 系统热键 clsSysHotKey 类模块的支持模块 ----------------------------
'需要 clsSysHotKey 类模块的支持
'需要 clsHashLK、clsStack、clsSubClass 类模块 和 mdlSubClass 标准模块的支持

'#国际化：无提示字符串常量

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_HOTKEY = &H312

'“登记”工程中所有的系统热键。Key=热键ID，Data=一个clsSysHotKey对象的地址
Private mHashSysHotkeys As New clsHashLK

'分配系统热键ID用（变量值逐渐增长，但不超过 mc_MaxHotKeyID）
Private mIDInc As Long
Private Const mc_MaxHotKeyID As Long = &H7FFF
'回收系统热键ID（使用堆栈）
Private mStackRecycleIDs As New clsStack

'多个 clsSysHotKey 对象为同一个窗口 hWnd 设置热键时，要重复对该窗口子类化 _
 '和取消子类化，且自定义窗口程序都是同一个 SHKWndProc
'为避免重复子类化，或取消尚在使用的子类化，子类化/取消子类化要由本模块统一进行
'"登记"工程中所有为设置系统热键而子类化的窗口 hWnd
Private mHashSubClsedHwnds As New clsHashLK    'Key=hWnd，Data=该窗口子类化计数，即被重复子类了几次（为0时自动取消子类化）

Public Function SHKRegOneObject(ByVal addrClsSysHotKey As Long) As Long
'“登记”一个 clsSysHotKey 对象，分配一个可以使用的系统热键ID，记录该对 _
  '象与该ID对应（在 clsSysHotKey 的 Class_Initialize 时调用）

'函数返回该系统热键 ID，失败返回 0
'addrClsSysHotKey 为一个clsSysHotKey对象的地址
'API函数要求的热键 id 范围在 0x0000～0xBFFF， _
  '但本程序规定范围在0x0001～0x7FFF（不使用0和负数）

    Dim lngNewID As Long
    '用 mIDInc 变量加1 的方式分配一个新的热键 ID：lngNewID
    mIDInc = mIDInc + 1
    If mIDInc > mc_MaxHotKeyID Then
        '//////////  mIDInc 变量增长到头 //////////
        '恢复 mIDInc 变量的值
        mIDInc = mIDInc - 1
        '如果有回收ID，使用回收ID
        If mStackRecycleIDs.IsEmpty Then
            '没有回收ID，返回失败
            SHKRegOneObject = 0
            Exit Function
        Else
            '使用回收的栈顶ID
            lngNewID = mStackRecycleIDs.PopLong
        End If
    Else
        '////////// 用 mIDInc 变量的值作为新热键 ID //////////
        lngNewID = mIDInc
    End If

    '在哈希表 mHashSysHotkeys 中记录 addrClsSysHotKey 地址和 lngNewID 对应
    If mHashSysHotkeys.Add(addrClsSysHotKey, lngNewID, 0, "", False) Then
        SHKRegOneObject = lngNewID
    Else
        SHKRegOneObject = 0
    End If
End Function

Public Function SHKUnRegOneObject(ByVal IDSysHotKey As Long) As Boolean
'取消“登记”一个 clsSysHotKey 对象，即删除 mHashSysHotkeys 中的一条 _
  '记录（在 clsSysHotKey 的 Class_Terminate 时调用）

    If mHashSysHotkeys.Remove(IDSysHotKey, False) Then
        '回收释放的热键 ID
        mStackRecycleIDs.PushLong IDSysHotKey
        '返回成功
        SHKUnRegOneObject = True
    Else
        '返回失败
        SHKUnRegOneObject = False
    End If
End Function

Public Function SHKSubClassHwnd(ByVal hwnd As Long) As Boolean
'请求将窗口 hwnd 子类处理为 SHKWndProc 自定义窗口程序
'由 clsSysHotKey 类模块调用

    If mHashSubClsedHwnds.IsKeyExist(hwnd) Then
        '该窗口已被子类化为 SHKWndProc 自定义窗口程序，不需重复子类化
        '但要增长子类化计数，更新其中的 Item 值为现有值 +1
        Dim lngCt As Long
        lngCt = mHashSubClsedHwnds.Item(hwnd, False)
        lngCt = lngCt + 1
        mHashSubClsedHwnds.Remove hwnd, False
        mHashSubClsedHwnds.Add lngCt, hwnd, 0, "", False
        SHKSubClassHwnd = True
    Else
        '该窗口还没被子类化为 SHKWndProc 自定义窗口程序
        '现在子类化
        SHKSubClassHwnd = SCCreateSubClass(hwnd, AddressOf SHKWndProc)
        'mHashSubClsedHwnds 中记录1次
        mHashSubClsedHwnds.Add 1, hwnd, 0, "", False
    End If
End Function

Public Sub SHKUnSubClassHwnd(ByVal hwnd As Long)
'请求将窗口 hwnd 取消子类处理为 SHKWndProc 自定义窗口程序
'由 clsSysHotKey 类模块调用

    If mHashSubClsedHwnds.IsKeyExist(hwnd) Then
        '该窗口已被子类化为 SHKWndProc 自定义窗口程序
        '子类化计数减1，子类化计数为0时才真正取消子类化
        Dim lngCt As Long
        lngCt = mHashSubClsedHwnds.Item(hwnd, False)
        lngCt = lngCt - 1
        mHashSubClsedHwnds.Remove hwnd, False
        If lngCt <= 0 Then
            '取消子类化
            SCRestoreSubClassOne hwnd, AddressOf SHKWndProc
        Else
            '不真正取消子类化，把减 1 的 lngCt 再存入 mHashSubClsedHwnds
            mHashSubClsedHwnds.Add lngCt, hwnd, 0, "", False
        End If
    Else
        'mHashSubClsedHwnds 中没记录该窗口子类化过，为容错，直接取消它的子类化
        SCRestoreSubClassOne hwnd, AddressOf SHKWndProc
    End If
End Sub

Public Function SHKWndProc(ByVal hwnd As Long, _
                           ByVal Msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
'被设置系统热键的窗口被子类化的自定义窗口程序


    If Msg = WM_HOTKEY Then
        '////////// 系统热键被按下，wParam 参数为热键 ID //////////

        Dim lngAddrObj As Long    '对应clsSysHotKey对象地址

        '从哈希表获得 clsSysHotKey 对象地址
        lngAddrObj = mHashSysHotkeys.Item(wParam, False)

        If lngAddrObj Then    '如果 lngAddrObj 为0表示失败或本VB程序中根本没有此热键
            '通过弱引用调用 clsSysHotKey 对象的 RaiseSysKeyPressedEvent 方法生成事件
            Dim objCls As clsSysHotKey
            CopyMemory objCls, lngAddrObj, 4
            objCls.RaiseSysKeyPressedEvent
            CopyMemory objCls, 0&, 4
        End If
    End If

    '返回值，“子类处理通用模块”会自动调用默认窗口程序处理
    SHKWndProc = gc_lngEventsGenDefautRet
End Function



