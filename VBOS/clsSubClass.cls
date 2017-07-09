VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------“子类化”记录类----------------------------
'使用本类记录一个“子类化”（记录其 hwnd， 默认窗口程序的地址 和 _
 '用户自定义窗口程序的地址）
'本类对象保存 用户自定义窗口程序的地址 的列表，而不是一个，这样可支持 _
 '对同一个窗口的多次子类化（每次子类化有不同的自定义窗口程序）
'本对象卸载时，自动恢复“子类化”（设置窗口程序地址为 默认窗口程序的地址）

'# 国际化：无提示信息字符串常量
'------------------------------------------------------------------------

'-------------------------------用法：-----------------------------------
'用 CreateSubClass 建立并记录一个子类，返回本对象的地址，出错返回0 _
 '若对同一个窗口多次子类，重复调用本函数即可，但 hwnd 参数每次必须一致； _
 '若每次 hwnd 不一致，将当作第一次子类化，自动恢复以前的 hwnd 子类处理， _
 '并建立一个新的子类。
'本对象还可防止 多次子类化时 同一个“自定义窗口程序”的地址重复添加到列表
'用 UserProcAddrsCount 获得 自定义窗口程序的 个数，然后从1循环到个数， _
 用 UserProcAddr(i) 获得 每个自定义窗口程序的地址，逐一调用自定义窗口程序。 _
 最后用 AddrPrevWndProc 获得默认窗口程序的地址，再调用默认的窗口程序。
'用 UnSubclassOne 仅恢复一个 自定义窗口程序 的子类（对同一窗口多次子类时）， _
 '只有在 自定义窗口程序 列表为空时才设置窗口程序为“默认窗口程序”，彻底恢复子类。
'用 UnSubclassWhole 彻底恢复子类，恢复窗口程序为“默认窗口程序”。本对象销毁时 _
 '自动调用本函数
'------------------------------------------------------------------------

'mUserProcAddrs() 只保存“自定义窗口程序”的地址，并没保存“默认 _
 窗口程序”的地址。“默认窗口程序”的地址由 lAddrPrevWndProc 保存

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub SetLastError Lib "kernel32" (ByVal dwErrCode As Long)
Private Declare Function GetLastError Lib "kernel32" () As Long

Private Const GWL_WNDPROC = (-4)



Private lHwnd As Long
Private lAddrPrevWndProc As Long
Private mUserProcAddrs() As Long, mUserProcAddrsCount As Long    '所有 自定义窗口程序 的地址的列表
Attribute mUserProcAddrsCount.VB_VarUserMemId = 1073938434

Public Function CreateSubClass(ByVal hwnd As Long, _
                               ByVal currUserProcAddr As Long) As Long
'用本对象建立并记录一个子类，返回本对象的地址，出错返回0
'hwnd为要子类处理的窗口的句柄；currUserProcAddr 为自定义窗口程序的地址
'记录 Hwnd 、默认窗口程序的地址、当前子类的自定义窗口程序的地址到本 _
  '对象的私有成员变量，并完成子类化的操作
'本函数会判断是否是对窗口的第一次子类化，只有第一次子类化的时 _
  '才记录 Hwnd 和 默认窗口程序的地址，也才会改变 hwnd 的窗口程序的地址， _
  '以后多次子类化只记录自定义窗口程序的地址
'若是第二次以后的子类化，hwnd 虽不再重复记录，但给出的 hwnd 参数必须 _
  '和第一次的一致，否则将删除以前的子类化记录，当作第一次子类重新记录

    If lAddrPrevWndProc <> 0 And lHwnd <> 0 Then
        '//////// 本对象已记录了一次子类化 ////////
        '看现在的 hwnd 是否是记录的 lHwnd
        If lHwnd = hwnd Then
            '//////// 现在的 hwnd 就是记录的 lHwnd，说明是对一个窗口的多次 _
             '子类化，只添加自定义窗口程序的地址到数组即可 ////////
            'AddOneUserProcAddr 函数会判断如果列表中已存在该地址则不再添加，直接返回
            AddOneUserProcAddr currUserProcAddr
            CreateSubClass = ObjPtr(Me)
            Exit Function
        Else
            '//////// 现在的 hwnd 与记录的 lHwnd 不符，说明是对一个新的 _
             '窗口子类化，删除以前的子类化，重新记录子类化 ////////
            If Not UnSubclassWhole Then GoTo errH    '删除以前的子类化
            '重新建立子类执行下面的程序
        End If
    End If

    '**************** 以下程序必是第一次子类化执行的程序 ***************
    '获取"默认窗口程序"的地址，并保存到模块级变量 lAddrPrevWndProc
    lAddrPrevWndProc = GetWindowLong(hwnd, GWL_WNDPROC)
    '如果所获取的"默认窗口程序"地址无效，就跳转 errH 处执行，拒绝子类处理该窗口
    If lAddrPrevWndProc = 0 Then GoTo errH
    '将窗口句柄保存到模块级变量 lHwnd
    lHwnd = hwnd

    '建立子类化
    SetLastError 0&
    If SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SCMyWndProc) = 0 Then
        'SetWindowLong 如果调用失败，就恢复子类处理，并跳转到 errH 处执行
        If GetLastError Then UnSubclassWhole: GoTo errH  'If GetLastError <> 0 Then
    End If

    '添加当前 自定义窗口程序地址 到数组
    AddOneUserProcAddr currUserProcAddr
    CreateSubClass = ObjPtr(Me)
    Exit Function
errH:
    CreateSubClass = 0
End Function

Public Function UnSubclassOne(addrUserProc As Long) As Boolean
'恢复一个 自定义窗口程序 的子类，只是从 自定义窗口程序 列表中删除此项
'恢复后判断若 自定义窗口程序 列表不为空，则不完全恢复子类，否则 _
  '将调用 UnSubclassWhole 完全恢复子类
'返回 false 时两种情况：if 列表为空，就是调用 UnSubclassWhole 失败 _
  'if 列表不为空，就是在列表中没有找到地址，此时应继续使用
'返回 true 时两种情况：if 列表为空，就是调用 UnSubclassWhole 成功 _
  'if 列表不为空，就是只删除一次子类成功，此时应继续使用
'总之，if 列表不为空，就什么都不做，继续使用；否则，列表为空，判断 _
  '本函数返回值，返回值表示了调用 UnSubclassWhole 的成功与否

    Dim i As Long
    '-------- 这一块先设置 UnSubclassOne 返回值，下面的程序还可再改 --------
    For i = 1 To mUserProcAddrsCount
        If mUserProcAddrs(i) = addrUserProc Then Exit For
    Next i
    If i <= mUserProcAddrsCount Then
        '找到确实有一个窗口程序＝addrUserProc
        '从 自定义窗口程序 列表中删除此项
        Dim j As Long
        For j = i To mUserProcAddrsCount - 1
            mUserProcAddrs(j) = mUserProcAddrs(j + 1)
        Next j
        mUserProcAddrsCount = mUserProcAddrsCount - 1
        If mUserProcAddrsCount > 0 _
           Then ReDim Preserve mUserProcAddrs(1 To mUserProcAddrsCount)
        UnSubclassOne = True
    Else
        '没有窗口程序＝addrUserProc，返回 false
        UnSubclassOne = False
    End If

    '-------- 如果列表此时不为空，函数返回上一块的返回值；否则返回下面的返回值 --------
    '判断是否 自定义窗口程序 列表为空，若为空则彻底恢复子类
    If mUserProcAddrsCount <= 0 Then UnSubclassOne = UnSubclassWhole
End Function

Public Function UnSubclassWhole() As Boolean
'恢复所有子类化（恢复窗口的默认窗口程序的地址）
'UnSubclassOne 是恢复一个 自定义窗口程序，若 自定义窗口程序 列表 _
  还有其他程序，则不完全恢复子类
    Dim ret As Long

    If lHwnd = 0 Or lAddrPrevWndProc = 0 Then
        ClearUserProcAddr
        UnSubclassWhole = False
        Exit Function
    Else
        ClearUserProcAddr
        SetLastError 0&
        ret = SetWindowLong(lHwnd, GWL_WNDPROC, lAddrPrevWndProc)
        '此时若恢复不成功，不能设置 lHwnd = 0 和 lAddrPrevWndProc = 0 _
         '因为主调程序还要获得这两个变量保存恢复不成功的信息以便后续处理
        If ret = 0 And GetLastError <> 0 Then UnSubclassWhole = False: Exit Function

        lHwnd = 0
        lAddrPrevWndProc = 0
        UnSubclassWhole = True
    End If
End Function

Private Sub AddOneUserProcAddr(ByVal oneProcAddr As Long)
'添加一个新的 自定义窗口程序的地址 到数组
'注意：如果数组中已有该项则不添加
    Dim i As Long
    '检查数组中是否已有该项，有则不添加直接退出
    For i = 1 To mUserProcAddrsCount
        If mUserProcAddrs(i) = oneProcAddr Then Exit Sub
    Next i
    '在数组中添加一项
    mUserProcAddrsCount = mUserProcAddrsCount + 1
    ReDim Preserve mUserProcAddrs(1 To mUserProcAddrsCount)
    mUserProcAddrs(mUserProcAddrsCount) = oneProcAddr
End Sub

Private Sub ClearUserProcAddr()
'清除 自定义窗口程序的地址 列表
'由于 mUserProcAddrs() 只保存自定义窗口程序的地址，并没保存默认 _
  窗口程序的地址，所以可大胆删除
    mUserProcAddrsCount = 0
    Erase mUserProcAddrs
End Sub




Public Property Get hwnd() As Long
    hwnd = lHwnd
End Property

Public Property Get PrevWndProcAddr() As Long
    PrevWndProcAddr = lAddrPrevWndProc
End Property

'获取当前有多少个 自定义窗口程序
Public Property Get UserProcAddrsCount() As Long
    UserProcAddrsCount = mUserProcAddrsCount
End Property

'获取一个 自定义窗口程序 的地址
Public Property Get UserProcAddr(index As Long) As Long
    If index > 0 And index <= mUserProcAddrsCount Then UserProcAddr = mUserProcAddrs(index)
End Property

Private Sub Class_Initialize()
    lHwnd = 0
    lAddrPrevWndProc = 0
    mUserProcAddrsCount = 0
End Sub

Private Sub Class_Terminate()
'Terminate 时，将 UnSubclassWhole，如果 UnSubclassWhole 没有成功，也销毁对象
    If lHwnd <> 0 And lAddrPrevWndProc <> 0 Then UnSubclassWhole
End Sub
