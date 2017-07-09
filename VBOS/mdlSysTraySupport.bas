Attribute VB_Name = "mdlSysTraySupport"
Option Explicit

'-------------------------- 系统托盘 clsSysTray 类模块的支持模块 ----------------------------

'需要 clsSysTray 类模块的支持
'需要 clsHashLK、clsSubClass 类模块 和 mdlSubClass 标准模块的支持

'#国际化：无提示字符串常量

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private Const WM_USER = &H400
Public Const ST_NOTI_MSG = WM_USER + 1001&    '自定义系统托盘的消息

'“登记”工程中所有的系统托盘。Key=对应窗体的hWnd，Data=一个clsSysTray对象的地址
'一个窗体最多只能使用一个 clsSysTray 对象设置托盘
Private mHashSysTrays As New clsHashLK


Public Function STRegOneObject(ByVal hwnd As Long, _
                               ByVal addrClsSysTray As Long) As Boolean
'“登记”一个 clsSysTray 对象
'一个窗体最多只能使用一个 clsSysTray 对象设置托盘
    If mHashSysTrays.IsKeyExist(hwnd) Then
        '该窗体已经被登记过“设置了系统托盘”，不能重复设置系统托盘
        STRegOneObject = False
    Else
        STRegOneObject = mHashSysTrays.Add(addrClsSysTray, hwnd, 0, "", False)
    End If
End Function

Public Function STUnRegOneObject(ByVal hwnd As Long) As Boolean
'取消"登记"一个窗体和 clsSysTray 对象，即删除 mHashSysTrays 中的一条记录
    mHashSysTrays.Remove hwnd
End Function


Public Function STWndProc(ByVal hwnd As Long, _
                          ByVal Msg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

'被设置系统托盘的窗口被子类化的自定义窗口程序
    If Msg = ST_NOTI_MSG Then
        '////////// 获得系统托盘消息，传递给对应的 clsSysTray 对象 //////////
        Dim lngAddrObj As Long    '对应 clsSysTray 对象地址

        '从哈希表获得 clsSysTray 对象地址
        lngAddrObj = mHashSysTrays.Item(hwnd, False)

        If lngAddrObj Then    '如果 lngAddrObj 为0表示失败或本VB程序中根本没有此窗口对应的托盘
            '通过弱引用调用 clsSysTray 对象的 EventsGen 方法生成事件
            Dim objCls As clsSysTray
            CopyMemory objCls, lngAddrObj, 4
            objCls.EventsGen wParam, lParam
            CopyMemory objCls, 0&, 4
        End If
    End If

    '返回值，“子类处理通用模块”会自动调用默认窗口程序处理
    STWndProc = gc_lngEventsGenDefautRet
End Function



