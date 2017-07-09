VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QClass_QEverydayTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'----------------------------------ShowGraphics()----------------------------------
Private Const MaxPicNum As Integer = 2048    '设置自动获取文件夹中的最大图片数量
Private sPath As String    '图片文件夹路径
Private PicList(1 To MaxPicNum) As String    '图片列表
Private OrderList() As String    '图片顺序列表
Private PicCount As Integer    '图片数量
Private IsGraphicsInit As Boolean    '是否已经初始化图像

Function InitFile(FileName As String, OutputFileName As String) As Boolean
    On Error GoTo Err
    'Doc: 将普通文本文件转换成CQAF使用的格式
    If Len(Dir(FileName)) = 0 Then
        InitFile = False
        Exit Function
    End If
    Dim FileNum As Integer, strTmp As String, i As Integer
    i = 1
    FileNum = FreeFile
    QApp.QSettings.SaveToFile OutputFileName, "def", "title", Mid(FileName, InStrRev(FileName, "\") + 1)    '保存文件标题
    Open FileName For Input As FileNum
    Do While Not EOF(FileNum)
loadFile:
        Line Input #FileNum, strTmp
        If Len(Trim(strTmp)) = 0 Then    '跳过空行
            GoTo loadFile
        End If
        QApp.QSettings.SaveToFile OutputFileName, "line", Str(i), Mid(Trim(strTmp), 7)
        i = i + 1
    Loop
    Close FileNum
    QApp.QSettings.SaveToFile OutputFileName, "def", "line_count", Trim(Str(i - 1))    '保存文件行数
    Exit Function
Err:
    QDB.Runtime_Error "QClass_QEverydayTips->InitFile()", Err.Description, Err.Number
End Function

Function ShowText(SelectItem As Integer) As String     '显示
    On Error GoTo Err
    Dim FilePath As String
    Dim strTmp As String, i As Integer, line As Integer

    FilePath = App.Path & "\Resource\"
    Dim FileName As String
    Select Case SelectItem
        Case 1: FileName = "《那些年，我们一起追的女孩》经典语录"
        Case 2: FileName = "安妮宝贝经典语录"
        Case 3: FileName = "佛教经典语录"
        Case 4: FileName = "何以笙箫默语录"
        Case 5: FileName = "唯美古诗词"
        Case 6: FileName = "无下限段子"
        Case 7: FileName = "小学逗比英语"
        Case 8: FileName = "致青春经典台词22条"
    End Select

    FilePath = FilePath & FileName & ".qdat1"

    Randomize
    line = Val(QApp.QSettings.ReadFromFile(FilePath, "def", "line_count"))
    i = Int((line - 0 + 1) * Rnd + 0)    '取随机数
    strTmp = QApp.QSettings.ReadFromFile(FilePath, "line", Str(i))
    QApp.Tray_ShowText QApp.QSettings.ReadFromFile(FilePath, "def", "title"), strTmp
    ShowText = strTmp
    Exit Function
Err:
    QDB.Runtime_Error "QClass_QEverydayTips->ShowText()", Err.Description, Err.Number
End Function

Function ShowGraphics() As StdPicture
    On Error GoTo Err
    Static k As Long    '当前排到的图片序号
    Dim Order As Long  '随机图片顺序
    Dim PicName As String

    If Not IsGraphicsInit Then
        Exit Function
    End If

LoopI:
    k = 1
    If k > PicCount Then    '全部图片播完，重置顺序列表
        k = 1
        Dim sb As Long
        For sb = 1 To PicCount
            OrderList(sb) = ""
        Next
    End If

    '获取随机顺序
    Randomize
    Order = Int((PicCount + 1) * Rnd) + 1

    '不重复算法
    Dim II As Long
    For II = 1 To PicCount
        If PicList(Order) = OrderList(II) Then
            GoTo LoopI
        End If
    Next

    PicName = PicList(Order)    '从图片列表中读取文件名
    If Len(PicName) = 0 Then Exit Function    '没有文件名就退出
    QDB.Log "Order=" & Order & " " & " PicName=" & PicName
    OrderList(k) = PicName    '用过的图片保存起来
    k = k + 1

    Dim tmpPic As StdPicture
    Set tmpPic = LoadPicture(PicName)
    QFrm_Tray.ShowTrayGraphics "无下限表情包", "", tmpPic
    Set ShowGraphics = tmpPic    '返回图片
    Set tmpPic = Nothing    '最后释放内存
    Exit Function
Err:
    QDB.Runtime_Error "QClass_QEverydayTips->ShowGraphics()", Err.Description, Err.Number
End Function

Function GraphicsInit()    '图像模式初始化
    On Error GoTo Err
    sPath = App.Path & "\Resource\pic"    '设置图片文件夹

    If Dir(sPath, vbDirectory) <> "" Then    '检测文件夹是否存在
        '获取图片列表
        Dim FileName As String
        Dim ItemIndex As Long
        FileName = Dir(sPath & "\", vbNormal)
        Do While Len(FileName) > 0
            ItemIndex = ItemIndex + 1
            PicCount = PicCount + 1
            PicList(ItemIndex) = sPath & "\" & FileName
            DoEvents    '防止程序假死
            FileName = Dir    '再获取一次
            QDB.Log "QClass_QEverydayTips->ShowGraphics()->GetPicList,PicFileName(" & ItemIndex & ")=" & FileName
        Loop
        If PicCount > 0 Then
            ReDim OrderList(1 To PicCount)
        End If
    End If
    IsGraphicsInit = True    '状态：图像模式已经完成初始化
    Exit Function
Err:
    QDB.Runtime_Error "QClass_QEverydayTips->GraphicsInit()", Err.Description, Err.Number
End Function
