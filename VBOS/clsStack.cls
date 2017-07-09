VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------堆栈类模块--------------------------------------------
'默认为使用字符串型数据的堆栈，还支持使用 Long, Single, Double 类型的 _
 '堆栈，必须使用对应的函数版本，如 PopLong，PushLong，PeekLong等。
'但 IsEmpty 和 Clear 方法对所有版本的数据类型都适用
'---------------------------------------------------------------------
'#国际化：无提示字符串

Private Enum STK_DataType
    stkUnDefined = 0
    stkString = 1
    stkLong = 2
    stkSingle = 3
    stkDouble = 4
End Enum

Private lStackData() As String
Private lStackDataLong() As Long
Private lStackDataSingle() As Single
Private lStackDataDouble() As Double
Private lStackDataCount As Long
Private mStackDataUbound As Long
Private mDataType As STK_DataType

Private Const mcArrExpandPer As Long = 100    '每次扩增数组的大小

Public Function Pop() As String
    If lStackDataCount < 1 Or mDataType <> stkString Then Exit Function

    Pop = lStackData(lStackDataCount)
    lStackDataCount = lStackDataCount - 1
End Function

Public Function PopLong() As Long
    If lStackDataCount < 1 Or mDataType <> stkLong Then Exit Function

    PopLong = lStackDataLong(lStackDataCount)
    lStackDataCount = lStackDataCount - 1
End Function

Public Function PopSng() As Single
    If lStackDataCount < 1 Or mDataType <> stkSingle Then Exit Function

    PopSng = lStackDataSingle(lStackDataCount)
    lStackDataCount = lStackDataCount - 1
End Function

Public Function PopDbl() As Double
    If lStackDataCount < 1 Or mDataType <> stkDouble Then Exit Function

    PopDbl = lStackDataDouble(lStackDataCount)
    lStackDataCount = lStackDataCount - 1
End Function







Public Function Push(DataToPush As String) As Long
'返回 Push 后的栈顶序号
    If mDataType = stkUnDefined Then mDataType = stkString
    If mDataType <> stkString Then Exit Function

    ExpandArr
    lStackDataCount = lStackDataCount + 1
    lStackData(lStackDataCount) = DataToPush
    Push = lStackDataCount
End Function

Public Function PushLong(DataToPush As Long) As Long
'返回 Push 后的栈顶序号
    If mDataType = stkUnDefined Then mDataType = stkLong
    If mDataType <> stkLong Then Exit Function

    ExpandArr
    lStackDataCount = lStackDataCount + 1
    lStackDataLong(lStackDataCount) = DataToPush
    PushLong = lStackDataCount
End Function

Public Function PushSng(DataToPush As Single) As Long
'返回 Push 后的栈顶序号
    If mDataType = stkUnDefined Then mDataType = stkSingle
    If mDataType <> stkSingle Then Exit Function

    ExpandArr
    lStackDataCount = lStackDataCount + 1
    lStackDataSingle(lStackDataCount) = DataToPush
    PushSng = lStackDataCount
End Function

Public Function PushDbl(DataToPush As Double) As Long
'返回 Push 后的栈顶序号
    If mDataType = stkUnDefined Then mDataType = stkDouble
    If mDataType <> stkDouble Then Exit Function

    ExpandArr
    lStackDataCount = lStackDataCount + 1
    lStackDataDouble(lStackDataCount) = DataToPush
    PushDbl = lStackDataCount
End Function





Public Function Peek() As String
    If lStackDataCount < 1 Or mDataType <> stkString Then Exit Function
    Peek = lStackData(lStackDataCount)
End Function

Public Function PeekLong() As Long
    If lStackDataCount < 1 Or mDataType <> stkLong Then Exit Function
    PeekLong = lStackDataLong(lStackDataCount)
End Function

Public Function PeekSng() As Single
    If lStackDataCount < 1 Or mDataType <> stkSingle Then Exit Function
    PeekSng = lStackDataSingle(lStackDataCount)
End Function

Public Function PeekDbl() As Double
    If lStackDataCount < 1 Or mDataType <> stkDouble Then Exit Function
    PeekDbl = lStackDataDouble(lStackDataCount)
End Function


Public Function IsEmpty() As Boolean
    IsEmpty = (lStackDataCount < 1)
End Function

Public Sub Clear()
    Erase lStackData()
    Erase lStackDataLong()
    Erase lStackDataSingle()
    Erase lStackDataDouble()

    lStackDataCount = 0
    mStackDataUbound = 0
    mDataType = stkUnDefined
End Sub



Private Sub ExpandArr()
'扩增数组，使得 lStackDataCount + 1 的空间可用
    If lStackDataCount + 1 > mStackDataUbound Then
        mStackDataUbound = mStackDataUbound + mcArrExpandPer
        Select Case mDataType
        Case stkString
            ReDim Preserve lStackData(1 To mStackDataUbound)
        Case stkLong
            ReDim Preserve lStackDataLong(1 To mStackDataUbound)
        Case stkSingle
            ReDim Preserve lStackDataSingle(1 To mStackDataUbound)
        Case stkDouble
            ReDim Preserve lStackDataDouble(1 To mStackDataUbound)
        End Select
    End If
End Sub

Public Property Get StackDataCount() As Long
    StackDataCount = lStackDataCount
End Property

Private Sub Class_Initialize()
    lStackDataCount = 0
    mStackDataUbound = 0
    mDataType = stkUnDefined
End Sub

Private Sub Class_Terminate()
    Clear
End Sub
