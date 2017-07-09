VERSION 5.00
Begin VB.Form Frm_Main 
   BackColor       =   &H00FFCC00&
   BorderStyle     =   0  'None
   Caption         =   "定时改壁纸"
   ClientHeight    =   7395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   20.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   6750
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Txt_Time 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFCC00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4080
      TabIndex        =   4
      Text            =   "001"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   240
   End
   Begin VB.Label lbl_weibo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "微博 @DLT_DA"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   3720
      TabIndex        =   14
      Top             =   6720
      Width           =   2880
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Deali-Axy"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   525
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   2445
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[DA开发]不占用资源，绿色无广告"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   525
      Left            =   240
      TabIndex        =   12
      Top             =   6120
      Width           =   6135
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   120
      X2              =   5760
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "使用方法"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   360
      TabIndex        =   11
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件夹里即可"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   360
      TabIndex        =   10
      Top             =   3000
      Width           =   2430
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "放在程序目录下面的""wall"""
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   4755
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "将需要自动设置壁纸的图片"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   4860
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2655
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   5655
   End
   Begin VB.Label Lbl_AutoStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "开机自启动(开启)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1440
      TabIndex        =   7
      Top             =   5160
      Width           =   3165
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   1200
      X2              =   6600
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   6000
      TabIndex        =   5
      Top             =   4560
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "自动更换间隔"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1440
      TabIndex        =   3
      Top             =   4560
      Width           =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape_Frame 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   495
      Left            =   6000
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Lbl_hide 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "隐藏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lbl_Exit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   2175
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto-WAll2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   30
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3570
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Shape_Frame
        .Left = 1
        .Top = 1
        .Width = Me.Width
        .Height = Me.Height
    End With
    WallCount = 0
    sPath = App.Path & "\wall\"
    If Dir(sPath, vbDirectory) <> "" Then flag = True Else flag = False
    If flag Then
        GetList
        GetConfig
        Run
    Else
        MsgBox "Config Error"
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl_hide.ForeColor = vbWhite
    lbl_Exit.ForeColor = vbWhite
    Lbl_AutoStart.ForeColor = vbWhite
    lbl_weibo.ForeColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveConfig
    If AppExit Then
    Else
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Lbl_AutoStart_Click()
    Dim tmp As Object
    Set tmp = CreateObject("WScript.Shell")
    If AutoStart Then
        AutoStart = False
        Lbl_AutoStart.Caption = "开机自启动(关闭)"
        tmp.regDelete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\DA-Autowall"
    Else
        AutoStart = True
        Lbl_AutoStart.Caption = "开机自启动(开启)"
        tmp.regWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\DA-Autowall", App.Path & "\" & App.EXEName & ".exe boot", "REG_SZ"
    End If
End Sub

Private Sub Lbl_AutoStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl_AutoStart.ForeColor = &HFFFF&
End Sub

Private Sub Lbl_AutoStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl_AutoStart_MouseState = MouseOver Then
        Lbl_AutoStart.ForeColor = &HFFC0FF
    Else
        lbl_AutoStart_MouseState = MouseOver
        Lbl_AutoStart.ForeColor = &HFFC0FF
    End If
End Sub

Private Sub Lbl_AutoStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl_AutoStart_MouseState = MouseOver Then
        Lbl_AutoStart.ForeColor = &HFFC0FF
    End If
End Sub

Private Sub lbl_Exit_Click()
    AppExit = True
    Unload Me
End Sub

Private Sub lbl_Exit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Lbl_hide.ForeColor = vbWhite
    lbl_Exit.ForeColor = &HFF00&
End Sub

Private Sub Lbl_hide_Click()
    Frm_Msg.ShowMsg "程序将会隐藏到后台运行"
    Me.Hide
End Sub

Private Sub Lbl_hide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbl_Exit.ForeColor = vbWhite
    Lbl_hide.ForeColor = &HFF00&
End Sub

Private Sub lbl_weibo_Click()
    ShellExecute Me.hwnd, "open", "http://www.weibo.com/dealiaxy", "", "", 5
End Sub

Private Sub lbl_weibo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lbl_weibo.ForeColor = &HFFFF&
End Sub

Private Sub Tmr_Timer()
    On Error Resume Next
    Dim PicName As String, TmpFile As String, Order As Long
    Static i As Integer, k As Integer
    i = i + 1
    If i >= iTime Then
        Randomize
        Order = Int(WallCount * Rnd + 1)
        PicName = WallList(Order)
        TmpFile = App.Path & "\DA-AutoWall-Tmp"
        If Dir(TmpFile) <> "" Then
            Kill TmpFile
        End If
        SavePicture LoadPicture(PicName), TmpFile
        Call SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, TmpFile, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)
        i = 0
        k = k + 1
        If k >= WallCount Then
            k = 0
        End If
    End If
End Sub

Private Sub Txt_Time_Change()
    iTime = Val(Txt_Time.Text)
End Sub

Private Sub Txt_Time_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then Exit Sub
        KeyAscii = 0
    End If
End Sub

Private Sub GetList()
    Dim FileName As String
    FileName = Dir(sPath, vbHidden Or vbNormal Or vbReadOnly Or vbSystem)
    Do While FileName <> ""
        AddItem sPath & FileName
        DoEvents
        FileName = Dir
        WallCount = WallCount + 1
        Debug.Print FileName
    Loop
End Sub

Private Sub GetConfig()
    If GetSetting("DA_Autowall", "Config", "Autostart") = "true" Then
        AutoStart = True
        Lbl_AutoStart.Caption = "开机自启动(开启)"
    Else
        AutoStart = False
        Lbl_AutoStart.Caption = "开机自启动(关闭)"
    End If
    If GetSetting("DA_Autowall", "Config", "time") <> "" Then
        iTime = Val(Trim(GetSetting("DA_Autowall", "Config", "time")))
        flag = True
        Txt_Time = Trim(Str(iTime))
    Else
        iTime = 0
        flag = False
        Txt_Time = "0"
    End If
End Sub

Private Sub SaveConfig()
    iTime = Val(Txt_Time)
    SaveSetting "DA_Autowall", "Config", "time", Str(iTime)
    SaveSetting "DA_Autowall", "Info", "build", Str(App.Revision)
    If AutoStart Then
        SaveSetting "DA_Autowall", "Config", "Autostart", "true"
    Else
        SaveSetting "DA_Autowall", "Config", "Autostart", "false"
    End If
End Sub

Private Sub AddItem(FileName As String)
    Static i As Integer
    WallList(i) = FileName
    i = i + 1
End Sub

Private Sub Run()
    Tmr.Enabled = True
End Sub
