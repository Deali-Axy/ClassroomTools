VERSION 5.00
Begin VB.Form Frm_Menu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3015
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu Menu_FrmMain_主标题菜单 
      Caption         =   "主标题菜单"
      Begin VB.Menu Menu_FrmMain_每日一句 
         Caption         =   "每日一句"
      End
      Begin VB.Menu Menu_FrmMain_h1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FrmMain_Halt 
         Caption         =   "关机"
      End
      Begin VB.Menu Menu_FrmMain_Reboot 
         Caption         =   "重启"
      End
      Begin VB.Menu Menu_FrmMain_h2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FrmMain_关于 
         Caption         =   "关于"
      End
   End
   Begin VB.Menu Menu_托盘菜单 
      Caption         =   "托盘菜单"
      Begin VB.Menu Menu_Tray_打开教室管理 
         Caption         =   "打开教室管理"
      End
      Begin VB.Menu Menu_Tray_随机抽号 
         Caption         =   "随机抽号"
      End
      Begin VB.Menu Menu_Tray_h1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Tray_彩色随机推荐 
         Caption         =   "彩色随机推荐"
      End
      Begin VB.Menu Menu_Tray_简约随机推荐 
         Caption         =   "简约随机推荐"
      End
      Begin VB.Menu Menu_Tray_图片随机推荐 
         Caption         =   "图片随机推荐"
      End
      Begin VB.Menu Menu_Tray_自定义推荐 
         Caption         =   "自定义推荐"
         Begin VB.Menu Menu_Tray_无下限段子 
            Caption         =   "无下限段子"
         End
         Begin VB.Menu Menu_Tray_台词和语录 
            Caption         =   "台词 And 语录"
         End
      End
      Begin VB.Menu Menu_Tray_h2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Tray_About 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "Frm_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Menu_FrmMain_Halt_Click()
    Shell "shutdown -s -t 0"
End Sub

Private Sub Menu_FrmMain_Reboot_Click()
    Shell "shutdown -r -t 0"
End Sub

Private Sub Menu_FrmMain_每日一句_Click()
    Call Menu_Tray_简约随机推荐_Click
End Sub

Private Sub Menu_FrmMain_关于_Click()
    QFrm_About.Show
End Sub

Private Sub Menu_Tray_About_Click()
    Call Menu_FrmMain_关于_Click
End Sub

Private Sub Menu_Tray_彩色随机推荐_Click()
    Call Menu_Tray_简约随机推荐_Click
    QFrm_Tray.RandomColor
End Sub

Private Sub Menu_Tray_打开教室管理_Click()
    If QFrm_Main.Visible Then
        QFrm_Main.Visible = False
    Else
        QFrm_Main.Visible = True
    End If
End Sub

Private Sub Menu_Tray_简约随机推荐_Click()
    Const ItemCount As Integer = 12
    Randomize
    Dim i As Integer
    i = Int((ItemCount + 1) * Rnd + 0)
    Select Case i
        Case Is <= 8
            QApp.QEverydayTips.ShowText i
        Case Is > 8
            QApp.QEverydayTips.ShowText 6
    End Select
End Sub

Private Sub Menu_Tray_随机抽号_Click()
    Frm_Random.Show
End Sub

Private Sub Menu_Tray_台词和语录_Click()
    Const ItemCount As Integer = 3
    Randomize
    Dim i As Integer
    i = Int((ItemCount + 1) * Rnd + 0)
    Select Case i
        Case 1
            QApp.QEverydayTips.ShowText 1
        Case 2
            QApp.QEverydayTips.ShowText 4
        Case 3
            QApp.QEverydayTips.ShowText 8
    End Select
End Sub

Private Sub Menu_Tray_图片随机推荐_Click()
    QApp.QEverydayTips.ShowGraphics
End Sub

Private Sub Menu_Tray_无下限段子_Click()
    QApp.QEverydayTips.ShowText 6
End Sub

Public Sub MenuFunc_简约随机推荐()
    Call Menu_Tray_简约随机推荐_Click
End Sub

Public Sub MenuFunc_彩色随机推荐()
    Call Menu_Tray_彩色随机推荐_Click
End Sub

Public Sub MenuFunc_图片随机推荐()
    Call Menu_Tray_图片随机推荐_Click
End Sub

Public Sub MenuFunc_显示无下限语录()
    Call Menu_Tray_无下限段子_Click
End Sub
