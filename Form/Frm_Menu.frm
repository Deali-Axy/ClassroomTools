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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu Menu_FrmMain_������˵� 
      Caption         =   "������˵�"
      Begin VB.Menu Menu_FrmMain_ÿ��һ�� 
         Caption         =   "ÿ��һ��"
      End
      Begin VB.Menu Menu_FrmMain_h1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FrmMain_Halt 
         Caption         =   "�ػ�"
      End
      Begin VB.Menu Menu_FrmMain_Reboot 
         Caption         =   "����"
      End
      Begin VB.Menu Menu_FrmMain_h2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FrmMain_���� 
         Caption         =   "����"
      End
   End
   Begin VB.Menu Menu_���̲˵� 
      Caption         =   "���̲˵�"
      Begin VB.Menu Menu_Tray_�򿪽��ҹ��� 
         Caption         =   "�򿪽��ҹ���"
      End
      Begin VB.Menu Menu_Tray_������ 
         Caption         =   "������"
      End
      Begin VB.Menu Menu_Tray_h1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Tray_��ɫ����Ƽ� 
         Caption         =   "��ɫ����Ƽ�"
      End
      Begin VB.Menu Menu_Tray_��Լ����Ƽ� 
         Caption         =   "��Լ����Ƽ�"
      End
      Begin VB.Menu Menu_Tray_ͼƬ����Ƽ� 
         Caption         =   "ͼƬ����Ƽ�"
      End
      Begin VB.Menu Menu_Tray_�Զ����Ƽ� 
         Caption         =   "�Զ����Ƽ�"
         Begin VB.Menu Menu_Tray_�����޶��� 
            Caption         =   "�����޶���"
         End
         Begin VB.Menu Menu_Tray_̨�ʺ���¼ 
            Caption         =   "̨�� And ��¼"
         End
      End
      Begin VB.Menu Menu_Tray_h2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Tray_About 
         Caption         =   "����"
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

Private Sub Menu_FrmMain_ÿ��һ��_Click()
    Call Menu_Tray_��Լ����Ƽ�_Click
End Sub

Private Sub Menu_FrmMain_����_Click()
    QFrm_About.Show
End Sub

Private Sub Menu_Tray_About_Click()
    Call Menu_FrmMain_����_Click
End Sub

Private Sub Menu_Tray_��ɫ����Ƽ�_Click()
    Call Menu_Tray_��Լ����Ƽ�_Click
    QFrm_Tray.RandomColor
End Sub

Private Sub Menu_Tray_�򿪽��ҹ���_Click()
    If QFrm_Main.Visible Then
        QFrm_Main.Visible = False
    Else
        QFrm_Main.Visible = True
    End If
End Sub

Private Sub Menu_Tray_��Լ����Ƽ�_Click()
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

Private Sub Menu_Tray_������_Click()
    Frm_Random.Show
End Sub

Private Sub Menu_Tray_̨�ʺ���¼_Click()
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

Private Sub Menu_Tray_ͼƬ����Ƽ�_Click()
    QApp.QEverydayTips.ShowGraphics
End Sub

Private Sub Menu_Tray_�����޶���_Click()
    QApp.QEverydayTips.ShowText 6
End Sub

Public Sub MenuFunc_��Լ����Ƽ�()
    Call Menu_Tray_��Լ����Ƽ�_Click
End Sub

Public Sub MenuFunc_��ɫ����Ƽ�()
    Call Menu_Tray_��ɫ����Ƽ�_Click
End Sub

Public Sub MenuFunc_ͼƬ����Ƽ�()
    Call Menu_Tray_ͼƬ����Ƽ�_Click
End Sub

Public Sub MenuFunc_��ʾ��������¼()
    Call Menu_Tray_�����޶���_Click
End Sub
