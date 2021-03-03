VERSION 5.00
Begin VB.Form Form14 
   BorderStyle     =   0  'None
   Caption         =   "Form12"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16080
   Enabled         =   0   'False
   Icon            =   "启动界面半透明遮罩背景 .frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   8895
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Me.BackColor = &H80000
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 70, LWA_ALPHA                           '   窗体透明
End Sub
