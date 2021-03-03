VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form19"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.BackColor = &H80000
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hWnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hWnd, &H80000, 70, LWA_COLORKEY
    
End Sub


