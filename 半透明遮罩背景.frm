VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   0  'None
   Caption         =   "Form12"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16080
   Enabled         =   0   'False
   Icon            =   "°ëÍ¸Ã÷ÕÚÕÖ±³¾°.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   8895
   ScaleWidth      =   16080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    If Form1.formback = 1 Then
        Form1.Show
    ElseIf Form1.formback = 2 Then
        Form2.Show
    ElseIf Form1.formback = 5 Then
        Form5.Show
    ElseIf Form1.formback = 8 Then
        Form8.Show
    ElseIf Form1.formback = 10 Then
        Form10.Show
    ElseIf Form1.formback = 11 Then
        Form11.Show
    ElseIf Form1.formback = 23 Then
        Form23.Show
    ElseIf Form1.formback = 24 Then
        Form24.Show
    ElseIf Form1.formback = 25 Then
        Form25.Show
    ElseIf Form1.formback = 26 Then
        Form26.Show
    ElseIf Form1.formback = 27 Then
        Form27.Show
    ElseIf Form1.formback = 28 Then
        Form28.Show
    End If
End Sub

Private Sub Form_DblClick()
    If Form1.formback = 1 Then
        Form1.Show
    ElseIf Form1.formback = 2 Then
        Form2.Show
    ElseIf Form1.formback = 5 Then
        Form5.Show
    ElseIf Form1.formback = 8 Then
        Form8.Show
    ElseIf Form1.formback = 10 Then
        Form10.Show
    ElseIf Form1.formback = 11 Then
        Form11.Show
    ElseIf Form1.formback = 23 Then
        Form23.Show
    ElseIf Form1.formback = 24 Then
        Form24.Show
    ElseIf Form1.formback = 25 Then
        Form25.Show
    ElseIf Form1.formback = 26 Then
        Form26.Show
    ElseIf Form1.formback = 27 Then
        Form27.Show
    ElseIf Form1.formback = 28 Then
        Form28.Show
    End If
End Sub



Private Sub Form_Load()
    Me.Height = Form1.Height
    Me.Width = Form1.Width
    Me.BackColor = &H80000
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    '²»Í¨Í¸,    °ëÍ¸Ã÷
    SetWindowLong hwnd, GWL_EXSTYLE, rtn Or WS_EX_LAYERED
    SetLayeredWindowAttributes hwnd, 0, 210, LWA_ALPHA                          '(¾ä±ú  ,ÑÚÂëÑÕÉ«[RGB]  ,  Í¸Ã÷³Ì¶È[0-255],  )
End Sub
