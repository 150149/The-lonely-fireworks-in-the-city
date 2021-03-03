VERSION 5.00
Begin VB.UserControl PUIMgr 
   ClientHeight    =   750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   3300
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "PUIMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event MoveSmlyComplete(Control As Object)
Public Event SizeSmlyComplete(Control As Object)
Public Event ColorSmlyIng(nColor As Long)
Public Event ColorSmlyComplete()
'█████████████████████████████████████████████████████████████
Dim C1 As Object
Dim gLeft As Integer
Dim gTop As Integer
'█████████████████████████████████████████████████████████████
Dim C2 As Object
Dim gWidth As Integer
Dim gHeight As Integer
'█████████████████████████████████████████████████████████████
Dim C3 As Long
Dim gColor As Long
Dim gCGSPD As Integer
'█████████████████████████████████████████████████████████████
Private Sub Timer1_Timer()
    If Abs(C1.Left - gLeft) > 15 Then
        C1.Left = C1.Left - (C1.Left - gLeft) / 10
    Else
        C1.Left = gLeft
    End If
    If Abs(C1.Top - gTop) > 15 Then
        C1.Top = C1.Top - (C1.Top - gTop) / 10
    Else
        C1.Top = gTop
    End If
    If (C1.Left = gLeft) And (C1.Top = gTop) Then
        RaiseEvent MoveSmlyComplete(C1)
        Timer1.Enabled = False
    End If
End Sub
'█████████████████████████████████████████████████████████████
Private Sub Timer2_Timer()
    If Abs(C2.Width - gWidth) > 15 Then
        C2.Width = C2.Width - (C2.Width - gWidth) / 10
    Else
        C2.Width = gWidth
    End If
    If Abs(C2.Height - gHeight) > 15 Then
        C2.Height = C2.Height - (C2.Height - gHeight) / 10
    Else
        C2.Height = gHeight
    End If
    If (C2.Width = gWidth) And (C2.Height = gHeight) Then
        RaiseEvent SizeSmlyComplete(C2)
        Timer2.Enabled = False
    End If
End Sub

Private Sub Timer3_Timer()
    Dim r1 As Integer, g1 As Integer, B1 As Integer                             '定义起始颜色的RGB值
    Dim R2 As Integer, g2 As Integer, B2 As Integer                             '定义终止颜色的RGB值
    r1 = C3 Mod 256                                                             '计算起始颜色的RGB值
    g1 = (C3 Mod 65536) \ 256
    B1 = C3 \ 65536
    R2 = gColor Mod 256                                                         '计算终止颜色的RGB值
    g2 = (gColor Mod 65536) \ 256
    B2 = gColor \ 65536
    If r1 < R2 Then                                                             '更新颜色的RGB值
        If r1 + gCGSPD < R2 Then
            r1 = r1 + gCGSPD
        Else
            r1 = R2
        End If
    End If
    If r1 > R2 Then
        If r1 - gCGSPD > R2 Then
            r1 = r1 - gCGSPD
        Else
            r1 = R2
        End If
    End If
    If g1 < g2 Then
        If g1 + gCGSPD < g2 Then
            g1 = g1 + gCGSPD
        Else
            g1 = g2
        End If
    End If
    If g1 > g2 Then
        If g1 - gCGSPD > g2 Then
            g1 = g1 - gCGSPD
        Else
            g1 = g2
        End If
    End If
    If B1 < B2 Then
        If B1 + gCGSPD < B2 Then
            B1 = B1 + gCGSPD
        Else
            B1 = B2
        End If
    End If
    If B1 > B2 Then
        If B1 - gCGSPD > B2 Then
            B1 = B1 - gCGSPD
        Else
            B1 = B2
        End If
    End If
    C3 = RGB(r1, g1, B1)
    RaiseEvent ColorSmlyIng(RGB(r1, g1, B1))
    If (r1 = R2) And (g1 = g2) And (B1 = B2) Then                               '如果颜色渐变完成(即起始颜色等于终止颜色)
        Timer3.Enabled = False
        RaiseEvent ColorSmlyComplete
    End If
End Sub

'█████████████████████████████████████████████████████████████
Private Sub UserControl_Resize()
    Width = 480
    Height = 480
End Sub
'█████████████████████████████████████████████████████████████
Public Sub MoveSmly(ByRef Control As Object, ByVal nLeft As Integer, ByVal nTop As Integer, ByVal Delay As Integer)
    If (Delay > 1000) Or (Delay < 1) Then Exit Sub
    Timer1.Interval = Delay
    gLeft = nLeft
    gTop = nTop
    Set C1 = Control
    Timer1.Enabled = True
End Sub
'█████████████████████████████████████████████████████████████
Public Sub SizeSmly(ByRef Control As Object, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal Delay As Integer)
    If (Delay > 1000) Or (Delay < 1) Then Exit Sub
    Timer2.Interval = Delay
    gWidth = nWidth
    gHeight = nHeight
    Set C2 = Control
    Timer2.Enabled = True
End Sub
'█████████████████████████████████████████████████████████████
Public Sub ColorSmly(ByVal CurrentColor As Long, ByVal GoalColor As Long, ByVal CGSPD As Integer, ByVal Delay As Integer)
    If (Delay > 1000) Or (Delay < 1) Then Exit Sub
    Timer3.Interval = Delay
    C3 = CurrentColor
    gColor = GoalColor
    gCGSPD = CGSPD
    Timer3.Enabled = True
End Sub
'█████████████████████████████████████████████████████████████
Public Sub ControlTransparent(ByRef Container As Object, ByRef Control As Object, ByVal Transparency As Integer)
    If (Transparency < 0) Or (Transparency > 255) Then Exit Sub
    Container.AutoRedraw = True
    Control.AutoRedraw = True
    Dim bf As BLENDFUNCTION, IBF As Long
    With bf
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Transparency
        .AlphaFormat = 0
    End With
    RtlMoveMemory IBF, bf, 4
    AlphaBlend Control.hdc, 0, 0, Control.ScaleWidth / 15, Control.ScaleHeight / 15, Container.hdc, Control.Left / 15, Control.Top / 15, Control.ScaleWidth / 15, Control.ScaleHeight / 15, IBF
End Sub
'█████████████████████████████████████████████████████████████
'Public Sub FormTransparent(ByRef Frm As Form, ByVal Transparency As Integer)
'    Dim R As Long
'    R = GetWindowLong(Frm.hWnd, GWL_EXSTYLE)
'    R = R Or WS_EX_LAYERED
'    SetWindowLong Frm.hWnd, GWL_EXSTYLE, R
'    SetLayeredWindowAttributes Frm.hWnd, 0, Transparency, LWA_ALPHA
'End Sub
