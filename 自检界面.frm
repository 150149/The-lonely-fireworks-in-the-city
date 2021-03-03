VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form21"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   Icon            =   "自检界面.frx":0000
   LinkTopic       =   "Form21"
   ScaleHeight     =   4590
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   840
      Width           =   10215
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Timemode As Integer

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 0, LWA_COLORKEY
    If Dir(App.Path & "\picture\shape\start.gif") <> "" Then
        Image1.Picture = LoadPicture(App.Path & "\picture\shape\start.gif")
    Else
        MsgBox "缺少\picture\start.gif文件，请重新安装此程序"
    End If
    Shell "reg.bat"
    Timemode = 0
    Timer1.Enabled = True
    Timer1.Interval = 100
End Sub
'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim i     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '保存窗体宽度缩放比例
    scaleY = FormName.ScaleHeight / FormOldHeight                               '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For i = 0 To 4
            '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If
            '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next i
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me)                                                         '确保窗体改变时控件随之改变
End Sub

'在调用ResizeForm前先调用本函数
Public Sub ResizeInit(FormName As Form)
    Dim Obj     As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
        Obj.Tag = Obj.Left & "   " & Obj.Top & "   " & Obj.Width & "   " & Obj.Height & "   "
    Next Obj
    On Error GoTo 0
End Sub



Private Sub Timer1_Timer()
    Dim X As Integer
    Dim xa As Integer
    Dim xb As Integer
    If Timemode = 0 Then
        Label1.Caption = "正在检查\item.ini文件"
    ElseIf Timemode = 1 Then
        Label1.Caption = "正在检查\1.mp4文件"
    ElseIf Timemode = 2 Then
        Label1.Caption = "正在检查\picture\shape\baoshidu.gif文件"
    ElseIf Timemode = 3 Then
        Label1.Caption = "正在检查\picture\shape\yinshui.gif文件"
    ElseIf Timemode = 4 Then
        Label1.Caption = "正在检查\picture\shape\tili.gif文件"
    ElseIf Timemode = 5 Then
        Label1.Caption = "正在检查\music\background\1.mp3"
    ElseIf Timemode = 6 Then
        Label1.Caption = "正在检查\music\background\2.mp3"
    ElseIf Timemode = 7 Then
        Label1.Caption = "正在检查\music\background\3.mp3"
    ElseIf Timemode = 8 Then
        Label1.Caption = "正在检查\music\background\4.mp3"
    ElseIf Timemode = 9 Then
        Label1.Caption = "正在检查\music\background\5.mp3"
    ElseIf Timemode = 10 Then
        Label1.Caption = "正在检查\music\background\6.mp3"
    ElseIf Timemode = 11 Then
        Label1.Caption = "正在检查\music\background\7.mp3"
    ElseIf Timemode = 12 Then
        Label1.Caption = "正在检查\music\background\8.mp3"
    ElseIf Timemode = 13 Then
        Label1.Caption = "正在检查\music\background\9.mp3"
    ElseIf Timemode = 14 Then
        Label1.Caption = "正在检查\music\background\10.mp3"
    ElseIf Timemode = 15 Then
        Label1.Caption = "正在检查\music\background\11.mp3"
    ElseIf Timemode = 16 Then
        Label1.Caption = "正在检查\music\background\12.mp3"
    ElseIf Timemode = 17 Then
        Label1.Caption = "正在检查\buildings\" & X & ".ini"
    ElseIf Timemode = 18 Then
        Label1.Caption = "正在检查\music\begin.mp3文件"
    ElseIf Timemode = 19 Then
        Label1.Caption = "正在检查\picture\logo.gif"
    ElseIf Timemode = 20 Then
        Label1.Caption = "正在检查\picture\item\" & xa & ".gif文件"
    ElseIf Timemode = 21 Then
        Label1.Caption = "正在检查\picture\item\" & xb
    ElseIf Timemode = 22 Then
        Label1.Caption = "正在检查\picture\shape\del.gif文件"
    ElseIf Timemode = 23 Then
        Label1.Caption = "正在检查\picture\shape\1600.jpg文件"
    ElseIf Timemode = 24 Then
        Label1.Caption = "正在检查\picture\shape\800.jpg文件"
    ElseIf Timemode = 25 Then
        Label1.Caption = "正在检查\picture\shape\1024.jpg文件"
    ElseIf Timemode = 26 Then
        Label1.Caption = "正在检查\picture\shape\1152.jpg文件"
    ElseIf Timemode = 27 Then
        Label1.Caption = "正在检查\picture\shape\1280.jpg文件"
    ElseIf Timemode = 28 Then
        Label1.Caption = "正在检查\picture\shape\1360.jpg文件"
    ElseIf Timemode = 29 Then
        Label1.Caption = "正在检查\picture\shape\1366.jpg文件"
    ElseIf Timemode = 30 Then
        Label1.Caption = "正在检查\picture\shape\1440.jpg文件"
    ElseIf Timemode = 31 Then
        Label1.Caption = "正在检查\picture\shape\1920.jpg文件"
    ElseIf Timemode = 32 Then
        Label1.Caption = "检查完成，正在启动游戏"
    End If
    Label1.Caption = "正在检查\item.ini文件"
    DoEvents
    If Timemode = 0 And Dir(App.Path & "\item.ini") = "" Then
        MsgBox "缺少item.ini文件，请重新安装该程序"
        End
    ElseIf Timemode = 0 And Dir(App.Path & "\item.ini") <> "" Then
        Timemode = 1
    End If
    Label1.Caption = "正在检查\1.mp4文件"
    DoEvents
    If Timemode = 1 And Dir(App.Path & "\video\1.mp4") = "" Then
        MsgBox "缺少1.mp4文件，请重新安装此程序"
        End
    ElseIf Timemode = 1 And Dir(App.Path & "\video\1.mp4") <> "" Then
        Timemode = 2
    End If
    Label1.Caption = "正在检查\picture\shape\baoshidu.gif文件"
    DoEvents
    If Timemode = 2 And Dir(App.Path & "\picture\shape\baoshidu.gif") = "" Then
        MsgBox "缺少\picture\shape\baoshidu.gif文件，请重新安装此程序"
        End
    ElseIf Timemode = 2 And Dir(App.Path & "\picture\shape\baoshidu.gif") <> "" Then
        Timemode = 3
    End If
    Label1.Caption = "正在检查\picture\shape\yinshui.gif文件"
    DoEvents
    If Timemode = 3 And Dir(App.Path & "\picture\shape\yinshui.gif") = "" Then
        MsgBox "缺少\picture\shape\yinshui.gif文件，请重新安装此程序"
        End
    ElseIf Timemode = 3 And Dir(App.Path & "\picture\shape\yinshui.gif") <> "" Then
        Timemode = 4
    End If
    Label1.Caption = "正在检查\picture\shape\tili.gif文件"
    DoEvents
    If Timemode = 4 And Dir(App.Path & "\picture\shape\tili.gif") = "" Then
        MsgBox "缺少\picture\shape\tili.gif文件，请重新安装此程序"
        End
    ElseIf Timemode = 4 And Dir(App.Path & "\picture\shape\tili.gif") <> "" Then
        Timemode = 5
    End If
    Label1.Caption = "正在检查\music\background\1.mp3"
    DoEvents
    'If Timemode = 5 And Dir(App.Path & "\picture\shape\money.gif") = "" Then
    'MsgBox "缺少\picture\shape\money.gif文件，请重新安装此程序"
    'End
    'ElseIf Timemode = 2 And Dir(App.Path & "\picture\shape\money.gif") <> "" Then
    'Timemode = 3
    'End If
    
    If Timemode = 5 And Dir(App.Path & "\music\background\1.mp3") = "" Then
        MsgBox "缺少\music\background\1.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 5 And Dir(App.Path & "\music\background\1.mp3") <> "" Then
        Timemode = 6
    End If
    Label1.Caption = "正在检查\music\background\2.mp3"
    DoEvents
    If Timemode = 6 And Dir(App.Path & "\music\background\2.mp3") = "" Then
        MsgBox "缺少\music\background\2.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 6 And Dir(App.Path & "\music\background\2.mp3") <> "" Then
        Timemode = 7
    End If
    Label1.Caption = "正在检查\music\background\3.mp3"
    DoEvents
    If Timemode = 7 And Dir(App.Path & "\music\background\3.mp3") = "" Then
        MsgBox "缺少\music\background\3.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 7 And Dir(App.Path & "\music\background\3.mp3") <> "" Then
        Timemode = 9
    End If
    Label1.Caption = "正在检查\music\background\4.mp3"
    DoEvents
    If Timemode = 9 And Dir(App.Path & "\music\background\4.mp3") = "" Then
        MsgBox "缺少\music\background\4.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 9 And Dir(App.Path & "\music\background\4.mp3") <> "" Then
        Timemode = 10
    End If
    Label1.Caption = "正在检查\music\background\5.mp3"
    DoEvents
    If Timemode = 10 And Dir(App.Path & "\music\background\5.mp3") = "" Then
        MsgBox "缺少\music\background\5.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 10 And Dir(App.Path & "\music\background\5.mp3") <> "" Then
        Timemode = 11
    End If
    Label1.Caption = "正在检查\music\background\6.mp3"
    DoEvents
    If Timemode = 11 And Dir(App.Path & "\music\background\6.mp3") = "" Then
        MsgBox "缺少\music\background\6.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 11 And Dir(App.Path & "\music\background\6.mp3") <> "" Then
        Timemode = 12
    End If
    Label1.Caption = "正在检查\music\background\7.mp3"
    DoEvents
    If Timemode = 11 And Dir(App.Path & "\music\background\7.mp3") = "" Then
        MsgBox "缺少\music\background\7.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 11 And Dir(App.Path & "\music\background\7.mp3") <> "" Then
        Timemode = 12
    End If
    Label1.Caption = "正在检查\music\background\8.mp3"
    DoEvents
    If Timemode = 12 And Dir(App.Path & "\music\background\8.mp3") = "" Then
        MsgBox "缺少\music\background\8.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 12 And Dir(App.Path & "\music\background\8.mp3") <> "" Then
        Timemode = 13
    End If
    Label1.Caption = "正在检查\music\background\9.mp3"
    DoEvents
    If Timemode = 13 And Dir(App.Path & "\music\background\9.mp3") = "" Then
        MsgBox "缺少\music\background\9.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 13 And Dir(App.Path & "\music\background\9.mp3") <> "" Then
        Timemode = 14
    End If
    Label1.Caption = "正在检查\music\background\10.mp3"
    DoEvents
    If Timemode = 14 And Dir(App.Path & "\music\background\10.mp3") = "" Then
        MsgBox "缺少\music\background\10.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 14 And Dir(App.Path & "\music\background\10.mp3") <> "" Then
        Timemode = 15
    End If
    Label1.Caption = "正在检查\music\background\11.mp3"
    DoEvents
    If Timemode = 15 And Dir(App.Path & "\music\background\11.mp3") = "" Then
        MsgBox "缺少\music\background\11.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 15 And Dir(App.Path & "\music\background\11.mp3") <> "" Then
        Timemode = 16
    End If
    Label1.Caption = "正在检查\music\background\12.mp3"
    DoEvents
    If Timemode = 16 And Dir(App.Path & "\music\background\12.mp3") = "" Then
        MsgBox "缺少\music\background\12.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 16 And Dir(App.Path & "\music\background\12.mp3") <> "" Then
        Timemode = 17
    End If
    If Timemode = 17 Then
        For X = 1 To 14
            Label1.Caption = "正在检查\buildings\" & X & ".ini"
            DoEvents
            If Dir(App.Path & "\buildings\" & X & ".ini") = "" Then MsgBox "缺少\buildings\" & X & ".ini文件，请重新安装此程序"
            If Dir(App.Path & "\buildings\" & X & ".ini") = "" Then End
        Next
        Timemode = 18
    End If
    Label1.Caption = "正在检查\music\begin.mp3文件"
    DoEvents
    If Timemode = 18 And Dir(App.Path & "\music\begin.mp3") = "" Then
        MsgBox "缺少\music\begin.mp3文件，请重新安装此程序"
        End
    ElseIf Timemode = 18 And Dir(App.Path & "\music\begin.mp3") <> "" Then
        Timemode = 19
    End If
    Label1.Caption = "正在检查\picture\logo.gif"
    DoEvents
    If Timemode = 19 And Dir(App.Path & "\picture\logo.gif") = "" Then
        MsgBox "缺少\picture\logo.gif文件，请重新安装此程序"
        End
    ElseIf Timemode = 19 And Dir(App.Path & "\picture\logo.gif") <> "" Then
        Timemode = 20
    End If
    
    If Timemode = 20 Then
        
        For xa = 1 To 38
            Label1.Caption = "正在检查\picture\item\" & xa & ".gif文件"
            DoEvents
            If Dir(App.Path & "\picture\item\" & xa & ".gif") = "" Then MsgBox "缺少\picture\item\" & xa & ".gif文件，请重新安装此程序"
            If Dir(App.Path & "\picture\item\" & xa & ".gif") = "" Then End
        Next
        Timemode = 21
    End If
    
    If Timemode = 21 Then
        
        For xb = 1 To 38
            Label1.Caption = "正在检查\picture\item\" & xb
            DoEvents
            If Dir(App.Path & "\picture\item\" & xb) = "" Then MsgBox "缺少\picture\item\" & xb & "文件，请重新安装此程序"
            If Dir(App.Path & "\picture\item\" & xb) = "" Then End
        Next
        Timemode = 22
    End If
    Label1.Caption = "正在检查\picture\shape\del.gif文件"
    DoEvents
    If Timemode = 22 And Dir(App.Path & "\picture\shape\del.gif") = "" Then
        MsgBox "缺少\picture\shape\del.gif文件，请重新安装此程序"
        End
    ElseIf Timemode = 22 And Dir(App.Path & "\picture\shape\del.gif") <> "" Then
        Timemode = 23
    End If
    Label1.Caption = "正在检查\picture\shape\1600.jpg文件"
    DoEvents
    If Timemode = 23 And Dir(App.Path & "\picture\shape\1600.jpg") = "" Then
        MsgBox "缺少\picture\shape\1600.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 23 And Dir(App.Path & "\picture\shape\1600.jpg") <> "" Then
        Timemode = 24
    End If
    Label1.Caption = "正在检查\picture\shape\800.jpg文件"
    DoEvents
    If Timemode = 24 And Dir(App.Path & "\picture\shape\800.jpg") = "" Then
        MsgBox "缺少\picture\shape\800.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 24 And Dir(App.Path & "\picture\shape\800.jpg") <> "" Then
        Timemode = 25
    End If
    Label1.Caption = "正在检查\picture\shape\1024.jpg文件"
    DoEvents
    If Timemode = 25 And Dir(App.Path & "\picture\shape\1024.jpg") = "" Then
        MsgBox "缺少\picture\shape\1024.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 25 And Dir(App.Path & "\picture\shape\1024.jpg") <> "" Then
        Timemode = 26
    End If
    Label1.Caption = "正在检查\picture\shape\1152.jpg文件"
    DoEvents
    If Timemode = 26 And Dir(App.Path & "\picture\shape\1152.jpg") = "" Then
        MsgBox "缺少\picture\shape\1152.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 26 And Dir(App.Path & "\picture\shape\1152.jpg") <> "" Then
        Timemode = 27
    End If
    Label1.Caption = "正在检查\picture\shape\1280.jpg文件"
    DoEvents
    If Timemode = 27 And Dir(App.Path & "\picture\shape\1280.jpg") = "" Then
        MsgBox "缺少\picture\shape\1280.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 27 And Dir(App.Path & "\picture\shape\1280.jpg") <> "" Then
        Timemode = 28
    End If
    Label1.Caption = "正在检查\picture\shape\1360.jpg文件"
    DoEvents
    If Timemode = 28 And Dir(App.Path & "\picture\shape\1360.jpg") = "" Then
        MsgBox "缺少\picture\shape\1360.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 28 And Dir(App.Path & "\picture\shape\1360.jpg") <> "" Then
        Timemode = 29
    End If
    Label1.Caption = "正在检查\picture\shape\1366.jpg文件"
    DoEvents
    If Timemode = 29 And Dir(App.Path & "\picture\shape\1366.jpg") = "" Then
        MsgBox "缺少\picture\shape\1366.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 29 And Dir(App.Path & "\picture\shape\1366.jpg") <> "" Then
        Timemode = 30
    End If
    Label1.Caption = "正在检查\picture\shape\1440.jpg文件"
    DoEvents
    If Timemode = 30 And Dir(App.Path & "\picture\shape\1440.jpg") = "" Then
        MsgBox "缺少\picture\shape\1440.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 30 And Dir(App.Path & "\picture\shape\1440.jpg") <> "" Then
        Timemode = 31
    End If
    Label1.Caption = "正在检查\picture\shape\1920.jpg文件"
    DoEvents
    If Timemode = 31 And Dir(App.Path & "\picture\shape\1920.jpg") = "" Then
        MsgBox "缺少\picture\shape\1920.jpg文件，请重新安装此程序"
        End
    ElseIf Timemode = 31 And Dir(App.Path & "\picture\shape\1920.jpg") <> "" Then
        Timemode = 32
    End If
    Label1.Caption = "检查完成，正在启动游戏"
    DoEvents
    If Timemode = 32 Then
        Form3.Show
        Unload Me
    End If
End Sub
