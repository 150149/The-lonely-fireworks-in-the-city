VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form21"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13200
   Icon            =   "�Լ����.frx":0000
   LinkTopic       =   "Form21"
   ScaleHeight     =   4590
   ScaleWidth      =   13200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�
Private Timemode As Integer

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������
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
        MsgBox "ȱ��\picture\start.gif�ļ��������°�װ�˳���"
    End If
    Shell "reg.bat"
    Timemode = 0
    Timer1.Enabled = True
    Timer1.Interval = 100
End Sub
'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim i     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '���洰�������ű���
    scaleY = FormName.ScaleHeight / FormOldHeight                               '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For i = 0 To 4
            '��ȡ�ؼ���ԭʼλ�����С
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If
            '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next i
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me)                                                         'ȷ������ı�ʱ�ؼ���֮�ı�
End Sub

'�ڵ���ResizeFormǰ�ȵ��ñ�����
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
        Label1.Caption = "���ڼ��\item.ini�ļ�"
    ElseIf Timemode = 1 Then
        Label1.Caption = "���ڼ��\1.mp4�ļ�"
    ElseIf Timemode = 2 Then
        Label1.Caption = "���ڼ��\picture\shape\baoshidu.gif�ļ�"
    ElseIf Timemode = 3 Then
        Label1.Caption = "���ڼ��\picture\shape\yinshui.gif�ļ�"
    ElseIf Timemode = 4 Then
        Label1.Caption = "���ڼ��\picture\shape\tili.gif�ļ�"
    ElseIf Timemode = 5 Then
        Label1.Caption = "���ڼ��\music\background\1.mp3"
    ElseIf Timemode = 6 Then
        Label1.Caption = "���ڼ��\music\background\2.mp3"
    ElseIf Timemode = 7 Then
        Label1.Caption = "���ڼ��\music\background\3.mp3"
    ElseIf Timemode = 8 Then
        Label1.Caption = "���ڼ��\music\background\4.mp3"
    ElseIf Timemode = 9 Then
        Label1.Caption = "���ڼ��\music\background\5.mp3"
    ElseIf Timemode = 10 Then
        Label1.Caption = "���ڼ��\music\background\6.mp3"
    ElseIf Timemode = 11 Then
        Label1.Caption = "���ڼ��\music\background\7.mp3"
    ElseIf Timemode = 12 Then
        Label1.Caption = "���ڼ��\music\background\8.mp3"
    ElseIf Timemode = 13 Then
        Label1.Caption = "���ڼ��\music\background\9.mp3"
    ElseIf Timemode = 14 Then
        Label1.Caption = "���ڼ��\music\background\10.mp3"
    ElseIf Timemode = 15 Then
        Label1.Caption = "���ڼ��\music\background\11.mp3"
    ElseIf Timemode = 16 Then
        Label1.Caption = "���ڼ��\music\background\12.mp3"
    ElseIf Timemode = 17 Then
        Label1.Caption = "���ڼ��\buildings\" & X & ".ini"
    ElseIf Timemode = 18 Then
        Label1.Caption = "���ڼ��\music\begin.mp3�ļ�"
    ElseIf Timemode = 19 Then
        Label1.Caption = "���ڼ��\picture\logo.gif"
    ElseIf Timemode = 20 Then
        Label1.Caption = "���ڼ��\picture\item\" & xa & ".gif�ļ�"
    ElseIf Timemode = 21 Then
        Label1.Caption = "���ڼ��\picture\item\" & xb
    ElseIf Timemode = 22 Then
        Label1.Caption = "���ڼ��\picture\shape\del.gif�ļ�"
    ElseIf Timemode = 23 Then
        Label1.Caption = "���ڼ��\picture\shape\1600.jpg�ļ�"
    ElseIf Timemode = 24 Then
        Label1.Caption = "���ڼ��\picture\shape\800.jpg�ļ�"
    ElseIf Timemode = 25 Then
        Label1.Caption = "���ڼ��\picture\shape\1024.jpg�ļ�"
    ElseIf Timemode = 26 Then
        Label1.Caption = "���ڼ��\picture\shape\1152.jpg�ļ�"
    ElseIf Timemode = 27 Then
        Label1.Caption = "���ڼ��\picture\shape\1280.jpg�ļ�"
    ElseIf Timemode = 28 Then
        Label1.Caption = "���ڼ��\picture\shape\1360.jpg�ļ�"
    ElseIf Timemode = 29 Then
        Label1.Caption = "���ڼ��\picture\shape\1366.jpg�ļ�"
    ElseIf Timemode = 30 Then
        Label1.Caption = "���ڼ��\picture\shape\1440.jpg�ļ�"
    ElseIf Timemode = 31 Then
        Label1.Caption = "���ڼ��\picture\shape\1920.jpg�ļ�"
    ElseIf Timemode = 32 Then
        Label1.Caption = "�����ɣ�����������Ϸ"
    End If
    Label1.Caption = "���ڼ��\item.ini�ļ�"
    DoEvents
    If Timemode = 0 And Dir(App.Path & "\item.ini") = "" Then
        MsgBox "ȱ��item.ini�ļ��������°�װ�ó���"
        End
    ElseIf Timemode = 0 And Dir(App.Path & "\item.ini") <> "" Then
        Timemode = 1
    End If
    Label1.Caption = "���ڼ��\1.mp4�ļ�"
    DoEvents
    If Timemode = 1 And Dir(App.Path & "\video\1.mp4") = "" Then
        MsgBox "ȱ��1.mp4�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 1 And Dir(App.Path & "\video\1.mp4") <> "" Then
        Timemode = 2
    End If
    Label1.Caption = "���ڼ��\picture\shape\baoshidu.gif�ļ�"
    DoEvents
    If Timemode = 2 And Dir(App.Path & "\picture\shape\baoshidu.gif") = "" Then
        MsgBox "ȱ��\picture\shape\baoshidu.gif�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 2 And Dir(App.Path & "\picture\shape\baoshidu.gif") <> "" Then
        Timemode = 3
    End If
    Label1.Caption = "���ڼ��\picture\shape\yinshui.gif�ļ�"
    DoEvents
    If Timemode = 3 And Dir(App.Path & "\picture\shape\yinshui.gif") = "" Then
        MsgBox "ȱ��\picture\shape\yinshui.gif�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 3 And Dir(App.Path & "\picture\shape\yinshui.gif") <> "" Then
        Timemode = 4
    End If
    Label1.Caption = "���ڼ��\picture\shape\tili.gif�ļ�"
    DoEvents
    If Timemode = 4 And Dir(App.Path & "\picture\shape\tili.gif") = "" Then
        MsgBox "ȱ��\picture\shape\tili.gif�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 4 And Dir(App.Path & "\picture\shape\tili.gif") <> "" Then
        Timemode = 5
    End If
    Label1.Caption = "���ڼ��\music\background\1.mp3"
    DoEvents
    'If Timemode = 5 And Dir(App.Path & "\picture\shape\money.gif") = "" Then
    'MsgBox "ȱ��\picture\shape\money.gif�ļ��������°�װ�˳���"
    'End
    'ElseIf Timemode = 2 And Dir(App.Path & "\picture\shape\money.gif") <> "" Then
    'Timemode = 3
    'End If
    
    If Timemode = 5 And Dir(App.Path & "\music\background\1.mp3") = "" Then
        MsgBox "ȱ��\music\background\1.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 5 And Dir(App.Path & "\music\background\1.mp3") <> "" Then
        Timemode = 6
    End If
    Label1.Caption = "���ڼ��\music\background\2.mp3"
    DoEvents
    If Timemode = 6 And Dir(App.Path & "\music\background\2.mp3") = "" Then
        MsgBox "ȱ��\music\background\2.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 6 And Dir(App.Path & "\music\background\2.mp3") <> "" Then
        Timemode = 7
    End If
    Label1.Caption = "���ڼ��\music\background\3.mp3"
    DoEvents
    If Timemode = 7 And Dir(App.Path & "\music\background\3.mp3") = "" Then
        MsgBox "ȱ��\music\background\3.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 7 And Dir(App.Path & "\music\background\3.mp3") <> "" Then
        Timemode = 9
    End If
    Label1.Caption = "���ڼ��\music\background\4.mp3"
    DoEvents
    If Timemode = 9 And Dir(App.Path & "\music\background\4.mp3") = "" Then
        MsgBox "ȱ��\music\background\4.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 9 And Dir(App.Path & "\music\background\4.mp3") <> "" Then
        Timemode = 10
    End If
    Label1.Caption = "���ڼ��\music\background\5.mp3"
    DoEvents
    If Timemode = 10 And Dir(App.Path & "\music\background\5.mp3") = "" Then
        MsgBox "ȱ��\music\background\5.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 10 And Dir(App.Path & "\music\background\5.mp3") <> "" Then
        Timemode = 11
    End If
    Label1.Caption = "���ڼ��\music\background\6.mp3"
    DoEvents
    If Timemode = 11 And Dir(App.Path & "\music\background\6.mp3") = "" Then
        MsgBox "ȱ��\music\background\6.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 11 And Dir(App.Path & "\music\background\6.mp3") <> "" Then
        Timemode = 12
    End If
    Label1.Caption = "���ڼ��\music\background\7.mp3"
    DoEvents
    If Timemode = 11 And Dir(App.Path & "\music\background\7.mp3") = "" Then
        MsgBox "ȱ��\music\background\7.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 11 And Dir(App.Path & "\music\background\7.mp3") <> "" Then
        Timemode = 12
    End If
    Label1.Caption = "���ڼ��\music\background\8.mp3"
    DoEvents
    If Timemode = 12 And Dir(App.Path & "\music\background\8.mp3") = "" Then
        MsgBox "ȱ��\music\background\8.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 12 And Dir(App.Path & "\music\background\8.mp3") <> "" Then
        Timemode = 13
    End If
    Label1.Caption = "���ڼ��\music\background\9.mp3"
    DoEvents
    If Timemode = 13 And Dir(App.Path & "\music\background\9.mp3") = "" Then
        MsgBox "ȱ��\music\background\9.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 13 And Dir(App.Path & "\music\background\9.mp3") <> "" Then
        Timemode = 14
    End If
    Label1.Caption = "���ڼ��\music\background\10.mp3"
    DoEvents
    If Timemode = 14 And Dir(App.Path & "\music\background\10.mp3") = "" Then
        MsgBox "ȱ��\music\background\10.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 14 And Dir(App.Path & "\music\background\10.mp3") <> "" Then
        Timemode = 15
    End If
    Label1.Caption = "���ڼ��\music\background\11.mp3"
    DoEvents
    If Timemode = 15 And Dir(App.Path & "\music\background\11.mp3") = "" Then
        MsgBox "ȱ��\music\background\11.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 15 And Dir(App.Path & "\music\background\11.mp3") <> "" Then
        Timemode = 16
    End If
    Label1.Caption = "���ڼ��\music\background\12.mp3"
    DoEvents
    If Timemode = 16 And Dir(App.Path & "\music\background\12.mp3") = "" Then
        MsgBox "ȱ��\music\background\12.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 16 And Dir(App.Path & "\music\background\12.mp3") <> "" Then
        Timemode = 17
    End If
    If Timemode = 17 Then
        For X = 1 To 14
            Label1.Caption = "���ڼ��\buildings\" & X & ".ini"
            DoEvents
            If Dir(App.Path & "\buildings\" & X & ".ini") = "" Then MsgBox "ȱ��\buildings\" & X & ".ini�ļ��������°�װ�˳���"
            If Dir(App.Path & "\buildings\" & X & ".ini") = "" Then End
        Next
        Timemode = 18
    End If
    Label1.Caption = "���ڼ��\music\begin.mp3�ļ�"
    DoEvents
    If Timemode = 18 And Dir(App.Path & "\music\begin.mp3") = "" Then
        MsgBox "ȱ��\music\begin.mp3�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 18 And Dir(App.Path & "\music\begin.mp3") <> "" Then
        Timemode = 19
    End If
    Label1.Caption = "���ڼ��\picture\logo.gif"
    DoEvents
    If Timemode = 19 And Dir(App.Path & "\picture\logo.gif") = "" Then
        MsgBox "ȱ��\picture\logo.gif�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 19 And Dir(App.Path & "\picture\logo.gif") <> "" Then
        Timemode = 20
    End If
    
    If Timemode = 20 Then
        
        For xa = 1 To 38
            Label1.Caption = "���ڼ��\picture\item\" & xa & ".gif�ļ�"
            DoEvents
            If Dir(App.Path & "\picture\item\" & xa & ".gif") = "" Then MsgBox "ȱ��\picture\item\" & xa & ".gif�ļ��������°�װ�˳���"
            If Dir(App.Path & "\picture\item\" & xa & ".gif") = "" Then End
        Next
        Timemode = 21
    End If
    
    If Timemode = 21 Then
        
        For xb = 1 To 38
            Label1.Caption = "���ڼ��\picture\item\" & xb
            DoEvents
            If Dir(App.Path & "\picture\item\" & xb) = "" Then MsgBox "ȱ��\picture\item\" & xb & "�ļ��������°�װ�˳���"
            If Dir(App.Path & "\picture\item\" & xb) = "" Then End
        Next
        Timemode = 22
    End If
    Label1.Caption = "���ڼ��\picture\shape\del.gif�ļ�"
    DoEvents
    If Timemode = 22 And Dir(App.Path & "\picture\shape\del.gif") = "" Then
        MsgBox "ȱ��\picture\shape\del.gif�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 22 And Dir(App.Path & "\picture\shape\del.gif") <> "" Then
        Timemode = 23
    End If
    Label1.Caption = "���ڼ��\picture\shape\1600.jpg�ļ�"
    DoEvents
    If Timemode = 23 And Dir(App.Path & "\picture\shape\1600.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1600.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 23 And Dir(App.Path & "\picture\shape\1600.jpg") <> "" Then
        Timemode = 24
    End If
    Label1.Caption = "���ڼ��\picture\shape\800.jpg�ļ�"
    DoEvents
    If Timemode = 24 And Dir(App.Path & "\picture\shape\800.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\800.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 24 And Dir(App.Path & "\picture\shape\800.jpg") <> "" Then
        Timemode = 25
    End If
    Label1.Caption = "���ڼ��\picture\shape\1024.jpg�ļ�"
    DoEvents
    If Timemode = 25 And Dir(App.Path & "\picture\shape\1024.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1024.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 25 And Dir(App.Path & "\picture\shape\1024.jpg") <> "" Then
        Timemode = 26
    End If
    Label1.Caption = "���ڼ��\picture\shape\1152.jpg�ļ�"
    DoEvents
    If Timemode = 26 And Dir(App.Path & "\picture\shape\1152.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1152.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 26 And Dir(App.Path & "\picture\shape\1152.jpg") <> "" Then
        Timemode = 27
    End If
    Label1.Caption = "���ڼ��\picture\shape\1280.jpg�ļ�"
    DoEvents
    If Timemode = 27 And Dir(App.Path & "\picture\shape\1280.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1280.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 27 And Dir(App.Path & "\picture\shape\1280.jpg") <> "" Then
        Timemode = 28
    End If
    Label1.Caption = "���ڼ��\picture\shape\1360.jpg�ļ�"
    DoEvents
    If Timemode = 28 And Dir(App.Path & "\picture\shape\1360.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1360.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 28 And Dir(App.Path & "\picture\shape\1360.jpg") <> "" Then
        Timemode = 29
    End If
    Label1.Caption = "���ڼ��\picture\shape\1366.jpg�ļ�"
    DoEvents
    If Timemode = 29 And Dir(App.Path & "\picture\shape\1366.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1366.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 29 And Dir(App.Path & "\picture\shape\1366.jpg") <> "" Then
        Timemode = 30
    End If
    Label1.Caption = "���ڼ��\picture\shape\1440.jpg�ļ�"
    DoEvents
    If Timemode = 30 And Dir(App.Path & "\picture\shape\1440.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1440.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 30 And Dir(App.Path & "\picture\shape\1440.jpg") <> "" Then
        Timemode = 31
    End If
    Label1.Caption = "���ڼ��\picture\shape\1920.jpg�ļ�"
    DoEvents
    If Timemode = 31 And Dir(App.Path & "\picture\shape\1920.jpg") = "" Then
        MsgBox "ȱ��\picture\shape\1920.jpg�ļ��������°�װ�˳���"
        End
    ElseIf Timemode = 31 And Dir(App.Path & "\picture\shape\1920.jpg") <> "" Then
        Timemode = 32
    End If
    Label1.Caption = "�����ɣ�����������Ϸ"
    DoEvents
    If Timemode = 32 Then
        Form3.Show
        Unload Me
    End If
End Sub
