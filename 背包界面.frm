VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   9105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13545
   Icon            =   "��������.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   9105
   ScaleWidth      =   13545
   StartUpPosition =   2  '��Ļ����
   Begin �о��й°����̻�.PProgressBar PProgressBar4 
      Height          =   255
      Left            =   11040
      TabIndex        =   14
      Top             =   3960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   16777215
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar3 
      Height          =   255
      Left            =   11040
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   16777215
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar2 
      Height          =   255
      Left            =   11040
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   16777215
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar1 
      Height          =   255
      Left            =   11040
      TabIndex        =   7
      Top             =   2880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   16777215
   End
   Begin VB.Timer Timer2 
      Left            =   12960
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   12960
      Top             =   240
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   600
      ScaleHeight     =   6855
      ScaleWidth      =   8415
      TabIndex        =   4
      Top             =   840
      Width           =   8415
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   9495
         Left            =   0
         ScaleHeight     =   9495
         ScaleWidth      =   8415
         TabIndex        =   5
         Top             =   0
         Width           =   8415
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1005
            Left            =   6000
            TabIndex        =   6
            Top             =   3240
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin �о��й°����̻�.PVScrollBar PVScrollBar1 
      Height          =   7215
      Left            =   9240
      TabIndex        =   3
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   12726
      Color_Top       =   14737632
      Color_Back      =   16777215
      Size            =   0.4
      Color_Border    =   16777215
   End
   Begin �о��й°����̻�.PButton PButton2 
      Height          =   855
      Left            =   10200
      TabIndex        =   2
      Top             =   6720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "����"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   855
      Left            =   10200
      TabIndex        =   1
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "ʹ��"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  ���� "
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   7935
      Left            =   -120
      Picture         =   "��������.frx":08CA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   10095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ܵ�4"
      Height          =   180
      Left            =   9960
      TabIndex        =   13
      Top             =   3960
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ܵ�3"
      Height          =   180
      Left            =   9960
      TabIndex        =   11
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ܵ�2"
      Height          =   180
      Left            =   9960
      TabIndex        =   9
      Top             =   3240
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ܵ�1"
      Height          =   180
      Left            =   9960
      TabIndex        =   8
      Top             =   2880
      Width           =   630
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   9960
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackColor       =   &H000080FF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12600
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   8655
      Left            =   -240
      Picture         =   "��������.frx":10E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private beibaomax As String
Private xuanzhong As Integer
Private zhizhen As Integer
Private shubiaox As Integer
Private shubiaoy As Integer
Private shubiaox2 As Integer
Private shubiaoy2 As Integer
Dim I&, J&, LeftPos&, TopPos&                                                   '���������̬ &����Long ���� Dim i& �൱�� Dim i As Long
Public form5open As Integer
Private Lastzhizhen As Integer

Private Sub Form_Load()
    Me.BorderStyle = 0
    KeyPreview = True
    form5open = 0
    For J = 0 To 7                                                              '����10�ŵ�ѭ��
        For I = 1 To 6                                                          '����10�ŵ�ѭ��
            Load Image1(I + J * 6)                                              'װ�ر�ǩ����Ԫ��
            Image1(I + J * 6).Top = TopPos                                      '���������ǩ�Ķ���TOPλ��
            Image1(I + J * 6).Left = LeftPos                                    '���������ǩ�����LEFTλ��
            Image1(I + J * 6).Visible = True                                    '��̬��ӵ�LabelĬ��Ϊ���ɼ� ��������Ҫ�����ɼ�
            LeftPos = LeftPos + Image1(0).Width + 300                           '��ǩLabel��X����������һ��Label1(1)�Ŀ���ټ��ϼ��300
        Next I
        TopPos = TopPos + Image1(0).Height + 300                                '��ǩLabel��Y����������һ��Label1(1)�ĸ߶��ټ��ϼ��300
        LeftPos = Image1(1).Left                                                '��ʼѭ��ǰ�������LeftPosΪ ��һ����ǩ��Left����λ��
    Next J
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
    Dim read_OK As Long
    beibaomax = String(10, 0)
    read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", beibaomax, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    Dim beibaomaxa As Integer
    Dim beibaomaxb As Integer
    beibaomaxb = CInt(beibaomax)
    Dim picnumber As String
    picnumber = String(10, 0)
    For beibaomaxa = 1 To beibaomaxb
        read_OK = GetPrivateProfileString("beibao", "beibao" & beibaomaxa, "0", picnumber, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If picnumber = 0 Then
        Else
            Image1(beibaomaxa).Picture = LoadPicture(App.Path & "\picture\item\" & CInt(picnumber) & ".gif")
        End If
    Next
    xuanzhong = -1
    zhizhen = -1
    Image1(0).Visible = False
    Timer1.Enabled = True
    Timer1.Interval = 300
    
    PProgressBar1.Value = CInt(Form1.Shejiao) / 1000
    PProgressBar2.Value = CInt(Form1.Zhili) / 1000
    PProgressBar3.Value = CInt(Form1.Guanchali) / 1000
    PProgressBar4.Value = CInt(Form1.Naili) / 1000
    Label3.Caption = "�罻������" & CInt(Form1.Shejiao)
    Label4.Caption = "������" & CInt(Form1.Zhili)
    Label5.Caption = "�۲�����" & CInt(Form1.Guanchali)
    Label6.Caption = "������" & CInt(Form1.Naili)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.BackColor = &H80FF&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    For I = 1 To 30                                                             '�ͷŶ�̬��ӵ�100��Label�ؼ�
        Unload Image1(I)                                                        'ж��������ӵı�ǩ����ؼ�
    Next I
    Unload Form12
    Unload Form5
End Sub

Private Sub Image1_Click(Index As Integer)
    If xuanzhong = -1 Then
    Else
        Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50), vbWhite
        Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
        Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
        Picture1.Line (Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
    End If
    xuanzhong = Index
    Picture1.DrawWidth = 3
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50), vbRed
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
    Picture1.Line (Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = Index
    If zhizhen = -1 Then
    Else
        shubiaox = X
        shubiaoy = Y
    End If
    Label2.Visible = True
    Label2.Left = shubiaox + shubiaox2
    Label2.Top = shubiaoy + shubiaoy2
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.BackColor = &H80FF&
End Sub

Private Sub Label7_Click()
    Dim first1 As String
    Dim read_OK As Long
    Dim write1 As Long
    first1 = String(10, 0)
    read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Replace(first1, Chr(0), "") = "9" And Form1.moshi = 0 Then
        write1 = WritePrivateProfileString("guide", "1", "10", App.Path & "\save\save" & Form3.saveid & ".fsave")
    ElseIf Replace(first1, Chr(0), "") = "7" And Form1.moshi = 2 Then
        write1 = WritePrivateProfileString("guide", "1", "8", App.Path & "\save\save" & Form3.saveid & ".fsave")
    End If
    
    KeyPreview = False
    Label7.BackColor = &H80FF&
    Form1.KeyPreview = True
    form5open = 1
    Timer1.Enabled = False
    Form1.formback = 1
    Form12.Hide
    Form5.Hide
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.BackColor = &H40C0&
End Sub

Private Sub PButton1_Click()
    Dim read_OK As Long
    Dim picnumber As String
    picnumber = String(10, 0)
    If xuanzhong = -1 Then
        Tishim = Tishi("��ʾ", "δѡ����Ʒ")
        Form3.Tishiback = 5
        Exit Sub
    Else
        read_OK = GetPrivateProfileString("beibao", "beibao" & xuanzhong, "0", picnumber, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If picnumber = 0 Then
            Tishim = Tishi("��ʾ", "�ø�������Ʒ")
            Form3.Tishiback = 5
        Else
            Dim xg1 As String
            Dim xg2 As String
            Dim xg3 As String
            Dim xg4 As String
            xg1 = String(10, 0)                                                 '��ʳ��
            xg2 = String(10, 0)                                                 '��ˮ
            xg3 = String(10, 0)                                                 '����
            xg4 = String(10, 0)                                                 '��Ǯ
            read_OK = GetPrivateProfileString(picnumber, "baoshidu", "0", xg1, 256, App.Path & "\item.ini")
            read_OK = GetPrivateProfileString(picnumber, "yinshui", "0", xg2, 256, App.Path & "\item.ini")
            read_OK = GetPrivateProfileString(picnumber, "tili", "0", xg3, 256, App.Path & "\item.ini")
            read_OK = GetPrivateProfileString(picnumber, "moneyadd", "0", xg4, 256, App.Path & "\item.ini")
            
            If Form1.baoshidu + CInt(xg1) > 1000 Then
                Tishim = Tishi("��ʾ", "��Բ�����")
                Form3.Tishiback = 5
            ElseIf Form1.yinshui + CInt(xg2) > 1000 Then
                Tishim = Tishi("��ʾ", "�㲢���ڿ�")
                Form3.Tishiback = 5
            ElseIf Form1.tili + CInt(xg3) > 1000 Then
                Tishim = Tishi("��ʾ", "��Բ�����")
                Form3.Tishiback = 5
            Else
                Dim write1 As Long
                write1 = WritePrivateProfileString("beibao", "beibao" & xuanzhong, "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                Form1.baoshidu = Form1.baoshidu + CInt(xg1)
                Form1.yinshui = Form1.yinshui + CInt(xg2)
                Form1.tili = Form1.tili + CInt(xg3)
                Form1.money = Form1.money + CInt(xg4)
                Image1(xuanzhong).Picture = LoadPicture("")
                Tishim = Tishi("��ʾ", "  ʹ�óɹ�")
                Form3.Tishiback = 5
            End If
            
            
            
            
            
        End If
    End If
    Picture1.DrawWidth = 3
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50), vbWhite
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
    Picture1.Line (Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), vbWhite
    xuanzhong = 0
End Sub

Private Sub PButton2_Click()
    Dim read_OK As Long
    Dim picnumber As String
    picnumber = String(10, 0)
    If xuanzhong = 0 Then
        Tishim = Tishi("��ʾ", "δѡ����Ʒ")
        Form3.Tishiback = 5
    ElseIf xuanzhong <> -1 Then
        read_OK = GetPrivateProfileString("beibao", "beibao" & xuanzhong, "0", picnumber, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        Dim write1 As Long
        write1 = WritePrivateProfileString("beibao", "beibao" & xuanzhong, "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        Image1(xuanzhong).Picture = LoadPicture("")
        Tishim = Tishi("��ʾ", "�����ɹ�")
        Form3.Tishiback = 5
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = -1
    shubiaox2 = X
    shubiaoy2 = Y
End Sub



Private Sub PVScrollBar1_Change(nValue As Single)
    Picture1.Top = -(PVScrollBar1.Value * PVScrollBar1.Height / 1)
    If xuanzhong = 0 Then
    Else
        Dim Index As Integer
        Index = xuanzhong
        Picture1.DrawWidth = 3
        If Index = -1 Then
        Else
            Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50), vbRed
            Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
            Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
            Picture1.Line (Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
        End If
    End If
End Sub

Private Sub Readiteminformation()
    
    Dim read_OK As Long
    Dim xg0 As Integer
    Dim ritem As String
    ritem = String(10, 0)
    read_OK = GetPrivateProfileString("beibao", "beibao" & zhizhen, "0", ritem, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    Dim xg1 As String
    Dim xg2 As String
    Dim xg3 As String
    Dim xg4 As String
    Dim xg5 As String
    xg0 = Replace(ritem, Chr(0), "")
    xg1 = String(10, 0)
    xg2 = String(10, 0)
    xg3 = String(10, 0)
    xg4 = String(10, 0)
    xg5 = String(10, 0)
    read_OK = GetPrivateProfileString(xg0, "money", "0", xg1, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "name", "0", xg2, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "baoshidu", "0", xg3, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "yinshui", "0", xg4, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "tili", "0", xg5, 256, App.Path & "\item.ini")
    
    Dim xg6 As String
    Dim xg7 As String
    Dim xg8 As String
    Dim xg9 As String
    Dim xg10 As String
    xg6 = CInt(xg1)
    xg7 = CInt(xg5)
    xg8 = CInt(xg3)
    xg9 = CInt(xg4)
    xg2 = Replace(xg2, Chr(0), "")
    Label2.Caption = xg2 & Chr(13) & "��Ǯ:" & xg6 & Chr(13) & "�ɼӱ�ʳ��" & xg8 & Chr(13) & "�ɼ���ˮ��" & xg9 & Chr(13) & "�ɼӾ���ֵ" & xg7
    
    
End Sub

Private Sub Timer1_Timer()
    If zhizhen <> -1 And zhizhen <> Lastzhizhen Then
        Readiteminformation
        Lastzhizhen = zhizhen
    ElseIf zhizhen = -1 Then
        Label2.Visible = False
    End If
End Sub


