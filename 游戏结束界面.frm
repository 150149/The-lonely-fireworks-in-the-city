VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form31 
   BorderStyle     =   0  'None
   Caption         =   "Form31"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "��Ϸ��������.frx":0000
   LinkTopic       =   "Form31"
   Picture         =   "��Ϸ��������.frx":08CA
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   4455
      Left            =   6000
      OleObjectBlob   =   "��Ϸ��������.frx":7B048
      TabIndex        =   17
      Top             =   3360
      Width           =   10095
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   600
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   615
      Left            =   6840
      TabIndex        =   10
      Top             =   8280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "�ص���ʼ����"
      Style_Border    =   0
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar4 
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   -2147483633
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar3 
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   -2147483633
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar2 
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   3840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   -2147483633
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar1 
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   -2147483633
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   720
      TabIndex        =   11
      Top             =   360
      Width           =   2055
      ExtentX         =   3625
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image Image2 
      Height          =   4935
      Left            =   5160
      Picture         =   "��Ϸ��������.frx":7CB49
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   11415
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      Height          =   255
      Left            =   2280
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ǯ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�罻����"
      Height          =   180
      Left            =   2280
      TabIndex        =   8
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�۲���"
      Height          =   180
      Left            =   2280
      TabIndex        =   6
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   2280
      TabIndex        =   5
      Top             =   4560
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϸ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   6960
      TabIndex        =   0
      Top             =   2160
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   1680
      Picture         =   "��Ϸ��������.frx":7D366
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3975
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long
Private Const WS_VERSION_REQD = &H101
Private cv As String
Public szMyText, szMyText2 As String
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private timermode As Integer


Public Function IsConnectedState() As Boolean
    Dim udtWSAD As WSADATA
    Call WSAStartup(WS_VERSION_REQD, udtWSAD)
    IsConnectedState = CBool(gethostbyname("www.baidu.com"))
    Call WSACleanup
End Function
Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������
    Form31.Height = Screen.Height
    Form31.Width = Screen.Width
    PProgressBar1.Value = CInt(Form1.Shejiao) / 1000
    PProgressBar2.Value = CInt(Form1.Zhili) / 1000
    PProgressBar3.Value = CInt(Form1.Guanchali) / 1000
    PProgressBar4.Value = CInt(Form1.Naili) / 1000
    Label3.Caption = "�罻������" & CStr(CInt(Form1.Shejiao))
    Label4.Caption = "������" & CStr(CInt(Form1.Zhili))
    Label5.Caption = "�۲�����" & CStr(CInt(Form1.Guanchali))
    Label6.Caption = "������" & CStr(CInt(Form1.Naili))
    Label2.Caption = "��Ǯ:" & Form1.money
    Label8.Caption = "����Ʒ�ʣ�" & Form1.Daodepingzhi
    Label9.Caption = "��ҵ����" & Form1.jingye
    Label10.Caption = "����Ʒ��" & Form1.dandang
    Label11.Caption = "����Ʒ��" & Form1.chengxin
    Label12.Caption = "��ѧ�̶�" & Form1.haoxue
    MSChart2.Row = 1
    MSChart2.data = Form1.Shejiao
    MSChart2.Row = 2
    MSChart2.data = Form1.Zhili
    MSChart2.Row = 3
    MSChart2.data = Form1.Guanchali
    MSChart2.Row = 4
    MSChart2.data = Form1.Naili
    MSChart2.Row = 5
    MSChart2.data = Form1.Daodepingzhi
    MSChart2.Row = 6
    MSChart2.data = Form1.jingye
    MSChart2.Row = 7
    MSChart2.data = Form1.dandang
    MSChart2.Row = 8
    MSChart2.data = Form1.chengxin
    MSChart2.Row = 9
    MSChart2.data = Form1.haoxue
    MSChart2.Plot.Backdrop.Fill.Brush.FillColor.Set 255, 255, 255               '���ñ���ˢ����ɫ
    MSChart2.Plot.Backdrop.Fill.Style = VtFillStyleBrush                        'ˢ����
    MSChart2.Plot.Light.EdgeVisible = False
End Sub
'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim I     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '���洰�������ű���
    scaleY = FormName.ScaleHeight / FormOldHeight                               '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For I = 0 To 4
            '��ȡ�ؼ���ԭʼλ�����С
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(I) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(I) = 0
            End If
            '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next I
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me)
    'ȷ������ı�ʱ�ؼ���֮�ı�
    Me.WebBrowser1.Left = 0
    Me.WebBrowser1.Top = 0
    Me.WebBrowser1.Width = 0
    Me.WebBrowser1.Height = 0
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


Private Sub PButton1_Click()
    Dim rbaoshidu As String
    Dim ryinshui As String
    Dim rtili As String
    Dim rmoney As String
    Dim rbaoshidux As String
    Dim ryinshuix As String
    Dim rtilix As String
    Dim rmoneyx As String
    Dim rdaya As String
    Dim rhoura As String
    Dim rmina As String
    Dim rseca As String
    Dim write1 As Long
    Dim rShejiao As String
    Dim rZhili As String
    Dim rGuanchali As String
    Dim rNaili As String
    Dim rdaodepingzhi As String
    Dim rjingye As String
    Dim rhaoxue As String
    Dim rdandang As String
    Dim rchengxin As String
    Dim rdj As String
    Dim rzhiwei As String
    rzhiwei = Form1.Zhiwei
    rdj = CStr(Form1.Dj)
    rbaoshidu = CStr(Form1.baoshidu)
    ryinshui = CStr(Form1.yinshui)
    rtili = CStr(Form1.tili)
    rmoney = CStr(Form1.money)
    rbaoshidux = CStr(Form3.baoshidux)
    ryinshuix = CStr(Form3.yinshuix)
    rtilix = CStr(Form3.tilix)
    rmoneyx = CStr(Form3.moneyx)
    rdaya = CStr(Form3.daya)
    rhoura = CStr(Form1.houra)
    rmina = CStr(Form1.mina)
    rseca = CStr(Form1.seca)
    rShejiao = CStr(Form1.Shejiao)
    rZhili = CStr(Form1.Zhili)
    rGuanchali = CStr(Form1.Guanchali)
    rNaili = CStr(Form1.Naili)
    rdaodepingzhi = CStr(Form1.Daodepingzhi)
    rjingye = CStr(Form1.jingye)
    rhaoxue = CStr(Form1.haoxue)
    rdandang = CStr(Form1.dandang)
    rchengxin = CStr(Form1.chengxin)
    If Form1.moshi = 3 And IsConnectedState Then
        WebBrowser1.Navigate "http://notepad.live/kbrt6"
        write1 = WritePrivateProfileString("human", "baoshidu", rbaoshidu, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "yinshui", ryinshui, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "tili", rtili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "money", rmoney, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "baoshidux", rbaoshidux, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "yinshuix", ryinshuix, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "tilix", rtilix, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "moneyx", rmoneyx, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "daya", rdaya, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "houra", rhoura, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "mina", rmina, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "seca", rseca, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "top", CStr(Form1.Label1.Top), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "left", CStr(Form1.Label1.Left), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("save", "changetime1", CStr(Date), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("save", "changetime2", CStr(Time), App.Path & "\save\save" & Form3.saveid & ".fsave")
        
        write1 = WritePrivateProfileString("human", "Shejiao", rShejiao, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Zhili", rZhili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Guanchali", rGuanchali, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Naili", rNaili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "daodepingzhi", rdaodepingzhi, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "jingye", rjingye, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "haoxue", rhaoxue, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "dandang", rdandang, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "chengxin", rchengxin, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("text", "dj", rdj, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "zhiwei", rzhiwei, App.Path & "\save\save" & Form3.saveid & ".fsave")
        
        Timer1.Enabled = True
        Timer1.Interval = 300
        Exit Sub
    End If
    
    write1 = WritePrivateProfileString("human", "baoshidu", rbaoshidu, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "yinshui", ryinshui, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "tili", rtili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "money", rmoney, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "baoshidux", rbaoshidux, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "yinshuix", ryinshuix, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "tilix", rtilix, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "moneyx", rmoneyx, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "daya", rdaya, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "houra", rhoura, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "mina", rmina, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "seca", rseca, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "top", CStr(Form1.Label1.Top), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "left", CStr(Form1.Label1.Left), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("save", "changetime1", CStr(Date), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("save", "changetime2", CStr(Time), App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    write1 = WritePrivateProfileString("human", "Shejiao", rShejiao, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Zhili", rZhili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Guanchali", rGuanchali, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Naili", rNaili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "daodepingzhi", rdaodepingzhi, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "jingye", rjingye, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "haoxue", rhaoxue, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "dandang", rdandang, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "chengxin", rchengxin, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("text", "dj", rdj, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "zhiwei", rzhiwei, App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    Form3.baoshidux = 0
    Form3.yinshuix = 0
    Form3.tilix = 0
    
    Form4.Show
    Form18.Show
    Form18.Top = Form3.Label1.Top
    Form18.Left = Form3.Label1.Left
    Form18.Height = Form3.Label1.Height
    Form18.Width = Form3.Label1.Width
    Form3.Show
    Form3.music = 0
    Form3.firstrise
    Unload Form1
    Unload Form7
    Unload Form6
    Unload Form31
End Sub

Private Sub Timer1_Timer()
    
    '--------------------------------------------��������------------------------------------------
    If WebBrowser1.busy Then
        Debug.Print ("��ҳδ�������")
        Exit Sub
    Else
        Debug.Print ("��ҳ�������")
        Debug.Print "-------------------------------------------------------------------------------"
        Timer1.Enabled = False
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        Dim vDoc, X, VTag
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '������б�ǩ
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Value = "�ύ" Then VTag.Click                          '����ύ�ˣ�һ�ж�OK��
            End If
        Next X
        Timer3.Enabled = True
        Timer3.Interval = 1500
        Debug.Print ("��¼״̬��" & "yes")
    End If
    If timermode = 1 Then
        Form4.Show
        Form18.Show
        Form18.Top = Form3.Label1.Top
        Form18.Left = Form3.Label1.Left
        Form18.Height = Form3.Label1.Height
        Form18.Width = Form3.Label1.Width
        Form3.Show
        Form3.music = 0
        Form3.firstrise
        Unload Form1
        Unload Form7
        Unload Form6
        Unload Form31
        Exit Sub
    End If
End Sub

Private Sub Timer3_Timer()
    If WebBrowser1.busy Then
        Debug.Print ("��ҳδ�������timer3")
        Exit Sub
    Else
        Timer3.Enabled = False
        Debug.Print ("��ʼ��ȡ����timer3")
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "��ʼ"                                                 '����Ҫ���ҵ��ַ�����ͷ
        szFindStrEnd = "����"                                                   '����Ҫ���ҵ��ַ�����β
        
        szText = WebBrowser1.Document.body.innerText                            '�õ��������֣���ʱ��ģ�壬ʵ��ʹ���л���ȥWebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '�ҿ�ͷ�ַ���
        If nBegin > 0 Then                                                      '���������ҵ���ͷ�˲ż���
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '�ҽ�β�ַ���
            If nEnd > nBegin Then                                               '��β����ȿ�ͷ��λ�ô�
                
                '�������ҵ��ַ���ģʽ��ע�͵������2��
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '������Ҫ��ȡ���ַ�������,���Ҫ�������ҵ��ַ�������1�У�ע������2��
                
                '���������ҵ��ַ���ģʽ
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '������������ҵ��ַ���������2��
                nBegin = nBegin + Len(szFindStrBegin)                           '������������ҵ��ַ���������2��
                
                szMyText = Mid(szText, nBegin, nLength)                         'ȡ����before then.���� "test" �м�Ķ���
                Debug.Print ("��ȡ���ݣ�" & szMyText)
            End If
        End If
        Randomize
        cv = Int(Rnd * (10000 - 1 + 1)) + 1
        Open App.Path & "\" & cv For Output As #1
        Print #1, szMyText
        Close #1
        Dim write1 As Long
        Dim mAES As New mAES                                                    ''AES���������
        Dim fasongneirong As String
        Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Binary As #1
        fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        Kill App.Path & "\save\save8.fsave"
        fasongneirong = mAES.EncryptStr(fasongneirong, "150149", Bit128, Bit128, False)
        write1 = WritePrivateProfileString("save", mAES.EncryptStr(Form30.mingzi, "150149", Bit128, Bit128, False), fasongneirong, App.Path & "\" & cv)
        Open App.Path & "\" & cv For Binary As #1
        fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        Kill App.Path & "\" & cv
        Dim vDoc, VTag, mType As String, mTagName As String
        Dim ia As Integer
        Set vDoc = WebBrowser1.Document
        For ia = 0 To vDoc.All.Length - 1
            Select Case UCase(vDoc.All(ia).tagName)
            Case "TEXTAREA"                                                     '"TEXTAREA" ��ǩ,�ı������д
                Set VTag = vDoc.All(ia)
                VTag.Value = "��ʼ" & fasongneirong & "����"                    '��Text1�е���������
                Timer1.Enabled = True
                Timer1.Interval = 300
                WebBrowser1.Navigate "http://notepad.live/kbrt6"
                Debug.Print "�������ݳɹ�"
                timermode = 1
            End Select
            
        Next ia
        
        
    End If
End Sub


