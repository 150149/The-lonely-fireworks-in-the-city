VERSION 5.00
Begin VB.Form Form28 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form28"
   ClientHeight    =   12765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16335
   LinkTopic       =   "Form28"
   ScaleHeight     =   12765
   ScaleWidth      =   16335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PButton PButton13 
      Height          =   735
      Left            =   12120
      TabIndex        =   17
      Top             =   8040
      Width           =   1935
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "极高难度关卡9"
      Font_Size       =   15
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton12 
      Height          =   735
      Left            =   12240
      TabIndex        =   16
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "高难度关卡8"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton11 
      Height          =   735
      Left            =   12240
      TabIndex        =   15
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "高难度关卡7"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton7 
      Height          =   735
      Left            =   12240
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "中难度关卡6"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Left            =   8400
      TabIndex        =   13
      Top             =   9600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "中难度关卡5"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton5 
      Height          =   735
      Left            =   8400
      TabIndex        =   12
      Top             =   8040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "中难度关卡4"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton4 
      Height          =   735
      Left            =   8400
      TabIndex        =   11
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2778
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12874308
      Text            =   "低难度关卡3"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   735
      Left            =   8520
      TabIndex        =   9
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "低难度关卡1"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   240
   End
   Begin 市井中孤傲的烟火.PButton PButton8 
      Height          =   1215
      Left            =   1680
      TabIndex        =   0
      Top             =   9960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2778
      Text            =   "返回"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   735
      Left            =   8400
      TabIndex        =   10
      Top             =   4920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "低难度关卡2"
      Font_Size       =   15
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton14 
      Height          =   735
      Left            =   12000
      TabIndex        =   18
      Top             =   9600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Back      =   12874308
      Color_Back_Down =   12874308
      Color_Begin     =   12874308
      Color_End       =   12874308
      Color_Text      =   16777215
      Color_Text_MouseMoved=   12632256
      Text            =   "极高难度关卡10"
      Font_Size       =   15
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF13 
      Height          =   1725
      Left            =   10800
      Top             =   9120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":0000
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF12 
      Height          =   1725
      Left            =   10800
      Top             =   7560
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":6E06
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF11 
      Height          =   1725
      Left            =   10800
      Top             =   6000
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":DC0C
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF10 
      Height          =   1725
      Left            =   10800
      Top             =   4440
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":14A12
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF9 
      Height          =   1725
      Left            =   10800
      Top             =   2880
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":1B818
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF8 
      Height          =   1725
      Left            =   6960
      Top             =   9120
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":2261E
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF7 
      Height          =   1725
      Left            =   6960
      Top             =   7560
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":29424
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF6 
      Height          =   1725
      Left            =   6960
      Top             =   6000
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":3022A
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF5 
      Height          =   1725
      Left            =   6960
      Top             =   4440
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":37030
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF4 
      Height          =   1725
      Left            =   6960
      Top             =   2880
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   3043
      GIF             =   "竞技大厅.frx":3DE36
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "金融职业技能测试"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   8640
      Width           =   2280
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF3 
      Height          =   1770
      Left            =   1800
      Top             =   8040
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   3122
      GIF             =   "竞技大厅.frx":44C3C
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "驾驶职业技能测试"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   6720
      Width           =   2280
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF2 
      Height          =   1770
      Left            =   1800
      Top             =   6120
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   3122
      GIF             =   "竞技大厅.frx":4C4EA
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "烹饪职业技能测试"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MouseIcon       =   "竞技大厅.frx":53D98
      MousePointer    =   1  'Arrow
      TabIndex        =   19
      Top             =   4800
      Width           =   2280
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF1 
      Height          =   1770
      Left            =   1800
      Top             =   4200
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   3122
      GIF             =   "竞技大厅.frx":540A2
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "社交能力"
      Height          =   180
      Left            =   1800
      TabIndex        =   8
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "智力"
      Height          =   180
      Left            =   1800
      TabIndex        =   7
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "观察力"
      Height          =   180
      Left            =   1800
      TabIndex        =   6
      Top             =   3360
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "耐力"
      Height          =   180
      Left            =   1800
      TabIndex        =   5
      Top             =   3720
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   12585
      Left            =   -4200
      Picture         =   "竞技大厅.frx":5B950
      Stretch         =   -1  'True
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Jingji As Boolean
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
Public Mark As Integer


Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    PProgressBar1.Value = CInt(Form1.Shejiao) / 1000
    PProgressBar2.Value = CInt(Form1.Zhili) / 1000
    PProgressBar3.Value = CInt(Form1.Guanchali) / 1000
    PProgressBar4.Value = CInt(Form1.Naili) / 1000
    Label12.Caption = "社交能力：" & CInt(Form1.Shejiao)
    Label11.Caption = "智力：" & CInt(Form1.Zhili)
    Label10.Caption = "观察力：" & CInt(Form1.Guanchali)
    Label9.Caption = "耐力：" & CInt(Form1.Naili)
    If Form1.Pengren = False Then Label1.Caption = "烹饪技能职业测试" & Chr(13) & "(未通过)"
    If Form1.Pengren = True Then Label1.Caption = "烹饪技能职业测试" & Chr(13) & "(已通过)"
    If Form1.Driven = False Then Label2.Caption = "驾驶技能职业测试" & Chr(13) & "(未通过)"
    If Form1.Driven = True Then Label2.Caption = "驾驶技能职业测试" & Chr(13) & "(已通过)"
    If Form1.Jinrong = False Then Label3.Caption = "金融技能职业测试" & Chr(13) & "(未通过)"
    If Form1.Jinrong = True Then Label3.Caption = "金融技能职业测试" & Chr(13) & "(已通过)"
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
End Sub
'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim I     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '保存窗体宽度缩放比例
    scaleY = FormName.ScaleHeight / FormOldHeight                               '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For I = 0 To 4
            '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(I) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(I) = 0
            End If
            '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next I
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

Public Function IsConnectedState() As Boolean
    Dim udtWSAD As WSADATA
    Call WSAStartup(WS_VERSION_REQD, udtWSAD)
    IsConnectedState = CBool(gethostbyname("www.baidu.com"))
    Call WSACleanup
End Function

Private Sub Label1_Click()
    Timer1.Enabled = True
    Timer1.Interval = 500
    Mark = 0
End Sub

Private Sub Label2_Click()
    Timer2.Enabled = True
    Timer2.Interval = 500
    Mark = 0
End Sub

Private Sub Label3_Click()
    Timer3.Enabled = True
    Timer3.Interval = 500
    Mark = 0
End Sub

Private Sub PButton11_Click()
    Form29.Aiblood = 800
    Form41.Show
End Sub

Private Sub PButton12_Click()
    Form29.Aiblood = 1000
    Form41.Show
End Sub

Private Sub PButton13_Click()
    Form29.Aiblood = 2000
    Form41.Show
End Sub

Private Sub PButton14_Click()
    Form29.Aiblood = 5000
    Form41.Show
End Sub

Private Sub PButton2_Click()
    Form29.Aiblood = 20
    Form41.Show
    
    Dim first1 As String
    first1 = String(10, 0)
    read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Replace(first1, Chr(0), "") = "12" Then
        Form37.Show
        Form37.Width = Form37.PPictureBox15.Width
        Form37.Left = Screen.Width - Form37.Width
        Form37.Top = 500
        Form37.PPictureBox1.Visible = False
        Form37.PPictureBox2.Visible = False
        Form37.PPictureBox3.Visible = False
        Form37.PPictureBox4.Visible = False
        Form37.PPictureBox5.Visible = False
        Form37.PPictureBox6.Visible = False
        Form37.PPictureBox7.Visible = False
        Form37.PPictureBox8.Visible = False
        Form37.PPictureBox9.Visible = False
        Form37.PPictureBox10.Visible = False
        Form37.PPictureBox11.Visible = False
        Form37.PPictureBox12.Visible = False
        Form37.PPictureBox13.Visible = False
        Form37.PPictureBox14.Visible = False
        Form37.PPictureBox15.Visible = True
        Form37.PPictureBox16.Visible = False
        Form37.PPictureBox17.Visible = False
        write1 = WritePrivateProfileString("guide", "1", "13", App.Path & "\save\save" & Form3.saveid & ".fsave")
    End If
End Sub

Private Sub PButton3_Click()
    Form29.Aiblood = 40
    Form41.Show
End Sub

Private Sub PButton4_Click()
    Form29.Aiblood = 70
     Form41.Show
End Sub

Private Sub PButton5_Click()
    Form29.Aiblood = 100
    Form41.Show
End Sub

Private Sub PButton6_Click()
    Form29.Aiblood = 200
     Form41.Show
End Sub

Private Sub PButton7_Click()
    Form29.Aiblood = 400
    Form41.Show
End Sub

Private Sub PButton8_Click()

    Form1.Timer1.Enabled = True
    Form1.Timer3.Enabled = True
    Unload Form12
    Unload Me
End Sub

Private Sub Timer1_Timer()
    If Form1.Accident = 0 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 51
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 52 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 53
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 54 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 55
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 56 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 57
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 58 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 59
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 60 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 61
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 62 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 63
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 64 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 65
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 66 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 67
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 68 And Timer2.Enabled = False And Timer3.Enabled = False Then
        Timer1.Enabled = False
        If Mark <= 4 Then
            Tishim = Tishi("提示", "考试未通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        ElseIf Mark > 4 Then
            Form1.Pengren = True
            If Form1.Pengren = True Then Label1.Caption = "烹饪技能职业测试" & Chr(13) & "(已通过)"
            Tishim = Tishi("提示", "考试已通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    If Form1.Accident = 0 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 71
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 72 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 73
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 74 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 75
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 76 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 77
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 78 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 79
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 80 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 81
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 82 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 83
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 84 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 85
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 86 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Form1.Accident = 87
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 88 And Timer1.Enabled = False And Timer3.Enabled = False Then
        Timer2.Enabled = False
        If Mark <= 4 Then
            Tishim = Tishi("提示", "考试未通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        ElseIf Mark > 4 Then
            Form1.Driven = True
            If Form1.Driven = True Then Label2.Caption = "驾驶技能职业测试" & Chr(13) & "(已通过)"
            Tishim = Tishi("提示", "考试已通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        End If
    End If
End Sub

Private Sub Timer3_Timer()
    If Form1.Accident = 0 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 91
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 92 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 93
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 94 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 95
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 96 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 97
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 98 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 99
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 100 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 101
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 102 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 103
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 104 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 105
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 106 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Form1.Accident = 107
        Form12.Show
        Form22.Show
        Form22.Label25.Caption = "测试"
    ElseIf Form1.Accident = 108 And Timer2.Enabled = False And Timer1.Enabled = False Then
        Timer3.Enabled = False
        If Mark <= 4 Then
            Tishim = Tishi("提示", "考试未通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        ElseIf Mark > 4 Then
            Form1.Jinrong = True
            If Form1.Jinrong = True Then Label3.Caption = "金融技能职业测试" & Chr(13) & "(已通过)"
            Tishim = Tishi("提示", "考试已通过，分数" & Mark & "/9")
            Form3.Tishiback = 28
            Form1.Accident = 0
        End If
    End If
End Sub

Private Sub ucAniGIF1_Click()
    Timer1.Enabled = True
    Timer1.Interval = 500
    Mark = 0
End Sub

Private Sub ucAniGIF2_Click()
    Timer2.Enabled = True
    Timer2.Interval = 500
    Mark = 0
End Sub

Private Sub ucAniGIF3_Click()
    Timer3.Enabled = True
    Timer3.Interval = 500
    Mark = 0
End Sub
