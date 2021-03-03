VERSION 5.00
Begin VB.Form Form29 
   BorderStyle     =   0  'None
   Caption         =   "Form37"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17475
   Icon            =   "竞技场景.frx":0000
   LinkTopic       =   "Form37"
   ScaleHeight     =   10380
   ScaleWidth      =   17475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer4 
      Left            =   120
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   600
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   1200
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   7320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Text            =   "释放职业技能"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   8280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Text            =   "离开战斗"
      Style_Border    =   0
   End
   Begin 市井中孤傲的烟火.PListBox PListBox1 
      Height          =   2175
      Left            =   1320
      TabIndex        =   2
      Top             =   7200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
   End
   Begin 市井中孤傲的烟火.PTab PTab1 
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   6480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Font_Size       =   11
      Color_Selector_Moved=   16777215
      Texts           =   "饭店|便利店|建筑公司|房地产公司|证券公司"
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   375
      Left            =   10560
      TabIndex        =   4
      Top             =   1320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Color_Top       =   255
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   255
      Left            =   11640
      TabIndex        =   6
      Top             =   8520
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   255
      Left            =   11640
      TabIndex        =   7
      Top             =   7800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar5 
      Height          =   255
      Left            =   11640
      TabIndex        =   8
      Top             =   8160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar6 
      Height          =   255
      Left            =   11640
      TabIndex        =   9
      Top             =   7440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   10080
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4935
   End
   Begin 市井中孤傲的烟火.ucAniGIF ucAniGIF1 
      Height          =   4080
      Left            =   10440
      Top             =   2040
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   7197
      GIF             =   "竞技场景.frx":08CA
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   16
      Top             =   6720
      Width           =   75
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10560
      TabIndex        =   14
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "耐力"
      Height          =   180
      Left            =   10680
      TabIndex        =   13
      Top             =   8520
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "观察力"
      Height          =   180
      Left            =   10680
      TabIndex        =   12
      Top             =   8160
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "智力"
      Height          =   180
      Left            =   10680
      TabIndex        =   11
      Top             =   7800
      Width           =   360
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "社交能力"
      Height          =   180
      Left            =   10680
      TabIndex        =   10
      Top             =   7440
      Width           =   720
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   2520
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   3000
      Left            =   9600
      Picture         =   "竞技场景.frx":2FE8
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   5400
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   1080
      Picture         =   "竞技场景.frx":3805
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   13260
      Left            =   -1440
      Picture         =   "竞技场景.frx":4BC0
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   20145
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Dianji As Integer
Private xuanzhong As Integer
Private Meblood As Integer
Public Aiblood As Integer
Public Aidamage As Integer
Private Tshejiao As Integer
Private Tguanchali As Integer
Private Tnaili As Integer
Private Tzhili As Integer
Private Aimaxblood As Integer
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private p As Integer
Private Rel As Boolean
Private zhizhen As Integer
Private Lastzhizhen As Integer
Public Checktime As Integer

Private Sub Form_Load()
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    Dim read_OK As Long
    Dim rworkmax As String
    rworkmax = String(10, 0)
    read_OK = GetPrivateProfileString("work", "max", "1", rworkmax, 256, App.Path & "\buildings\1.ini")
    Dim rwork As String
    rwork = String(256, 0)
    Dim X As Integer
    For X = 1 To CInt(rworkmax)
        read_OK = GetPrivateProfileString("work", CStr(X), "无工作选项", rwork, 256, App.Path & "\buildings\1.ini")
        If rwork = "" Then
        Else
            PListBox1.AddItem (Replace(rwork, Chr(0), ""))
        End If
    Next
    Dianji = -1
    Tshejiao = Form1.Shejiao
    Tguanchali = Form1.Guanchali
    Tnaili = Form1.Naili
    Tzhili = Form1.Zhili
    Label12.Caption = "社交能力：" & CInt(Tshejiao)
    Label11.Caption = "智力：" & CInt(Tzhili)
    Label10.Caption = "观察力：" & CInt(Tguanchali)
    Label9.Caption = "耐力：" & CInt(Tnaili)
    Aimaxblood = Aiblood
    Meblood = 100
    Rel = True
    Label24.Visible = False
    Shape1.Visible = False
    zhizhen = -1
    '  Call Picture1_paint
    Timer4.Enabled = True
    Timer4.Interval = 1000
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label24.Visible = True Then Label24.Visible = False
    If Shape1.Visible = True Then Shape1.Visible = False
End Sub

Private Sub PButton1_Click()
    If Rel = False Then Exit Sub
    If Tshejiao < 0 Or Tzhili < 0 Or Tguanchali < 0 Or Tnaili < 0 Then Exit Sub
    If PListBox1.ListIndex = -1 Then
        Exit Sub
    Else
        Dim read_OK As Long
        Dim rdamage As String
        Dim rShejiao As String
        Dim rGuanchali As String
        Dim rNaili As String
        Dim rZhili As String
        Dim runlock As String
        rShejiao = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        rZhili = String(10, 0)
        rdamage = String(10, 0)
        runlock = String(10, 0)
        read_OK = GetPrivateProfileString(Dianji + 1, "damage", "0", rdamage, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "useshejiao", "0", rShejiao, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "useguanchali", "0", rGuanchali, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "usenaili", "0", rNaili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "usezhili", "0", rZhili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(CStr("lock" & CStr(xuanzhong + 1)), CStr(PListBox1.ListIndex + 1), "0", runlock, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If CInt(runlock) = 0 Then Exit Sub
        If Tshejiao - CInt(rShejiao) < 0 Or Tzhili - CInt(rZhili) < 0 Or Tguanchali - CInt(rGuanchali) < 0 Or Tnaili - CInt(rNaili) < 0 Then Exit Sub
        Aiblood = Aiblood - CInt(rdamage)
        Label1.Caption = Aiblood
        PProgressBar2.Value = CSng(Aiblood / Aimaxblood)
        If Aiblood <= 0 Then
            Form47.Show
            Form1.money = Form1.money + Aimaxblood
        End If
        Tshejiao = Tshejiao - CInt(rShejiao)
        Tguanchali = Tguanchali - CInt(rGuanchali)
        Tnaili = Tnaili - CInt(rNaili)
        Tzhili = Tzhili - CInt(rZhili)
        If Form1.Shejiao <> 0 Then PProgressBar6.Value = CSng(Tshejiao / Form1.Shejiao)
        If Form1.Zhili <> 0 Then PProgressBar3.Value = CSng(Tzhili / Form1.Zhili)
        If Form1.Guanchali <> 0 Then PProgressBar5.Value = CSng(Tguanchali / Form1.Guanchali)
        If Form1.Naili <> 0 Then PProgressBar4.Value = CSng(Tnaili / Form1.Naili)
        Label12.Caption = "社交能力：" & CInt(Tshejiao)
        Label11.Caption = "智力：" & CInt(Tzhili)
        Label10.Caption = "观察力：" & CInt(Tguanchali)
        Label9.Caption = "耐力：" & CInt(Tnaili)
        Randomize
        If Int(Rnd * (2 - 1 + 1)) + 1 = 2 Then
            Timer1.Enabled = True
            Rel = False
            Timer1.Interval = 90
            p = 8
            Me.BackColor = &H0&
            Image2.Picture = LoadPicture("")
        Else
            Timer3.Enabled = True
            Rel = False
            Timer3.Interval = 90
            p = 9
            Me.BackColor = &H0&
            Image2.Picture = LoadPicture("")
        End If
    End If
    
    Dim first1 As String
    first1 = String(10, 0)
    read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Replace(first1, Chr(0), "") = "13" Then
        Form37.Show
        Form37.Width = Form37.PPictureBox16.Width
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
        Form37.PPictureBox15.Visible = False
        Form37.PPictureBox16.Visible = True
        Form37.PPictureBox17.Visible = False
    End If
End Sub


Private Sub PButton2_Click()
    Form45.Show
    Form1.baoshidu = Form1.baoshidu - 100
    Form1.yinshui = Form1.yinshui - 100
    Form1.tili = Form1.tili - 100
End Sub

Private Sub PListBox1_ListClick(Index As Long)
    Dianji = Index
End Sub

Private Sub PListBox1_ListMouseMove(Index As Long)
    zhizhen = Index
    If Label24.Visible = False Then Label24.Visible = True
    If Shape1.Visible = False Then Shape1.Visible = True
    If zhizhen = -1 Then Exit Sub
    If zhizhen <> Lastzhizhen Then
        Lastzhizhen = zhizhen
        Dim read_OK As Long
        Dim rdamage As String
        Dim rShejiao As String
        Dim rGuanchali As String
        Dim rNaili As String
        Dim rZhili As String
        Dim runlock As String
        rShejiao = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        rZhili = String(10, 0)
        rdamage = String(10, 0)
        runlock = String(10, 0)
        read_OK = GetPrivateProfileString(zhizhen + 1, "damage", "0", rdamage, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(zhizhen + 1, "useshejiao", "0", rShejiao, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(zhizhen + 1, "useguanchali", "0", rGuanchali, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(zhizhen + 1, "usenaili", "0", rNaili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(zhizhen + 1, "usezhili", "0", rZhili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(CStr("lock" & CStr(xuanzhong + 1)), CStr(PListBox1.ListIndex + 1), "0", runlock, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If CInt(runlock) = 0 Then
            Label24.Caption = "未培训"
            Exit Sub
        End If
        If Tshejiao - CInt(rShejiao) < 0 Then Label24.Caption = "社交能力不足"
        If Tguanchali - CInt(rGuanchali) < 0 Then Label24.Caption = "观察力不足"
        If Tnaili - CInt(rNaili) < 0 Then Label24.Caption = "耐力不足"
        If Tzhili - CInt(rZhili) < 0 Then Label24.Caption = "智力不足"
        If Not Tshejiao - CInt(rShejiao) < 0 And Not Tguanchali - CInt(rGuanchali) < 0 And Not Tnaili - CInt(rNaili) < 0 And Not Tzhili - CInt(rZhili) < 0 Then Label24.Caption = "可进行此工作"
    End If
End Sub


Private Sub PTab1_ItemSelected(NewIndex As Integer)
    PListBox1.Clear
    xuanzhong = NewIndex
    If NewIndex = 4 Then xuanzhong = 8
    Dim read_OK As Long
    Dim rworkmax As String
    rworkmax = String(10, 0)
    read_OK = GetPrivateProfileString("work", "max", "1", rworkmax, 256, App.Path & "\buildings\" & xuanzhong + 1 & ".ini")
    Dim rwork As String
    rwork = String(256, 0)
    Dim X As Integer
    For X = 1 To CInt(rworkmax)
        read_OK = GetPrivateProfileString("work", CStr(X), "无工作选项", rwork, 256, App.Path & "\buildings\" & xuanzhong + 1 & ".ini")
        If rwork = "" Then
        Else
            PListBox1.AddItem (Replace(rwork, Chr(0), ""))
        End If
    Next
End Sub

Private Sub Aiattack()
    
    Meblood = Meblood - Aidamage
    Label2.Caption = Meblood
End Sub

Private Sub Timer1_Timer()
    p = p - 1
    If p <= 1 Then
        Timer1.Enabled = False
        Image5.Picture = LoadPicture("")
        p = 7
        Timer2.Enabled = True
        Timer2.Interval = 90
        Meblood = Meblood - 5
        Label2.Caption = Meblood
        PProgressBar1.Value = CSng(Meblood / 100)
        If Meblood <= 0 Then
            Form48.Show
        End If
        Exit Sub
    End If
    Image5.Picture = LoadPicture(App.Path & "\picture\shape\s1\a" & p & ".gif")
End Sub

Private Sub Timer2_Timer()
    p = p - 1
    If p <= 1 Then
        Timer2.Enabled = False
        Image6.Picture = LoadPicture("")
        Rel = True
        Me.BackColor = &HC0FFFF
        Image2.Picture = LoadPicture(App.Path & "\picture\shape\zhandoubackground.gif")
        Exit Sub
    End If
    Image6.Picture = LoadPicture(App.Path & "\picture\shape\s2\" & p & ".gif")
End Sub

Private Sub Timer3_Timer()
    p = p - 1
    If p <= 0 Then
        Timer3.Enabled = False
        Image5.Picture = LoadPicture("")
        Rel = True
        Me.BackColor = &HC0FFFF
        Image2.Picture = LoadPicture(App.Path & "\picture\shape\zhandoubackground.gif")
        Exit Sub
    End If
    Image5.Picture = LoadPicture(App.Path & "\picture\shape\s3\" & p & ".gif")
End Sub

'Private Sub Picture1_paint()

'  Dim m_token As Long
'  Dim pImg As Long
'  Dim pGraphics As Long
' Dim w As Long, H As Long
' Call GdipCreateFromHDC(Me.Picture1.hDC, pGraphics)
' Call GdipLoadImageFromFile(StrConv(App.Path & "\picture\shape\tuoyuan.png", vbUnicode), pImg)
' Call GdipGetImageWidth(pImg, w)
' Call GdipGetImageHeight(pImg, H)
' Call GdipDrawImageRect(pGraphics, pImg, 0, 0, w, H)

' Call GdipDisposeImage(pImg)
'  Call GdipDeleteGraphics(pGraphics)
'End Sub


Private Sub Timer4_Timer()
    Dim read_OK  As Long
    Dim write1 As Long
    Checktime = Checktime + 1
    If Checktime = 3 Then
        Timer4.Enabled = False
        Dim first1 As String
        first1 = String(10, 0)
        read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If Replace(first1, Chr(0), "") = "3" Then
            Form12.Show
            Form37.Show
            Form37.PPictureBox1.Visible = False
            Form37.PPictureBox2.Visible = False
            Form37.PPictureBox3.Visible = False
            Form37.PPictureBox4.Visible = True
            Form37.PPictureBox5.Visible = False
            Form37.PPictureBox6.Visible = False
            write1 = WritePrivateProfileString("guide", "1", "4", App.Path & "\save\save" & Form3.saveid & ".fsave")
        End If
    End If
End Sub
