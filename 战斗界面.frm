VERSION 5.00
Begin VB.Form Form29 
   BorderStyle     =   0  'None
   Caption         =   "Form29"
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16755
   LinkTopic       =   "Form37"
   ScaleHeight     =   9795
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Left            =   16320
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Left            =   16320
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Left            =   -600
      Top             =   -120
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Text            =   "释放职业技能"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   615
      Left            =   6240
      TabIndex        =   1
      Top             =   8160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Text            =   "离开战斗"
   End
   Begin 市井中孤傲的烟火.PListBox PListBox1 
      Height          =   2175
      Left            =   600
      TabIndex        =   2
      Top             =   7080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   3836
   End
   Begin 市井中孤傲的烟火.PTab PTab1 
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   6360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1296
      Font_Size       =   11
      Color_Selector  =   0
      Texts           =   "饭店|便利店|建筑公司|房地产公司|证券公司"
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   375
      Left            =   9840
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      Color_Top       =   255
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   255
      Left            =   10920
      TabIndex        =   6
      Top             =   8400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   255
      Left            =   10920
      TabIndex        =   7
      Top             =   7680
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar5 
      Height          =   255
      Left            =   10920
      TabIndex        =   8
      Top             =   8040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar6 
      Height          =   255
      Left            =   10920
      TabIndex        =   9
      Top             =   7320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      Color_Back      =   -2147483633
   End
   Begin VB.Image Image6 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   4335
      Left            =   360
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Image Image2 
      Height          =   4335
      Left            =   9840
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "社交能力"
      Height          =   180
      Left            =   9960
      TabIndex        =   15
      Top             =   7320
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "智力"
      Height          =   180
      Left            =   9960
      TabIndex        =   14
      Top             =   7680
      Width           =   360
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "观察力"
      Height          =   180
      Left            =   9960
      TabIndex        =   13
      Top             =   8040
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "耐力"
      Height          =   180
      Left            =   9960
      TabIndex        =   12
      Top             =   8400
      Width           =   360
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.Image Image5 
      Enabled         =   0   'False
      Height          =   4335
      Left            =   9360
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   4935
   End
   Begin VB.Image Image3 
      Height          =   2280
      Left            =   9000
      Picture         =   "战斗界面.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   5400
   End
   Begin VB.Image Image4 
      Height          =   9840
      Left            =   -720
      Picture         =   "战斗界面.frx":081D
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   16920
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
    Rel = True
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
        rShejiao = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        rZhili = String(10, 0)
        rdamage = String(10, 0)
        read_OK = GetPrivateProfileString(Dianji + 1, "damage", "0", rdamage, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "useshejiao", "0", rShejiao, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "useguanchali", "0", rGuanchali, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "usenaili", "0", rNaili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        read_OK = GetPrivateProfileString(Dianji + 1, "usezhili", "0", rZhili, 256, App.Path & "\buildings\" & CStr(xuanzhong + 1) & ".ini")
        If Tshejiao - CInt(rShejiao) < 0 Or Tzhili - CInt(rZhili) < 0 Or Tguanchali - CInt(rGuanchali) < 0 Or Tnaili - CInt(rNaili) < 0 Then Exit Sub
        Aiblood = Aiblood - CInt(rdamage)
        Label1.Caption = Aiblood
        If Aiblood <= 0 Then
            
            Form1.money = Form1.money + Aimaxblood / 10
        End If
        Tshejiao = Tshejiao - CInt(rShejiao)
        Tguanchali = Tguanchali - CInt(rGuanchali)
        Tnaili = Tnaili - CInt(rNaili)
        Tzhili = Tzhili - CInt(rZhili)
        PProgressBar6.Value = CSng(Tshejiao / Form1.Shejiao)
        PProgressBar3.Value = CSng(Tzhili / Form1.Zhili)
        PProgressBar5.Value = CSng(Tguanchali / Form1.Guanchali)
        PProgressBar4.Value = CSng(Tnaili / Form1.Naili)
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
        Else
            Timer3.Enabled = True
            Rel = False
            Timer3.Interval = 90
            p = 9
        End If
    End If
End Sub


Private Sub PButton2_Click()
    Form45.Show
    Form1.baoshidu = Form1.baoshidu - 300
    Form1.yinshui = Form1.yinshui - 300
    Form1.tili = Form1.tili - 300
End Sub

Private Sub PListBox1_ListClick(Index As Long)
    Dianji = Index
End Sub

Private Sub PTab1_ItemSelected(NewIndex As Integer)
    PListBox1.Clear
    xuanzhong = NewIndex
    If NewIndex = 3 Then xuanzhong = 8
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
        If Meblood <= 0 Then Label2.Caption = Meblood
        Exit Sub
    End If
    Image5.Picture = LoadPicture(App.Path & "\picture\shape\s1\a" & p & ".gif")
End Sub

Private Sub Timer2_Timer()
    p = p - 1
    If p <= 1 Then Timer2.Enabled = False
    If p <= 1 Then Image6.Picture = LoadPicture("")
    If p <= 1 Then Rel = True
    If p <= 1 Then Exit Sub
    Image6.Picture = LoadPicture(App.Path & "\picture\shape\s2\" & p & ".gif")
End Sub

Private Sub Timer3_Timer()
    p = p - 1
    If p <= 0 Then Timer3.Enabled = False
    If p <= 0 Then Image5.Picture = LoadPicture("")
    If p <= 0 Then Rel = True
    If p <= 0 Then Exit Sub
    Image5.Picture = LoadPicture(App.Path & "\picture\shape\s3\" & p & ".gif")
End Sub
