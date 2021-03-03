VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "启动界面2.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   6360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   4
      Left            =   10560
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   5
      Left            =   10560
      TabIndex        =   5
      Top             =   4200
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   6
      Left            =   10560
      TabIndex        =   6
      Top             =   5280
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   735
      Index           =   7
      Left            =   10560
      TabIndex        =   7
      Top             =   6360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   ""
      Font_Size       =   15
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton5 
      Height          =   1815
      Left            =   6840
      TabIndex        =   8
      Top             =   5280
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3201
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text_MouseMoved=   0
      Text            =   "返回"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   1815
      Left            =   6840
      TabIndex        =   9
      Top             =   3120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   3201
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text_MouseMoved=   0
      Text            =   "编辑模式"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "点击存档位可修改存档"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6360
      TabIndex        =   10
      Top             =   2400
      Width           =   2910
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   13320
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   13320
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   13320
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   13320
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   735
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pb1e As Double
Private pb2e As Double
Private pb3e As Double
Private pb4e As Double
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private bianji As Integer
Private zhizhen As Integer


Private Sub Form_Load()
    Form7.BackColor = &H80000
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 70, LWA_ALPHA                           '   窗体透明
    pb1e = PButton6(0).Top / Me.Height
    pb2e = PButton6(1).Top / Me.Height
    pb3e = PButton6(2).Top / Me.Height
    pb4e = PButton6(3).Top / Me.Height
    Form7.Width = Screen.Width
    Form7.Height = Screen.Height
    Dim cd As Integer
    For cd = 0 To 7
        PButton6(cd).Visible = False
        PButton6(cd).Top = Form3.Height
        PButton6(cd).Visible = True
    Next
    PButton5.Visible = False
    
    Dim v As Integer
    For v = 0 To 7
        If v <= 3 Then
            PButton6(v).Left = (Me.Width * 1 / 4) - (1 / 2 * PButton6(v).Width)
        ElseIf v >= 4 Then
            PButton6(v).Left = (Me.Width * 3 / 4) - (1 / 2 * PButton6(v).Width)
        End If
    Next
    PButton5.Left = (Me.Width * 1 / 2) - (1 / 2 * PButton5.Width)
    PButton1.Left = (Me.Width * 1 / 2) - (1 / 2 * PButton1.Width)
    Label1.Left = Me.Width / 2 - Label1.Width / 2
    PButton5.Top = Form3.Height
    PButton1.Top = Form3.Height
    PButton5.Visible = True
    Label1.Visible = False
    Timer3.Enabled = True
    Timer3.Interval = 5
    Dim c As Integer
    For c = 0 To 7
        Image1(c).Picture = LoadPicture(App.Path & "\picture\shape\del.gif")
        Image1(c).Visible = False
    Next
    
    zhizhen = -1
    Label1.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = -1
    Label1.Visible = False
End Sub

Private Sub Image1_Click(Index As Integer)
    If Dir(App.Path & "\save\save" & Index & ".fsave") = "" Then
        Tishim = Tishi("提示", "该存档位为空")
        Form3.Tishiback = 7
    Else
        Kill (App.Path & "\save\save" & Index & ".fsave")
        Tishim = Tishi("提示", "删除成功")
        Form3.Tishiback = 7
        Form7.PButton6(Index).Text = "空" & Chr(13) & "new" & " " & "new"
    End If
End Sub

Private Sub PButton1_Click()
    Dim c As Integer
    If bianji = 0 Then
        For c = 0 To 7
            Image1(c).Top = PButton6(c).Top
            Image1(c).Left = PButton6(c).Left + PButton6(c).Width
            Image1(c).Visible = True
        Next
        bianji = 1
        PButton1.Text = "退出编辑模式"
        Label1.Visible = True
    Else
        bianji = 0
        For c = 0 To 7
            Image1(c).Visible = False
        Next
        PButton1.Text = "编辑模式"
        PButton1.Font_Size = 17
        Label1.Visible = False
    End If
End Sub

Private Sub PButton5_Click()
    Dim c As Integer
    For c = 0 To 7
        Image1(c).Visible = False
    Next
    Label1.Visible = False
    Timer4.Enabled = True
    Timer4.Interval = 5
End Sub

Private Sub PButton6_Click(Index As Integer)
    Form3.saveid = Index
    If bianji = 1 Then
        Form16.Show
        Unload Form7
    Else
        If Dir(App.Path & "\save\save" & Form3.saveid & ".fsave") = "" Then
            Form13.Show
            Form7.Hide
        ElseIf Dir(App.Path & "\save\save" & Form3.saveid & ".fsave") <> "" Then
            Form20.Show
            Dim read_OK As Long
            Dim rbaoshidu  As String
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
            Dim rbeibaomax As String
            Dim rx As String
            Dim ry As String
            Dim rShejiao As String
            Dim rZhili As String
            Dim rGuanchali As String
            Dim rNaili As String
            Dim rmoshi As String
            Dim rdaodepingzhi As String
            Dim rjingye As String
            Dim rhaoxue As String
            Dim rdandang As String
            Dim rchengxin As String
            Dim rdj As String
            Dim rzhiwei As String
            Dim rpengren As String
            Dim rdriven As String
            Dim rjinrong As String
            rpengren = String(10, 0)
            rdriven = String(10, 0)
            rjinrong = String(10, 0)
            rzhiwei = String(10, 0)
            rdj = String(10, 0)
            rdandang = String(10, 0)
            rchengxin = String(10, 0)
            rdaodepingzhi = String(10, 0)
            rjingye = String(10, 0)
            rhaoxue = String(10, 0)
            rShejiao = String(10, 0)
            rZhili = String(10, 0)
            rGuanchali = String(10, 0)
            rNaili = String(10, 0)
            rx = String(10, 0)
            ry = String(10, 0)
            rbaoshidu = String(10, 0)
            ryinshui = String(10, 0)
            rtili = String(10, 0)
            rmoney = String(10, 0)
            rdaya = String(10, 0)
            rhoura = String(10, 0)
            rmina = String(10, 0)
            rseca = String(10, 0)
            rbeibaomax = String(10, 0)
            rbaoshidux = String(10, 0)
            ryinshuix = String(10, 0)
            rtilix = String(10, 0)
            rmoneyx = String(10, 0)
            rmoshi = String(10, 0)
            read_OK = GetPrivateProfileString("human", "baoshidu", "100", rbaoshidu, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "yinshui", "100", ryinshui, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "tili", "1000", rtili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "money", "100", rmoney, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("extrahuman", "baoshidux", "0", rbaoshidux, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("extrahuman", "yinshuix", "0", ryinshuix, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("extrahuman", "tilix", "0", rtilix, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("extrahuman", "moneyx", "0", rmoneyx, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("time", "daya", "1", rdaya, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("time", "houra", "6", rhoura, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("time", "mina", "0", rmina, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("time", "seca", "0", rseca, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", rbeibaomax, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "top", "4440", rx, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "left", "10680", ry, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "Shejiao", "0", rShejiao, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "Zhili", "0", rZhili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "Guanchali", "0", rGuanchali, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "Naili", "0", rNaili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("save", "mode", "0", rmoshi, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("mark", "daodepingzhi", "0", rdaodepingzhi, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("mark", "jingye", "0", rjingye, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("mark", "haoxue", "0", rhaoxue, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("mark", "dandang", "0", rdandang, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("mark", "chengxin", "0", rchengxin, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("text", "dj", "1", rdj, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "pengren", "False", rpengren, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "driven", "False", rdriven, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            read_OK = GetPrivateProfileString("human", "jinrong", "False", rjinrong, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            
            Form1.Dj = CInt(rdj)
            Form1.Shejiao = CInt(rShejiao)
            Form1.Zhili = CInt(rZhili)
            Form1.Guanchali = CInt(rGuanchali)
            Form1.Naili = CInt(rNaili)
            Form1.baoshidu = CInt(rbaoshidu)
            Form1.yinshui = CInt(ryinshui)
            Form1.tili = CInt(rtili)
            Form1.money = CInt(rmoney)
            ' Form3.baoshidux = CInt(rbaoshidux)
            ' Form3.yinshuix = CInt(ryinshuix)
            ' Form3.tilix = CInt(rtilix)
            ' Form3.moneyx = CInt(rmoneyx)
            Form3.daya = CInt(rdaya)
            Form1.houra = CInt(rhoura)
            Form1.mina = CInt(rmina)
            Form1.seca = CInt(rseca)
            Form20.rx = CInt(rx)
            Form20.ry = CInt(ry)
            If Replace(rpengren, Chr(0), "") = "False" Then Form1.Pengren = False
            If Replace(rpengren, Chr(0), "") = "True" Then Form1.Pengren = True
            If Replace(rdriven, Chr(0), "") = "False" Then Form1.Driven = False
            If Replace(rdriven, Chr(0), "") = "True" Then Form1.Driven = True
            If Replace(rjinrong, Chr(0), "") = "False" Then Form1.Jinrong = False
            If Replace(rjinrong, Chr(0), "") = "True" Then Form1.Jinrong = True
            '   'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "模式读取为" & rmoshi
            Form1.moshi = CInt(rmoshi)
            '   'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "模式为" & Form1.moshi
            If Form1.moshi = 2 Then Form1.mina = 20
            Form1.Daodepingzhi = CInt(rdaodepingzhi)
            Form1.jingye = CInt(rjingye)
            Form1.haoxue = CInt(rhaoxue)
            Form1.dandang = CInt(rdandang)
            Form1.chengxin = CInt(rchengxin)
            Form1.moshi = CInt(rmoshi)
            Form20.Show
        End If
    End If
End Sub

Private Sub PButton6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = Index
    If PButton1.Text = "退出编辑模式" Then
        Label1.Top = Y + PButton6(Index).Top
        Label1.Left = X + PButton6(Index).Left
        Label1.Visible = True
    End If
End Sub

Private Sub Timer3_Timer()
    
    If PButton6(0).Top <= pb1e * Form3.Height Then
    Else
        PButton6(0).Top = PButton6(0).Top - 220
    End If
    
    If PButton6(1).Top <= pb2e * Form3.Height Then
    Else
        PButton6(1).Top = PButton6(1).Top - 210
    End If
    
    If PButton6(2).Top <= pb3e * Form3.Height Then
    Else
        PButton6(2).Top = PButton6(2).Top - 200
    End If
    
    If PButton6(3).Top <= pb4e * Form3.Height Then
    Else
        PButton6(3).Top = PButton6(3).Top - 190
    End If
    
    If PButton6(4).Top <= pb1e * Form3.Height Then
    Else
        PButton6(4).Top = PButton6(4).Top - 220
    End If
    
    If PButton6(5).Top <= pb2e * Form3.Height Then
    Else
        PButton6(5).Top = PButton6(5).Top - 210
    End If
    
    If PButton6(6).Top <= pb3e * Form3.Height Then
    Else
        PButton6(6).Top = PButton6(6).Top - 200
    End If
    
    If PButton6(7).Top <= pb4e * Form3.Height Then
    Else
        PButton6(7).Top = PButton6(7).Top - 190
    End If
    
    If PButton5.Top <= pb3e * Form3.Height Then
    Else
        PButton5.Top = PButton5.Top - 190
    End If
    
    If PButton1.Top <= pb1e * Form3.Height Then
    Else
        PButton1.Top = PButton1.Top - 190
    End If
    
    If PButton6(0).Top <= pb1e * Form3.Height And PButton6(1).Top <= pb2e * Form3.Height And PButton6(2).Top <= pb3e * Form3.Height And PButton6(3).Top <= pb4e * Form3.Height And PButton6(5).Top <= pb2e * Form3.Height And PButton6(6).Top <= pb3e * Form3.Height And PButton6(7).Top <= pb4e * Form3.Height And PButton6(4).Top <= pb1e * Form3.Height And PButton5.Top <= pb3e * Form3.Height And PButton1.Top <= pb1e * Form3.Height Then
        Timer3.Enabled = False
        Dim n As Integer
        Dim read_OK As Long
        Dim name As String
        Dim t1 As String
        Dim t2 As String
        For n = 0 To 7
            t1 = String(10, 0)
            t2 = String(10, 0)
            name = String(10, 0)
            read_OK = GetPrivateProfileString("save", "name", "空", name, 256, App.Path & "\save\save" & n & ".fsave")
            read_OK = GetPrivateProfileString("save", "changetime1", "new", t1, 256, App.Path & "\save\save" & n & ".fsave")
            read_OK = GetPrivateProfileString("save", "changetime2", "new", t2, 256, App.Path & "\save\save" & n & ".fsave")
            Form7.PButton6(n).Text = Replace(name, Chr(0), "") & Chr(13) & Replace(t1, Chr(0), "") & " " & Replace(t2, Chr(0), "")
        Next
    End If
End Sub

Private Sub Timer4_Timer()
    If PButton6(0).Top >= Form3.Height - 200 Then
        PButton6(0).Visible = False
    Else
        PButton6(0).Top = PButton6(0).Top + 220
    End If
    
    If PButton6(1).Top >= Form3.Height - 200 Then
        PButton6(1).Visible = False
    Else
        PButton6(1).Top = PButton6(1).Top + 210
    End If
    
    If PButton6(2).Top >= Form3.Height - 200 Then
        PButton6(2).Visible = False
    Else
        PButton6(2).Top = PButton6(2).Top + 200
    End If
    
    If PButton6(3).Top >= Form3.Height - 200 Then
        PButton6(3).Visible = False
    Else
        PButton6(3).Top = PButton6(3).Top + 190
    End If
    
    If PButton6(4).Top >= Form3.Height - 200 Then
        PButton6(4).Visible = False
    Else
        PButton6(4).Top = PButton6(4).Top + 220
    End If
    
    If PButton6(5).Top >= Form3.Height - 200 Then
        PButton6(5).Visible = False
    Else
        PButton6(5).Top = PButton6(5).Top + 210
    End If
    
    If PButton6(6).Top >= Form3.Height - 200 Then
        PButton6(6).Visible = False
    Else
        PButton6(6).Top = PButton6(6).Top + 200
    End If
    
    If PButton6(7).Top >= Form3.Height - 200 Then
        PButton6(7).Visible = False
    Else
        PButton6(7).Top = PButton6(7).Top + 190
    End If
    
    If PButton5.Top >= Form3.Height - 200 Then
        PButton5.Visible = False
    Else
        PButton5.Top = PButton5.Top + 190
    End If
    
    If PButton1.Top >= Form3.Height - 200 Then
        PButton1.Visible = False
    Else
        PButton1.Top = PButton1.Top + 190
    End If
    
    If PButton6(0).Top >= Form3.Height - 200 And PButton6(1).Top >= Form3.Height - 200 And PButton6(2).Top >= Form3.Height - 200 And PButton6(3).Top >= Form3.Height - 200 And PButton6(5).Top >= Form3.Height - 200 And PButton6(6).Top >= Form3.Height - 200 And PButton6(7).Top >= Form3.Height - 200 And PButton6(4).Top >= Form3.Height - 200 And PButton5.Top >= Form3.Height - 200 And PButton1.Top >= Form3.Height - 200 Then
        Timer4.Enabled = False
        Form3.music = 0
        Form3.Show
        Form18.Show
        Form18.Top = Form3.Label1.Top
        Form18.Left = Form3.Label1.Left
        Form18.Height = Form3.Label1.Height
        Form18.Width = Form3.Label1.Width
        Form3.firstrise
        Unload Form7
    End If
End Sub
