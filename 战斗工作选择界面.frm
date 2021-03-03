VERSION 5.00
Begin VB.Form Form43 
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
   Icon            =   "战斗工作选择界面.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   7170
   ScaleWidth      =   11220
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   1560
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Color_Back      =   3644654
      Color_Back_Down =   10079487
      Color_Begin     =   3644654
      Color_End       =   10079487
      Text            =   "确定"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Color_Back      =   3644654
      Color_Back_Down =   10079487
      Color_Begin     =   3644654
      Color_End       =   10079487
      Text            =   "<<<<"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Color_Back      =   3644654
      Color_Back_Down =   10079487
      Color_Begin     =   3644654
      Color_End       =   10079487
      Text            =   ">>>"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PListBox PListBox2 
      Height          =   5415
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   9551
      Color_Top_ScrollBar=   10079487
      Color_Back_ScrollBar=   3644654
      Color_Back_Selected=   10079487
      Color_Back_Moved=   3644654
   End
   Begin 市井中孤傲的烟火.PListBox PListBox1 
      Height          =   5415
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   9551
      Color_Top_ScrollBar=   10079487
      Color_Back_ScrollBar=   3644654
      Color_Back_Selected=   10079487
      Color_Back_Moved=   3644654
   End
   Begin VB.Label Label3 
      BackColor       =   &H0099CCFF&
      BackStyle       =   0  'Transparent
      Caption         =   "选择工作"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   7
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   5895
      Left            =   6600
      Picture         =   "战斗工作选择界面.frx":08CA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   5040
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   360
      Picture         =   "战斗工作选择界面.frx":10E7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image Image3 
      Height          =   6960
      Left            =   0
      Picture         =   "战斗工作选择界面.frx":1904
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10800
   End
End
Attribute VB_Name = "Form43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private zhizhen As Integer
Private Allbaoshidux As Integer
Private Allyinshuix As Integer
Private Alltilix As Integer
Private Allmoneyx As Integer
Dim Workid(25) As Integer

Private Sub Form_Load()
    KeyPreview = True
    Dim read_OK As Long
    Dim rworkmax As String
    rworkmax = String(10, 0)
    read_OK = GetPrivateProfileString("work", "max", "1", rworkmax, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    Dim rwork As String
    rwork = String(256, 0)
    Dim X As Integer
    For X = 1 To CInt(rworkmax)
        read_OK = GetPrivateProfileString("work", CStr(X), "无工作选项", rwork, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        If rwork = "" Then
        Else
            PListBox1.AddItem (Replace(rwork, Chr(0), ""))
            
        End If
    Next
    zhizhen = -1
    Timer1.Enabled = True
    Timer1.Interval = 300
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form8
End Sub

Private Sub PButton1_Click()
    Dim read_OK As Long
    If PListBox1.ListIndex = -1 Then
        Tishim = Tishi("提示", "未选择要做的事")
        Form3.Tishiback = 8
    Else
        
        Dim xg1 As String
        Dim xg2 As String
        Dim xg3 As String
        Dim xg4 As String
        xg1 = String(10, 0)
        xg2 = String(10, 0)
        xg3 = String(10, 0)
        xg4 = String(10, 0)
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "shejiao", "0", xg1, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "zhili", "0", xg2, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "guanchali", "0", xg3, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "naili", "0", xg4, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        If CInt(xg1) > Form1.Shejiao Or CInt(xg2) > Form1.Zhili Or CInt(xg3) > Form1.Guanchali Or CInt(xg4) > Form1.Naili Then
            Tishim = Tishi("提示", "你没有足够的能力去做这件事")
            Form3.Tishiback = 8
            Exit Sub
        End If
        
        If Position = 9 And Form1.Jinrong = False Then
            Tishim = Tishi("提示", "你没有通过金融技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        End If
        If PListBox1.Text = "司机-输送材料" And Form1.Driven = False Then
            Tishim = Tishi("提示", "你没有通过驾驶技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "司机-挖土填坑" And Form1.Driven = False Then
            Tishim = Tishi("提示", "你没有通过驾驶技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "司机-打好地基" And Form1.Driven = False Then
            Tishim = Tishi("提示", "你没有通过驾驶技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "厨师- 准备好食材" And Form1.Pengren = False Then
            Tishim = Tishi("提示", "你没有通过烹饪技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "厨师- 烹饪好菜肴" And Form1.Pengren = False Then
            Tishim = Tishi("提示", "你没有通过烹饪技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "厨师- 制作凉拌菜" And Form1.Pengren = False Then
            Tishim = Tishi("提示", "你没有通过烹饪技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        ElseIf PListBox1.Text = "厨师- 准备好调料" And Form1.Pengren = False Then
            Tishim = Tishi("提示", "你没有通过烹饪技能测试，不能做这件事")
            Form3.Tishiback = 8
            Exit Sub
        End If
        
        PListBox2.AddItem (PListBox1.Text)
        Workid(PListBox2.ListCount) = PListBox1.ListIndex
        
        
        Dim rbaoshidux As String
        Dim ryinshuix As String
        Dim rtilix As String
        Dim rmoneyx As String
        Dim rtime As String
        Dim rShejiao As String
        Dim rZhili As String
        Dim rGuanchali As String
        Dim rNaili As String
        rbaoshidux = String(10, 0)
        ryinshuix = String(10, 0)
        rtilix = String(10, 0)
        rmoneyx = String(10, 0)
        rtime = String(10, 0)
        rShejiao = String(10, 0)
        rZhili = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "baoshidux", "0", rbaoshidux, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "yinshuix", "0", ryinshuix, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "tilix", "0", rtilix, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "moneyx", "0", rmoneyx, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "time", "0", rtime, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "addshejiao", "0", rShejiao, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "addzhili", "0", rZhili, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "addguanchali", "0", rGuanchali, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "addnaili", "0", rNaili, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        Form1.Alltime = Form1.Alltime + CInt(rtime)
        Allbaoshidux = Allbaoshidux + CInt(rbaoshidux)
        Allyinshuix = Allyinshuix + CInt(ryinshuix)
        Alltilix = Alltilix + CInt(rtilix)
        Allmoneyx = Allmoneyx + CInt(rmoneyx)
        Debug.Print "alltime=" & Form1.Alltime
        Form1.allShejiao = CInt(rShejiao) + Form1.allShejiao
        Form1.allZhili = CInt(rZhili) + Form1.allZhili
        Form1.allGuanchali = CInt(rGuanchali) + Form1.allGuanchali
        Form1.allNaili = CInt(rNaili) + Form1.allNaili
        'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加社交能力" & rShejiao
        'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加智力" & rZhili
        'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加观察力" & rGuanchali
        'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加耐力" & rNaili
        Debug.Print "Form1.alltime=" & Form1.Alltime
    End If
End Sub

Private Sub PButton2_Click()
    If PListBox2.ListIndex = -1 Then
        Tishim = Tishi("提示", "没有选择要做的事")
        Form3.Tishiback = 8
    Else
        
        Dim read_OK As Long
        Dim rbaoshidux As String
        Dim ryinshuix As String
        Dim rtilix As String
        Dim rmoneyx As String
        Dim rtime As String
        Dim rShejiao As String
        Dim rZhili As String
        Dim rGuanchali As String
        Dim rNaili As String
        rbaoshidux = String(10, 0)
        ryinshuix = String(10, 0)
        rtilix = String(10, 0)
        rmoneyx = String(10, 0)
        rtime = String(10, 0)
        rShejiao = String(10, 0)
        rZhili = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "baoshidux", "0", rbaoshidux, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "yinshuix", "0", ryinshuix, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "tilix", "0", rtilix, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "moneyx", "0", rmoneyx, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "time", "0", rtime, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "addshejiao", "0", rShejiao, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "addzhili", "0", rZhili, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "addguanchali", "0", rGuanchali, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(Workid(PListBox2.ListIndex) + 1, "addnaili", "0", rNaili, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        Form1.Alltime = Form1.Alltime - CInt(rtime)
        Allbaoshidux = Allbaoshidux - CInt(rbaoshidux)
        Allyinshuix = Allyinshuix - CInt(ryinshuix)
        Alltilix = Alltilix - CInt(rtilix)
        Allmoneyx = Allmoneyx - CInt(rmoneyx)
        Form1.allShejiao = Form1.allShejiao - CInt(rShejiao)
        Form1.allZhili = Form1.allZhili - CInt(rZhili)
        Form1.allGuanchali = Form1.allGuanchali - CInt(rGuanchali)
        Form1.allNaili = Form1.allNaili - CInt(rNaili)
        PListBox2.RemoveItem (PListBox2.ListIndex)
    End If
End Sub

Private Sub PButton3_Click()
    If Form1.Alltime <> 0 And Form1.Timer2.Enabled = True Then
        Tishim = Tishi("提示", "请先停止当前工作")
        Form3.Tishiback = 8
        Form1.allGuanchali = 0
        Form1.allNaili = 0
        Form1.allShejiao = 0
        Form1.allZhili = 0
        Unload Form12
        Unload Form8
        Exit Sub
    End If
    KeyPreview = False
    Form1.KeyPreview = True
    Form3.baoshidux = Allbaoshidux
    Form3.yinshuix = Allyinshuix
    Form3.tilix = Alltilix
    Form3.moneyx = Allmoneyx
    Allbaoshidux = 0
    Allyinshuix = 0
    Allmoneyx = 0
    Alltilix = 0
    Form1.Lastposition = Form1.Position
    If Not Form1.Timer2.Enabled = True And Form1.Alltime <> 0 Then
        Load Form1.PProgressBar4(1)                                             '装载标签数组元素
        Form1.PProgressBar4(1).Visible = True
    End If
    If Form1.moshi = 2 Then Form1.PButton6.Visible = True
    Form1.Timer2.Enabled = True
    Form1.Timer2.Interval = 1000
    Unload Form12
    Unload Form8
End Sub

Private Sub PListBox1_ListMouseMove(Index As Long)
    zhizhen = Index
End Sub

Private Sub Timer1_Timer()
    If zhizhen = -1 Then
        Label2.Visible = False
    Else
        Readiteminformation
        Label2.Visible = True
        Label2.Left = PListBox1.Left + PListBox1.Width
        Label2.Top = PListBox1.Top
    End If
End Sub

Private Sub Readiteminformation()
    If PListBox1.ListIndex = -1 Then
    Else
        Dim read_OK As Long
        Dim xg0 As Integer
        
        Dim xg1 As String
        Dim xg2 As String
        Dim xg3 As String
        Dim xg4 As String
        Dim xg5 As String
        Dim xg6 As String
        Dim xg7 As String
        Dim xg8 As String
        Dim xg9 As String
        Dim xg10 As String
        Dim xg11 As String
        Dim xg12 As String
        Dim xg13 As String
        Dim xg14 As String
        Dim xg20 As String
        xg0 = zhizhen + 1
        xg1 = String(10, 0)
        xg2 = String(10, 0)
        xg3 = String(10, 0)
        xg4 = String(10, 0)
        xg10 = String(10, 0)
        xg11 = String(10, 0)
        xg12 = String(10, 0)
        xg13 = String(10, 0)
        xg14 = String(10, 0)
        xg20 = String(10, 0)
        read_OK = GetPrivateProfileString(xg0, "shejiao", "0", xg1, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "zhili", "0", xg2, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "guanchali", "0", xg3, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "naili", "0", xg4, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "baoshidux", "0", xg10, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "yinshuix", "0", xg11, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "tilix", "0", xg12, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "moneyx", "0", xg13, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(xg0, "time", "0", xg14, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        read_OK = GetPrivateProfileString(PListBox1.ListIndex + 1, "time", "0", xg20, 10, App.Path & "\buildings\" & Form1.Position & ".ini")
        Dim xg15 As String
        Dim xg16 As String
        Dim xg17 As String
        Dim xg18 As String
        Dim xg19 As String
        Dim xg21 As String
        xg6 = CInt(xg1)
        xg7 = CInt(xg2)
        xg8 = CInt(xg3)
        xg9 = CInt(xg4)
        xg15 = CInt(xg10)
        xg16 = CInt(xg11)
        xg17 = CInt(xg12)
        xg18 = CInt(xg13)
        xg19 = CInt(xg14)
        xg21 = CInt(xg20)
        Debug.Print PListBox1.Text & Chr(13) & "所需社交能力:" & xg6 & Chr(13) & "所需智力" & xg8 & Chr(13) & "所需观察力" & xg9 & Chr(13) & "所需耐力" & xg7 & Chr(13) & "每分钟降低饱食度" & xg15 & Chr(13) & "每分钟降低饮水度" & xg16 & Chr(13) & "每分钟降低体力值" & xg17 & Chr(13) & "每分钟增加金钱" & xg18
        Label2.Caption = PListBox1.Text & Chr(13) & "所需社交能力:" & xg6 & Chr(13) & "所需智力" & xg8 & Chr(13) & "所需观察力" & xg9 & Chr(13) & "所需耐力" & xg7 & Chr(13) & "每分钟降低饱食度" & xg15 & Chr(13) & "每分钟降低饮水度" & xg16 & Chr(13) & "每分钟降低体力值" & xg17 & Chr(13) & "每分钟增加金钱" & xg18 & Chr(13) & "所需时间" & xg21
    End If
    
End Sub
