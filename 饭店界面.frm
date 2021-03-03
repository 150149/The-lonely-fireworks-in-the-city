VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   0  'None
   Caption         =   "Form10"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16320
   LinkTopic       =   "Form2"
   ScaleHeight     =   9465
   ScaleWidth      =   16320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Index           =   6
      Left            =   5640
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Index           =   5
      Left            =   5160
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Index           =   4
      Left            =   4680
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Index           =   3
      Left            =   4200
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Index           =   2
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Index           =   1
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Left            =   2760
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Left            =   2280
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   135
      Index           =   0
      Left            =   3600
      TabIndex        =   19
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   238
      Color_Back      =   0
   End
   Begin VB.Timer Timer4 
      Left            =   1800
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PButton PButton4 
      Height          =   855
      Left            =   360
      TabIndex        =   10
      Top             =   7680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      Text            =   "背包"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   495
      Left            =   12000
      TabIndex        =   9
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Text            =   "暂停"
      Can_Text_Move   =   0   'False
   End
   Begin VB.Timer Timer3 
      Left            =   1320
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PScreen PScreen1 
      Height          =   735
      Left            =   13200
      TabIndex        =   8
      Top             =   600
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Text      =   16777215
      Color_Text_Back =   0
      Text            =   "06:00"
   End
   Begin VB.Timer Timer2 
      Index           =   0
      Left            =   840
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Top       =   8454143
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Top       =   33023
      Color_Back      =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "xxx建筑"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
      Begin 市井中孤傲的烟火.PButton PButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   661
         Text            =   "打工"
         Can_Text_Move   =   0   'False
      End
      Begin 市井中孤傲的烟火.PButton PButton2 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
         _ExtentX        =   1931
         _ExtentY        =   661
         Text            =   "交易"
         Can_Text_Move   =   0   'False
      End
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   120
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "大门"
      Height          =   615
      Left            =   3000
      TabIndex        =   20
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      Caption         =   "客人A"
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   18
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      Caption         =   "人物贴图以后搞"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "厨房取餐区"
      Height          =   1095
      Left            =   5280
      TabIndex        =   2
      Top             =   2280
      Width           =   7335
   End
   Begin VB.Label Label2 
      Caption         =   "状态"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "1号桌"
      Height          =   1215
      Index           =   0
      Left            =   5400
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "2号桌"
      Height          =   1215
      Index           =   1
      Left            =   7560
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "3号桌"
      Height          =   1215
      Index           =   2
      Left            =   9960
      TabIndex        =   14
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "6号桌"
      Height          =   1215
      Index           =   5
      Left            =   9960
      TabIndex        =   17
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "5号桌"
      Height          =   1215
      Index           =   4
      Left            =   7440
      TabIndex        =   16
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "4号桌"
      Height          =   1215
      Index           =   3
      Left            =   5400
      TabIndex        =   15
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H000080FF&
      Caption         =   "收银台"
      Height          =   1335
      Left            =   13080
      TabIndex        =   11
      Top             =   3600
      Width           =   495
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RetVal As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Const MAX_FILENAME_LEN = 256
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Addai As Integer
Private Arriveai As Integer
Private Position As Integer
Private leaveid As Integer
Private tili2 As Integer
Private timermode(1 To 6) As Integer
Private Aipatient(1 To 6) As Integer
Private Checktime(1 To 6) As Integer
Private Firstcome(1 To 6) As Integer
Private Firstcomea(1 To 6) As Integer
Private Cooktime(1 To 6) As Integer
Private Eattime(1 To 6) As Integer




Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    tili2 = 1
End Sub

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Form10.Width = Screen.Width
    Form10.Height = Screen.Height
    Frame1.Visible = False
    KeyPreview = True
    Timer1.Enabled = True
    Timer1.Interval = 1000
    Timer3.Enabled = True
    Timer3.Interval = 1000
    Timer4.Enabled = True
    Timer4.Interval = 500
    Image1.Picture = LoadPicture(App.Path & "\picture\shape\baoshidu.gif")
    Image2.Picture = LoadPicture(App.Path & "\picture\shape\yinshui.gif")
    Image3.Picture = LoadPicture(App.Path & "\picture\shape\tili.gif")
    'Image4.Picture = LoadPicture(App.Path & "\picture\shape\money.gif")
    Form6.Show
    Form10.Show
    Label1.Top = Me.Height - Label7.Height - 100 - (Me.Height - Label7.Top)
    Label1.Left = Me.Width / 4
    Label6(0).Visible = False
    Addai = -1
    
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

Private Sub Form_Unload(Cancel As Integer)
    
    Dim c As Integer
    For c = 1 To 6
        Timer2(c).Enabled = False
        If timermode(c) <> 0 Then
            Unload Label6(c)
        End If
    Next
    Unload Me
End Sub

Private Sub PButton3_Click()
    Timer3.Enabled = False
    Timer1.Enabled = False
    Form6.Show
End Sub

Private Sub PButton4_Click()
    Form12.Show
    Form5.form5open = 1
    Form5.Show
    Form5.Timer1.Enabled = True
    Form5.Timer1.Interval = 300
End Sub

'----------------------------------------人物生命指标模块--------------------------------------
Private Sub Timer1_Timer()
    
    
    
    Form1.baoshidu = Form1.baoshidu - 3
    Form1.yinshui = Form1.yinshui - 6
    
    If tili2 = 1 And Form1.tili < 1000 Then
        Form1.tili = Form1.tili + 1
    Else
    End If
    
    PProgressBar1.Value = CSng(Form1.baoshidu / 1000)
    PProgressBar2.Value = CSng(Form1.yinshui / 1000)
    PProgressBar3.Value = CSng(Form1.tili / 1000)
    Label11.Caption = Form1.money
    Label2.Caption = "value: " & PProgressBar3.Value & "  " & "form1.tili: " & Form1.tili / 100
    
    If Form1.baoshidu <= 0 Or Form1.yinshui <= 0 Then
        MsgBox "Gameover"
        Unload Me
        End
    End If
End Sub
'----------------------------------------人物生命指标模块--------------------------------------


'----------------------------------控制模块------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Label1.Top < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Top = Label1.Top + 300
    ElseIf Label1.Left < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Left = Label1.Left + 300
    ElseIf Form1.Height - Label1.Top < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Top = Label1.Top - 300
    ElseIf Form1.Width - Label1.Left < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Left = Label1.Left - 300
    ElseIf Form1.tili = 0 Then
        Label2.Caption = "你没力气了"
    Else
        If KeyCode = 85 Then
            Label1.Left = Label1.Left - 100
            tili2 = 0
            Form1.tili = Form1.tili - 1
            PProgressBar3.Value = CSng(Form1.tili / 1000)
        ElseIf KeyCode = 87 Then
            Label1.Top = Label1.Top - 100
            tili2 = 0
            Form1.tili = Form1.tili - 1
            PProgressBar3.Value = CSng(Form1.tili / 1000)
        ElseIf KeyCode = 65 Then
            Label1.Left = Label1.Left - 100
            tili2 = 0
            Form1.tili = Form1.tili - 1
            PProgressBar3.Value = CSng(Form1.tili / 1000)
        ElseIf KeyCode = 83 Then
            Label1.Top = Label1.Top + 100
            tili2 = 0
            Form1.tili = Form1.tili - 1
            PProgressBar3.Value = CSng(Form1.tili / 1000)
        ElseIf KeyCode = 68 Then
            Label1.Left = Label1.Left + 100
            tili2 = 0
            Form1.tili = Form1.tili - 1
            PProgressBar3.Value = CSng(Form1.tili / 1000)
        ElseIf KeyCode = 13 Then
            Unload Me
            End
        Else
            Select Case KeyCode
            Case vbKeyUp
                Label1.Top = Label1.Top - 100
            Case vbKeyDown
                Label1.Top = Label1.Top + 100
            Case vbKeyLeft
                Label1.Left = Label1.Left - 100
            Case vbKeyRight
                Label1.Left = Label1.Left + 100
            End Select
        End If
    End If
    If Label1.Top >= Label3.Top And Label1.Top <= Label3.Top + Label3.Height And Label1.Left >= Label3.Left And Label1.Left <= Label3.Left + Label3.Width Then
        Label2.Caption = "你现在在厨房取餐区内"
        Frame1.Visible = True
        Position = 1
    ElseIf Label1.Top >= Label5(0).Top And Label1.Top <= Label5(0).Top + Label5(0).Height And Label1.Left >= Label5(0).Left And Label1.Left <= Label5(0).Left + Label5(0).Width Then
        Label2.Caption = "你现在在1号桌内"
        Frame1.Visible = True
        Position = 11
    ElseIf Label1.Top >= Label5(1).Top And Label1.Top <= Label5(1).Top + Label5(1).Height And Label1.Left >= Label5(1).Left And Label1.Left <= Label5(1).Left + Label5(1).Width Then
        Label2.Caption = "你现在在2号桌内"
        Frame1.Visible = True
        Position = 12
    ElseIf Label1.Top >= Label5(2).Top And Label1.Top <= Label5(2).Top + Label5(2).Height And Label1.Left >= Label5(2).Left And Label1.Left <= Label5(2).Left + Label5(2).Width Then
        Label2.Caption = "你现在在3号桌内"
        Frame1.Visible = True
        Position = 13
    ElseIf Label1.Top >= Label5(3).Top And Label1.Top <= Label5(3).Top + Label5(3).Height And Label1.Left >= Label5(3).Left And Label1.Left <= Label5(3).Left + Label5(3).Width Then
        Label2.Caption = "你现在在4号桌内"
        Frame1.Visible = True
        Position = 14
    ElseIf Label1.Top >= Label5(4).Top And Label1.Top <= Label5(4).Top + Label5(4).Height And Label1.Left >= Label5(4).Left And Label1.Left <= Label5(4).Left + Label5(4).Width Then
        Label2.Caption = "你现在在5号桌内"
        Frame1.Visible = True
        Position = 15
    ElseIf Label1.Top >= Label5(5).Top And Label1.Top <= Label5(5).Top + Label5(5).Height And Label1.Left >= Label5(5).Left And Label1.Left <= Label5(5).Left + Label5(5).Width Then
        Label2.Caption = "你现在在6号桌内"
        Frame1.Visible = True
        Position = 16
    ElseIf Label1.Top >= Label7.Top And Label1.Top <= Label7.Top + Label7.Height And Label1.Left >= Label7.Left And Label1.Left <= Label7.Left + Label7.Width Then
        Form1.Show
        Form1.Timer1.Enabled = True
        Form1.Timer3.Enabled = True
        Form1.Timer5.Enabled = True
        Form1.Label1.Top = (Form1.Height - Form1.Label13.Top - Form1.Label13.Height)
        Unload Form10
    Else
        Label2.Caption = "你现在在饭店内"
        Frame1.Visible = False
        Position = 0
    End If
End Sub
'--------------------------------------控制模块--------------------------------------------------------

Private Sub Timer2_Timer(Index As Integer)
    If Label6(Index).Top >= Label5(Index - 1).Top And timermode(Index) = 1 Then
        Label6(Index).Top = Label6(Index).Top - 100
    ElseIf Label6(Index).Left <= Label5(Index - 1).Left And timermode(Index) = 1 Then
        Label6(Index).Left = Label6(Index).Left + 100
    End If
    If Not Label6(Index).Top >= Label5(Index - 1).Top And Not Label6(Index).Left <= Label5(Index - 1).Left And timermode(Index) = 1 Then
        Arriveai = Index
        Load PProgressBar4(Index)                                               '装载标签数组元素
        PProgressBar4(Index).Top = Label6(Index).Top - 300                      '定义数组标签的顶端TOP位置
        PProgressBar4(Index).Left = Label6(Index).Left                          '定义数组标签的左边LEFT位置
        PProgressBar4(Index).Visible = True                                     '动态添加的Label默认为不可见 所以我们要让它可见
        Aipatient(Index) = 30
        timermode(Index) = 2
        Timer2(Index).Interval = 1000
    End If
    
    If timermode(Index) = 2 And Aipatient(Index) <> 0 And (Position - 10) <> Index Then
        Aipatient(Index) = Aipatient(Index) - 1
        PProgressBar4(Index).Value = CSng(Aipatient(Index) / 30)
        Firstcome(Index) = 0
        Timer2(Index).Interval = 1000
    ElseIf timermode(Index) = 2 And Aipatient(Index) = 0 Then
        MsgBox "一位顾客不耐心地离开了"
        Timer2(Index).Interval = 50
        timermode(Index) = 3
        Unload PProgressBar4(Index)
    End If
    If timermode(Index) = 2 And Aipatient(Index) <> 0 And (Position - 10) = Index And Firstcome(Index) = 0 Then
        Checktime(Index) = 0
        Firstcome(Index) = 1
        Timer2(Index).Interval = 100
    ElseIf timermode(Index) = 2 And Aipatient(Index) <> 0 And (Position - 10) = Index And Firstcome(Index) = 1 And Checktime(Index) <> 50 Then
        Checktime(Index) = Checktime(Index) + 1
        PProgressBar4(Index).Value = CSng(Checktime(Index) / 50)
    ElseIf timermode(Index) = 2 And Aipatient(Index) <> 0 And (Position - 10) = Index And Firstcome(Index) = 1 And Checktime(Index) = 50 Then
        timermode(Index) = 4
        Unload PProgressBar4(Index)
        Firstcome(Index) = 0
        Checktime(Index) = 0
    End If
    
    If timermode(Index) = 3 And Label6(Index).Left >= Me.Width / 4 Then
        Label6(Index).Left = Label6(Index).Left - 100
    ElseIf Label6(Index).Top <= Me.Height - 300 And timermode(Index) = 3 Then
        Label6(Index).Top = Label6(Index).Top + 100
    ElseIf Not Label6(Index).Left >= Me.Width / 4 And Not Label6(Index).Top <= Me.Height - 300 And timermode(Index) = 3 Then
        timermode(Index) = 1
        Unload Label6(Index)
        Timer2(Index).Enabled = False
    End If
    
    If timermode(Index) = 4 And Position = 1 And Firstcomea(Index) = 0 Then
        Firstcomea(Index) = 1
        timermode(Index) = 5
        Timer2(Index).Interval = 1000
    End If
    
    If timermode(Index) = 5 And Cooktime(Index) <> 15 Then
        Cooktime(Index) = Cooktime(Index) + 1
    ElseIf timermode(Index) = 5 And Cooktime(Index) = 15 Then
        Label3.Caption = "厨房取餐区" & "(" & CStr(Index) & "号桌菜等待端上)"
        timermode(Index) = 9
        Load PProgressBar4(Index)                                               '装载标签数组元素
        PProgressBar4(Index).Top = Label6(Index).Top - 300                      '定义数组标签的顶端TOP位置
        PProgressBar4(Index).Left = Label6(Index).Left                          '定义数组标签的左边LEFT位置
        PProgressBar4(Index).Visible = True                                     '动态添加的Label默认为不可见 所以我们要让它可见
        Aipatient(Index) = 30
        Timer2(Index).Interval = 1000
    End If
    
    If timermode(Index) = 9 And Position = 1 Then
        Label3.Caption = "厨房取餐区"
        timermode(Index) = 6
    ElseIf timermode(Index) = 9 And Aipatient(Index) = 0 Then
        MsgBox "一位顾客不耐心地离开了"
        Timer2(Index).Interval = 50
        timermode(Index) = 3
        Unload PProgressBar4(Index)
    ElseIf timermode(Index) = 9 And Aipatient(Index) <> 0 And (Position - 10) <> Index Then
        Aipatient(Index) = Aipatient(Index) - 1
        PProgressBar4(Index).Value = CSng(Aipatient(Index) / 30)
    End If
    
    If timermode(Index) = 6 And Aipatient(Index) = 0 Then
        MsgBox "一位顾客不耐心地离开了"
        Timer2(Index).Interval = 50
        timermode(Index) = 3
        Unload PProgressBar4(Index)
    ElseIf timermode(Index) = 6 And Aipatient(Index) <> 0 And (Position - 10) <> Index Then
        Aipatient(Index) = Aipatient(Index) - 1
        PProgressBar4(Index).Value = CSng(Aipatient(Index) / 30)
    ElseIf timermode(Index) = 6 And Aipatient(Index) <> 0 And (Position - 10) = Index Then
        timermode(Index) = 7
        Eattime(Index) = 15
        Unload PProgressBar4(Index)
    End If
    
    If timermode(Index) = 7 And Eattime(Index) <> 0 Then
        Eattime(Index) = Eattime(Index) - 1
    ElseIf timermode(Index) = 7 And Eattime(Index) = 0 Then
        timermode(Index) = 10
        Timer2(Index).Interval = 1000
        Load PProgressBar4(Index)                                               '装载标签数组元素
        PProgressBar4(Index).Top = Label6(Index).Top - 300                      '定义数组标签的顶端TOP位置
        PProgressBar4(Index).Left = Label6(Index).Left                          '定义数组标签的左边LEFT位置
        PProgressBar4(Index).Visible = True                                     '动态添加的Label默认为不可见 所以我们要让它可见
        Aipatient(Index) = 15
    End If
    
    If timermode(Index) = 10 And (Position - 10) = Index And Aipatient(Index) <> 0 Then
        timermode(Index) = 8
        Unload PProgressBar4(Index)
        Timer2(Index).Interval = 50
        Form1.money = Form1.money + 2
    ElseIf timermode(Index) = 10 And Aipatient(Index) = 0 And (Position - 10) <> Index Then
        MsgBox "一位顾客不耐心地离开了"
        Timer2(Index).Interval = 50
        timermode(Index) = 3
        Unload PProgressBar4(Index)
    ElseIf timermode(Index) = 10 And Aipatient(Index) <> 0 Then
        Aipatient(Index) = Aipatient(Index) - 1
        PProgressBar4(Index).Value = CSng(Aipatient(Index) / 15)
    End If
    
    If timermode(Index) = 8 And Label6(Index).Left >= Me.Width / 4 Then
        Label6(Index).Left = Label6(Index).Left - 100
    ElseIf Label6(Index).Top <= Me.Height - 300 And timermode(Index) = 8 Then
        Label6(Index).Top = Label6(Index).Top + 100
    ElseIf Not Label6(Index).Left >= Me.Width / 4 And Not Label6(Index).Top <= Me.Height - 300 And timermode(Index) = 8 Then
        timermode(Index) = 0
        Unload Label6(Index)
        Timer2(Index).Enabled = False
    End If
    
    
End Sub

Private Sub Timer3_Timer()
    If Form1.moshi = 0 Then                                                     '模式检测 0为经典无尽，1为限时模式，2为评测模式
        Form1.mina = Form1.mina + 1
        If Form1.mina = 60 Then
            Form1.houra = Form1.houra + 1
            Form1.mina = 0
        End If
        If Form1.houra = 24 Then
            Form1.houra = 0
            Form3.daya = Form3.daya + 1
        End If
        
        Dim b As String
        b = "0" & Form1.houra
        b = Right(b, 2)
        Dim c As String
        c = "0" & Form1.mina
        c = Right(c, 2)
        PScreen1.Text = b & ":" & c
        
        
    ElseIf Form1.moshi = 1 Then
        Form1.seca = Form1.seca + 1
        If Form1.seca = 60 Then
            Form1.seca = 0
            Form1.mina = Form1.mina + 1
        End If
        If Form1.mina = 60 Then
            Form1.mina = 0
            Form1.houra = Form1.houra + 1
        End If
        If Form1.houra = 24 Then
            Form1.houra = 0
            Form3.daya = Form3.daya + 1
        End If
        
    ElseIf Form1.moshi = 2 Then
    End If
End Sub



Private Sub Timer4_Timer()
    Dim X As Integer
    For X = 0 To 5
        Label5(X).Caption = X + 1 & "号桌" & "(" & timermode(X + 1) & ")"
        
    Next
End Sub

Private Sub Timer5_Timer()
    Randomize
    If Int(Rnd * (15 - 1 + 1)) + 1 = 3 Then
        
        Addai = Int(Rnd * (6 - 1 + 1)) + 1
        If timermode(Addai) <> 0 Then Exit Sub
        
        Load Label6(Addai)                                                      '装载标签数组元素
        Label6(Addai).Top = Me.Height - 300                                     '定义数组标签的顶端TOP位置
        Label6(Addai).Left = Me.Width / 4                                       '定义数组标签的左边LEFT位置
        Label6(Addai).Visible = True                                            '动态添加的Label默认为不可见 所以我们要让它可见
        timermode(Addai) = 1
        Timer2(Addai).Enabled = True
        Timer2(Addai).Interval = 50
        Addai = -1
    End If
End Sub

'-----------------------------------工作模块----------------------------------------------
Private Sub PButton1_Click()
    Form1.Position = 1
    Form12.Show
    Form8.Show
End Sub
'-----------------------------------工作模块---------------------------------------------------


Private Sub PButton2_Click()
    Form12.Show
    Form23.Show
End Sub
