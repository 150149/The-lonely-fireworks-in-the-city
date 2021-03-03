VERSION 5.00
Begin VB.Form Form36 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form36"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "人物技能点选择界面.frx":0000
   LinkTopic       =   "Form36"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin 市井中孤傲的烟火.PButton PButton9 
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   6720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "确定"
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   4080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton4 
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton5 
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   5520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   16777215
      Color_Back_Down =   8421504
      Color_Begin     =   16777215
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   615
      Left            =   6840
      TabIndex        =   6
      Top             =   4080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton7 
      Height          =   615
      Left            =   6840
      TabIndex        =   7
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PButton PButton8 
      Height          =   615
      Left            =   6840
      TabIndex        =   8
      Top             =   5520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   495
      Left            =   9960
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   495
      Left            =   9960
      TabIndex        =   20
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   495
      Left            =   9960
      TabIndex        =   21
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "耐力"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   17
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "观察力"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   16
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "智力"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   15
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "社交能力"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   14
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "可分配的技能点数"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   6240
      Width           =   2415
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Point1 As Integer
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Point1 = 20
    Label1.Caption = "可分配的技能点数：" & Point1
    Label2.Caption = "社交能力：" & Text1.Text
    Label3.Caption = "智力：" & Text2.Text
    Label4.Caption = "观察力：" & Text3.Text
    Label5.Caption = "耐力：" & Text4.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    
    If Screen.Width / Screen.TwipsPerPixelX = 1920 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1920.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1600 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1600.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1440 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1440.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1366 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1366.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1360 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1360.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1280 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1280.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1152 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1152.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1024 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1024.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 800 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\800.jpg")
    Else
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1600.jpg")
    End If
End Sub

Private Sub PButton1_Click()
    If CInt(Text1.Text) <= 0 Then Exit Sub
    Text1.Text = CStr(CInt(Text1.Text) - 1)
    Point1 = Point1 + 1
    Label2.Caption = "社交能力：" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton2_Click()
    If Point1 = 0 Then Exit Sub
    Text1.Text = CStr(CInt(Text1.Text) + 1)
    Point1 = Point1 - 1
    Label2.Caption = "社交能力：" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton3_Click()
    If CInt(Text2.Text) <= 0 Then Exit Sub
    Text2.Text = CStr(CInt(Text2.Text) - 1)
    Point1 = Point1 + 1
    Label3.Caption = "智力：" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton4_Click()
    If CInt(Text3.Text) <= 0 Then Exit Sub
    Text3.Text = CStr(CInt(Text3.Text) - 1)
    Point1 = Point1 + 1
    Label4.Caption = "观察力：" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton5_Click()
    If CInt(Text4.Text) <= 0 Then Exit Sub
    Text4.Text = CStr(CInt(Text4.Text) - 1)
    Point1 = Point1 + 1
    Label5.Caption = "耐力：" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton6_Click()
    If Point1 = 0 Then Exit Sub
    Text2.Text = CStr(CInt(Text2.Text) + 1)
    Point1 = Point1 - 1
    Label3.Caption = "智力：" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton7_Click()
    If Point1 = 0 Then Exit Sub
    Text3.Text = CStr(CInt(Text3.Text) + 1)
    Point1 = Point1 - 1
    Label4.Caption = "观察力：" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton8_Click()
    If Point1 = 0 Then Exit Sub
    Text4.Text = CStr(CInt(Text4.Text) + 1)
    Point1 = Point1 - 1
    Label5.Caption = "耐力：" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub PButton9_Click()
    Form1.Shejiao = CInt(Text1.Text)
    Form1.Zhili = CInt(Text2.Text)
    Form1.Guanchali = CInt(Text3.Text)
    Form1.Naili = CInt(Text4.Text)
    If Form1.Shejiao < 6 Then Form1.Shejiao = 6
    If Form1.Guanchali < 2 Then Form1.Guanchali = 2
    If Form1.Naili < 2 Then Form1.Naili = 2
    Form20.Show
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) = False And Not Text1.Text = "" Then Text1.Text = "0"
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = "" Then Text1.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text1.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label2.Caption = "社交能力：" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub Text2_Change()
    If IsNumeric(Text2.Text) = False And Not Text2.Text = "" Then Text2.Text = "0"
End Sub

Private Sub Text3_Change()
    If IsNumeric(Text3.Text) = False And Not Text3.Text = "" Then Text3.Text = "0"
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4.Text) = False And Not Text4.Text = "" Then Text4.Text = "0"
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text = "" Then Text2.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text2.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label3.Caption = "智力：" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub Text3_LostFocus()
    If Text3.Text = "" Then Text3.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text3.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label4.Caption = "观察力：" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
End Sub

Private Sub Text4_LostFocus()
    If Text4.Text = "" Then Text4.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text4.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label5.Caption = "耐力：" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "可分配的技能点数：" & Point1
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
