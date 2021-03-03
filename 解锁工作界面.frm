VERSION 5.00
Begin VB.Form Form44 
   BorderStyle     =   0  'None
   Caption         =   "Form8"
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11220
   Icon            =   "解锁工作界面.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   7170
   ScaleWidth      =   11220
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   5520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Color_Back      =   3644654
      Color_Back_Down =   16576
      Color_Begin     =   3644654
      Color_End       =   16576
      Text            =   "花钱培训"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   1560
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Color_Back      =   3644654
      Color_Back_Down =   10079487
      Color_Begin     =   3644654
      Color_End       =   10079487
      Text            =   "返回"
      Can_Text_Move   =   0   'False
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   6840
      TabIndex        =   4
      Top             =   960
      Width           =   3375
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
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   5895
      Left            =   6600
      Picture         =   "解锁工作界面.frx":08CA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   6135
      Left            =   360
      Picture         =   "解锁工作界面.frx":10E7
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3855
   End
   Begin VB.Image Image3 
      Height          =   6960
      Left            =   0
      Picture         =   "解锁工作界面.frx":1904
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10800
   End
End
Attribute VB_Name = "Form44"
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
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    PButton1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form8
End Sub

Private Sub PButton1_Click()
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
    xg1 = String(10, 0)
    xg2 = String(10, 0)
    xg3 = String(10, 0)
    xg4 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr("lock" & Form1.Position), CStr(PListBox1.ListIndex + 1), "0", xg1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If CInt(xg1) = 1 Then
        Exit Sub
    End If
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "lockmoney", "0", xg2, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    If Form1.money - CInt(xg2) < 0 Then
        Tishim = Tishi("提示", "培训失败，金钱不足")
        Form3.Tishiback = 44
        Exit Sub
    ElseIf Form1.money - CInt(xg2) >= 0 Then
        Form1.money = Form1.money - CInt(xg2)
        Dim write1 As Long
        write1 = WritePrivateProfileString(CStr("lock" & Form1.Position), CStr(PListBox1.ListIndex + 1), "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
        Tishim = Tishi("提示", "培训成功")
        Form3.Tishiback = 44
        
        Dim xg15 As String
        xg15 = String(10, 0)
        read_OK = GetPrivateProfileString(CStr("lock" & Form1.Position), CStr(PListBox1.ListIndex + 1), "0", xg15, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If CInt(xg15) = 1 Then
            PButton1.Text = "已培训"
        ElseIf CInt(xg15) = 0 Then
            PButton1.Text = "花钱培训"
        End If
        PButton1.Visible = True
        Dim xg16 As String
        Dim xg17 As String
        Dim xg18 As String
        xg16 = String(10, 0)
        read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "lockmoney", "0", xg16, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
        If PButton1.Text = "花钱培训" Then xg17 = "未培训"
        If PButton1.Text = "已培训" Then xg17 = "已培训"
        If Form1.money - CInt(xg16) < 0 Then xg18 = "不可培训，金钱不足"
        If Form1.money - CInt(xg16) >= 0 Then xg18 = "可培训"
        Label2.Caption = PListBox1.Text & Chr(13) & "培训花费金钱：" & Replace(xg16, Chr(0), "") & Chr(13) & "状态：" & xg17 & Chr(13) & "是否可以培训：" & xg18
    End If
End Sub

Private Sub PButton3_Click()
    Dim first1 As String
    Dim read_OK As Long
    Dim write1 As Long
    first1 = String(10, 0)
    read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Replace(first1, Chr(0), "") = "3" Then
        write1 = WritePrivateProfileString("guide", "1", "4", App.Path & "\save\save" & Form3.saveid & ".fsave")
    End If
    Unload Form12
    Form1.formback = 1
    Unload Form44
End Sub

Private Sub PListBox1_ListClick(Index As Long)
    Dim xg1 As String
    xg1 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr("lock" & Form1.Position), CStr(PListBox1.ListIndex + 1), "0", xg1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If CInt(xg1) = 1 Then
        PButton1.Text = "已培训"
    ElseIf CInt(xg1) = 0 Then
        PButton1.Text = "花钱培训"
    End If
    PButton1.Visible = True
    Dim xg2 As String
    Dim xg3 As String
    Dim xg4 As String
    Dim xg5 As String
    Dim xg6 As String
    Dim xg7 As String
    Dim xg8 As String
    Dim xg9 As String
    xg2 = String(10, 0)
    xg5 = String(10, 0)
    xg6 = String(10, 0)
    xg7 = String(10, 0)
    xg8 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "lockmoney", "0", xg2, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "shejiao", "0", xg5, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "zhili", "0", xg6, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "guanchali", "0", xg7, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    read_OK = GetPrivateProfileString(CStr(PListBox1.ListIndex + 1), "naili", "0", xg8, 256, App.Path & "\buildings\" & Form1.Position & ".ini")
    If PButton1.Text = "花钱培训" Then xg3 = "未培训"
    If PButton1.Text = "已培训" Then xg3 = "已培训"
    If Form1.money - CInt(xg2) < 0 Then xg4 = "不可培训，金钱不足"
    If Form1.money - CInt(xg2) >= 0 Then xg4 = "可培训"
    If Form1.Shejiao < CInt(xg5) Then
        xg9 = "不可工作，社交能力不足"
    ElseIf Form1.Shejiao >= CInt(xg5) Then
        xg9 = "培训完成后可以工作"
    End If
    If Form1.Zhili < CInt(xg6) Then
        xg9 = "不可工作，智力不足"
    ElseIf Form1.Zhili >= CInt(xg6) Then
        xg9 = "培训完成后可以工作"
    End If
    If Form1.Guanchali < CInt(xg7) Then
        xg9 = "不可工作，观察力不足"
    ElseIf Form1.Guanchali >= CInt(xg7) Then
        xg9 = "培训完成后可以工作"
    End If
    If Form1.Naili < CInt(xg8) Then
        xg9 = "不可工作，耐力不足"
    ElseIf Form1.Naili >= CInt(xg8) Then
        xg9 = "培训完成后可以工作"
    End If
    Label2.Caption = PListBox1.Text & Chr(13) & "培训花费金钱：" & Replace(xg2, Chr(0), "") & Chr(13) & "状态：" & xg3 & Chr(13) & "是否可以培训：" & xg4 & Chr(13) & xg9 & Chr(13) & "小提示：工作也能提高技能点哦"
End Sub


