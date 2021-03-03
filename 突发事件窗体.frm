VERSION 5.00
Begin VB.Form Form22 
   BorderStyle     =   0  'None
   Caption         =   "Form22"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "突发事件窗体.frx":0000
   LinkTopic       =   "Form22"
   ScaleHeight     =   6465
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "A."
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "B."
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   5040
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "C."
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "突发事件"
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
      Left            =   600
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   2280
      Left            =   -240
      Picture         =   "突发事件窗体.frx":08CA
      Stretch         =   -1  'True
      Top             =   600
      Width           =   9360
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
      Height          =   2565
      Left            =   6960
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   3240
      Left            =   -240
      Picture         =   "突发事件窗体.frx":10E7
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   9360
   End
   Begin VB.Image Image1 
      Height          =   6120
      Left            =   -240
      Picture         =   "突发事件窗体.frx":1904
      Stretch         =   -1  'True
      Top             =   240
      Width           =   9600
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private zhizhen As Integer
Private Lastzhizhen As Integer
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    If Form1.Accident = 0 Then Exit Sub
    Dim Accident As String
    Dim ChooseA As String
    Dim ChooseB As String
    Dim ChooseC As String
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
    Dim read_OK As Long
    Accident = String(256, 0)
    ChooseA = String(256, 0)
    ChooseB = String(256, 0)
    ChooseC = String(256, 0)
    read_OK = GetPrivateProfileString(Form1.Accident, "descriptions", "", Accident, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "choosea", "", ChooseA, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "chooseb", "", ChooseB, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "choosec", "", ChooseC, 256, App.Path & "\accident.ini")
    
    Label1.Caption = Replace(Accident, Chr(0), "")
    PButton1.Text = "A." & Replace(ChooseA, Chr(0), "")
    PButton2.Text = "B." & Replace(ChooseB, Chr(0), "")
    PButton3.Text = "C." & Replace(ChooseC, Chr(0), "")
    Timer1.Enabled = True
    Timer1.Interval = 300
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = 0
End Sub

Private Sub PButton1_Click()
    Dim read_OK As Long
    Dim xg10 As String
    Dim xg11 As String
    Dim xg12 As String
    Dim xg13 As String
    Dim xg14 As String
    Dim rShejiao As String
    Dim rZhili As String
    Dim rGuanchali As String
    Dim rNaili As String
    Dim rdaodepingzhi As String
    Dim rjingye As String
    Dim rhaoxue As String
    Dim rdandang As String
    Dim rchengxin As String
    rdandang = String(10, 0)
    rchengxin = String(10, 0)
    rdaodepingzhi = String(10, 0)
    rjingye = String(10, 0)
    rhaoxue = String(10, 0)
    xg10 = String(10, 0)
    xg11 = String(10, 0)
    xg12 = String(10, 0)
    xg13 = String(10, 0)
    xg14 = String(10, 0)
    rShejiao = String(10, 0)
    rZhili = String(10, 0)
    rGuanchali = String(10, 0)
    rNaili = String(10, 0)
    read_OK = GetPrivateProfileString(Form1.Accident, "baoshidux" & zhizhen, "0", xg10, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "yinshuix" & zhizhen, "0", xg11, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "tilix" & zhizhen, "0", xg12, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "moneyx" & zhizhen, "0", xg13, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "time" & zhizhen, "0", xg14, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "shejiao" & zhizhen, "0", rShejiao, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "zhili" & zhizhen, "0", rZhili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "guanchali" & zhizhen, "0", rGuanchali, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "naili" & zhizhen, "0", rNaili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "daodepingzhi" & zhizhen, "0", rdaodepingzhi, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "jingye" & zhizhen, "0", rjingye, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "haoxue" & zhizhen, "0", rhaoxue, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "dandang" & zhizhen, "0", rdandang, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "chengxin" & zhizhen, "0", rchengxin, 256, App.Path & "\accident.ini")
    
    Dim addacc As String
    addacc = String(10, 0)
    read_OK = GetPrivateProfileString("accident", CStr(Form1.Accident), "0", addacc, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "form3.moshi=" & Form1.moshi
    ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "cint(addacc)=" & CInt(addacc)
    If Form1.moshi = 2 And CInt(addacc) = 1 Then
        '     'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "评测模式下，重复事件不加测评点"
        Unload Form12
        Unload Form22
        Exit Sub
    End If
    write1 = WritePrivateProfileString("accident", CStr(Form1.Accident), "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饱食度修正值" & CStr(CInt(xg10))
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饮水度修正值" & CStr(CInt(xg11))
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加体力值修正值" & CStr(CInt(xg12))
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加金钱修正值" & CStr(CInt(xg13))
    Form1.Shejiao = CInt(rjineng) + Form1.Shejiao
    Form1.Zhili = CInt(rjineng) + Form1.Zhili
    Form1.Guanchali = CInt(rjineng) + Form1.Guanchali
    Form1.Naili = CInt(rjineng) + Form1.Naili
    Form3.baoshidux = Form3.baoshidux + CInt(xg10)
    Form3.yinshuix = Form3.yinshuix + CInt(xg11)
    Form1.tili = Form1.tili + CInt(xg12)
    Form1.money = Form1.money + CInt(xg13)
    Form1.Alltime = Form1.Alltime + CInt(xg14)
    Form1.Daodepingzhi = CInt(rdaodepingzhi) + Form1.Daodepingzhi
    Form1.jingye = CInt(rjingye) + Form1.jingye
    Form1.haoxue = CInt(rhaoxue) + Form1.haoxue
    Form1.dandang = CInt(rdandang) + Form1.dandang
    Form1.chengxin = CInt(rchengxin) + Form1.chengxin
    Dim rmark As String
    If Form1.Accident = 29 Then
        Form1.n = 0
        Form1.Timer10.Enabled = True
        Form1.Timer10.Interval = 30
        Form1.xianshi = 1
    ElseIf Form1.Accident > 50 And Form1.Accident < 70 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 70 And Form1.Accident < 90 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 90 And Form1.Accident < 110 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    Else
        Form1.Accident = 0
    End If
    ' Form32.Label11.Caption = "道德品质:" & Form1.Daodepingzhi
    ' Form32.Label12.Caption = "敬业水平:" & Form1.jingye
    '  Form32.Label13.Caption = "工作经验:" & Form1.haoxue
    ' Form32.Label14.Caption = "dandang:" & Form1.dandang
    '  Form32.Label15.Caption = "chengxin:" & Form1.chengxin
    
    If FindWindow(vbNullString, "Form5") <> 0 Then
        Form5.Show
    ElseIf FindWindow(vbNullString, "Form8") <> 0 Then
        Form8.Show
    ElseIf FindWindow(vbNullString, "Form23") <> 0 Then
        Form23.Show
    ElseIf FindWindow(vbNullString, "Form24") <> 0 Then
        Form24.Show
    ElseIf FindWindow(vbNullString, "Form25") <> 0 Then
        Form25.Show
    ElseIf FindWindow(vbNullString, "Form26") <> 0 Then
        Form26.Show
    ElseIf FindWindow(vbNullString, "Form27") <> 0 Then
        Form27.Show
    ElseIf FindWindow(vbNullString, "Form28") <> 0 Then
        Form28.Show
    Else
        Unload Form12
    End If
    Unload Form22
End Sub

Private Sub PButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = 1
End Sub

Private Sub PButton2_Click()
    Dim read_OK As Long
    Dim xg10 As String
    Dim xg11 As String
    Dim xg12 As String
    Dim xg13 As String
    Dim xg14 As String
    Dim rShejiao As String
    Dim rZhili As String
    Dim rGuanchali As String
    Dim rNaili As String
    Dim rdaodepingzhi As String
    Dim rjingye As String
    Dim rhaoxue As String
    Dim rdandang As String
    Dim rchengxin As String
    rdandang = String(10, 0)
    rchengxin = String(10, 0)
    rdaodepingzhi = String(10, 0)
    rjingye = String(10, 0)
    rhaoxue = String(10, 0)
    xg10 = String(10, 0)
    xg11 = String(10, 0)
    xg12 = String(10, 0)
    rShejiao = String(10, 0)
    rZhili = String(10, 0)
    rGuanchali = String(10, 0)
    rNaili = String(10, 0)
    xg13 = String(10, 0)
    xg14 = String(10, 0)
    read_OK = GetPrivateProfileString(Form1.Accident, "baoshidux" & zhizhen, "0", xg10, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "yinshuix" & zhizhen, "0", xg11, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "tilix" & zhizhen, "0", xg12, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "moneyx" & zhizhen, "0", xg13, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "time" & zhizhen, "0", xg14, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "shejiao" & zhizhen, "0", rShejiao, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "zhili" & zhizhen, "0", rZhili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "guanchali" & zhizhen, "0", rGuanchali, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "naili" & zhizhen, "0", rNaili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "daodepingzhi" & zhizhen, "0", rdaodepingzhi, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "jingye" & zhizhen, "0", rjingye, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "haoxue" & zhizhen, "0", rhaoxue, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "dandang" & zhizhen, "0", rdandang, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "chengxin" & zhizhen, "0", rchengxin, 256, App.Path & "\accident.ini")
    
    Dim addacc As String
    addacc = String(10, 0)
    read_OK = GetPrivateProfileString("accident", CStr(Form1.Accident), "0", addacc, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Form1.moshi = 2 And CInt(addacc) = 1 Then
        Unload Form12
        Unload Form22
        Exit Sub
    End If
    write1 = WritePrivateProfileString("accident", CStr(Form1.Accident), "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饱食度修正值" & CStr(CInt(xg10))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饮水度修正值" & CStr(CInt(xg11))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加体力值修正值" & CStr(CInt(xg12))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加金钱修正值" & CStr(CInt(xg13))
    Form1.Shejiao = CInt(rjineng) + Form1.Shejiao
    Form1.Zhili = CInt(rjineng) + Form1.Zhili
    Form1.Guanchali = CInt(rjineng) + Form1.Guanchali
    Form1.Naili = CInt(rjineng) + Form1.Naili
    Form3.baoshidux = Form3.baoshidux + CInt(xg10)
    Form3.yinshuix = Form3.yinshuix + CInt(xg11)
    Form1.tili = Form1.tili + CInt(xg12)
    Form1.money = Form1.money + CInt(xg13)
    Form1.Alltime = Form1.Alltime + CInt(xg14)
    Form1.Daodepingzhi = CInt(rdaodepingzhi) + Form1.Daodepingzhi
    Form1.jingye = CInt(rjingye) + Form1.jingye
    Form1.haoxue = CInt(rhaoxue) + Form1.haoxue
    Form1.dandang = CInt(rdandang) + Form1.dandang
    Form1.chengxin = CInt(rchengxin) + Form1.chengxin
    ' Form32.Label11.Caption = "道德品质:" & Form1.Daodepingzhi
    'Form32.Label12.Caption = "敬业水平:" & Form1.jingye
    'Form32.Label13.Caption = "工作经验:" & Form1.haoxue
    'Form32.Label14.Caption = "dandang:" & Form1.dandang
    'Form32.Label15.Caption = "chengxin:" & Form1.chengxin
    Dim rmark As String
    If Form1.Accident = 29 Then
        Form1.n = 0
        Form1.Timer10.Enabled = True
        Form1.Timer10.Interval = 30
        Form1.xianshi = 1
    ElseIf Form1.Accident > 50 And Form1.Accident < 70 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 70 And Form1.Accident < 90 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 90 And Form1.Accident < 110 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    Else
        Form1.Accident = 0
    End If
    If FindWindow(vbNullString, "Form5") <> 0 Then
        Form5.Show
    ElseIf FindWindow(vbNullString, "Form8") <> 0 Then
        Form8.Show
    ElseIf FindWindow(vbNullString, "Form23") <> 0 Then
        Form23.Show
    ElseIf FindWindow(vbNullString, "Form24") <> 0 Then
        Form24.Show
    ElseIf FindWindow(vbNullString, "Form25") <> 0 Then
        Form25.Show
    ElseIf FindWindow(vbNullString, "Form26") <> 0 Then
        Form26.Show
    ElseIf FindWindow(vbNullString, "Form27") <> 0 Then
        Form27.Show
    ElseIf FindWindow(vbNullString, "Form28") <> 0 Then
        Form28.Show
    Else
        Unload Form12
    End If
    Unload Form22
End Sub

Private Sub PButton2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = 2
End Sub

Private Sub PButton3_Click()
    Dim read_OK As Long
    Dim xg10 As String
    Dim xg11 As String
    Dim xg12 As String
    Dim xg13 As String
    Dim xg14 As String
    Dim rShejiao As String
    Dim rZhili As String
    Dim rGuanchali As String
    Dim rdaodepingzhi As String
    Dim rjingye As String
    Dim rhaoxue As String
    Dim rdandang As String
    Dim rchengxin As String
    rdandang = String(10, 0)
    rchengxin = String(10, 0)
    rdaodepingzhi = String(10, 0)
    rjingye = String(10, 0)
    rhaoxue = String(10, 0)
    Dim rNaili As String
    xg10 = String(10, 0)
    xg11 = String(10, 0)
    xg12 = String(10, 0)
    xg13 = String(10, 0)
    xg14 = String(10, 0)
    rShejiao = String(10, 0)
    rZhili = String(10, 0)
    rGuanchali = String(10, 0)
    rNaili = String(10, 0)
    read_OK = GetPrivateProfileString(Form1.Accident, "baoshidux" & zhizhen, "0", xg10, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "yinshuix" & zhizhen, "0", xg11, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "tilix" & zhizhen, "0", xg12, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "moneyx" & zhizhen, "0", xg13, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "time" & zhizhen, "0", xg14, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "shejiao" & zhizhen, "0", rShejiao, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "zhili" & zhizhen, "0", rZhili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "guanchali" & zhizhen, "0", rGuanchali, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "naili" & zhizhen, "0", rNaili, 10, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "daodepingzhi" & zhizhen, "0", rdaodepingzhi, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "jingye" & zhizhen, "0", rjingye, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "haoxue" & zhizhen, "0", rhaoxue, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "dandang" & zhizhen, "0", rdandang, 256, App.Path & "\accident.ini")
    read_OK = GetPrivateProfileString(Form1.Accident, "chengxin" & zhizhen, "0", rchengxin, 256, App.Path & "\accident.ini")
    
    Dim addacc As String
    addacc = String(10, 0)
    read_OK = GetPrivateProfileString("accident", CStr(Form1.Accident), "0", addacc, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Form1.moshi = 2 And CInt(addacc) = 1 Then
        Unload Form12
        Unload Form22
        Exit Sub
    End If
    write1 = WritePrivateProfileString("accident", CStr(Form1.Accident), "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饱食度修正值" & CStr(CInt(xg10))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加饮水度修正值" & CStr(CInt(xg11))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加体力值修正值" & CStr(CInt(xg12))
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加金钱修正值" & CStr(CInt(xg13))
    Form1.Shejiao = CInt(rjineng) + Form1.Shejiao
    Form1.Zhili = CInt(rjineng) + Form1.Zhili
    Form1.Guanchali = CInt(rjineng) + Form1.Guanchali
    Form1.Naili = CInt(rjineng) + Form1.Naili
    Form3.baoshidux = Form3.baoshidux + CInt(xg10)
    Form3.yinshuix = Form3.yinshuix + CInt(xg11)
    Form1.tili = Form1.tili + CInt(xg12)
    Form1.money = Form1.money + CInt(xg13)
    Form1.Alltime = Form1.Alltime + CInt(xg14)
    Form1.Daodepingzhi = CInt(rdaodepingzhi) + Form1.Daodepingzhi
    Form1.jingye = CInt(rjingye) + Form1.jingye
    Form1.haoxue = CInt(rhaoxue) + Form1.haoxue
    Form1.dandang = CInt(rdandang) + Form1.dandang
    Form1.chengxin = CInt(rchengxin) + Form1.chengxin
    'Form32.Label11.Caption = "道德品质:" & Form1.Daodepingzhi
    'Form32.Label12.Caption = "敬业水平:" & Form1.jingye
    'Form32.Label13.Caption = "工作经验:" & Form1.haoxue
    'Form32.Label14.Caption = "dandang:" & Form1.dandang
    'Form32.Label15.Caption = "chengxin:" & Form1.chengxin
    Dim rmark As String
    If Form1.Accident = 29 Then
        Form1.n = 0
        Form1.Timer10.Enabled = True
        Form1.Timer10.Interval = 30
        Form1.xianshi = 1
    ElseIf Form1.Accident > 50 And Form1.Accident < 70 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 70 And Form1.Accident < 90 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    ElseIf Form1.Accident > 90 And Form1.Accident < 110 Then
        rmark = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "mark" & zhizhen, "0", rmark, 256, App.Path & "\accident.ini")
        Form28.Mark = Form28.Mark + CInt(rmark)
        Form1.Accident = Form1.Accident + 1
    Else
        Form1.Accident = 0
    End If
    If FindWindow(vbNullString, "Form5") <> 0 Then
        Form5.Show
    ElseIf FindWindow(vbNullString, "Form8") <> 0 Then
        Form8.Show
    ElseIf FindWindow(vbNullString, "Form23") <> 0 Then
        Form23.Show
    ElseIf FindWindow(vbNullString, "Form24") <> 0 Then
        Form24.Show
    ElseIf FindWindow(vbNullString, "Form25") <> 0 Then
        Form25.Show
    ElseIf FindWindow(vbNullString, "Form26") <> 0 Then
        Form26.Show
    ElseIf FindWindow(vbNullString, "Form27") <> 0 Then
        Form27.Show
    ElseIf FindWindow(vbNullString, "Form28") <> 0 Then
        Form28.Show
    Else
        Unload Form12
    End If
    Unload Form22
End Sub

Private Sub PButton3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = 3
End Sub

Private Sub Timer1_Timer()
    If zhizhen <> lastzhizhgen And zhizhen <> 0 Then
        Dim read_OK As Long
        Dim xg10 As String
        Dim xg11 As String
        Dim xg12 As String
        Dim xg13 As String
        Dim rShejiao As String
        Dim rZhili As String
        Dim rGuanchali As String
        Dim rNaili As String
        Dim xg14 As String
        Dim rdaodepingzhi As String
        Dim rjingye As String
        Dim rhaoxue As String
        Dim rdandang As String
        Dim rchengxin As String
        rdandang = String(10, 0)
        rchengxin = String(10, 0)
        rdaodepingzhi = String(10, 0)
        rjingye = String(10, 0)
        rhaoxue = String(10, 0)
        xg10 = String(10, 0)
        xg11 = String(10, 0)
        xg12 = String(10, 0)
        xg13 = String(10, 0)
        xg14 = String(10, 0)
        rShejiao = String(10, 0)
        rZhili = String(10, 0)
        rGuanchali = String(10, 0)
        rNaili = String(10, 0)
        read_OK = GetPrivateProfileString(Form1.Accident, "baoshidux" & zhizhen, "0", xg10, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "yinshuix" & zhizhen, "0", xg11, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "tilix" & zhizhen, "0", xg12, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "moneyx" & zhizhen, "0", xg13, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "time" & zhizhen, "0", xg14, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "shejiao" & zhizhen, "0", rShejiao, 10, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "zhili" & zhizhen, "0", rZhili, 10, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "guanchali" & zhizhen, "0", rGuanchali, 10, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "naili" & zhizhen, "0", rNaili, 10, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "daodepingzhi" & zhizhen, "0", rdaodepingzhi, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "jingye" & zhizhen, "0", rjingye, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "haoxue" & zhizhen, "0", rhaoxue, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "dandang" & zhizhen, "0", rdandang, 256, App.Path & "\accident.ini")
        read_OK = GetPrivateProfileString(Form1.Accident, "chengxin" & zhizhen, "0", rchengxin, 256, App.Path & "\accident.ini")
        
        Dim xg15 As String
        Dim xg16 As String
        Dim xg17 As String
        Dim xg18 As String
        Dim xg19 As String
        
        Dim xg20 As String
        Dim xg21 As String
        Dim xg22 As String
        Dim xg23 As String
        Dim xg24 As String
        
        xg15 = CInt(xg10)
        xg16 = CInt(xg11)
        xg17 = CInt(xg12)
        xg18 = CInt(xg13)
        xg19 = CInt(xg14)
        xg20 = CInt(rdaodepingzhi)
        xg21 = CInt(rjingye)
        xg22 = CInt(haoxue)
        xg23 = CInt(dandang)
        xg24 = CInt(chengxin)
        Label2.Caption = "每分钟降低饱食度" & xg15 & Chr(13) & "每分钟降低饮水度" & xg16 & Chr(13) & "每分钟降低体力值" & xg17 & Chr(13) & "每分钟增加金钱" & xg18
        Label2.Visible = True
    End If
    If zhizhen = 0 Then Label2.Visible = False
End Sub
