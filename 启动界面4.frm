VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "启动界面4.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C2C2C&
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Text            =   "新的生活"
      Top             =   2880
      Width           =   9135
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Back      =   12648447
      Color_Back_Down =   12648447
      Color_Begin     =   12648447
      Color_End       =   12648447
      Color_Text      =   16777215
      Text            =   "经典无尽模式"
      Font_Size       =   15
      Font_Bold       =   -1  'True
      Color_Border    =   12632256
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PSwitch PSwitch1 
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   4920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Color_Top       =   12632256
      Color_Back      =   8421504
      Color_Back_True =   14737632
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   735
      Left            =   10680
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Back      =   12648447
      Color_Back_Down =   12648447
      Color_Begin     =   12648447
      Color_End       =   12648447
      Color_Text      =   16777215
      Text            =   "个人评分模式"
      Font_Size       =   15
      Font_Bold       =   -1  'True
      Color_Border    =   12632256
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   3840
      TabIndex        =   7
      Top             =   7200
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3720
      TabIndex        =   6
      Top             =   7080
      Width           =   9255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "评测玩家社会生活能力指标、联网分数比较"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "探索生活、收集资源、提高自身素质、快乐生活"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "存档名称"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "Form16"
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


Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    'Form13.BackColor = &H80000
    'Dim rtn As Long
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, 0, 70, LWA_ALPHA '   窗体透明
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
    'Dim rtn As Long
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, vbBlack, 70, LWA_COLORKEY
    
    Form16.Width = Screen.Width
    Form16.Height = Screen.Height
    Form14.Show
    Form16.Show
    Form17.Show
    Form17.Height = Form16.Label6.Height
    Form17.Width = Form16.Label6.Width
    Form17.Top = Form16.Label6.Top
    Form17.Left = Form16.Label6.Left
    Dim read_OK As Long
    Dim r1 As String
    r1 = String(10, 0)
    read_OK = GetPrivateProfileString("save", "name", "新的生活", r1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    Text1.Text = Replace(r1, Chr(0), "")
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
