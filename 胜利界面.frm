VERSION 5.00
Begin VB.Form Form47 
   BorderStyle     =   0  'None
   Caption         =   "Form47"
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   Icon            =   "胜利界面.frx":0000
   LinkTopic       =   "Form47"
   ScaleHeight     =   4755
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer Timer5 
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Left            =   120
      Top             =   120
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text_MouseMoved=   0
      Text            =   "确定"
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "你打败了时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   4695
      Left            =   0
      Picture         =   "胜利界面.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "Form47"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ib As Integer
Private ic As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public rx As Integer
Public ry As Integer

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Timer4.Interval = 10
    Timer5.Interval = 10
    Timer4.Enabled = True
    Timer5.Enabled = False
    Timer2.Enabled = True
    Timer2.Interval = 1
    ib = 0
End Sub

Private Sub PButton1_Click()
    Timer5.Enabled = True
    Dim first1 As String
    first1 = String(10, 0)
    read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    If Replace(first1, Chr(0), "") = "13" Then
        write1 = WritePrivateProfileString("guide", "1", "14", App.Path & "\save\save" & Form3.saveid & ".fsave")
    End If
End Sub

Private Sub Timer2_Timer()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Timer4_Timer()
    If ib + 5 <= 255 Then
        ib = ib + 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib >= 255 Then
            Timer4.Enabled = False
        End If
    End If
End Sub

Private Sub Timer5_Timer()
    If Timer4.Enabled = True Then Timer4.Enabled = False
    If ib - 5 >= 0 Then
        ib = ib - 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib <= 0 Then
            Timer4.Enabled = False
            Timer5.Enabled = False
            Form45.Show
            Form1.baoshidu = Form1.baoshidu - 100
            Form1.yinshui = Form1.yinshui - 100
            Form1.tili = Form1.tili - 100
            Unload Form47
        End If
    ElseIf ib - 5 < 0 Then
        Form45.Show
        Form1.baoshidu = Form1.baoshidu - 100
        Form1.yinshui = Form1.yinshui - 100
        Form1.tili = Form1.tili - 100
        Unload Form47
    Else
    End If
End Sub

