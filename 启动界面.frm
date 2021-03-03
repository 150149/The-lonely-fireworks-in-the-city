VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "启动界面.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   840
   End
   Begin 市井中孤傲的烟火.PButton PButton4 
      Height          =   735
      Left            =   6360
      TabIndex        =   3
      Top             =   7200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "退出软件"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Style_Border    =   0
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   735
      Left            =   6360
      TabIndex        =   2
      Top             =   6120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "关于"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Style_Border    =   0
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   735
      Left            =   6360
      TabIndex        =   1
      Top             =   5040
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "联机对战"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Style_Border    =   0
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   735
      Left            =   6360
      TabIndex        =   0
      Top             =   3960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "单人模式"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Style_Border    =   0
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   9975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pb1 As Long
Private pb2 As Long
Private pb3 As Long
Private pb4 As Long
Private pb1e As Double
Private pb2e As Double
Private pb3e As Double
Private pb4e As Double
Private fh As Long
Private fw As Long
Private risea As Integer
Public music As Integer
Public back As Integer
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^公共变量区^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Public saveid As Integer                                                        '存档ID
Public beibaomax As Integer                                                     '背包格子最大值
Public daya As Integer                                                          '天数
Public baoshidux As Integer                                                     '饱食度修正值
Public yinshuix As Integer                                                      '饮水修正值
Public tilix As Integer                                                         '体力修正值
Public moneyx As Integer                                                        '金钱修正值
Public moshi As Integer                                                         '模式
Public Tishiback As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call PButton4_Click
    End If
End Sub

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^公共变量区^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub Form_Load()
    
    pb1 = PButton1.Top
    pb2 = PButton2.Top
    pb3 = PButton3.Top
    pb4 = PButton4.Top
    
    fw = Form3.Width
    fh = Form3.Height
    pb1e = pb1 / fh
    pb2e = pb2 / fh
    pb3e = pb3 / fh
    pb4e = pb4 / fh
    
    
    Form3.Height = Screen.Height
    Form3.Width = Screen.Width
    
    firstrise
    PButton1.Left = (1 / 2) * Form3.Width - (1 / 2) * PButton1.Width
    PButton2.Left = (1 / 2) * Form3.Width - (1 / 2) * PButton2.Width
    PButton3.Left = (1 / 2) * Form3.Width - (1 / 2) * PButton3.Width
    PButton4.Left = (1 / 2) * Form3.Width - (1 / 2) * PButton4.Width
    Label1.Left = (1 / 2) * Form3.Width - (1 / 2) * Label1.Width
    
    
    
    
    'Me.BackColor = &HFF0000
    'Dim rtn As Long
    'Dim BorderStyler
    'BorderStyler = 0
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, &HFF0000, 0, LWA_COLORKEY
    
    Me.BackColor = &H80000
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 70, LWA_ALPHA                           '   窗体透明
    
    Form4.Show
    Form3.Show
    Form18.Show
    Form18.Top = Label1.Top
    Form18.Left = Label1.Left
    Form18.Height = Label1.Height
    Form18.Width = Label1.Width
End Sub

Public Sub firstrise()
    PButton1.Visible = False
    PButton2.Visible = False
    PButton3.Visible = False
    PButton4.Visible = False
    
    PButton1.Top = Form3.Height
    PButton2.Top = Form3.Height
    PButton3.Top = Form3.Height
    PButton4.Top = Form3.Height
    PButton1.Visible = True
    PButton2.Visible = True
    PButton3.Visible = True
    PButton4.Visible = True
    Timer1.Enabled = True
    Timer1.Interval = 5
End Sub

Private Sub PButton1_Click()
    firstdown
End Sub

Private Sub PButton2_Click()
    'MsgBox "预览版不开放此功能"
    Form38.Show
End Sub

Private Sub PButton3_Click()
    Form53.Show
End Sub

Private Sub PButton4_Click()
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form5
    Unload Form6
    Unload Form7
    Unload Form8
    Unload Form9
    Unload Form10
    Unload Form11
    Unload Form12
    Unload Form13
    Unload Form14
    Unload Form15
    Unload Form16
    Unload Form17
    Unload Form18
    Unload Me
    End
End Sub







Private Sub Timer1_Timer()
    If music >= 1 Then
        If PButton1.Top <= pb1e * Form3.Height Then
        Else
            PButton1.Top = PButton1.Top - 220
        End If
        If PButton2.Top <= pb2e * Form3.Height Then
        Else
            PButton2.Top = PButton2.Top - 210
        End If
        If PButton3.Top <= pb3e * Form3.Height Then
        Else
            PButton3.Top = PButton3.Top - 200
        End If
        If PButton4.Top <= pb4e * Form3.Height Then
        Else
            PButton4.Top = PButton4.Top - 190
        End If
        If PButton1.Top <= pb1e * Form3.Height And PButton2.Top <= pb2e * Form3.Height And PButton3.Top <= pb3e * Form3.Height And PButton4.Top <= pb4e * Form3.Height Then
            Timer1.Enabled = False
            
        End If
    End If
End Sub

Private Sub firstdown()
    
    Timer2.Enabled = True
    Timer2.Interval = 5
    
End Sub

Private Sub Timer2_Timer()
    
    If PButton1.Top >= Form3.Height - 300 Then
        PButton1.Visible = False
    Else
        PButton1.Top = PButton1.Top + 220
    End If
    If PButton2.Top >= Form3.Height - 300 Then
        PButton2.Visible = False
    Else
        PButton2.Top = PButton2.Top + 210
    End If
    If PButton3.Top >= Form3.Height - 300 Then
        PButton3.Visible = False
    Else
        PButton3.Top = PButton3.Top + 200
    End If
    If PButton4.Top >= Form3.Height - 300 Then
        PButton4.Visible = False
    Else
        PButton4.Top = PButton4.Top + 190
    End If
    If PButton1.Top >= Form3.Height - 300 And PButton2.Top >= Form3.Height - 300 And PButton3.Top >= Form3.Height - 300 And PButton4.Top >= Form3.Height - 300 Then
        Timer2.Enabled = False
        Form3.Hide
        Unload Form18
        Form7.Show
        If Screen.Width / Screen.TwipsPerPixelX <= 1024 Then
            Unload Form18
        End If
    End If
End Sub



