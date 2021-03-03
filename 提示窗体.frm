VERSION 5.00
Begin VB.Form Form34 
   BorderStyle     =   0  'None
   Caption         =   "Form34"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   Icon            =   "提示窗体.frx":0000
   LinkTopic       =   "Form34"
   ScaleHeight     =   4305
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      Color_Back      =   3644654
      Color_Back_Down =   8438015
      Color_Begin     =   3644654
      Color_End       =   8438015
      Text            =   "确定"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   3615
      Left            =   0
      Picture         =   "提示窗体.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub Form_Load()
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub PButton1_Click()
    If Form3.Tishiback = 5 Then
        Form5.Enabled = True
    ElseIf Form3.Tishiback = 7 Then
        Form7.Enabled = True
    ElseIf Form3.Tishiback = 8 Then
        Form8.Enabled = True
    ElseIf Form3.Tishiback = 9 Then
        Form9.Enabled = True
    ElseIf Form3.Tishiback = 23 Then
        Form23.Enabled = True
    ElseIf Form3.Tishiback = 24 Then
        Form24.Enabled = True
    ElseIf Form3.Tishiback = 25 Then
        Form25.Enabled = True
    ElseIf Form3.Tishiback = 26 Then
        Form26.Enabled = True
    ElseIf Form3.Tishiback = 27 Then
        Form27.Enabled = True
    ElseIf Form3.Tishiback = 28 Then
        Form28.Enabled = True
    ElseIf Form3.Tishiback = 29 Then
        Form29.Enabled = True
        Form45.Show
        Form1.baoshidu = Form1.baoshidu - 300
        Form1.yinshui = Form1.yinshui - 300
        Form1.tili = Form1.tili - 300
    ElseIf Form3.Tishiback = 44 Then
        Form44.Enabled = True
    End If
    Unload Form34
End Sub
