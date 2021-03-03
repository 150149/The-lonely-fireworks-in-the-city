VERSION 5.00
Begin VB.Form Form25 
   BorderStyle     =   0  'None
   Caption         =   "Form25"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11640
   Icon            =   "小吃商店界面2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11640
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   840
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   12640511
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   6
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   7
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   4
      Left            =   8640
      TabIndex        =   8
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   5
      Left            =   1440
      TabIndex        =   9
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   6
      Left            =   3240
      TabIndex        =   10
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   7
      Left            =   5040
      TabIndex        =   11
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   8
      Left            =   6840
      TabIndex        =   12
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "购买"
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   13
      Top             =   4920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "敬请期待"
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image13 
      Height          =   735
      Left            =   5640
      Picture         =   "小吃商店界面2.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $0"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   8640
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $7"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   20
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $8"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   5040
      TabIndex        =   19
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $6"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   18
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   17
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $15"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8640
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   15
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $5"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   8
      Left            =   6840
      Picture         =   "小吃商店界面2.frx":73A9
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   5
      Left            =   5040
      Picture         =   "小吃商店界面2.frx":847D
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Image Image11 
      Height          =   2175
      Left            =   8280
      Picture         =   "小吃商店界面2.frx":CE85
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image10 
      Height          =   2175
      Left            =   6360
      Picture         =   "小吃商店界面2.frx":D6A2
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   4680
      Picture         =   "小吃商店界面2.frx":DEBF
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   7
      Left            =   3240
      Picture         =   "小吃商店界面2.frx":E6DC
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Image Image9 
      Height          =   2175
      Left            =   2880
      Picture         =   "小吃商店界面2.frx":FAFF
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   6
      Left            =   1440
      Picture         =   "小吃商店界面2.frx":1031C
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1005
   End
   Begin VB.Image Image8 
      Height          =   2175
      Left            =   1080
      Picture         =   "小吃商店界面2.frx":10E4F
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   4
      Left            =   8640
      Picture         =   "小吃商店界面2.frx":1166C
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image7 
      Height          =   2175
      Left            =   8280
      Picture         =   "小吃商店界面2.frx":1366C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   3
      Left            =   6840
      Picture         =   "小吃商店界面2.frx":13E89
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image6 
      Height          =   2175
      Left            =   6480
      Picture         =   "小吃商店界面2.frx":16286
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   2
      Left            =   5040
      Picture         =   "小吃商店界面2.frx":16AA3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   2175
      Left            =   4680
      Picture         =   "小吃商店界面2.frx":17747
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "     $8"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   1
      Left            =   3120
      Picture         =   "小吃商店界面2.frx":17F64
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image4 
      Height          =   2175
      Left            =   2880
      Picture         =   "小吃商店界面2.frx":19EBF
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $10"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10440
      TabIndex        =   2
      Top             =   600
      Width           =   495
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
      Left            =   10080
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   0
      Left            =   1440
      Picture         =   "小吃商店界面2.frx":1A6DC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   1080
      Picture         =   "小吃商店界面2.frx":1C322
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1695
   End
   Begin VB.Image Image12 
      Height          =   6000
      Left            =   360
      Picture         =   "小吃商店界面2.frx":1CB3F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   10680
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private xuanzhong As Integer
Private zhizhen As Integer
Private Lastzhizhen As Integer

Private Sub Form_Load()
    
    KeyPreview = True
    
    '   Me.BackColor = &H80000
    'Dim rtn As Long
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA '   窗体透明
    
    'Me.BackColor = &H80000
    'Dim rtn As Long
    'Dim BorderStyler
    'BorderStyler = 0
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    '不通透,    半透明
    '    SetWindowLong hwnd, GWL_EXSTYLE, rtn Or WS_EX_LAYERED
    '    SetLayeredWindowAttributes hwnd, 0, 225, LWA_ALPHA   '(句柄  ,掩码颜色[RGB]  ,  透明程度[0-255],  )
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    Label3.Caption = "$" & Form1.money
    zhizhen = -1
    Timer1.Enabled = True
    Timer1.Interval = 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.BackColor = &H80FF&
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = Index
    If zhizhen <= 3 Then
        Label2.Left = X + Image1(zhizhen).Width
        Label2.Top = Y + Label2.Height
    ElseIf zhizhen > 3 Then
        Label2.Left = X + Image1(zhizhen).Width
        Label2.Top = Image1(Index).Top - Label2.Height
    End If
    
End Sub



Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zhizhen = -1
    Label10.BackColor = &H80FF&
End Sub



Private Sub Label10_Click()
    KeyPreview = False
    Label10.BackColor = &H80FF&
    Form1.KeyPreview = True
    Unload Form12
    Unload Form25
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.BackColor = &H40C0&
End Sub

Private Sub PButton1_Click(Index As Integer)
    Dim read_OK As Long
    Dim xg0 As Integer
    
    If Index = 0 Then
        xg0 = 1
    ElseIf Index = 1 Then
        xg0 = 3
    ElseIf Index = 2 Then
        xg0 = 5
    ElseIf Index = 3 Then
        xg0 = 12
    ElseIf Index = 4 Then
        xg0 = 14
    ElseIf Index = 5 Then
        xg0 = 16
    ElseIf Index = 6 Then
        xg0 = 32
    ElseIf Index = 7 Then
        xg0 = 37
    ElseIf Index = 8 Then
        xg0 = 29
    ElseIf Index = 9 Then
        Exit Sub
    End If
    
    Dim xg1 As String
    xg1 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr(xg0), "money", "0", xg1, 256, App.Path & "\item.ini")
    If Form1.money < CInt(xg1) Then
        Tishim = Tishi("提示", "购买失败，你没有足够的钱")
        Form3.Tishiback = 25
    Else
        Dim xg2 As String
        xg2 = String(10, 0)
        read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", xg2, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        Dim xg3 As Integer
        Dim xg4 As Integer
        Dim xg5 As String
        Dim xg6 As Integer
        Dim xg7 As Integer
        Dim xg8 As Integer
        xg5 = String(10, 0)
        xg3 = CInt(xg2)
        xg4 = 1
        For xg4 = 1 To xg3
            read_OK = GetPrivateProfileString("beibao", "beibao" & xg4, "", xg5, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            If xg5 = 0 And xg6 = 0 Then
                xg7 = CInt(xg5)
                xg6 = xg6 + 1
                Form1.money = Form1.money - CInt(xg1)
                Label3.Caption = "$" & Form1.money
                Dim write1 As Long
                
                write1 = WritePrivateProfileString("beibao", "beibao" & xg4, CStr(xg0), App.Path & "\save\save" & Form3.saveid & ".fsave")
                Tishim = Tishi("提示", "购买成功")
                Form3.Tishiback = 25
            ElseIf xg5 <> 0 Then
                xg8 = xg8 + 1
            End If
        Next
        If xg8 = CInt(xg2) Then
            Tishim = Tishi("提示", "购买失败，背包没有足够空间")
            Form3.Tishiback = 25
        End If
        
    End If
End Sub

Private Sub Timer1_Timer()
    If zhizhen <> -1 And zhizhen <> Lastzhizhen Then
        Lastzhizhen = zhizhen
        Readiteminformation
        Label2.Visible = True
    ElseIf zhizhen = -1 Then
        Label2.Visible = False
    End If
End Sub

Private Sub Readiteminformation()
    Dim read_OK As Long
    Dim xg0 As Integer
    
    If zhizhen = 0 Then
        xg0 = 1
    ElseIf zhizhen = 1 Then
        xg0 = 3
    ElseIf zhizhen = 2 Then
        xg0 = 5
    ElseIf zhizhen = 3 Then
        xg0 = 12
    ElseIf zhizhen = 4 Then
        xg0 = 14
    ElseIf zhizhen = 5 Then
        xg0 = 16
    ElseIf zhizhen = 6 Then
        xg0 = 32
    ElseIf zhizhen = 7 Then
        xg0 = 37
    ElseIf zhizhen = 8 Then
        xg0 = 29
    End If
    
    Dim xg1 As String
    Dim xg2 As String
    Dim xg3 As String
    Dim xg4 As String
    Dim xg5 As String
    
    xg1 = String(10, 0)
    xg2 = String(10, 0)
    xg3 = String(10, 0)
    xg4 = String(10, 0)
    xg5 = String(10, 0)
    read_OK = GetPrivateProfileString(xg0, "money", "0", xg1, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "name", "0", xg2, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "baoshidu", "0", xg3, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "yinshui", "0", xg4, 256, App.Path & "\item.ini")
    read_OK = GetPrivateProfileString(xg0, "tili", "0", xg5, 256, App.Path & "\item.ini")
    Dim xg6 As String
    Dim xg7 As String
    Dim xg8 As String
    Dim xg9 As String
    Dim xg10 As String
    xg6 = CInt(xg1)
    xg7 = CInt(xg5)
    xg8 = CInt(xg3)
    xg9 = CInt(xg4)
    xg2 = Replace(xg2, Chr(0), "")
    Label2.Caption = xg2 & Chr(13) & "金钱:" & xg6 & Chr(13) & "可加饱食度" & xg8 & Chr(13) & "可加饮水量" & xg9 & Chr(13) & "可加精力值" & xg7
    
    
End Sub
