VERSION 5.00
Begin VB.Form Form36 
   BorderStyle     =   0  'None
   Caption         =   "Form27"
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9015
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.Timer Timer1 
         Left            =   6720
         Top             =   1080
      End
      Begin 市井中孤傲的烟火.PButton PButton1 
         Height          =   735
         Left            =   6960
         TabIndex        =   1
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1296
         Text            =   "购买"
         Can_Text_Move   =   0   'False
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
         Left            =   5880
         TabIndex        =   4
         Top             =   3240
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label7 
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
         Left            =   8520
         TabIndex        =   2
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H000080FF&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   9015
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   9000
         X2              =   9000
         Y1              =   360
         Y2              =   5520
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   0
         Y1              =   360
         Y2              =   5520
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   9000
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Image Image1 
         Height          =   1000
         Index           =   0
         Left            =   360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1000
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   1
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   2
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   3
         Left            =   360
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   4
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   6
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   7
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Image Image1 
         Height          =   1005
         Index           =   8
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "Form36"
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
    Image1(0).Picture = LoadPicture(App.Path & "\picture\item\5.gif")
    Image1(1).Picture = LoadPicture(App.Path & "\picture\item\14.gif")
    Image1(2).Picture = LoadPicture(App.Path & "\picture\item\16.gif")
    Image1(3).Picture = LoadPicture(App.Path & "\picture\item\12.gif")
    Image1(4).Picture = LoadPicture(App.Path & "\picture\item\27.gif")
    Image1(5).Picture = LoadPicture(App.Path & "\picture\item\28.gif")
    Image1(6).Picture = LoadPicture(App.Path & "\picture\item\32.gif")
    Image1(7).Picture = LoadPicture(App.Path & "\picture\item\33.gif")
    Image1(8).Picture = LoadPicture(App.Path & "\picture\item\34.gif")
    
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
    
    Line2.y1 = Label1.Height
    Line2.y2 = Picture1.Height
    Line1.x1 = Picture1.Width - 10
    Line1.x2 = Picture1.Width - 10
    Line1.y1 = Label1.Height - 10
    Line1.y2 = Picture1.Height - 10
    Line3.x1 = 0
    Line3.y2 = Picture1.Height - 10
    Line3.y1 = Picture1.Height - 10
    Line3.x2 = Picture1.Width - 10
    
    zhizhen = -1
    Timer1.Enabled = True
    Timer1.Interval = 300
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Label7.BackColor = &H80FF&
End Sub

Private Sub Image1_Click(Index As Integer)
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50), Picture1.BackColor
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), Picture1.BackColor
    Picture1.Line (Image1(xuanzhong).Left - 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), Picture1.BackColor
    Picture1.Line (Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top - 50)-(Image1(xuanzhong).Left + Image1(xuanzhong).Width + 50, Image1(xuanzhong).Top + Image1(xuanzhong).Height + 50), Picture1.BackColor
    xuanzhong = Index
    Picture1.DrawWidth = 3
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50), vbRed
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top - 50)-(Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
    Picture1.Line (Image1(Index).Left - 50, Image1(Index).Top + Image1(Index).Height + 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
    Picture1.Line (Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top - 50)-(Image1(Index).Left + Image1(Index).Width + 50, Image1(Index).Top + Image1(Index).Height + 50), vbRed
    
End Sub



Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
    zhizhen = Index
    If zhizhen <= 3 Then
        Label2.Left = X + Image1(zhizhen).Width
        Label2.Top = y + Label2.Height
    ElseIf zhizhen > 3 Then
        Label2.Left = X + Image1(zhizhen).Width
        Label2.Top = Image1(Index).Top - Label2.Height
    End If
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Label7.BackColor = &H80FF&
End Sub





Private Sub Label7_Click()
    KeyPreview = False
    Label7.BackColor = &H80FF&
    Form1.KeyPreview = True
    Unload Form12
    Unload Form27
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Label7.BackColor = &H40C0&
End Sub

Private Sub PButton1_Click()
    Dim read_OK As Long
    Dim xg0 As Integer
    
    If xuanzhong = 0 Then
        xg0 = 5
    ElseIf xuanzhong = 1 Then
        xg0 = 14
    ElseIf xuanzhong = 2 Then
        xg0 = 16
    ElseIf xuanzhong = 3 Then
        xg0 = 12
    ElseIf xuanzhong = 4 Then
        xg0 = 27
    ElseIf xuanzhong = 5 Then
        xg0 = 28
    ElseIf xuanzhong = 6 Then
        xg0 = 32
    ElseIf xuanzhong = 7 Then
        xg0 = 33
    ElseIf xuanzhong = 8 Then
        xg0 = 34
    End If
    Dim xg1 As String
    xg1 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr(xg0), "money", "0", xg1, 256, App.Path & "\item.ini")
    If Form1.money < CInt(xg1) Then
        MsgBox "购买失败，你没有足够的钱"
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
                Dim write1 As Long
                
                write1 = WritePrivateProfileString("beibao", "beibao" & xg4, CStr(xg0), App.Path & "\save\save" & Form3.saveid & ".fsave")
                MsgBox "购买成功"
            ElseIf xg5 <> 0 Then
                xg8 = xg8 + 1
            End If
        Next
        If xg8 = CInt(xg2) Then
            MsgBox "购买失败，背包没有足够的空间"
        End If
        
    End If
End Sub




Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    zhizhen = -1
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
        xg0 = 5
    ElseIf zhizhen = 1 Then
        xg0 = 14
    ElseIf zhizhen = 2 Then
        xg0 = 16
    ElseIf zhizhen = 3 Then
        xg0 = 12
    ElseIf zhizhen = 4 Then
        xg0 = 27
    ElseIf zhizhen = 5 Then
        xg0 = 28
    ElseIf zhizhen = 6 Then
        xg0 = 32
    ElseIf zhizhen = 7 Then
        xg0 = 33
    ElseIf zhizhen = 8 Then
        xg0 = 34
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
    Debug.Print "读取" & xg2 & "金钱:" & xg1 & "可加饱食度" & xg3 & "可加饮水量" & xg4 & "可加精力值" & xg5
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
