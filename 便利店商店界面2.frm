VERSION 5.00
Begin VB.Form Form27 
   BorderStyle     =   0  'None
   Caption         =   "Form27"
   ClientHeight    =   11055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16035
   Icon            =   "�������̵����2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   16035
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   840
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   0
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   12640511
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   1
      Left            =   7680
      TabIndex        =   3
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   2
      Left            =   9480
      TabIndex        =   6
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   3
      Left            =   11280
      TabIndex        =   7
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   4
      Left            =   13080
      TabIndex        =   8
      Top             =   5160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   9
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   6
      Left            =   7680
      TabIndex        =   10
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   7
      Left            =   9480
      TabIndex        =   11
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   8
      Left            =   11280
      TabIndex        =   12
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "����"
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   375
      Index           =   9
      Left            =   12960
      TabIndex        =   13
      Top             =   7800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   8438015
      Color_Begin     =   33023
      Color_End       =   8438015
      Text            =   "�����ڴ�"
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   22
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image13 
      Height          =   615
      Left            =   9000
      Picture         =   "�������̵����2.frx":08CA
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $0"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   13080
      TabIndex        =   21
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $5"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   11280
      TabIndex        =   20
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $5"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9480
      TabIndex        =   19
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $6"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   7680
      TabIndex        =   18
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $3"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   17
      Top             =   7440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $3"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   13080
      TabIndex        =   16
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $10"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   11280
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "   $10"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   8
      Left            =   11280
      Picture         =   "�������̵����2.frx":73A9
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   5
      Left            =   9480
      Picture         =   "�������̵����2.frx":AFF0
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image11 
      Height          =   2175
      Left            =   12720
      Picture         =   "�������̵����2.frx":BF30
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image10 
      Height          =   2175
      Left            =   10920
      Picture         =   "�������̵����2.frx":C74D
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   2175
      Left            =   9120
      Picture         =   "�������̵����2.frx":CF6A
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   7
      Left            =   7680
      Picture         =   "�������̵����2.frx":D787
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image9 
      Height          =   2175
      Left            =   7320
      Picture         =   "�������̵����2.frx":EBAA
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   6
      Left            =   5880
      Picture         =   "�������̵����2.frx":F3C7
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Image Image8 
      Height          =   2175
      Left            =   5520
      Picture         =   "�������̵����2.frx":FF47
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   4
      Left            =   13080
      Picture         =   "�������̵����2.frx":10764
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Image Image7 
      Height          =   2175
      Left            =   12720
      Picture         =   "�������̵����2.frx":111ED
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   3
      Left            =   11280
      Picture         =   "�������̵����2.frx":11A0A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Image Image6 
      Height          =   2175
      Left            =   10920
      Picture         =   "�������̵����2.frx":13E07
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   2
      Left            =   9480
      Picture         =   "�������̵����2.frx":14624
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Image Image5 
      Height          =   2175
      Left            =   9120
      Picture         =   "�������̵����2.frx":15157
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $15"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7560
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   1
      Left            =   7560
      Picture         =   "�������̵����2.frx":15974
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Image Image4 
      Height          =   2175
      Left            =   7320
      Picture         =   "�������̵����2.frx":17974
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "    $5"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H000080FF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   13560
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4605
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Image Image1 
      Height          =   1005
      Index           =   0
      Left            =   5880
      Picture         =   "�������̵����2.frx":18191
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   5520
      Picture         =   "�������̵����2.frx":18E35
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image Image12 
      Height          =   9960
      Left            =   360
      Picture         =   "�������̵����2.frx":19652
      Stretch         =   -1  'True
      Top             =   240
      Width           =   14640
   End
End
Attribute VB_Name = "Form27"
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
    'SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA '   ����͸��
    
    'Me.BackColor = &H80000
    'Dim rtn As Long
    'Dim BorderStyler
    'BorderStyler = 0
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    '��ͨ͸,    ��͸��
    '    SetWindowLong hwnd, GWL_EXSTYLE, rtn Or WS_EX_LAYERED
    '    SetLayeredWindowAttributes hwnd, 0, 225, LWA_ALPHA   '(���  ,������ɫ[RGB]  ,  ͸���̶�[0-255],  )
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

End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.BackColor = &H80FF&
    zhizhen = -1
End Sub

Private Sub Label10_Click()
    KeyPreview = False
    Label10.BackColor = &H80FF&
    Form1.KeyPreview = True
    Unload Form12
    Unload Form27
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label10.BackColor = &H40C0&
End Sub

Private Sub PButton1_Click(Index As Integer)
    Dim read_OK As Long
    Dim xg0 As Integer
    
    If Index = 0 Then
        xg0 = 5
    ElseIf Index = 1 Then
        xg0 = 14
    ElseIf Index = 2 Then
        xg0 = 16
    ElseIf Index = 3 Then
        xg0 = 12
    ElseIf Index = 4 Then
        xg0 = 27
    ElseIf Index = 5 Then
        xg0 = 28
    ElseIf Index = 6 Then
        xg0 = 32
    ElseIf Index = 7 Then
        xg0 = 33
    ElseIf Index = 8 Then
        xg0 = 34
    End If
    
    Dim xg1 As String
    xg1 = String(10, 0)
    read_OK = GetPrivateProfileString(CStr(xg0), "money", "0", xg1, 256, App.Path & "\item.ini")
    If Form1.money < CInt(xg1) Then
        Tishim = Tishi("��ʾ", "����ʧ�ܣ���û���㹻��Ǯ")
        Form3.Tishiback = 27
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
                Label3.Caption = "$" & Form1.money
                write1 = WritePrivateProfileString("beibao", "beibao" & xg4, CStr(xg0), App.Path & "\save\save" & Form3.saveid & ".fsave")
                Tishim = Tishi("��ʾ", "����ɹ�")
                Form3.Tishiback = 27
            ElseIf xg5 <> 0 Then
                xg8 = xg8 + 1
            End If
        Next
        If xg8 = CInt(xg2) Then
            Tishim = Tishi("��ʾ", "����ʧ�ܣ�����û���㹻�ռ�")
            Form3.Tishiback = 27
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
    Label2.Caption = xg2 & Chr(13) & "��Ǯ:" & xg6 & Chr(13) & "�ɼӱ�ʳ��" & xg8 & Chr(13) & "�ɼ���ˮ��" & xg9 & Chr(13) & "�ɼӾ���ֵ" & xg7
    
    
End Sub
