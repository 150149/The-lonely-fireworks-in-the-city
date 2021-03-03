VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "主界面"
   ClientHeight    =   16200
   ClientLeft      =   6285
   ClientTop       =   3855
   ClientWidth     =   28800
   LinkTopic       =   "Form2"
   Picture         =   "主界面.frx":0000
   ScaleHeight     =   16200
   ScaleWidth      =   28800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer15 
      Left            =   19080
      Top             =   0
   End
   Begin VB.Timer Timer14 
      Left            =   10080
      Top             =   0
   End
   Begin VB.Timer Timer13 
      Left            =   9600
      Top             =   0
   End
   Begin VB.Timer Timer12 
      Left            =   9120
      Top             =   0
   End
   Begin VB.Timer Timer11 
      Left            =   8640
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton8 
      Height          =   375
      Left            =   720
      TabIndex        =   34
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "解锁"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Timer Timer10 
      Left            =   18600
      Top             =   0
   End
   Begin VB.Timer Timer9 
      Index           =   0
      Left            =   18120
      Top             =   0
   End
   Begin VB.Timer Timer8 
      Left            =   17640
      Top             =   0
   End
   Begin VB.Timer Timer7 
      Left            =   17160
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Left            =   16680
      Top             =   0
   End
   Begin VB.Timer Timer5 
      Left            =   15600
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Left            =   15120
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   375
      Left            =   19080
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
      _ExtentX        =   1720
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "暂停"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin VB.Timer Timer3 
      Left            =   14640
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   14160
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar3 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Top       =   8454143
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar2 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Top       =   33023
      Color_Back      =   0
   End
   Begin VB.Timer Timer1 
      Left            =   13680
      Top             =   0
   End
   Begin 市井中孤傲的烟火.PProgressBar PProgressBar4 
      Height          =   135
      Index           =   0
      Left            =   10440
      TabIndex        =   25
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   238
      Color_Back      =   0
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "打工"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton2 
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "交易"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton5 
      Height          =   375
      Left            =   720
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "睡觉"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton6 
      Height          =   375
      Left            =   19080
      TabIndex        =   33
      Top             =   1680
      Width           =   1695
      _ExtentX        =   1931
      _ExtentY        =   661
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "加速"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PScreen PScreen1 
      Height          =   855
      Left            =   17040
      TabIndex        =   26
      Top             =   1320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      Color_Text      =   16777215
      Color_Text_Back =   0
      Text            =   "06:00"
      Size            =   50
   End
   Begin VB.Image Label1 
      Height          =   1425
      Left            =   10320
      Picture         =   "主界面.frx":13C04F
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1440
   End
   Begin VB.Image Image12 
      Height          =   2295
      Left            =   12840
      Picture         =   "主界面.frx":13D422
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "超市"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   46
      Top             =   10200
      Width           =   480
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "家"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   21960
      TabIndex        =   45
      Top             =   9480
      Width           =   240
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "证券交易所"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   14880
      TabIndex        =   43
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "饭店"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9720
      TabIndex        =   42
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "0天6时0分"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   19320
      TabIndex        =   41
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "工作剩余时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10680
      TabIndex        =   40
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "体力值"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   39
      Top             =   1440
      Width           =   720
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "饮水度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   38
      Top             =   960
      Width           =   720
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "饱食度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3960
      TabIndex        =   37
      Top             =   480
      Width           =   720
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "战斗"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   36
      Top             =   9480
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "背包"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      TabIndex        =   35
      Top             =   6360
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   975
   End
   Begin VB.Image Image9 
      Height          =   1935
      Left            =   720
      Picture         =   "主界面.frx":13EB04
      Stretch         =   -1  'True
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   2415
      Left            =   480
      Picture         =   "主界面.frx":147910
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Label Label26 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   4680
      TabIndex        =   32
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxx建筑"
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
      Left            =   360
      TabIndex        =   31
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   2880
      Width           =   135
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   360
      Width           =   855
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   375
      Left            =   16080
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   661
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   855
      Left            =   3840
      TabIndex        =   20
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2295
      Left            =   7680
      TabIndex        =   19
      Top             =   10560
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Height          =   2055
      Left            =   6120
      TabIndex        =   18
      Top             =   6840
      Width           =   2775
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   20520
      TabIndex        =   17
      Top             =   0
      Width           =   8295
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2175
      Left            =   16440
      TabIndex        =   16
      Top             =   9600
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   4440
      TabIndex        =   15
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2175
      Left            =   21000
      TabIndex        =   12
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2535
      Left            =   23040
      TabIndex        =   11
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2175
      Left            =   22920
      TabIndex        =   10
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   1455
      Left            =   17040
      TabIndex        =   9
      Top             =   12120
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2295
      Left            =   18360
      TabIndex        =   8
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   16920
      TabIndex        =   7
      Top             =   14400
      Width           =   5055
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   240
      Picture         =   "主界面.frx":14B815
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   240
      Picture         =   "主界面.frx":15EF5A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Picture         =   "主界面.frx":15FA02
      Stretch         =   -1  'True
      Top             =   840
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Picture         =   "主界面.frx":16052B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   1935
      Left            =   8760
      TabIndex        =   1
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "状态"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   -120
      Picture         =   "主界面.frx":160E4E
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   615
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   19080
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14640
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   3495
      Left            =   14400
      TabIndex        =   14
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "房地产公司"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   19080
      TabIndex        =   44
      Top             =   8280
      Width           =   1200
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   18960
      Shape           =   4  'Rounded Rectangle
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   21720
      Shape           =   4  'Rounded Rectangle
      Top             =   9360
      Width           =   735
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   2895
      Left            =   21480
      TabIndex        =   13
      Top             =   9000
      Width           =   1455
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   7680
      Shape           =   4  'Rounded Rectangle
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "便利店"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   17520
      TabIndex        =   47
      Top             =   11760
      Width           =   720
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   17280
      Shape           =   4  'Rounded Rectangle
      Top             =   11640
      Width           =   1215
   End
   Begin VB.Label Label38 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "建筑公司"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   19320
      TabIndex        =   48
      Top             =   13920
      Width           =   960
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   19080
      Shape           =   4  'Rounded Rectangle
      Top             =   13800
      Width           =   1335
   End
   Begin VB.Image Image11 
      Height          =   3615
      Left            =   16200
      Picture         =   "主界面.frx":1618F6
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3615
   End
   Begin VB.Shape Shape15 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   5775
   End
   Begin VB.Shape Shape16 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   3855
   End
   Begin VB.Shape Shape17 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RetVal As Long
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Const MAX_FILENAME_LEN = 256
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public baoshidu As Integer                                                      '定义饱食度
Public yinshui As Integer                                                       '定义饮水度
Public tili As Integer                                                          '定义体力值
Public tili2 As Integer                                                         '定义体力是否恢复
Public money As Integer                                                         '定义金钱
Public mina As Integer                                                          '定义分钟
Public seca As Integer                                                          '定义秒
Public houra As Integer                                                         '定义小时
Public Position As Integer                                                      '定义位置ID
Public workok As Integer                                                        '定义工作状态
Public Shejiao As Integer                                                       '定义社交技能
Public Zhili As Integer                                                         '定义智力技能
Public Guanchali As Integer                                                     '定义观察力技能
Public Naili As Integer                                                         '定义耐力技能
Public difficult As Integer                                                     '定义难度系数
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public formback As Integer                                                      '定义返回窗体ID
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Const LWA_COLORKEY = &H1
Private ib As Integer                                                           '定义透明度
Private Addai As Integer                                                        '定义AI增加的ID
Public work As Integer                                                          '定义工作ID
Public Worktime As Integer                                                      '定义已工作时间
Public Accident As Integer                                                      '定义突发事件ID
Private beibaomax As String                                                     '定义背包最大格数
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public allNaili As Integer                                                      '定义工作增加的耐力值
Public allZhili As Integer                                                      '定义工作增加的智力值
Public allShejiao As Integer                                                    '定义工作增加的社交值
Public allGuanchali As Integer                                                  '定义工作增加的观察力值
Public Alltime As Integer                                                       '定义工作总时间
Public Daodepingzhi As Integer                                                  '定义道德品质值
Public jingye As Integer                                                        '定义敬业值
Public haoxue As Integer                                                        '定义好学值
Public dandang As Integer                                                       '定义担当值
Public chengxin As Integer                                                      '定义诚信值
Public Lastposition As Integer                                                  '定义工作开始时所在的位置ID
Public n As Integer                                                             '定义Gal显示字符长度
Public xianshi As Integer                                                       '定义Gal完全显示内容
Dim zd As String                                                                '定义Gal当前显示的内容
Dim dq As String                                                                '定义Gal读取的内容
Public Dj As Integer                                                            '定义Gal点击次数（显示的字段ID）
Private Speedup As Boolean                                                      '定义是否加速
Public Zhiwei As String                                                         '定义职位
Public moshi As Integer                                                         '定义游戏模式
Public Pengren As Boolean                                                       '定义是否可以烹饪
Public Driven As Boolean                                                        '定义是否可以驾驶
Public Jinrong As Boolean                                                       '定义是否可以进入金融领域
Private timermode(10) As Integer                                                '定义时钟控件模式
Private HumanMove As Integer
Public Checktime As Integer

Private Sub Form_GotFocus()
    If Form1.Timer2.Enabled = False And Form1.workok = 1 Then
        Form1.Timer2.Enabled = True
        Form1.workok = 2
    End If
    If Form1.formback = 1 Then
        Form1.Show
    ElseIf Form1.formback = 2 Then
        Form2.Show
    ElseIf Form1.formback = 5 Then
        Form5.Show
    ElseIf Form1.formback = 8 Then
        Form8.Show
    ElseIf Form1.formback = 10 Then
        Form10.Show
    ElseIf Form1.formback = 11 Then
        Form11.Show
    ElseIf Form1.formback = 23 Then
        Form23.Show
    ElseIf Form1.formback = 24 Then
        Form24.Show
    ElseIf Form1.formback = 25 Then
        Form25.Show
    ElseIf Form1.formback = 26 Then
        Form26.Show
    ElseIf Form1.formback = 27 Then
        Form27.Show
    ElseIf Form1.formback = 28 Then
        Form28.Show
    ElseIf Form1.formback = 34 Then
        Form34.Show
    ElseIf Form1.formback = 32 Then
        Form1.Show
    ElseIf Form1.formback = 37 Then
        Form37.Show
    Else
        Form1.Show
    End If
    Form1.Timer3.Enabled = True
    Form1.Timer1.Enabled = True
End Sub

'==========================当键盘按键松开，才能重新恢复体力==========================================
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    tili2 = 1                                                                   '修改是否恢复体力为true
    If Timer11.Enabled = True Then
        Timer11.Enabled = False
        Label1.Picture = LoadPicture(App.Path & "\picture\human\rise.gif")
    End If
    If Timer12.Enabled = True Then
        Timer12.Enabled = False
        Label1.Picture = LoadPicture(App.Path & "\picture\human\down.gif")
    End If
    If Timer13.Enabled = True Then
        Timer13.Enabled = False
        Label1.Picture = LoadPicture(App.Path & "\picture\human\left.gif")
    End If
    If Timer14.Enabled = True Then
        Timer14.Enabled = False
        Label1.Picture = LoadPicture(App.Path & "\picture\human\right.gif")
    End If
End Sub
'====================================================================================================
Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入,使窗体控件自动缩放
    Form1.Width = Screen.Width                                                  '改变窗体大小为屏幕大小
    Form1.Height = Screen.Height                                                '改变窗体大小为屏幕大小
    
    KeyPreview = True                                                           '确保窗体响应键盘按下事件
    Timer1.Enabled = True                                                       '开启生命指标模块
    Timer1.Interval = 1000                                                      '确定生命指标模块刷新频率
    Timer3.Enabled = True                                                       '开启世界时间模块
    Timer3.Interval = 1000                                                      '确定世界时间模块刷新频率
    Timer4.Enabled = False                                                      '开启窗体焦点判断模块
    Timer4.Interval = 1000                                                      '确定窗体焦点判断模块刷新频率
    Timer2.Enabled = False                                                      '关闭工作模块
    If Form1.moshi = 2 Then
        Timer5.Enabled = True                                                   '开启突发事件模块
        Timer5.Interval = 1000                                                  '确定突发事件刷新频率
    End If
    Timer10.Enabled = False                                                     '关闭AI模块
    Form6.Show                                                                  '当窗体失去焦点时（切换其他程序）,通过暂停界面返回游戏
    Form1.Show                                                                  '让主窗体在顶部
    'Form32.Show                                                                 '显示debug窗口
    ' 'unload form32
    Timer6.Enabled = True                                                       '开启背景音乐播完自动切换功能
    Timer6.Interval = 1000                                                      '确定检测音乐是否播完的频率
    'Timer10.Enabled = True'AI模块
    'Timer10.Interval = 1000'AI模块
    'Label20(0).Visible = False'AI 模块
    Randomize                                                                   '初始化VB随机数发生器
    WindowsMediaPlayer1.URL = App.Path & "\music\background\" & CStr(Int(Rnd * (12 - 1 + 1)) + 1) & ".mp3" '进入游戏播放背景音乐
    ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "当前播放的音乐为" & WindowsMediaPlayer1.URL 'Debng输出音乐播放路径
    PProgressBar4(0).Visible = False                                            '隐藏工作进度条
    If Dir(App.Path & "\text.ini") = "" Then                                    '判断text.ini（Gal文字）是否存在
        Open "text.ini" For Output As #11
        Print #11, ""
        Close #11
        Dim a As String
        Dim b As Integer
        Dim c As Long
        For b = 1 To 500
            c = WritePrivateProfileString("text", CStr(b), "", App.Path & "\text.ini")
        Next
    End If                                                                      '不存在即写入文件
    Dim read_OK As Long
    dq = String(256, 0)
    read_OK = GetPrivateProfileString("text", CStr(Dj), "", dq, 256, App.Path & "\text.ini") '读取Gal上次阅读位置
    zd = dq
    If Form1.moshi = 2 Then
        PButton6.Visible = False
        Image9.Visible = False
    ElseIf Form1.moshi = 0 Then
        PButton6.Visible = False
    End If
    
    Label24.Visible = False
    Label19.Visible = False
    Shape1.Visible = False
    Label27.Visible = False
    Label28.Visible = False
    Label29.Visible = False
    Label30.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
    Shape5.Visible = False
    Shape6.Visible = False
    Label31.Visible = False
    Shape7.Visible = False
    Label32.Visible = False
    Shape8.Visible = False
    Shape9.Visible = False
    Shape10.Visible = False
    Shape11.Visible = False
    Shape12.Visible = False
    Shape13.Visible = False
    Label33.Visible = False
    Label34.Visible = False
    Label35.Visible = False
    Label36.Visible = False
    Label37.Visible = False
    Shape14.Visible = False
    Label38.Visible = False
    
    
    Timer11.Enabled = False
    Timer12.Enabled = False
    Timer13.Enabled = False
    Timer14.Enabled = False
    Timer15.Enabled = True
    Timer15.Interval = 1000
End Sub
'=====================================自动缩放控件======================================================
'=======================================================================================================
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

Private Sub Form_LostFocus()
    Dim thWnd As Long
    thWnd = GetForegroundWindow
    If FindWindow(vbNullString, "Form2") <> 0 And thWnd <> Form1.hwnd Then
        formback = 2
        Debug.Print "2"
    ElseIf FindWindow(vbNullString, "Form5") <> 0 And thWnd <> Form1.hwnd Then
        formback = 5
        Debug.Print "5"
    ElseIf FindWindow(vbNullString, "Form8") <> 0 And thWnd <> Form1.hwnd Then
        formback = 8
        Debug.Print "8"
    ElseIf FindWindow(vbNullString, "Form6") <> 0 And thWnd = Form1.hwnd Then
        Debug.Print "6"
    ElseIf FindWindow(vbNullString, "Form10") <> 0 And thWnd <> Form1.hwnd Then
        formback = 10
        Debug.Print "10"
    ElseIf FindWindow(vbNullString, "Form11") <> 0 And thWnd <> Form1.hwnd Then
        formback = 11
        Debug.Print "11"
    ElseIf FindWindow(vbNullString, "Form23") <> 0 And thWnd <> Form1.hwnd Then
        formback = 23
        Debug.Print "23"
    ElseIf FindWindow(vbNullString, "Form24") <> 0 And thWnd <> Form1.hwnd Then
        formback = 24
        Debug.Print "24"
    ElseIf FindWindow(vbNullString, "Form25") <> 0 And thWnd <> Form1.hwnd Then
        formback = 25
        Debug.Print "25"
    ElseIf FindWindow(vbNullString, "Form26") <> 0 And thWnd <> Form1.hwnd Then
        formback = 26
        Debug.Print "26"
    ElseIf FindWindow(vbNullString, "Form27") <> 0 And thWnd <> Form1.hwnd Then
        formback = 27
        Debug.Print "27"
    ElseIf FindWindow(vbNullString, "Form28") <> 0 And thWnd <> Form1.hwnd Then
        formback = 28
        Debug.Print "28"
    ElseIf FindWindow(vbNullString, "Form31") <> 0 And thWnd <> Form1.hwnd Then
        formback = 31
        Debug.Print "31"
        ' ElseIf FindWindow(vbNullString, "Form32") <> 0 And thWnd <> Form1.hwnd Then
        '     formback = 32
        '     Debug.Print "32"
    ElseIf FindWindow(vbNullString, "Form34") <> 0 And thWnd <> Form1.hwnd Then
        formback = 34
        Debug.Print "34"
    ElseIf FindWindow(vbNullString, "Form37") <> 0 And thWnd <> Form1.hwnd Then
        formback = 37
        Debug.Print "37"
    ElseIf thWnd = Form1.hwnd Then
        formback = 1
        Debug.Print "1"
    Else
        Timer3.Enabled = False
        Timer1.Enabled = False
        If Timer2.Enabled = True Then
            Timer2.Enabled = False
            workok = 1
            Form6.Show
            formback = 1
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label24.Visible = True Then Label24.Visible = False
    If Shape1.Visible = True Then Shape1.Visible = False
    If Label27.Visible = True Then Label27.Visible = False
    If Label28.Visible = True Then Label28.Visible = False
    If Label29.Visible = True Then Label29.Visible = False
    If Label30.Visible = True Then Label30.Visible = False
    If Shape2.Visible = True Then Shape2.Visible = False
    If Shape3.Visible = True Then Shape3.Visible = False
    If Shape4.Visible = True Then Shape4.Visible = False
    If Shape6.Visible = True Then Shape6.Visible = False
    If Shape7.Visible = True Then Shape7.Visible = False
    If Shape5.Visible = True Then Shape5.Visible = False
    If Shape8.Visible = True Then Shape8.Visible = False
    If Shape9.Visible = True Then Shape9.Visible = False
    If Shape10.Visible = True Then Shape10.Visible = False
    If Shape11.Visible = True Then Shape11.Visible = False
    If Shape12.Visible = True Then Shape12.Visible = False
    If Shape13.Visible = True Then Shape13.Visible = False
    If Shape14.Visible = True Then Shape14.Visible = False
    If Label19.Visible = True Then Label19.Visible = False
    If Label31.Visible = True Then Label31.Visible = False
    If Shape7.Visible = True Then Shape7.Visible = False
    If Label32.Visible = True Then Label32.Visible = False
    If Label33.Visible = True Then Label33.Visible = False
    If Label34.Visible = True Then Label34.Visible = False
    If Label36.Visible = True Then Label36.Visible = False
    If Label35.Visible = True Then Label35.Visible = False
    If Label37.Visible = True Then Label37.Visible = False
    If Label38.Visible = True Then Label38.Visible = False
    If PButton3.Visible = True Then PButton3.Visible = False
    If Label32.Visible = True Then Label32.Visible = False
    If Shape7.Visible = True Then Shape7.Visible = False
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
'============================================================================================================

Private Sub Form_Unload(Cancel As Integer)
    Accident = 0                                                                '在退出窗体时清除突发事件ID，以便退出后游戏初始化
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape3.Visible = True
    Label28.Visible = True
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PButton3.Visible = False Then PButton3.Visible = True
    If Label32.Visible = False Then Label32.Visible = True
    If Shape7.Visible = False Then Shape7.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape4.Visible = True
    Label29.Visible = True
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape5.Visible = True
    Label30.Visible = True
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label24.Visible = False
    Shape1.Visible = False
    Label27.Visible = False
    Label28.Visible = False
    Label29.Visible = False
    Label30.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
    Shape5.Visible = False
End Sub

Private Sub Image8_Click()
    
                                                       '主界面不响应键盘事件
    Form12.Show                                                                 '打开半透明遮罩
    Form5.form5open = Form5.form5open + 1
    Form5.Show                                                                  '显示背包界面
    Form5.Timer1.Enabled = True                                                 '启动背包界面物品属性显示控件
    Form5.Timer1.Interval = 300                                                 '确定属性刷新频率
    If Form5.form5open = 2 Then                                                 '预加载背包界面内容
        Form5.PProgressBar1.Value = CInt(Form1.Shejiao) / 1000                  '预加载
        Form5.PProgressBar2.Value = CInt(Form1.Zhili) / 1000                    '预加载
        Form5.PProgressBar3.Value = CInt(Form1.Guanchali) / 1000                '预加载
        Form5.PProgressBar4.Value = CInt(Form1.Naili) / 1000                    '预加载
        Form5.Label3.Caption = "社交能力：" & CInt(Form1.Shejiao)               '预加载
        Form5.Label4.Caption = "智力：" & CInt(Form1.Zhili)                     '预加载
        Form5.Label5.Caption = "观察力：" & CInt(Form1.Guanchali)               '预加载
        Form5.Label6.Caption = "耐力：" & CInt(Form1.Naili)                     '预加载
        Dim read_OK As Long
        beibaomax = String(10, 0)
        read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", beibaomax, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        Dim beibaomaxa As Integer                                               '读取背包最大格数
        Dim beibaomaxb As Integer                                               '读取背包最大格数
        beibaomaxb = CInt(beibaomax)                                            '读取背包最大格数
        Dim picnumber As String                                                 '读取背包最大格数
        picnumber = String(10, 0)                                               '读取背包最大格数
        For beibaomaxa = 1 To beibaomaxb                                        '读取背包最大格数
            read_OK = GetPrivateProfileString("beibao", "beibao" & beibaomaxa, "0", picnumber, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            If picnumber <> 0 Then                                              '重新加载背包物品图片
                Form5.Image1(beibaomaxa).Picture = LoadPicture("")              '重新加载背包物品图片
                Form5.Image1(beibaomaxa).Picture = LoadPicture(App.Path & "\picture\item\" & Replace(picnumber, Chr(0), "") & ".gif")
             '   'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "加载物品图在" & CStr(beibaomaxa) & " " & "替换物品为" & Replace(picnumber, Chr(0), "")
            End If
        Next
    End If
    formback = 5
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label24.Visible = True
    Shape1.Visible = True
End Sub

Private Sub Image9_Click()
    Timer3.Enabled = False
    Timer1.Enabled = False
    Form46.Show
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape2.Visible = True
    Label27.Visible = True
End Sub


Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape11.Visible = True
    Label35.Visible = True
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape9.Visible = True
    Label33.Visible = True
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape12.Visible = True
    Label36.Visible = True
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape3.Visible = True
    Label28.Visible = True
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape4.Visible = True
    Label29.Visible = True
End Sub

Private Sub Label23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape5.Visible = True
    Label30.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape8.Visible = True
    Label19.Visible = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape14.Visible = True
    Label38.Visible = True
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape10.Visible = True
    Label34.Visible = True
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape13.Visible = True
    Label37.Visible = True
End Sub

'-----------------------------------工作模块----------------------------------------------
Private Sub PButton1_Click()
   ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "打开工作选择界面"         'debug输出
    Form12.Show                                                                 '显示半透明遮罩
    Form8.Show                                                                  '显示工作选择界面
End Sub
'-----------------------------------工作模块----------------------------------------------




Private Sub PButton2_Click()
    Form49.Show
End Sub

Private Sub PButton3_Click()                                                    '暂停键
    Timer3.Enabled = False                                                      '暂停世界世界模块                                              '
    Timer1.Enabled = False                                                      '暂停生命指标模块
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "timer2状态为" & CStr(Timer2.Enabled) 'debug输出
    If Timer2.Enabled = True Then                                               '当工作模块正在进行，暂停工作模块
        Timer2.Enabled = False
        workok = 1                                                              '保存工作模块工作状态
    End If
    Form6.Show                                                                  '显示暂停界面
End Sub



Private Sub PButton5_Click()
    If PButton5.Text = "起床" Then                                              '判断是否在睡觉状态
        Timer3.Enabled = True                                                   '恢复世界时间模块
        Timer3.Interval = 1000                                                  '恢复时间时间模块速率
        PButton5.Text = "睡觉"                                                  '更改状态
    Else
        PButton5.Text = "起床"                                                  '判断是否在睡觉状态
        Timer3.Enabled = True                                                   '加速世界时间模块
        Timer3.Interval = 100                                                   '确定加速倍率
    End If
End Sub



Private Sub PButton6_Click()
    If PButton6.Text = "普通速度" Then                                          '判断是否在工作加速状态
        Timer2.Enabled = True                                                   '恢复工作速度
        Timer2.Interval = 1000                                                  '恢复工作速度
        Timer1.Enabled = True                                                   '恢复生命指标模块速率
        Timer1.Interval = 1000                                                  '确定生命指标模块正常倍率
        Timer3.Enabled = True                                                   '恢复世界时间速度
        Timer3.Interval = 1000                                                  '确定世界时间速度
        PButton6.Text = "加快速度"                                              '更改工作加速状态
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "减速" & CStr(Timer3.Interval) 'debug输出
    Else                                                                        '判断是否在工作加速状态
        Timer2.Enabled = True                                                   '加速工作模块
        Timer2.Interval = 100                                                   '加速工作模块
        Timer1.Enabled = True                                                   '加速生命指标模块
        Timer1.Interval = 100                                                   '加速生命指标模块
        Timer3.Enabled = True                                                   '加速世界时间模块
        Timer3.Interval = 100                                                   '加速世界时间模块
        PButton6.Text = "普通速度"                                              '更改工作加速状态
      '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "快进" & CStr(Timer3.Interval) 'debug输出
    End If
End Sub

Private Sub PButton8_Click()
    Form12.Show
    Form44.Show
End Sub

Private Sub PProgressBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Single)
    Shape3.Visible = True
    Label28.Visible = True
End Sub

Private Sub PProgressBar2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Single)
    Shape4.Visible = True
    Label29.Visible = True
End Sub

Private Sub PProgressBar3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Single)
    Shape5.Visible = True
    Label30.Visible = True
End Sub

Private Sub PProgressBar4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Single)
    If Timer2.Enabled = False Then Exit Sub
    Shape6.Visible = True
    Label31.Visible = True
End Sub

Private Sub PScreen1_GotFocus()
    Shape7.Visible = True
    Label32.Visible = True
End Sub

'----------------------------------------人物生命指标模块--------------------------------------
Private Sub Timer1_Timer()
    
    
    If Form1.moshi = 0 Then
        baoshidu = baoshidu - 1 - Form3.baoshidux                               '每秒减少3/1000的饱食度
        yinshui = yinshui - 2 - Form3.yinshuix                                  '每秒减少6/1000的饮水度
    End If
    If Form1.moshi = 2 Then
        baoshidu = baoshidu - 3 - Form3.baoshidux                               '每秒减少3/1000的饱食度
        yinshui = yinshui - 6 - Form3.yinshuix                                  '每秒减少6/1000的饮水度
    End If
    If tili2 = 1 And tili < 1000 And Not Timer2.Enabled = True Then             '判断是否处于工作状态
        tili = tili + 1                                                         '非工作状态每秒恢复1/1000的体力值
        PProgressBar3.Value = CSng(tili / 1000)
        Label23.Caption = Format(CSng(tili / 1000), "0%")
    Else
        tili = tili - Form3.tilix                                               '处于工作状态，每秒减少工作所需体力
        PProgressBar3.Value = CSng(tili / 1000)
        Label23.Caption = Format(CSng(tili / 1000), "0%")
    End If
    
    PProgressBar1.Value = CSng(baoshidu / 1000)                                 '显示饱食度条
    PProgressBar2.Value = CSng(yinshui / 1000)                                  '显示饮水度条
    PProgressBar3.Value = CSng(tili / 1000)                                     '显示体力条
    Label11.Caption = money                                                     '显示金钱
    Label21.Caption = Format(CSng(baoshidu / 1000), "0%")                       '显示饱食度百分比
    Label22.Caption = Format(CSng(yinshui / 1000), "0%")                        '显示饮水度百分比
    
    
    If baoshidu <= 0 Or yinshui <= 0 Then                                       '判断饱食度或者饮水度是否为0
        Form12.Show                                                             '如果0则游戏结束
        Form31.Show
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Timer6.Enabled = False
        Timer10.Enabled = False
        Call Form31.ResizeInit(Me)                                              '在程序装入时必须加入
        Form31.PProgressBar1.Value = CInt(Form1.Shejiao) / 1000                 '预加载游戏结束界面
        Form31.PProgressBar2.Value = CInt(Form1.Zhili) / 1000
        Form31.PProgressBar3.Value = CInt(Form1.Guanchali) / 1000               '预加载
        Form31.PProgressBar4.Value = CInt(Form1.Naili) / 1000                   '预加载
        Form31.Label3.Caption = "社交能力：" & CStr(CInt(Form1.Shejiao))        '预加载
        Form31.Label4.Caption = "智力：" & CStr(CInt(Form1.Zhili))              '预加载
        Form31.Label5.Caption = "观察力：" & CStr(CInt(Form1.Guanchali))        '预加载
        Form31.Label6.Caption = "耐力：" & CStr(CInt(Form1.Naili))              '预加载
        Form31.Label2.Caption = "金钱:" & Form1.money                           '预加载
    End If                                                                      '预加载
    ' Form32.Label1.Caption = "饱食度：" & baoshidu                               '预加载
    ' Form32.Label2.Caption = "饮水度：" & yinshui                                '预加载
    ' Form32.Label3.Caption = "体力值：" & tili                                   '预加载
    ' Form32.Label4.Caption = "金钱：" & money                                    '预加载
    ' Form32.Label5.Caption = "饱食修正：" & Form3.baoshidux                      '预加载
    ' Form32.Label6.Caption = "饮水修正：" & Form3.yinshuix                       '预加载
    ' Form32.Label7.Caption = "体力修正：" & Form3.tilix                          '预加载
    ' Form32.Label8.Caption = "金钱修正：" & Form3.moneyx                         '预加载
    
    If Form1.tili <= 200 Then                                                   '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar3.Color_Top = &HFF&                                         '判断生命指标是否低于20%，如果是则显示条变红
        Label2.Caption = "我快要累死了"                                         '判断生命指标是否低于20%，如果是则显示条变红
    Else                                                                        '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar3.Color_Top = &H80FFFF                                      '判断生命指标是否低于20%，如果是则显示条变红
    End If                                                                      '判断生命指标是否低于20%，如果是则显示条变红
    If Form1.yinshui <= 200 Then                                                '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar2.Color_Top = &HFF&                                         '判断生命指标是否低于20%，如果是则显示条变红
        Label2.Caption = "我快要渴死了"                                         '判断生命指标是否低于20%，如果是则显示条变红
    Else                                                                        '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar2.Color_Top = &HFF7402                                      '判断生命指标是否低于20%，如果是则显示条变红
    End If                                                                      '判断生命指标是否低于20%，如果是则显示条变红
    If Form1.baoshidu <= 200 Then                                               '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar1.Color_Top = &HFF&                                         '判断生命指标是否低于20%，如果是则显示条变红
        Label2.Caption = "我快要饿死了"                                         '判断生命指标是否低于20%，如果是则显示条变红
    Else                                                                        '判断生命指标是否低于20%，如果是则显示条变红
        PProgressBar1.Color_Top = &H80FF&                                       '判断生命指标是否低于20%，如果是则显示条变红
    End If                                                                      '判断生命指标是否低于20%，如果是则显示条变红
End Sub                                                                         '判断生命指标是否低于20%，如果是则显示条变红
'----------------------------------------人物生命指标模块--------------------------------------


'----------------------------------控制模块------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If PButton5.Text = "起床" Then
        Call PButton5_Click
        Exit Sub
    End If
    
    If Label1.Top < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Top = Label1.Top + 300
    ElseIf Label1.Left < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Left = Label1.Left + 300
    ElseIf Form1.Height - Label1.Top < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Top = Label1.Top - 300
    ElseIf Form1.Width - Label1.Left < 100 Then
        Label2.Caption = "你不能再往前走了"
        Label1.Left = Label1.Left - 300
    ElseIf tili <= 0 Then
        Label2.Caption = "你没力气了"
    Else
        If KeyCode = 85 Then
            If Timer13.Enabled = False Then
                Timer13.Enabled = True
                Timer13.Interval = 100
                Timer11.Enabled = False
                Timer12.Enabled = False
                Timer14.Enabled = False
                HumanMove = 0
            End If
            Label1.Left = Label1.Left - 100
            Label1.Top = Label1.Top - 60
            tili2 = 0
            tili = tili - 1
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        ElseIf KeyCode = 87 Then
            Label1.Top = Label1.Top - 50
            Label1.Left = Label1.Left + 90
            If Timer11.Enabled = False Then
                Timer11.Enabled = True
                Timer11.Interval = 100
                Timer13.Enabled = False
                Timer12.Enabled = False
                Timer14.Enabled = False
                HumanMove = 0
            End If
            tili2 = 0
            tili = tili - 1
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        ElseIf KeyCode = 65 Then
            Label1.Top = Label1.Top - 51
            Label1.Left = Label1.Left - 90
            If Timer13.Enabled = False Then
                Timer13.Enabled = True
                Timer13.Interval = 100
                Timer11.Enabled = False
                Timer12.Enabled = False
                Timer14.Enabled = False
                HumanMove = 0
            End If
            tili2 = 0
            tili = tili - 1
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        ElseIf KeyCode = 83 Then
            Label1.Top = Label1.Top + 50
            Label1.Left = Label1.Left - 90
            If Timer12.Enabled = False Then
                Timer12.Enabled = True
                Timer12.Interval = 100
                Timer11.Enabled = False
                Timer13.Enabled = False
                Timer14.Enabled = False
                HumanMove = 0
            End If
            tili2 = 0
            tili = tili - 1
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        ElseIf KeyCode = 68 Then
            Label1.Left = Label1.Left + 70
            Label1.Top = Label1.Top + 45
            If Timer14.Enabled = False Then
                Timer14.Enabled = True
                Timer14.Interval = 100
                Timer11.Enabled = False
                Timer13.Enabled = False
                Timer12.Enabled = False
                HumanMove = 0
            End If
            tili2 = 0
            tili = tili - 1
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        ElseIf KeyCode = 13 Then
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
            Unload Form19
            Unload Form20
            Unload Form21
            Unload Form22
            Unload Form23
            Unload Form24
            Unload Form25
            Unload Form26
            Unload Form27
            Unload Form28
            Unload Form29
            Unload Form30
            Unload Form31
            'unload form32
            Unload Me
            End
        Else
            Select Case KeyCode
            Case vbKeyUp
                Label1.Top = Label1.Top - 100
            Case vbKeyDown
                Label1.Top = Label1.Top + 100
            Case vbKeyLeft
                Label1.Left = Label1.Left - 100
            Case vbKeyRight
                Label1.Left = Label1.Left + 100
            End Select
        End If
    End If
    If Label1.Top >= Label3.Top And Label1.Top <= Label3.Top + Label3.Height And Label1.Left >= Label3.Left And Label1.Left <= Label3.Left + Label3.Width Then
        Label2.Caption = "你现在在饭店内"
        Label25.Caption = "饭店"
        Position = 1
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton1.Visible = True
        PButton2.Visible = True
        Label25.Visible = True
        If Form1.moshi = 0 Then PButton8.Visible = True
        'pbutton8.Visible = True
        'Form10.Show
        'Timer1.Enabled = False
        'Timer3.Enabled = False
        'Timer5.Enabled = False
        'Form1.Hide
    ElseIf Label1.Top >= Label4.Top And Label1.Top <= Label4.Top + Label4.Height And Label1.Left >= Label4.Left And Label1.Left <= Label4.Left + Label4.Width Then
        Label2.Caption = "你现在在建筑公司内"
        Label25.Caption = "建筑公司"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton2.Visible = False
        PButton1.Visible = True
        If Form1.moshi = 0 Then PButton8.Visible = True
        'pbutton8.Visible = True
        PButton2.Visible = True
        Position = 2
    ElseIf Label1.Top >= Label5.Top And Label1.Top <= Label5.Top + Label5.Height And Label1.Left >= Label5.Left And Label1.Left <= Label5.Left + Label5.Width Then
        Label2.Caption = "你现在在房地产公司内"
        Label25.Caption = "房地产公司"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = False
        If Form1.moshi = 0 Then PButton8.Visible = True
        'pbutton8.Visible = True
        PButton2.Visible = False
        Position = 3
    ElseIf Label1.Top >= Label6.Top And Label1.Top <= Label6.Top + Label6.Height And Label1.Left >= Label6.Left And Label1.Left <= Label6.Left + Label6.Width Then
        Label2.Caption = "你现在在便利店内"
        Label25.Caption = "便利店"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        'pbutton8.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = True
        If Form1.moshi = 0 Then PButton8.Visible = True
        Position = 4
    ElseIf Label1.Top >= Label7.Top And Label1.Top <= Label7.Top + Label7.Height And Label1.Left >= Label7.Left And Label1.Left <= Label7.Left + Label7.Width Then
        Label2.Caption = "你现在在住宅楼1内"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 5
    ElseIf Label1.Top >= Label8.Top And Label1.Top <= Label8.Top + Label8.Height And Label1.Left >= Label8.Left And Label1.Left <= Label8.Left + Label8.Width Then
        Label2.Caption = "你现在在住宅楼2内"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 6
    ElseIf Label1.Top >= Label9.Top And Label1.Top <= Label9.Top + Label9.Height And Label1.Left >= Label9.Left And Label1.Left <= Label9.Left + Label9.Width Then
        Label2.Caption = "你现在在住宅楼3内"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 7
    ElseIf Label1.Top >= Label10.Top And Label1.Top <= Label10.Top + Label10.Height And Label1.Left >= Label10.Left And Label1.Left <= Label10.Left + Label10.Width Then
        Label2.Caption = "你现在在家内"
        Label25.Caption = "家"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = True
        PButton1.Visible = False
        PButton2.Visible = False
        'pbutton8.Visible = False
        Position = 8
    ElseIf Label1.Top >= Label12.Top And Label1.Top <= Label12.Top + Label12.Height And Label1.Left >= Label12.Left And Label1.Left <= Label12.Left + Label12.Width Then
        Label2.Caption = "你现在在证券交易所内"
        Label25.Caption = "证券交易所"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        'pbutton8.Visible = True
        PButton1.Visible = True
        If Form1.moshi = 0 Then PButton8.Visible = True
        Position = 9
    ElseIf Label1.Top >= Label13.Top And Label1.Top <= Label13.Top + Label13.Height And Label1.Left >= Label13.Left And Label1.Left <= Label13.Left + Label13.Width Then
        Label2.Caption = "你现在在教育机构内"
        Label25.Caption = "教育机构"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = False
        Position = 10
    ElseIf Label1.Top >= Label14.Top And Label1.Top <= Label14.Top + Label14.Height And Label1.Left >= Label14.Left And Label1.Left <= Label14.Left + Label14.Width Then
        Label2.Caption = "你现在在客运公司内"
        Label25.Caption = "客运公司"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = True
        Position = 11
    ElseIf Label1.Top / 2 >= Label15.Top And Label1.Top / 2 <= Label15.Top + Label15.Height And Label1.Left / 2 >= Label15.Left And Label1.Left / 2 <= Label15.Left + Label15.Width Then
        Label1.Left = Label1.Left - 500
    ElseIf Label1.Top >= Label16.Top And Label1.Top <= Label16.Top + Label16.Height And Label1.Left >= Label16.Left And Label1.Left <= Label16.Left + Label16.Width Then
        Label2.Caption = "你现在在公园内"
        Label25.Caption = "公园"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = False
        PButton2.Visible = False
        Position = 12
    ElseIf Label1.Top >= Label17.Top And Label1.Top <= Label17.Top + Label17.Height And Label1.Left >= Label17.Left And Label1.Left <= Label17.Left + Label17.Width Then
        Label2.Caption = "你现在在超市内"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Label25.Caption = "超市"
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = True
        Position = 13
    ElseIf Label1.Top >= Label18.Top And Label1.Top <= Label18.Top + Label18.Height And Label1.Left >= Label18.Left And Label1.Left <= Label18.Left + Label18.Width Then
        Label2.Caption = "你现在在软件公司内"
        Label25.Caption = "软件公司"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = False
        Position = 14
    Else
        Label2.Caption = "你现在在街道上"
        If PButton5.Visible = True Then PButton5.Visible = False
        If PButton1.Visible = True Then PButton1.Visible = False
        'pbutton8.Visible = False
        If PButton2.Visible = True Then PButton2.Visible = False
        If PButton8.Visible = True Then PButton8.Visible = False
        If Shape16.Visible = True Then Shape16.Visible = False
        If Label25.Visible = True Then Label25.Visible = False
        If Label20.Visible = True Then Label20.Visible = False
        Position = 0
    End If
End Sub
'--------------------------------------控制模块--------------------------------------------------------



Private Sub Timer11_Timer()
    HumanMove = HumanMove + 1
    Label1.Picture = LoadPicture(App.Path & "\picture\human\rise" & CStr(HumanMove) & ".gif")
    If HumanMove = 3 Then HumanMove = 1
End Sub

Private Sub Timer12_Timer()
    HumanMove = HumanMove + 1
    Label1.Picture = LoadPicture(App.Path & "\picture\human\down" & CStr(HumanMove) & ".gif")
    If HumanMove = 3 Then HumanMove = 1
End Sub

Private Sub Timer13_Timer()
    HumanMove = HumanMove + 1
    Label1.Picture = LoadPicture(App.Path & "\picture\human\left" & CStr(HumanMove) & ".gif")
    If HumanMove = 3 Then HumanMove = 1
End Sub

Private Sub Timer14_Timer()
    HumanMove = HumanMove + 1
    Label1.Picture = LoadPicture(App.Path & "\picture\human\right" & CStr(HumanMove) & ".gif")
    If HumanMove = 3 Then HumanMove = 1
End Sub

Private Sub Timer15_Timer()
    Dim read_OK  As Long
    Dim write1 As Long
    Checktime = Checktime + 1
    If Checktime = 3 Then
        Timer15.Enabled = False
        Dim first1 As String
        first1 = String(10, 0)
        read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        If Replace(first1, Chr(0), "") = "" Then
            Form12.Show
            Form37.Show
            Form37.PPictureBox1.Visible = True
            Form37.PPictureBox2.Visible = False
            Form37.PPictureBox3.Visible = False
            Form37.PPictureBox4.Visible = False
            Form37.PPictureBox5.Visible = False
            write1 = WritePrivateProfileString("guide", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        ElseIf Replace(first1, Chr(0), "") = "0" And Form1.moshi = 0 Then
            Form12.Show
            Form37.Show
            Form37.PPictureBox1.Visible = False
            Form37.PPictureBox2.Visible = True
            Form37.PPictureBox3.Visible = False
            Form37.PPictureBox4.Visible = False
            Form37.PPictureBox5.Visible = False
            write1 = WritePrivateProfileString("guide", "1", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
        ElseIf Replace(first1, Chr(0), "") = "0" And Form1.moshi = 2 Then
            Form12.Show
            Form37.Show
            Form37.PPictureBox1.Visible = False
            Form37.PPictureBox2.Visible = False
            Form37.PPictureBox3.Visible = True
            Form37.PPictureBox4.Visible = False
            Form37.PPictureBox5.Visible = False
            write1 = WritePrivateProfileString("guide", "1", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
        ElseIf Replace(first1, Chr(0), "") = "1" Then
            Form12.Show
            Form37.Show
            Form37.PPictureBox1.Visible = False
            Form37.PPictureBox2.Visible = False
            Form37.PPictureBox3.Visible = False
            Form37.PPictureBox4.Visible = False
            Form37.PPictureBox5.Visible = True
            write1 = WritePrivateProfileString("guide", "1", "2", App.Path & "\save\save" & Form3.saveid & ".fsave")
        End If
    End If
End Sub

Private Sub Timer2_Timer()
    Label23.Caption = Format(CSng(tili / 1000), "0%")
    If Alltime = 0 Then
        Timer2.Enabled = False
        ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "工作时间为0，停止timer2"
        Exit Sub
    Else
        
        Form1.PProgressBar4(1).Value = CSng(Worktime / Alltime)
    End If
    If Lastposition <> Position Then
        Form3.baoshidux = 1
        Form3.yinshuix = 1
        Form3.tilix = 1
        Form3.moneyx = 1
        allNaili = 0
        allGuanchali = 0
        allShejiao = 0
        allZhili = 0
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "离开工作地点，停止timer2"
        Timer2.Enabled = False
        Worktime = 0
        Alltime = 0
        If PButton6.Text = "普通速度" Then
            PButton6.Text = "加快速度"
            '    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "减速" & CStr(Timer3.Interval)
        End If
        PButton6.Visible = False
        Timer2.Enabled = False
        Timer2.Interval = 1000
        Timer1.Enabled = True
        Timer1.Interval = 1000
        Timer3.Enabled = True
        Timer3.Interval = 1000
        Unload Form1.PProgressBar4(1)
        Exit Sub
    End If
    Lastposition = Position
    Worktime = Worktime + 1
    If Alltime <> 0 Then
        Form1.PProgressBar4(1).Top = Form1.Label1.Top - 300
        Form1.PProgressBar4(1).Left = Form1.Label1.Left
        Label31.Top = Form1.Label1.Top - 320 - Shape6.Height
        Label31.Left = Form1.Label1.Left + 50
        Shape6.Top = Form1.Label1.Top - 400 - Shape6.Height
        Shape6.Left = Form1.Label1.Left
        Label31.Caption = "工作剩余时间：" & CStr(Alltime - Worktime) & "分钟"
    End If
    
    
    If tili <= 0 Then
        Form3.baoshidux = 1
        Form3.yinshuix = 1
        Form3.tilix = 1
        Form3.moneyx = 1
        allNaili = 0
        allGuanchali = 0
        allShejiao = 0
        allZhili = 0
        ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "体力值为0，停止timer2"
        Timer2.Enabled = False
        Worktime = 0
        Alltime = 0
        If PButton6.Text = "普通速度" Then
            PButton6.Text = "加快速度"
            '    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "减速" & CStr(Timer3.Interval)
        End If
        PButton6.Visible = False
        Timer2.Enabled = False
        Timer2.Interval = 1000
        Timer1.Enabled = True
        Timer1.Interval = 1000
        Timer3.Enabled = True
        Timer3.Interval = 1000
        Unload Form1.PProgressBar4(1)
        Exit Sub
    End If
    
    If Worktime = Alltime Then
        Form3.baoshidux = 1
        Form3.yinshuix = 1
        Form3.tilix = 1
        If Form1.moshi = 2 Then money = money + Form3.moneyx
        Form3.moneyx = 1
        Form1.Shejiao = allShejiao + Form1.Shejiao
        Form1.Zhili = allZhili + Form1.Zhili
        Form1.Guanchali = allGuanchali + Form1.Guanchali
        Form1.Naili = allNaili + Form1.Naili
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加社交能力" & allShejiao
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加智力" & allZhili
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加观察力" & allGuanchali
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "增加耐力" & allNaili
        allNaili = 0
        allGuanchali = 0
        allShejiao = 0
        allZhili = 0
        Worktime = 0
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "清除alltime"
        Alltime = 0
        Timer2.Enabled = False
        If PButton6.Text = "普通速度" Then
            PButton6.Text = "加快速度"
            '     'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "减速" & CStr(Timer3.Interval)
        End If
        PButton6.Visible = False
        Timer2.Enabled = False
        Timer2.Interval = 1000
        Timer1.Enabled = True
        Timer1.Interval = 1000
        Timer3.Enabled = True
        Timer3.Interval = 1000
        Unload Form1.PProgressBar4(1)
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "工作完成，停止timer2"
        Exit Sub
    End If
    
    
End Sub

Private Sub Timer3_Timer()
    
    If Form1.moshi = 0 Then                                                     '模式检测 0为经典无尽，1为限时模式，2为评测模式
        mina = mina + 1
        If mina = 60 Then
            houra = houra + 1
            mina = 0
        End If
        If houra = 24 Then
            houra = 0
            Form3.daya = Form3.daya + 1
        End If
        Dim b As String
        b = "0" & houra
        b = Right(b, 2)
        Dim c As String
        c = "0" & mina
        c = Right(c, 2)
        PScreen1.Text = b & ":" & c
        Label32.Caption = Form3.daya & "天" & houra & "时" & mina & "分" & seca & "秒"
        If PButton5.Text = "起床" And Form1.tili >= 1000 Then
            Call PButton5_Click
        End If
        
        If PButton5.Text = "起床" Then
            Form1.tili = Form1.tili + 3
            PProgressBar3.Value = CSng(tili / 1000)
            Label23.Caption = Format(CSng(tili / 1000), "0%")
        End If
        
    ElseIf Form1.moshi = 1 Then
        seca = seca + 1
        If seca = 60 Then
            seca = 0
            mina = mina + 1
        End If
        If mina = 60 Then
            mina = 0
            houra = houra + 1
        End If
        If houra = 24 Then
            houra = 0
            Form3.daya = Form3.daya + 1
        End If
        Label32.Caption = Form3.daya & "天" & houra & "时" & mina & "分"
    ElseIf Form1.moshi = 2 Then
        seca = seca - 1
        If seca <= 0 Then
            seca = 60
            mina = mina - 1
        End If
        If mina <= 0 And seca <= 0 Then
            Form12.Show
            Form31.Show
            Timer1.Enabled = False
            Timer2.Enabled = False
            Timer3.Enabled = False
            Timer4.Enabled = False
            Timer5.Enabled = False
            Timer6.Enabled = False
            Timer10.Enabled = False
            Call Form31.ResizeInit(Me)                                          '在程序装入时必须加入
            
        End If
        Dim d As String
        d = "0" & mina
        d = Right(d, 2)
        Dim E As String
        E = "0" & seca
        E = Right(E, 2)
        PScreen1.Text = d & ":" & E
        Label32.Caption = Form3.daya & "天" & houra & "时" & mina & "分" & seca & "秒"
    ElseIf Form1.moshi = 3 Then
        mina = mina + 1
        If mina = 60 Then
            houra = houra + 1
            mina = 0
        End If
        If houra = 24 Then
            houra = 0
            Form3.daya = Form3.daya + 1
        End If
        Dim f As String
        f = "0" & houra
        f = Right(f, 2)
        Dim G As String
        G = "0" & mina
        G = Right(G, 2)
        PScreen1.Text = b & ":" & c
        Label32.Caption = Form3.daya & "天" & houra & "时" & mina & "分"
    End If
    
    
End Sub

Private Sub Timer4_Timer()
    Dim thWnd As Long
    thWnd = GetForegroundWindow
    If FindWindow(vbNullString, "Form2") <> 0 And thWnd <> Form1.hwnd Then
        formback = 2
        Debug.Print "2"
    ElseIf FindWindow(vbNullString, "Form5") <> 0 And thWnd <> Form1.hwnd Then
        formback = 5
        Debug.Print "5"
    ElseIf FindWindow(vbNullString, "Form8") <> 0 And thWnd <> Form1.hwnd Then
        formback = 8
        Debug.Print "8"
    ElseIf FindWindow(vbNullString, "Form6") <> 0 And thWnd = Form1.hwnd Then
        Debug.Print "6"
    ElseIf FindWindow(vbNullString, "Form10") <> 0 And thWnd <> Form1.hwnd Then
        formback = 10
        Debug.Print "10"
    ElseIf FindWindow(vbNullString, "Form11") <> 0 And thWnd <> Form1.hwnd Then
        formback = 11
        Debug.Print "11"
    ElseIf FindWindow(vbNullString, "Form23") <> 0 And thWnd <> Form1.hwnd Then
        formback = 23
        Debug.Print "23"
    ElseIf FindWindow(vbNullString, "Form24") <> 0 And thWnd <> Form1.hwnd Then
        formback = 24
        Debug.Print "24"
    ElseIf FindWindow(vbNullString, "Form25") <> 0 And thWnd <> Form1.hwnd Then
        formback = 25
        Debug.Print "25"
    ElseIf FindWindow(vbNullString, "Form26") <> 0 And thWnd <> Form1.hwnd Then
        formback = 26
        Debug.Print "26"
    ElseIf FindWindow(vbNullString, "Form27") <> 0 And thWnd <> Form1.hwnd Then
        formback = 27
        Debug.Print "27"
    ElseIf FindWindow(vbNullString, "Form28") <> 0 And thWnd <> Form1.hwnd Then
        formback = 28
        Debug.Print "28"
    ElseIf FindWindow(vbNullString, "Form31") <> 0 And thWnd <> Form1.hwnd Then
        formback = 31
        Debug.Print "31"
        ' ElseIf FindWindow(vbNullString, "Form32") <> 0 And thWnd <> Form1.hwnd Then
        '     formback = 32
        '     Debug.Print "32"
    ElseIf FindWindow(vbNullString, "Form34") <> 0 And thWnd <> Form1.hwnd Then
        formback = 34
        Debug.Print "34"
    ElseIf FindWindow(vbNullString, "Form37") <> 0 And thWnd <> Form1.hwnd Then
        formback = 37
        Debug.Print "37"
    ElseIf thWnd = Form1.hwnd Then
        formback = 1
        Debug.Print "1"
    Else
        Timer3.Enabled = False
        Timer1.Enabled = False
        If Timer2.Enabled = True Then
            Timer2.Enabled = False
            workok = 1
            Form6.Show
            formback = 1
        End If
    End If
    ' Form32.Label9.Caption = "formback=" & formback
End Sub

Private Sub Timer5_Timer()
    Randomize
    If Int(Rnd * (100 - 1 + 1)) + 1 = 51 And money > 1000 And Accident = 0 Then '购置家具突发事件
        Form12.Show
        Form11.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Accident = 0 And Int(Rnd * (40 - 1 + 1)) + 1 = 5 And Naili <= 4 And Zhili <= 4 And Shejiao <= 4 And Guanchali <= 4 Then
        Accident = 1
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 2
        Form12.Show
        Form22.Show
    ElseIf Position = 1 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 3
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 5
        Form12.Show
        Form22.Show
    ElseIf Position = 1 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 6
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 7
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 8
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 11
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 12
        Form12.Show
        Form22.Show
    ElseIf Position <> 0 And Dj = 1 And Timer2.Enabled = True And Timer10.Enabled = False And Int(Rnd * (50 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 14
        
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 15
        Form12.Show
        Form22.Show
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 16
        Form12.Show
        Form22.Show
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 17
        Form12.Show
        Form22.Show
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 18
        Form12.Show
        Form22.Show
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 19
        Form12.Show
        Form22.Show
    ElseIf Position = 2 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 20
        Form12.Show
        Form22.Show
    ElseIf Position = 3 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 21
        Form12.Show
        Form22.Show
    ElseIf Position = 3 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 22
        Form12.Show
        Form22.Show
    ElseIf Position = 3 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 23
        Form12.Show
        Form22.Show
    ElseIf Position = 3 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 24
        Form12.Show
        Form22.Show
    ElseIf Position = 4 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 25
        Form12.Show
        Form22.Show
    ElseIf Position = 4 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 26
        Form12.Show
        Form22.Show
    ElseIf Position = 4 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 27
        Form12.Show
        Form22.Show
    ElseIf Position = 4 And Timer2.Enabled = True And Int(Rnd * (100 - 1 + 1)) + 1 = 12 And Accident = 0 Then
        Accident = 28
        Form12.Show
        Form22.Show
    End If
    '  Form32.Label10.Caption = "突发事件ID=" & Accident
End Sub

Private Sub Timer6_Timer()
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    If WindowsMediaPlayer1.playState = 1 Then                                   '1为停止(一曲播完)
        Randomize
        If Int(Rnd * (10 - 1 + 1)) + 1 = 3 Then
            WindowsMediaPlayer1.URL = App.Path & "\music\background\" & CStr(Int(Rnd * (12 - 1 + 1)) + 1) & ".mp3"
            WindowsMediaPlayer1.Controls.play
        End If
    End If
End Sub

Private Sub timer7_Timer()
    If ib + 5 <= 255 Then
        ib = ib + 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib = 255 Then Timer4.Enabled = False
    End If
End Sub
Private Sub Timer8_Timer()
    ib = ib - 5
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                        '150为透明度(0-255)
    If ib = 0 Then
        Timer4.Enabled = False
        Unload Me
    End If
End Sub


