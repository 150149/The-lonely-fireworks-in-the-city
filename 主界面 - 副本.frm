VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "������"
   ClientHeight    =   16200
   ClientLeft      =   6285
   ClientTop       =   3855
   ClientWidth     =   28800
   LinkTopic       =   "Form2"
   Picture         =   "������.frx":0000
   ScaleHeight     =   16200
   ScaleWidth      =   28800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
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
   Begin �о��й°����̻�.PButton PButton8 
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
      Text            =   "����"
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
   Begin �о��й°����̻�.PButton PButton3 
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
      Text            =   "��ͣ"
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
   Begin �о��й°����̻�.PProgressBar PProgressBar3 
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
   Begin �о��й°����̻�.PProgressBar PProgressBar2 
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Back      =   0
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar1 
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
   Begin �о��й°����̻�.PProgressBar PProgressBar4 
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
   Begin �о��й°����̻�.PButton PButton1 
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
      Text            =   "��"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton2 
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
      Text            =   "����"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton5 
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
      Text            =   "˯��"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PButton PButton6 
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
      Text            =   "����"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PScreen PScreen1 
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
      Picture         =   "������.frx":13C04F
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1440
   End
   Begin VB.Image Image12 
      Height          =   2295
      Left            =   12840
      Picture         =   "������.frx":13D422
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "֤ȯ������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "0��6ʱ0��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����ʣ��ʱ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����ֵ"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��ˮ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "��ʳ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "ս��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Picture         =   "������.frx":13EB04
      Stretch         =   -1  'True
      Top             =   9840
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   2415
      Left            =   480
      Picture         =   "������.frx":147910
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
      Caption         =   "xxx����"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
         Name            =   "΢���ź�"
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
      Picture         =   "������.frx":14B815
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Left            =   240
      Picture         =   "������.frx":15EF5A
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   240
      Picture         =   "������.frx":15FA02
      Stretch         =   -1  'True
      Top             =   840
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Picture         =   "������.frx":16052B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "״̬"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Picture         =   "������.frx":160E4E
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
      Caption         =   "���ز���˾"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "������˾"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Picture         =   "������.frx":1618F6
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
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�
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
Public baoshidu As Integer                                                      '���履ʳ��
Public yinshui As Integer                                                       '������ˮ��
Public tili As Integer                                                          '��������ֵ
Public tili2 As Integer                                                         '���������Ƿ�ָ�
Public money As Integer                                                         '�����Ǯ
Public mina As Integer                                                          '�������
Public seca As Integer                                                          '������
Public houra As Integer                                                         '����Сʱ
Public Position As Integer                                                      '����λ��ID
Public workok As Integer                                                        '���幤��״̬
Public Shejiao As Integer                                                       '�����罻����
Public Zhili As Integer                                                         '������������
Public Guanchali As Integer                                                     '����۲�������
Public Naili As Integer                                                         '������������
Public difficult As Integer                                                     '�����Ѷ�ϵ��
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public formback As Integer                                                      '���巵�ش���ID
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
'Const LWA_COLORKEY = &H1
Private ib As Integer                                                           '����͸����
Private Addai As Integer                                                        '����AI���ӵ�ID
Public work As Integer                                                          '���幤��ID
Public Worktime As Integer                                                      '�����ѹ���ʱ��
Public Accident As Integer                                                      '����ͻ���¼�ID
Private beibaomax As String                                                     '���屳��������
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public allNaili As Integer                                                      '���幤�����ӵ�����ֵ
Public allZhili As Integer                                                      '���幤�����ӵ�����ֵ
Public allShejiao As Integer                                                    '���幤�����ӵ��罻ֵ
Public allGuanchali As Integer                                                  '���幤�����ӵĹ۲���ֵ
Public Alltime As Integer                                                       '���幤����ʱ��
Public Daodepingzhi As Integer                                                  '�������Ʒ��ֵ
Public jingye As Integer                                                        '���徴ҵֵ
Public haoxue As Integer                                                        '�����ѧֵ
Public dandang As Integer                                                       '���嵣��ֵ
Public chengxin As Integer                                                      '�������ֵ
Public Lastposition As Integer                                                  '���幤����ʼʱ���ڵ�λ��ID
Public n As Integer                                                             '����Gal��ʾ�ַ�����
Public xianshi As Integer                                                       '����Gal��ȫ��ʾ����
Dim zd As String                                                                '����Gal��ǰ��ʾ������
Dim dq As String                                                                '����Gal��ȡ������
Public Dj As Integer                                                            '����Gal�����������ʾ���ֶ�ID��
Private Speedup As Boolean                                                      '�����Ƿ����
Public Zhiwei As String                                                         '����ְλ
Public moshi As Integer                                                         '������Ϸģʽ
Public Pengren As Boolean                                                       '�����Ƿ�������
Public Driven As Boolean                                                        '�����Ƿ���Լ�ʻ
Public Jinrong As Boolean                                                       '�����Ƿ���Խ����������
Private timermode(10) As Integer                                                '����ʱ�ӿؼ�ģʽ
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

'==========================�����̰����ɿ����������»ָ�����==========================================
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    tili2 = 1                                                                   '�޸��Ƿ�ָ�����Ϊtrue
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
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������,ʹ����ؼ��Զ�����
    Form1.Width = Screen.Width                                                  '�ı䴰���СΪ��Ļ��С
    Form1.Height = Screen.Height                                                '�ı䴰���СΪ��Ļ��С
    
    KeyPreview = True                                                           'ȷ��������Ӧ���̰����¼�
    Timer1.Enabled = True                                                       '��������ָ��ģ��
    Timer1.Interval = 1000                                                      'ȷ������ָ��ģ��ˢ��Ƶ��
    Timer3.Enabled = True                                                       '��������ʱ��ģ��
    Timer3.Interval = 1000                                                      'ȷ������ʱ��ģ��ˢ��Ƶ��
    Timer4.Enabled = False                                                      '�������役���ж�ģ��
    Timer4.Interval = 1000                                                      'ȷ�����役���ж�ģ��ˢ��Ƶ��
    Timer2.Enabled = False                                                      '�رչ���ģ��
    If Form1.moshi = 2 Then
        Timer5.Enabled = True                                                   '����ͻ���¼�ģ��
        Timer5.Interval = 1000                                                  'ȷ��ͻ���¼�ˢ��Ƶ��
    End If
    Timer10.Enabled = False                                                     '�ر�AIģ��
    Form6.Show                                                                  '������ʧȥ����ʱ���л���������,ͨ����ͣ���淵����Ϸ
    Form1.Show                                                                  '���������ڶ���
    'Form32.Show                                                                 '��ʾdebug����
    ' 'unload form32
    Timer6.Enabled = True                                                       '�����������ֲ����Զ��л�����
    Timer6.Interval = 1000                                                      'ȷ����������Ƿ����Ƶ��
    'Timer10.Enabled = True'AIģ��
    'Timer10.Interval = 1000'AIģ��
    'Label20(0).Visible = False'AI ģ��
    Randomize                                                                   '��ʼ��VB�����������
    WindowsMediaPlayer1.URL = App.Path & "\music\background\" & CStr(Int(Rnd * (12 - 1 + 1)) + 1) & ".mp3" '������Ϸ���ű�������
    ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "��ǰ���ŵ�����Ϊ" & WindowsMediaPlayer1.URL 'Debng������ֲ���·��
    PProgressBar4(0).Visible = False                                            '���ع���������
    If Dir(App.Path & "\text.ini") = "" Then                                    '�ж�text.ini��Gal���֣��Ƿ����
        Open "text.ini" For Output As #11
        Print #11, ""
        Close #11
        Dim a As String
        Dim b As Integer
        Dim c As Long
        For b = 1 To 500
            c = WritePrivateProfileString("text", CStr(b), "", App.Path & "\text.ini")
        Next
    End If                                                                      '�����ڼ�д���ļ�
    Dim read_OK As Long
    dq = String(256, 0)
    read_OK = GetPrivateProfileString("text", CStr(Dj), "", dq, 256, App.Path & "\text.ini") '��ȡGal�ϴ��Ķ�λ��
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
'=====================================�Զ����ſؼ�======================================================
'=======================================================================================================
'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim i     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '���洰�������ű���
    scaleY = FormName.ScaleHeight / FormOldHeight                               '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For i = 0 To 4
            '��ȡ�ؼ���ԭʼλ�����С
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(i) = 0
            End If
            '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
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
    Call ResizeForm(Me)                                                         'ȷ������ı�ʱ�ؼ���֮�ı�
End Sub

'�ڵ���ResizeFormǰ�ȵ��ñ�����
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
    Accident = 0                                                                '���˳�����ʱ���ͻ���¼�ID���Ա��˳�����Ϸ��ʼ��
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
    
                                                       '�����治��Ӧ�����¼�
    Form12.Show                                                                 '�򿪰�͸������
    Form5.form5open = Form5.form5open + 1
    Form5.Show                                                                  '��ʾ��������
    Form5.Timer1.Enabled = True                                                 '��������������Ʒ������ʾ�ؼ�
    Form5.Timer1.Interval = 300                                                 'ȷ������ˢ��Ƶ��
    If Form5.form5open = 2 Then                                                 'Ԥ���ر�����������
        Form5.PProgressBar1.Value = CInt(Form1.Shejiao) / 1000                  'Ԥ����
        Form5.PProgressBar2.Value = CInt(Form1.Zhili) / 1000                    'Ԥ����
        Form5.PProgressBar3.Value = CInt(Form1.Guanchali) / 1000                'Ԥ����
        Form5.PProgressBar4.Value = CInt(Form1.Naili) / 1000                    'Ԥ����
        Form5.Label3.Caption = "�罻������" & CInt(Form1.Shejiao)               'Ԥ����
        Form5.Label4.Caption = "������" & CInt(Form1.Zhili)                     'Ԥ����
        Form5.Label5.Caption = "�۲�����" & CInt(Form1.Guanchali)               'Ԥ����
        Form5.Label6.Caption = "������" & CInt(Form1.Naili)                     'Ԥ����
        Dim read_OK As Long
        beibaomax = String(10, 0)
        read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", beibaomax, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
        Dim beibaomaxa As Integer                                               '��ȡ����������
        Dim beibaomaxb As Integer                                               '��ȡ����������
        beibaomaxb = CInt(beibaomax)                                            '��ȡ����������
        Dim picnumber As String                                                 '��ȡ����������
        picnumber = String(10, 0)                                               '��ȡ����������
        For beibaomaxa = 1 To beibaomaxb                                        '��ȡ����������
            read_OK = GetPrivateProfileString("beibao", "beibao" & beibaomaxa, "0", picnumber, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            If picnumber <> 0 Then                                              '���¼��ر�����ƷͼƬ
                Form5.Image1(beibaomaxa).Picture = LoadPicture("")              '���¼��ر�����ƷͼƬ
                Form5.Image1(beibaomaxa).Picture = LoadPicture(App.Path & "\picture\item\" & Replace(picnumber, Chr(0), "") & ".gif")
             '   'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "������Ʒͼ��" & CStr(beibaomaxa) & " " & "�滻��ƷΪ" & Replace(picnumber, Chr(0), "")
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

'-----------------------------------����ģ��----------------------------------------------
Private Sub PButton1_Click()
   ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "�򿪹���ѡ�����"         'debug���
    Form12.Show                                                                 '��ʾ��͸������
    Form8.Show                                                                  '��ʾ����ѡ�����
End Sub
'-----------------------------------����ģ��----------------------------------------------




Private Sub PButton2_Click()
    Form49.Show
End Sub

Private Sub PButton3_Click()                                                    '��ͣ��
    Timer3.Enabled = False                                                      '��ͣ��������ģ��                                              '
    Timer1.Enabled = False                                                      '��ͣ����ָ��ģ��
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "timer2״̬Ϊ" & CStr(Timer2.Enabled) 'debug���
    If Timer2.Enabled = True Then                                               '������ģ�����ڽ��У���ͣ����ģ��
        Timer2.Enabled = False
        workok = 1                                                              '���湤��ģ�鹤��״̬
    End If
    Form6.Show                                                                  '��ʾ��ͣ����
End Sub



Private Sub PButton5_Click()
    If PButton5.Text = "��" Then                                              '�ж��Ƿ���˯��״̬
        Timer3.Enabled = True                                                   '�ָ�����ʱ��ģ��
        Timer3.Interval = 1000                                                  '�ָ�ʱ��ʱ��ģ������
        PButton5.Text = "˯��"                                                  '����״̬
    Else
        PButton5.Text = "��"                                                  '�ж��Ƿ���˯��״̬
        Timer3.Enabled = True                                                   '��������ʱ��ģ��
        Timer3.Interval = 100                                                   'ȷ�����ٱ���
    End If
End Sub



Private Sub PButton6_Click()
    If PButton6.Text = "��ͨ�ٶ�" Then                                          '�ж��Ƿ��ڹ�������״̬
        Timer2.Enabled = True                                                   '�ָ������ٶ�
        Timer2.Interval = 1000                                                  '�ָ������ٶ�
        Timer1.Enabled = True                                                   '�ָ�����ָ��ģ������
        Timer1.Interval = 1000                                                  'ȷ������ָ��ģ����������
        Timer3.Enabled = True                                                   '�ָ�����ʱ���ٶ�
        Timer3.Interval = 1000                                                  'ȷ������ʱ���ٶ�
        PButton6.Text = "�ӿ��ٶ�"                                              '���Ĺ�������״̬
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����" & CStr(Timer3.Interval) 'debug���
    Else                                                                        '�ж��Ƿ��ڹ�������״̬
        Timer2.Enabled = True                                                   '���ٹ���ģ��
        Timer2.Interval = 100                                                   '���ٹ���ģ��
        Timer1.Enabled = True                                                   '��������ָ��ģ��
        Timer1.Interval = 100                                                   '��������ָ��ģ��
        Timer3.Enabled = True                                                   '��������ʱ��ģ��
        Timer3.Interval = 100                                                   '��������ʱ��ģ��
        PButton6.Text = "��ͨ�ٶ�"                                              '���Ĺ�������״̬
      '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "���" & CStr(Timer3.Interval) 'debug���
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

'----------------------------------------��������ָ��ģ��--------------------------------------
Private Sub Timer1_Timer()
    
    
    If Form1.moshi = 0 Then
        baoshidu = baoshidu - 1 - Form3.baoshidux                               'ÿ�����3/1000�ı�ʳ��
        yinshui = yinshui - 2 - Form3.yinshuix                                  'ÿ�����6/1000����ˮ��
    End If
    If Form1.moshi = 2 Then
        baoshidu = baoshidu - 3 - Form3.baoshidux                               'ÿ�����3/1000�ı�ʳ��
        yinshui = yinshui - 6 - Form3.yinshuix                                  'ÿ�����6/1000����ˮ��
    End If
    If tili2 = 1 And tili < 1000 And Not Timer2.Enabled = True Then             '�ж��Ƿ��ڹ���״̬
        tili = tili + 1                                                         '�ǹ���״̬ÿ��ָ�1/1000������ֵ
        PProgressBar3.Value = CSng(tili / 1000)
        Label23.Caption = Format(CSng(tili / 1000), "0%")
    Else
        tili = tili - Form3.tilix                                               '���ڹ���״̬��ÿ����ٹ�����������
        PProgressBar3.Value = CSng(tili / 1000)
        Label23.Caption = Format(CSng(tili / 1000), "0%")
    End If
    
    PProgressBar1.Value = CSng(baoshidu / 1000)                                 '��ʾ��ʳ����
    PProgressBar2.Value = CSng(yinshui / 1000)                                  '��ʾ��ˮ����
    PProgressBar3.Value = CSng(tili / 1000)                                     '��ʾ������
    Label11.Caption = money                                                     '��ʾ��Ǯ
    Label21.Caption = Format(CSng(baoshidu / 1000), "0%")                       '��ʾ��ʳ�Ȱٷֱ�
    Label22.Caption = Format(CSng(yinshui / 1000), "0%")                        '��ʾ��ˮ�Ȱٷֱ�
    
    
    If baoshidu <= 0 Or yinshui <= 0 Then                                       '�жϱ�ʳ�Ȼ�����ˮ���Ƿ�Ϊ0
        Form12.Show                                                             '���0����Ϸ����
        Form31.Show
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Timer6.Enabled = False
        Timer10.Enabled = False
        Call Form31.ResizeInit(Me)                                              '�ڳ���װ��ʱ�������
        Form31.PProgressBar1.Value = CInt(Form1.Shejiao) / 1000                 'Ԥ������Ϸ��������
        Form31.PProgressBar2.Value = CInt(Form1.Zhili) / 1000
        Form31.PProgressBar3.Value = CInt(Form1.Guanchali) / 1000               'Ԥ����
        Form31.PProgressBar4.Value = CInt(Form1.Naili) / 1000                   'Ԥ����
        Form31.Label3.Caption = "�罻������" & CStr(CInt(Form1.Shejiao))        'Ԥ����
        Form31.Label4.Caption = "������" & CStr(CInt(Form1.Zhili))              'Ԥ����
        Form31.Label5.Caption = "�۲�����" & CStr(CInt(Form1.Guanchali))        'Ԥ����
        Form31.Label6.Caption = "������" & CStr(CInt(Form1.Naili))              'Ԥ����
        Form31.Label2.Caption = "��Ǯ:" & Form1.money                           'Ԥ����
    End If                                                                      'Ԥ����
    ' Form32.Label1.Caption = "��ʳ�ȣ�" & baoshidu                               'Ԥ����
    ' Form32.Label2.Caption = "��ˮ�ȣ�" & yinshui                                'Ԥ����
    ' Form32.Label3.Caption = "����ֵ��" & tili                                   'Ԥ����
    ' Form32.Label4.Caption = "��Ǯ��" & money                                    'Ԥ����
    ' Form32.Label5.Caption = "��ʳ������" & Form3.baoshidux                      'Ԥ����
    ' Form32.Label6.Caption = "��ˮ������" & Form3.yinshuix                       'Ԥ����
    ' Form32.Label7.Caption = "����������" & Form3.tilix                          'Ԥ����
    ' Form32.Label8.Caption = "��Ǯ������" & Form3.moneyx                         'Ԥ����
    
    If Form1.tili <= 200 Then                                                   '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar3.Color_Top = &HFF&                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
        Label2.Caption = "�ҿ�Ҫ������"                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
    Else                                                                        '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar3.Color_Top = &H80FFFF                                      '�ж�����ָ���Ƿ����20%�����������ʾ�����
    End If                                                                      '�ж�����ָ���Ƿ����20%�����������ʾ�����
    If Form1.yinshui <= 200 Then                                                '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar2.Color_Top = &HFF&                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
        Label2.Caption = "�ҿ�Ҫ������"                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
    Else                                                                        '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar2.Color_Top = &HFF7402                                      '�ж�����ָ���Ƿ����20%�����������ʾ�����
    End If                                                                      '�ж�����ָ���Ƿ����20%�����������ʾ�����
    If Form1.baoshidu <= 200 Then                                               '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar1.Color_Top = &HFF&                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
        Label2.Caption = "�ҿ�Ҫ������"                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
    Else                                                                        '�ж�����ָ���Ƿ����20%�����������ʾ�����
        PProgressBar1.Color_Top = &H80FF&                                       '�ж�����ָ���Ƿ����20%�����������ʾ�����
    End If                                                                      '�ж�����ָ���Ƿ����20%�����������ʾ�����
End Sub                                                                         '�ж�����ָ���Ƿ����20%�����������ʾ�����
'----------------------------------------��������ָ��ģ��--------------------------------------


'----------------------------------����ģ��------------------------------------
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If PButton5.Text = "��" Then
        Call PButton5_Click
        Exit Sub
    End If
    
    If Label1.Top < 100 Then
        Label2.Caption = "�㲻������ǰ����"
        Label1.Top = Label1.Top + 300
    ElseIf Label1.Left < 100 Then
        Label2.Caption = "�㲻������ǰ����"
        Label1.Left = Label1.Left + 300
    ElseIf Form1.Height - Label1.Top < 100 Then
        Label2.Caption = "�㲻������ǰ����"
        Label1.Top = Label1.Top - 300
    ElseIf Form1.Width - Label1.Left < 100 Then
        Label2.Caption = "�㲻������ǰ����"
        Label1.Left = Label1.Left - 300
    ElseIf tili <= 0 Then
        Label2.Caption = "��û������"
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
        Label2.Caption = "�������ڷ�����"
        Label25.Caption = "����"
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
        Label2.Caption = "�������ڽ�����˾��"
        Label25.Caption = "������˾"
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
        Label2.Caption = "�������ڷ��ز���˾��"
        Label25.Caption = "���ز���˾"
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
        Label2.Caption = "�������ڱ�������"
        Label25.Caption = "������"
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
        Label2.Caption = "��������סլ¥1��"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 5
    ElseIf Label1.Top >= Label8.Top And Label1.Top <= Label8.Top + Label8.Height And Label1.Left >= Label8.Left And Label1.Left <= Label8.Left + Label8.Width Then
        Label2.Caption = "��������סլ¥2��"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 6
    ElseIf Label1.Top >= Label9.Top And Label1.Top <= Label9.Top + Label9.Height And Label1.Left >= Label9.Left And Label1.Left <= Label9.Left + Label9.Width Then
        Label2.Caption = "��������סլ¥3��"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Position = 7
    ElseIf Label1.Top >= Label10.Top And Label1.Top <= Label10.Top + Label10.Height And Label1.Left >= Label10.Left And Label1.Left <= Label10.Left + Label10.Width Then
        Label2.Caption = "�������ڼ���"
        Label25.Caption = "��"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = True
        PButton1.Visible = False
        PButton2.Visible = False
        'pbutton8.Visible = False
        Position = 8
    ElseIf Label1.Top >= Label12.Top And Label1.Top <= Label12.Top + Label12.Height And Label1.Left >= Label12.Left And Label1.Left <= Label12.Left + Label12.Width Then
        Label2.Caption = "��������֤ȯ��������"
        Label25.Caption = "֤ȯ������"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        'pbutton8.Visible = True
        PButton1.Visible = True
        If Form1.moshi = 0 Then PButton8.Visible = True
        Position = 9
    ElseIf Label1.Top >= Label13.Top And Label1.Top <= Label13.Top + Label13.Height And Label1.Left >= Label13.Left And Label1.Left <= Label13.Left + Label13.Width Then
        Label2.Caption = "�������ڽ���������"
        Label25.Caption = "��������"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = False
        Position = 10
    ElseIf Label1.Top >= Label14.Top And Label1.Top <= Label14.Top + Label14.Height And Label1.Left >= Label14.Left And Label1.Left <= Label14.Left + Label14.Width Then
        Label2.Caption = "�������ڿ��˹�˾��"
        Label25.Caption = "���˹�˾"
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
        Label2.Caption = "�������ڹ�԰��"
        Label25.Caption = "��԰"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = False
        PButton2.Visible = False
        Position = 12
    ElseIf Label1.Top >= Label17.Top And Label1.Top <= Label17.Top + Label17.Height And Label1.Left >= Label17.Left And Label1.Left <= Label17.Left + Label17.Width Then
        Label2.Caption = "�������ڳ�����"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        Label25.Caption = "����"
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = True
        Position = 13
    ElseIf Label1.Top >= Label18.Top And Label1.Top <= Label18.Top + Label18.Height And Label1.Left >= Label18.Left And Label1.Left <= Label18.Left + Label18.Width Then
        Label2.Caption = "�������������˾��"
        Label25.Caption = "�����˾"
        If Shape16.Visible = False Then Shape16.Visible = True
        If Label25.Visible = False Then Label25.Visible = True
        If Label20.Visible = False Then Label20.Visible = True
        PButton5.Visible = False
        PButton1.Visible = True
        PButton2.Visible = False
        Position = 14
    Else
        Label2.Caption = "�������ڽֵ���"
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
'--------------------------------------����ģ��--------------------------------------------------------



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
        ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����ʱ��Ϊ0��ֹͣtimer2"
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
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "�뿪�����ص㣬ֹͣtimer2"
        Timer2.Enabled = False
        Worktime = 0
        Alltime = 0
        If PButton6.Text = "��ͨ�ٶ�" Then
            PButton6.Text = "�ӿ��ٶ�"
            '    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����" & CStr(Timer3.Interval)
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
        Label31.Caption = "����ʣ��ʱ�䣺" & CStr(Alltime - Worktime) & "����"
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
        ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����ֵΪ0��ֹͣtimer2"
        Timer2.Enabled = False
        Worktime = 0
        Alltime = 0
        If PButton6.Text = "��ͨ�ٶ�" Then
            PButton6.Text = "�ӿ��ٶ�"
            '    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����" & CStr(Timer3.Interval)
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
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "�����罻����" & allShejiao
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "��������" & allZhili
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "���ӹ۲���" & allGuanchali
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "��������" & allNaili
        allNaili = 0
        allGuanchali = 0
        allShejiao = 0
        allZhili = 0
        Worktime = 0
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "���alltime"
        Alltime = 0
        Timer2.Enabled = False
        If PButton6.Text = "��ͨ�ٶ�" Then
            PButton6.Text = "�ӿ��ٶ�"
            '     'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "����" & CStr(Timer3.Interval)
        End If
        PButton6.Visible = False
        Timer2.Enabled = False
        Timer2.Interval = 1000
        Timer1.Enabled = True
        Timer1.Interval = 1000
        Timer3.Enabled = True
        Timer3.Interval = 1000
        Unload Form1.PProgressBar4(1)
        '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "������ɣ�ֹͣtimer2"
        Exit Sub
    End If
    
    
End Sub

Private Sub Timer3_Timer()
    
    If Form1.moshi = 0 Then                                                     'ģʽ��� 0Ϊ�����޾���1Ϊ��ʱģʽ��2Ϊ����ģʽ
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
        Label32.Caption = Form3.daya & "��" & houra & "ʱ" & mina & "��" & seca & "��"
        If PButton5.Text = "��" And Form1.tili >= 1000 Then
            Call PButton5_Click
        End If
        
        If PButton5.Text = "��" Then
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
        Label32.Caption = Form3.daya & "��" & houra & "ʱ" & mina & "��"
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
            Call Form31.ResizeInit(Me)                                          '�ڳ���װ��ʱ�������
            
        End If
        Dim d As String
        d = "0" & mina
        d = Right(d, 2)
        Dim E As String
        E = "0" & seca
        E = Right(E, 2)
        PScreen1.Text = d & ":" & E
        Label32.Caption = Form3.daya & "��" & houra & "ʱ" & mina & "��" & seca & "��"
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
        Label32.Caption = Form3.daya & "��" & houra & "ʱ" & mina & "��"
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
    If Int(Rnd * (100 - 1 + 1)) + 1 = 51 And money > 1000 And Accident = 0 Then '���üҾ�ͻ���¼�
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
    '  Form32.Label10.Caption = "ͻ���¼�ID=" & Accident
End Sub

Private Sub Timer6_Timer()
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    If WindowsMediaPlayer1.playState = 1 Then                                   '1Ϊֹͣ(һ������)
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
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150Ϊ͸����(0-255)
        If ib = 255 Then Timer4.Enabled = False
    End If
End Sub
Private Sub Timer8_Timer()
    ib = ib - 5
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                        '150Ϊ͸����(0-255)
    If ib = 0 Then
        Timer4.Enabled = False
        Unload Me
    End If
End Sub


