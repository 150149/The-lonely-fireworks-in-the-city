VERSION 5.00
Begin VB.Form Form33 
   BorderStyle     =   0  'None
   Caption         =   "Form33"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   Icon            =   "设置界面.frx":0000
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PHScrollBar PHScrollBar1 
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      Color_Top       =   33023
      Color_Back      =   8438015
      Style_Border    =   2
   End
   Begin VB.Label Label26 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "音量"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   1080
      Picture         =   "设置界面.frx":08CA
      Stretch         =   -1  'True
      Top             =   720
      Width           =   6000
   End
End
Attribute VB_Name = "Form33"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ' WindowsMediaPlayer1.settings.volume = 0~100
End Sub

