VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "启动界面背景.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   9795
   ScaleWidth      =   16995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   1080
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer2 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
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
      _cy             =   873
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   10440
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16980
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
      _cx             =   29951
      _cy             =   18415
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Sub Form_GotFocus()
    Form3.Show
    Form3.SetFocus
End Sub

Private Sub Form_Load()
    Form4.Height = Screen.Height
    Form4.Width = Screen.Width
    WindowsMediaPlayer1.Height = Form3.Height
    WindowsMediaPlayer1.Width = Form3.Width
    WindowsMediaPlayer1.URL = App.Path & "\video\1.mp4"
    WindowsMediaPlayer1.Controls.play
    WindowsMediaPlayer1.uiMode = "none"
    WindowsMediaPlayer1.stretchToFit = True
    WindowsMediaPlayer1.Height = Form3.Height
    WindowsMediaPlayer1.Width = Form3.Width
    Timer1.Enabled = True
    Timer1.Interval = 2000
    WindowsMediaPlayer2.URL = App.Path & "\music\begin.mp3"
End Sub


Private Sub Timer1_Timer()
    If Form3.music >= 3 Then
    Else
        Form3.music = Form3.music + 1
    End If
    
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    If WindowsMediaPlayer1.playState = 1 Then                                   '1为停止(一曲播完)
        WindowsMediaPlayer1.URL = App.Path & "\video\1.mp4"
        WindowsMediaPlayer1.Controls.play
    End If
End Sub

