VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form35 
   BorderStyle     =   0  'None
   Caption         =   "Form35"
   ClientHeight    =   9840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16755
   Enabled         =   0   'False
   Icon            =   "多人游戏背景.frx":0000
   LinkTopic       =   "Form35"
   ScaleHeight     =   9840
   ScaleWidth      =   16755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Left            =   3600
      Top             =   6240
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13020
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
      _cx             =   22966
      _cy             =   9737
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long, ByVal dwMaximumWorkingSetSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
    Rem 转移输入焦点的声明
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Rem 禁止本窗体拥有输入焦点的常数
Private Const HWND_NOTOPMOST = -2
Private Const WS_DISABLED = &H8000000
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Sub Form_Click()
    Form30.Show
End Sub

Private Sub Form_DblClick()
    Form30.Show
End Sub

Private Sub Form_Load()
    Rem 窗体调用时置顶，且禁止拥有输入焦点
    SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_DISABLED
    Form35.Height = Screen.Height
    Form35.Width = Screen.Width
    WindowsMediaPlayer1.Height = Form35.Height
    WindowsMediaPlayer1.Width = Form35.Width
    WindowsMediaPlayer1.URL = App.Path & "\video\2.mov"
    WindowsMediaPlayer1.Controls.play
    WindowsMediaPlayer1.uiMode = "none"
    WindowsMediaPlayer1.stretchToFit = True
    WindowsMediaPlayer1.Height = Form3.Height
    WindowsMediaPlayer1.Width = Form3.Width
    Timer1.Enabled = True
    Timer1.Interval = 2000
End Sub
Private Sub Timer1_Timer()
    SetProcessWorkingSetSize GetCurrentProcess(), -1&, -1&
    If WindowsMediaPlayer1.playState = 1 Then                                   '1为停止(一曲播完)
        WindowsMediaPlayer1.URL = App.Path & "\video\2.mp4"
        WindowsMediaPlayer1.Controls.play
    End If
End Sub

Private Sub WindowsMediaPlayer1_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    Form30.Show
End Sub


Private Sub WindowsMediaPlayer1_DoubleClick(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    Form30.Show
End Sub
