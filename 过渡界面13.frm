VERSION 5.00
Begin VB.Form Form54 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form54"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Enabled         =   0   'False
   Icon            =   "过渡界面13.frx":0000
   LinkTopic       =   "Form53"
   ScaleHeight     =   10380
   ScaleMode       =   0  'User
   ScaleWidth      =   18065.35
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timer5 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   480
   End
End
Attribute VB_Name = "Form54"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ib As Integer
Private ic As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public rx As Integer
Public ry As Integer

Private Sub Form_Load()
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Timer4.Interval = 10
    Timer5.Interval = 10
    Timer4.Enabled = True
    Timer5.Enabled = False
    Timer2.Enabled = True
    Timer2.Interval = 1
    ib = 0
End Sub



Private Sub Timer4_Timer()
    If ib + 5 <= 255 Then
        ib = ib + 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib = 255 Then
            Timer4.Enabled = False
            Timer5.Enabled = True
            Unload Form52
            Form18.Show
            Form18.Top = Form3.Label1.Top
            Form18.Left = Form3.Label1.Left
            Form18.Height = Form3.Label1.Height
            Form18.Width = Form3.Label1.Width
        End If
    End If
End Sub

Private Sub Timer5_Timer()
    If ib - 5 >= 0 Then
        ib = ib - 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib = 0 Then
            Timer4.Enabled = False
            Timer5.Enabled = False
            Unload Form53
        End If
    ElseIf ib - 5 < 0 Then
        Unload Form53
    Else
        Unload Form53
    End If
End Sub
