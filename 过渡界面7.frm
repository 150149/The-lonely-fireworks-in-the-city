VERSION 5.00
Begin VB.Form Form46 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form20"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Enabled         =   0   'False
   Icon            =   "过渡界面7.frx":0000
   LinkTopic       =   "Form20"
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
Attribute VB_Name = "Form46"
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
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    Timer4.Interval = 10
    Timer5.Interval = 10
    Timer4.Enabled = True
    Timer5.Enabled = False
    Timer2.Enabled = True
    Timer2.Interval = 1
    ib = 0
End Sub



Private Sub Timer2_Timer()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Timer4_Timer()
    If ib + 5 <= 255 Then
        ib = ib + 5
        SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
        SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                    '150为透明度(0-255)
        If ib = 255 Then
            Timer4.Enabled = False
            Timer5.Enabled = True
            Form12.Show
            Form28.Show
            Dim first1 As String
            Dim read_OK As Long
            Dim write1 As Long
            first1 = String(10, 0)
            read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
            If Replace(first1, Chr(0), "") = "11" Then
                Form37.Show
                Form37.Width = Form37.PPictureBox14.Width
                Form37.Left = Screen.Width - Form37.Width
                Form37.Top = 500
                Form37.PPictureBox1.Visible = False
                Form37.PPictureBox2.Visible = False
                Form37.PPictureBox3.Visible = False
                Form37.PPictureBox4.Visible = False
                Form37.PPictureBox5.Visible = False
                Form37.PPictureBox6.Visible = False
                Form37.PPictureBox7.Visible = False
                Form37.PPictureBox8.Visible = False
                Form37.PPictureBox9.Visible = False
                Form37.PPictureBox10.Visible = False
                Form37.PPictureBox11.Visible = False
                Form37.PPictureBox12.Visible = False
                Form37.PPictureBox13.Visible = False
                Form37.PPictureBox14.Visible = True
                Form37.PPictureBox15.Visible = False
                Form37.PPictureBox16.Visible = False
                Form37.PPictureBox17.Visible = False
                Form37.PPictureBox18.Visible = False
                Form37.PPictureBox19.Visible = False
                write1 = WritePrivateProfileString("guide", "1", "12", App.Path & "\save\save" & Form3.saveid & ".fsave")
            End If
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
            Unload Form46
        End If
    ElseIf ib - 5 < 0 Then
        Unload Form46
    Else
        Unload Form46
    End If
End Sub
