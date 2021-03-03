VERSION 5.00
Begin VB.Form Form49 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form49"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Enabled         =   0   'False
   Icon            =   "过渡界面9.frx":0000
   LinkTopic       =   "Form49"
   LockControls    =   -1  'True
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
Attribute VB_Name = "Form49"
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
            If Form1.Position = 4 Then                                          '当位置ID为4（便利店）
                ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "当前位置为" & Form1.Position 'debug输出
                ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "打开" & Form1.Position 'debug输出
                KeyPreview = False                                              '主窗体不响应键盘事件
                Form12.Show
                Form27.Show
                Exit Sub
            ElseIf Form1.Position = 13 Then                                     '当位置ID为13超市
                ''Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "当前位置为" & Form1.Position 'debug输出
                ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "打开" & Form1.Position 'debug输出
                KeyPreview = False                                              '主窗体不响应键盘事件
                Form12.Show
                Form24.Show
                Exit Sub
            ElseIf Form1.Position = 1 Then                                      '当位置ID为1（饭店）
                ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "当前位置为" & Form1.Position 'debug输出
                ' 'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "打开" & Form1.Position 'debug输出
                KeyPreview = False                                              '主窗体不响应键盘事件
                Form12.Show
                Form23.Show
                Dim first1 As String
                Dim read_OK As Long
                Dim write1 As Long
                first1 = String(10, 0)
                read_OK = GetPrivateProfileString("guide", "1", "", first1, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                If Replace(first1, Chr(0), "") = "7" And Form1.moshi = 0 Then
                    Form37.Show
                    Form37.Width = Form37.PPictureBox10.Width
                    Form37.Left = 0
                    Form37.Top = 0
                    Form37.PPictureBox1.Visible = False
                    Form37.PPictureBox2.Visible = False
                    Form37.PPictureBox3.Visible = False
                    Form37.PPictureBox4.Visible = False
                    Form37.PPictureBox5.Visible = False
                    Form37.PPictureBox6.Visible = False
                    Form37.PPictureBox7.Visible = False
                    Form37.PPictureBox8.Visible = False
                    Form37.PPictureBox9.Visible = False
                    Form37.PPictureBox10.Visible = True
                    Form37.PPictureBox11.Visible = False
                    Form37.PPictureBox12.Visible = False
                    Form37.PPictureBox13.Visible = False
                    Form37.PPictureBox14.Visible = False
                    Form37.PPictureBox15.Visible = False
                    Form37.PPictureBox16.Visible = False
                    Form37.PPictureBox17.Visible = False
                    Form37.PPictureBox18.Visible = False
                    Form37.PPictureBox19.Visible = False
                ElseIf Replace(first1, Chr(0), "") = "5" And Form1.moshi = 2 Then
                    Form37.Show
                    Form37.Width = Form37.PPictureBox10.Width
                    Form37.Left = 0
                    Form37.Top = 0
                    Form37.PPictureBox1.Visible = False
                    Form37.PPictureBox2.Visible = False
                    Form37.PPictureBox3.Visible = False
                    Form37.PPictureBox4.Visible = False
                    Form37.PPictureBox5.Visible = False
                    Form37.PPictureBox6.Visible = False
                    Form37.PPictureBox7.Visible = False
                    Form37.PPictureBox8.Visible = False
                    Form37.PPictureBox9.Visible = False
                    Form37.PPictureBox10.Visible = True
                    Form37.PPictureBox11.Visible = False
                    Form37.PPictureBox12.Visible = False
                    Form37.PPictureBox13.Visible = False
                    Form37.PPictureBox14.Visible = False
                    Form37.PPictureBox15.Visible = False
                    Form37.PPictureBox16.Visible = False
                    Form37.PPictureBox17.Visible = False
                    Form37.PPictureBox18.Visible = False
                    Form37.PPictureBox19.Visible = False
                End If
                Exit Sub
            Else
                Unload Form12                                                   '当位置ID为0，卸载半透明遮罩
                Exit Sub
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
            Unload Form49
        End If
    ElseIf ib - 5 < 0 Then
        Unload Form49
    Else
        Unload Form49
    End If
End Sub
