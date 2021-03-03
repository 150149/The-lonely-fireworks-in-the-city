VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form30 
   BorderStyle     =   0  'None
   Caption         =   "Form12"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16080
   Icon            =   "多人游戏大厅.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   8895
   ScaleWidth      =   16080
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer5 
      Left            =   0
      Top             =   2280
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   1800
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   1215
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
      _extentx        =   4683
      _extenty        =   2143
      color_back      =   33023
      color_back_down =   16576
      color_begin     =   33023
      color_end       =   16576
      text            =   "进入你的网络存档"
      style_border    =   0
      can_text_move   =   0   'False
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   375
      ExtentX         =   661
      ExtentY         =   661
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "刷新"
      Height          =   375
      Left            =   13320
      TabIndex        =   3
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "发送"
      Height          =   375
      Left            =   9960
      TabIndex        =   2
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   9960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   7080
      Width           =   4695
   End
   Begin VB.Label Label8 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   9360
      TabIndex        =   13
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label Label20 
      BackColor       =   &H00379CEE&
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "存档体力值"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "存档饮水度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   9
      Top             =   6600
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "存档饱食度"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   6000
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "存档金钱"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Image Image3 
      Height          =   4455
      Left            =   720
      Picture         =   "多人游戏大厅.frx":08CA
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   3615
      Left            =   720
      Picture         =   "多人游戏大厅.frx":10E7
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   " "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   3450
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6660
      Left            =   9960
      TabIndex        =   0
      Top             =   480
      Width           =   4650
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   9000
      Picture         =   "多人游戏大厅.frx":1904
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public mingzi As String
Public suoyou As String
Public lishi As String
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Public denglu As String
Public doudong As String
Public jiazaicishu As Integer
Public dd2 As String
Public szMyText, szMyText2 As String
Public lishi2 As String
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long
Private Const WS_VERSION_REQD = &H101
'-------------------------------------------------------------------------------------------------------------------------
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const MAX_PATH As Integer = 260
Const TH32CS_SNAPPROCESS As Long = 2&
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Jilu As String
Private cv As String
Private page As Integer
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function IsConnectedState() As Boolean
    Dim udtWSAD As WSADATA
    Call WSAStartup(WS_VERSION_REQD, udtWSAD)
    IsConnectedState = CBool(gethostbyname("www.baidu.com"))
    Call WSACleanup
End Function

Private Sub Command1_Click()
    If IsConnectedState Then                                                    '检查网络连接
        WebBrowser1.Navigate "http://showtxt.cn/kbrt5"
        '-------------------------------------------发送中心-------------------------------------------------
        Dim mAES As New mAES                                                    ''AES加密类对象
        'Randomize
        'cv = Int(Rnd * (10000 - 1 + 1)) + 1
        'Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Output As #1
        'Print #1, mAES.DecryptStr(Jilu, "AE|set", Bit128, Bit128, False)
        'Print #1, mingzi & ": " & Text1.Text
        ' Close #1
        Dim fasongneirong As String
        'Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Binary As #1
        'fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
        'Close #1
        'Kill App.Path & "\save\save" & Form3.saveid & ".fsave"
        'cv = ""
        fasongneirong = mAES.DecryptStr(Jilu, "AE|set", Bit128, Bit128, False) & Chr(10) & Chr(13) & mingzi & ": " & Text1.Text
        Debug.Print "解算历史记录:" & fasongneirong
        fasongneirong = mAES.EncryptStr(fasongneirong, "AE|set", Bit128, Bit128, False)
        Debug.Print "合成发送内容:" & fasongneirong
        Dim vDoc, VTag, mType As String, mTagName As String
        Dim ia As Integer
        Set vDoc = WebBrowser1.Document
        For ia = 0 To vDoc.All.Length - 1
            Select Case UCase(vDoc.All(ia).tagName)
            Case "TEXTAREA"                                                     '"TEXTAREA" 标签,文本框的填写
                Set VTag = vDoc.All(ia)
                VTag.Value = "开始" & fasongneirong & "结束"                    '将Text1中的内容填入
                Text1.Text = ""
                Timer1.Enabled = True
                Timer1.Interval = 300
                jiazaicishu = 0
                Command1.Enabled = False
                Command3.Enabled = False
                WebBrowser1.Navigate "http://showtxt.cn/kbrt5"
                Debug.Print "发送内容成功"
                Exit Sub
            End Select
            
        Next ia
    Else
        Label4.Caption = "网络未连接"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Command3_Click()
    If IsConnectedState Then
        WebBrowser1.Navigate "http://showtxt.cn/kbrt5"
        Timer1.Enabled = True
        Timer1.Interval = 1000
        Command1.Enabled = False
        Command3.Enabled = False
    Else
        Label4.Caption = "网络未连接"
    End If
End Sub

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    
    '-------------------------------------------------检查网络连接------------------------------------
    If IsConnectedState Then
        WebBrowser1.Navigate "http://showtxt.cn/kbrt5"
        Label4.Caption = "正在获取内容"
        denglu = "no"
        Command1.Enabled = False
        Command3.Enabled = False
        Timer1.Enabled = True
        Timer1.Interval = 1000
    Else
        Label4.Caption = "网络未连接"
        Command1.Enabled = False
        Command3.Enabled = False
    End If
    WebBrowser1.Silent = True
    WebBrowser1.Visible = True
    Form35.Show
    Form30.Show
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
    Dim read_OK As Long
    Dim rbaoshidu  As String
    Dim ryinshui As String
    Dim rtili As String
    Dim rmoney As String
    rbaoshidu = String(10, 0)
    ryinshui = String(10, 0)
    rtili = String(10, 0)
    rmoney = String(10, 0)
    read_OK = GetPrivateProfileString("human", "baoshidu", "100", rbaoshidu, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    read_OK = GetPrivateProfileString("human", "yinshui", "100", ryinshui, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    read_OK = GetPrivateProfileString("human", "tili", "1000", rtili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    read_OK = GetPrivateProfileString("human", "money", "100", rmoney, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
    Label5.Caption = "存档饱食度：" & Replace(rbaoshidu, Chr(0), "")
    Label6.Caption = "存档饮水度：" & Replace(ryinshui, Chr(0), "")
    Label7.Caption = "存档体力值：" & Replace(rtili, Chr(0), "")
    Label2.Caption = "存档金钱：" & Replace(rmoney, Chr(0), "")
End Sub

Private Sub PButton1_Click()
    If IsConnectedState Then
        WebBrowser1.Navigate "http://showtxt.cn/kbrt6"
        page = 1
        jiazaicishu = 0
        Timer4.Enabled = True
        Timer4.Interval = 1000
    Else
        MsgBox "网络未连接"
    End If
End Sub

Private Sub Timer1_Timer()
    jiazaicishu = jiazaicishu + 1
    Debug.Print ("加载次数：" & jiazaicishu)
    '--------------------------------------------接收中心------------------------------------------
    If WebBrowser1.busy Then
        Debug.Print ("网页未加载完成")
        Exit Sub
    Else
        Debug.Print ("网页加载完成")
        Debug.Print "-------------------------------------------------------------------------------"
        Timer1.Enabled = False
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        Dim vDoc, X, VTag
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '检测所有标签
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '找到input标签
                Set VTag = vDoc.All(X)
                If VTag.Value = "提交" Then VTag.Click                          '点击提交了，一切都OK了
            End If
        Next X
        denglu = "yes"
        Timer3.Enabled = True
        Timer3.Interval = 1500
        Debug.Print ("登录状态：" & "yes")
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then                                                        '如果，是回车键按下
        If Command1.Enabled = True Then
            Call Command1_Click
        Else
        End If
    End If
End Sub



Private Sub Timer3_Timer()
    If WebBrowser1.busy Then
        Debug.Print ("网页未加载完成")
        Exit Sub
    Else
        Timer3.Enabled = False
        Debug.Print ("开始提取文字")
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "开始"                                                 '定义要查找的字符串开头
        szFindStrEnd = "结束"                                                   '定义要查找的字符串结尾
        
        szText = WebBrowser1.Document.body.innerText                            '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '找开头字符串
        If nBegin > 0 Then                                                      '必须有能找到开头了才继续
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '找结尾字符串
            If nEnd > nBegin Then                                               '结尾必须比开头的位置大
                
                '包含查找的字符串模式，注释掉下面的2行
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
                
                '不包含查找的字符串模式
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '如果不包含查找的字符串，用这2行
                nBegin = nBegin + Len(szFindStrBegin)                           '如果不包含查找的字符串，用这2行
                
                szMyText = Mid(szText, nBegin, nLength)                         '取出“before then.”到 "test" 中间的东西
                Debug.Print ("截取内容：" & szMyText)
            End If
        End If
        
        Label4.Caption = ""
        Jilu = szMyText
        Debug.Print "写入历史记录:" & szMyText
        Dim mAES As New mAES
        Open App.Path & "\history" For Output As #3
        Print #3, mAES.DecryptStr(szMyText, "AE|set", Bit128, Bit128, False)
        Close #3
        Dim showtext As String
        Open App.Path & "\history" For Binary As #1
        showtext = StrConv(InputB(LOF(1), 1), vbUnicode)
        Debug.Print "showtext=" & showtext
        Label3.Caption = Replace(showtext, Chr(0), "")
        Debug.Print "解算历史记录:" & StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        Kill App.Path & "\history"
        Text1.SetFocus
        Debug.Print "label3.caption=" & Label3.Caption
    End If
    Command1.Enabled = True
    Command3.Enabled = True
End Sub

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim I     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '保存窗体宽度缩放比例
    scaleY = FormName.ScaleHeight / FormOldHeight                               '保存窗体高度缩放比例
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For I = 0 To 4
            '读取控件的原始位置与大小
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(I) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(I) = 0
            End If
            '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next I
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me)                                                         '确保窗体改变时控件随之改变
    Me.WebBrowser1.Left = 0
    Me.WebBrowser1.Top = 0
    Me.WebBrowser1.Width = 0
    Me.WebBrowser1.Height = 0
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


Private Sub Timer4_Timer()
    jiazaicishu = jiazaicishu + 1
    Debug.Print ("加载次数：" & jiazaicishu)
    '--------------------------------------------接收中心------------------------------------------
    If WebBrowser1.busy Then
        Debug.Print ("网页未加载完成")
        Exit Sub
    Else
        Debug.Print ("网页加载完成")
        Debug.Print "-------------------------------------------------------------------------------"
        Timer4.Enabled = False
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        Dim vDoc, X, VTag
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '检测所有标签
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '找到input标签
                Set VTag = vDoc.All(X)
                If VTag.Value = "提交" Then VTag.Click                          '点击提交了，一切都OK了
            End If
        Next X
        denglu = "yes"
        Timer5.Enabled = True
        Timer5.Interval = 1500
        Debug.Print ("登录状态：" & "yes")
    End If
End Sub

Private Sub Timer5_Timer()
    If WebBrowser1.busy Then
        Debug.Print ("网页未加载完成")
        Exit Sub
    Else
        Timer5.Enabled = False
        Debug.Print ("开始提取文字")
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "开始"                                                 '定义要查找的字符串开头
        szFindStrEnd = "结束"                                                   '定义要查找的字符串结尾
        
        szText = WebBrowser1.Document.body.innerText                            '得到所有文字，临时用模板，实际使用切换回去WebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '找开头字符串
        If nBegin > 0 Then                                                      '必须有能找到开头了才继续
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '找结尾字符串
            If nEnd > nBegin Then                                               '结尾必须比开头的位置大
                
                '包含查找的字符串模式，注释掉下面的2行
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '计算需要提取的字符串长度,如果要包含查找的字符串用这1行，注释下面2行
                
                '不包含查找的字符串模式
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '如果不包含查找的字符串，用这2行
                nBegin = nBegin + Len(szFindStrBegin)                           '如果不包含查找的字符串，用这2行
                
                szMyText = Mid(szText, nBegin, nLength)                         '取出“before then.”到 "test" 中间的东西
                Debug.Print ("截取内容：" & szMyText)
                Randomize
                Dim cv As String
                cv = Int(Rnd * (10000 - 1 + 1)) + 1
                Form3.saveid = 8
                Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Output As #1
                Print #1, szMyText
                Close #1
                Dim read_OK As Long
                Dim mAES As New mAES                                            ''AES加密类对象
                Dim Saveindex As String
                Saveindex = String(2056, 0)
                read_OK = GetPrivateProfileString("save", mAES.EncryptStr(mingzi, "150149", Bit128, Bit128, False), "", Saveindex, 2056, App.Path & "\save\save" & Form3.saveid & ".fsave")
                Kill App.Path & "\save\save" & Form3.saveid & ".fsave"
                Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Output As #1
                Print #1, mAES.DecryptStr(Saveindex, "150149", Bit128, Bit128, False)
                Close #1
                
                If Replace(Saveindex, Chr(0), "") = "" Then
                    Debug.Print "本账户联网存档为空"
                    Dim write1 As Long
                    write1 = WritePrivateProfileString("human", "baoshidu", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "yinshui", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "tili", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "money", "100", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("extrahuman", "baoshidux", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("extrahuman", "yinshuix", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("extrahuman", "tilix", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("extrahuman", "moneyx", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("time", "daya", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("time", "houra", "6", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("time", "mina", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("time", "seca", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("beibao", "beibaomax", "20", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    Dim beibaomaxa As Integer
                    For beibaomaxa = 1 To 80
                        write1 = WritePrivateProfileString("beibao", "beibao" & beibaomaxa, "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    Next
                    write1 = WritePrivateProfileString("human", "Shejiao", "6", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "Zhili", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "Guanchali", "2", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "Naili", "2", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("save", "name", Form13.Text1.Text, App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("mark", "daodepingzhi", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("mark", "jingye", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("mark", "haoxue", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("mark", "dandang", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("mark", "chengxin", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("text", "dj", "1", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "zhiwei", "无业打工者", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "pengren", "False", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "driven", "False", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("human", "jinrong", "False", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "2", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "3", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "4", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "5", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "6", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "7", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "8", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "9", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "10", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "11", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "12", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "13", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "14", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "15", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "16", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock1", "17", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "2", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "3", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "4", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "5", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "6", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock2", "7", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock3", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock3", "2", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock3", "3", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock3", "4", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock3", "5", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock4", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock4", "2", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock4", "3", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock4", "4", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock9", "1", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock9", "2", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock9", "3", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock9", "4", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                    write1 = WritePrivateProfileString("lock9", "5", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
                End If
                
                
                Dim rbaoshidu  As String
                Dim ryinshui As String
                Dim rtili As String
                Dim rmoney As String
                Dim rbaoshidux As String
                Dim ryinshuix As String
                Dim rtilix As String
                Dim rmoneyx As String
                Dim rdaya As String
                Dim rhoura As String
                Dim rmina As String
                Dim rseca As String
                Dim rbeibaomax As String
                Dim rx As String
                Dim ry As String
                Dim rShejiao As String
                Dim rZhili As String
                Dim rGuanchali As String
                Dim rNaili As String
                Dim rmoshi As String
                Dim rdaodepingzhi As String
                Dim rjingye As String
                Dim rhaoxue As String
                Dim rdandang As String
                Dim rchengxin As String
                Dim rdj As String
                Dim rzhiwei As String
                Dim rpengren As String
                Dim rdriven As String
                Dim rjinrong As String
                rpengren = String(10, 0)
                rdriven = String(10, 0)
                rjinrong = String(10, 0)
                rzhiwei = String(10, 0)
                rdj = String(10, 0)
                rdandang = String(10, 0)
                rchengxin = String(10, 0)
                rdaodepingzhi = String(10, 0)
                rjingye = String(10, 0)
                rhaoxue = String(10, 0)
                rShejiao = String(10, 0)
                rZhili = String(10, 0)
                rGuanchali = String(10, 0)
                rNaili = String(10, 0)
                rx = String(10, 0)
                ry = String(10, 0)
                rbaoshidu = String(10, 0)
                ryinshui = String(10, 0)
                rtili = String(10, 0)
                rmoney = String(10, 0)
                rdaya = String(10, 0)
                rhoura = String(10, 0)
                rmina = String(10, 0)
                rseca = String(10, 0)
                rbeibaomax = String(10, 0)
                rbaoshidux = String(10, 0)
                ryinshuix = String(10, 0)
                rtilix = String(10, 0)
                rmoneyx = String(10, 0)
                rmoshi = String(10, 0)
                read_OK = GetPrivateProfileString("human", "baoshidu", "100", rbaoshidu, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "yinshui", "100", ryinshui, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "tili", "1000", rtili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "money", "100", rmoney, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("extrahuman", "baoshidux", "0", rbaoshidux, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("extrahuman", "yinshuix", "0", ryinshuix, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("extrahuman", "tilix", "0", rtilix, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("extrahuman", "moneyx", "0", rmoneyx, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("time", "daya", "1", rdaya, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("time", "houra", "6", rhoura, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("time", "mina", "0", rmina, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("time", "seca", "0", rseca, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("beibao", "beibaomax", "5", rbeibaomax, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "top", "4440", rx, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "left", "10680", ry, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "Shejiao", "0", rShejiao, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "Zhili", "0", rZhili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "Guanchali", "0", rGuanchali, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "Naili", "0", rNaili, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("save", "mode", "0", rmoshi, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("mark", "daodepingzhi", "0", rdaodepingzhi, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("mark", "jingye", "0", rjingye, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("mark", "haoxue", "0", rhaoxue, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("mark", "dandang", "0", rdandang, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("mark", "chengxin", "0", rchengxin, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("text", "dj", "1", rdj, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "zhiwei", "无业打工者", rzhiwei, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "pengren", "False", rpengren, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "driven", "False", rdriven, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                read_OK = GetPrivateProfileString("human", "jinrong", "False", rjinrong, 256, App.Path & "\save\save" & Form3.saveid & ".fsave")
                
                Form1.Dj = CInt(rdj)
                Form1.Zhiwei = Replace(rzhiwei, Chr(0), "")
                Form1.Shejiao = CInt(rShejiao)
                Form1.Zhili = CInt(rZhili)
                Form1.Guanchali = CInt(rGuanchali)
                Form1.Naili = CInt(rNaili)
                Form1.baoshidu = CInt(rbaoshidu)
                Form1.yinshui = CInt(ryinshui)
                Form1.tili = CInt(rtili)
                Form1.money = CInt(rmoney)
                Form3.baoshidux = 0
                Form3.yinshuix = 0
                Form3.tilix = 0
                Form3.moneyx = 0
                Form3.daya = CInt(rdaya)
                Form1.houra = CInt(rhoura)
                Form1.mina = CInt(rmina)
                Form1.seca = CInt(rseca)
                Form20.rx = CInt(rx)
                Form20.ry = CInt(ry)
                If Replace(rpengren, Chr(0), "") = "False" Then Form1.Pengren = False
                If Replace(rpengren, Chr(0), "") = "True" Then Form1.Pengren = True
                If Replace(rdriven, Chr(0), "") = "False" Then Form1.Driven = False
                If Replace(rdriven, Chr(0), "") = "True" Then Form1.Driven = True
                If Replace(rjinrong, Chr(0), "") = "False" Then Form1.Jinrong = False
                If Replace(rjinrong, Chr(0), "") = "True" Then Form1.Jinrong = True
                Form1.Daodepingzhi = CInt(rdaodepingzhi)
                Form1.jingye = CInt(rjingye)
                Form1.haoxue = CInt(rhaoxue)
                Form1.dandang = CInt(rdandang)
                Form1.chengxin = CInt(rchengxin)
                Form1.moshi = 3
                
                
                Form20.Show
            End If
        End If
    End If
    
End Sub
