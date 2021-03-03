VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   "暂停"
   ClientHeight    =   8715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13380
   Icon            =   "暂停界面.frx":0000
   LinkTopic       =   "Form6"
   ScaleHeight     =   8715
   ScaleWidth      =   13380
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   4095
      ExtentX         =   7223
      ExtentY         =   6588
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
   Begin 市井中孤傲的烟火.PButton PButton3 
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   4200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "存档并退出"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   735
      Left            =   5400
      TabIndex        =   0
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      Color_Back      =   33023
      Color_Back_Down =   16576
      Color_Begin     =   33023
      Color_End       =   16576
      Text            =   "返回"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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
Private cv As String
Public szMyText, szMyText2 As String
Private timermode As Integer

Public Function IsConnectedState() As Boolean
    Dim udtWSAD As WSADATA
    Call WSAStartup(WS_VERSION_REQD, udtWSAD)
    IsConnectedState = CBool(gethostbyname("www.baidu.com"))
    Call WSACleanup
End Function

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    Me.Height = Form1.Height
    Me.Width = Form1.Width
    
    Me.BackColor = &H80000
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA                          '   窗体透明
    Form1.Show
    WebBrowser1.Silent = True
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
    Call ResizeForm(Me)
    Me.WebBrowser1.Left = 0
    Me.WebBrowser1.Top = 0
    Me.WebBrowser1.Width = 0
    Me.WebBrowser1.Height = 0
    '确保窗体改变时控件随之改变
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

Private Sub PButton1_Click()
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
    If Form1.Timer2.Enabled = False And Form1.workok = 1 Then
        Form1.Timer2.Enabled = True
        Form1.workok = 2
    End If
    Form1.WindowsMediaPlayer1.Controls.play
End Sub

Private Sub PButton3_Click()
    Dim rbaoshidu As String
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
    Dim write1 As Long
    Dim rShejiao As String
    Dim rZhili As String
    Dim rGuanchali As String
    Dim rNaili As String
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
    rzhiwei = Form1.Zhiwei
    rdj = CStr(Form1.Dj)
    rbaoshidu = CStr(Form1.baoshidu)
    ryinshui = CStr(Form1.yinshui)
    rtili = CStr(Form1.tili)
    rmoney = CStr(Form1.money)
    rbaoshidux = CStr(Form3.baoshidux)
    ryinshuix = CStr(Form3.yinshuix)
    rtilix = CStr(Form3.tilix)
    rmoneyx = CStr(Form3.moneyx)
    rdaya = CStr(Form3.daya)
    rhoura = CStr(Form1.houra)
    rmina = CStr(Form1.mina)
    rseca = CStr(Form1.seca)
    rShejiao = CStr(Form1.Shejiao)
    rZhili = CStr(Form1.Zhili)
    rGuanchali = CStr(Form1.Guanchali)
    rNaili = CStr(Form1.Naili)
    rdaodepingzhi = CStr(Form1.Daodepingzhi)
    rjingye = CStr(Form1.jingye)
    rhaoxue = CStr(Form1.haoxue)
    rdandang = CStr(Form1.dandang)
    rchengxin = CStr(Form1.chengxin)
    If Form1.moshi = 3 And IsConnectedState Then
        WebBrowser1.Navigate "http://showtxt.cn/kbrt6"
        write1 = WritePrivateProfileString("human", "baoshidu", rbaoshidu, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "yinshui", ryinshui, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "tili", rtili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "money", rmoney, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "baoshidux", rbaoshidux, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "yinshuix", ryinshuix, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "tilix", rtilix, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "moneyx", rmoneyx, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "daya", rdaya, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "houra", rhoura, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "mina", rmina, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("time", "seca", rseca, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "top", CStr(Form1.Label1.Top), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "left", CStr(Form1.Label1.Left), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("save", "changetime1", CStr(Date), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("save", "changetime2", CStr(Time), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Shejiao", rShejiao, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Zhili", rZhili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Guanchali", rGuanchali, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "Naili", rNaili, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "daodepingzhi", rdaodepingzhi, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "jingye", rjingye, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "haoxue", rhaoxue, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "dandang", rdandang, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("mark", "chengxin", rchengxin, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("text", "dj", rdj, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "zhiwei", rzhiwei, App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "pengren", CStr(Form1.Pengren), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "driven", CStr(Form1.Driven), App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "jinrong", CStr(Form1.Jinrong), App.Path & "\save\save" & Form3.saveid & ".fsave")
        Timer1.Enabled = True
        Timer1.Interval = 300
        Exit Sub
    End If
    
    write1 = WritePrivateProfileString("human", "baoshidu", rbaoshidu, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "yinshui", ryinshui, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "tili", rtili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "money", rmoney, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "baoshidux", rbaoshidux, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "yinshuix", ryinshuix, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "tilix", rtilix, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("extrahuman", "moneyx", rmoneyx, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "daya", rdaya, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "houra", rhoura, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "mina", rmina, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("time", "seca", rseca, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "top", CStr(Form1.Label1.Top), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "left", CStr(Form1.Label1.Left), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("save", "changetime1", CStr(Date), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("save", "changetime2", CStr(Time), App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    write1 = WritePrivateProfileString("human", "Shejiao", rShejiao, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Zhili", rZhili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Guanchali", rGuanchali, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "Naili", rNaili, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "daodepingzhi", rdaodepingzhi, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "jingye", rjingye, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "haoxue", rhaoxue, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "dandang", rdandang, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("mark", "chengxin", rchengxin, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("text", "dj", rdj, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "zhiwei", rzhiwei, App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "pengren", CStr(Form1.Pengren), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "driven", CStr(Form1.Driven), App.Path & "\save\save" & Form3.saveid & ".fsave")
    write1 = WritePrivateProfileString("human", "jinrong", CStr(Form1.Jinrong), App.Path & "\save\save" & Form3.saveid & ".fsave")
    
    Form3.baoshidux = 0
    Form3.yinshuix = 0
    Form3.tilix = 0
    Form3.moneyx = 0
    Form1.formback = 0
    Form42.Show
End Sub

Private Sub Timer1_Timer()
    
    '--------------------------------------------接收中心------------------------------------------
    If WebBrowser1.busy Then
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
        Timer3.Enabled = True
        Timer3.Interval = 1500
        Debug.Print ("登录状态：" & "yes")
    End If
    If timermode = 1 Then
        Form4.Show
        Form18.Show
        Form18.Top = Form3.Label1.Top
        Form18.Left = Form3.Label1.Left
        Form18.Height = Form3.Label1.Height
        Form18.Width = Form3.Label1.Width
        Form42.Show
        Exit Sub
    End If
End Sub

Private Sub Timer3_Timer()
    If WebBrowser1.busy Then
        Debug.Print ("网页未加载完成timer3")
        Exit Sub
    Else
        Timer3.Enabled = False
        Debug.Print ("开始提取文字timer3")
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
        Randomize
        cv = Int(Rnd * (10000 - 1 + 1)) + 1
        Open App.Path & "\" & cv For Output As #1
        Print #1, szMyText
        Close #1
        Dim write1 As Long
        Dim mAES As New mAES                                                    ''AES加密类对象
        Dim fasongneirong As String
        Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Binary As #1
        fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        Kill App.Path & "\save\save8.fsave"
        fasongneirong = mAES.EncryptStr(fasongneirong, "150149", Bit128, Bit128, False)
        write1 = WritePrivateProfileString("save", mAES.EncryptStr(Form30.mingzi, "150149", Bit128, Bit128, False), fasongneirong, App.Path & "\" & cv)
        Open App.Path & "\" & cv For Binary As #1
        fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
        Close #1
        Kill App.Path & "\" & cv
        Dim vDoc, VTag, mType As String, mTagName As String
        Dim ia As Integer
        Set vDoc = WebBrowser1.Document
        For ia = 0 To vDoc.All.Length - 1
            Select Case UCase(vDoc.All(ia).tagName)
            Case "TEXTAREA"                                                     '"TEXTAREA" 标签,文本框的填写
                Set VTag = vDoc.All(ia)
                VTag.Value = "开始" & fasongneirong & "结束"                    '将Text1中的内容填入
                Timer1.Enabled = True
                Timer1.Interval = 300
                WebBrowser1.Navigate "http://showtxt.cn/kbrt6"
                Debug.Print "发送内容成功"
                timermode = 1
            End Select
            
        Next ia
        
        
    End If
End Sub
