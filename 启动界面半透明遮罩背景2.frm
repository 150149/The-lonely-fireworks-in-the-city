VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   0  'None
   Caption         =   "Form12"
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   Icon            =   "启动界面半透明遮罩背景2.frx":0000
   LinkTopic       =   "Form12"
   ScaleHeight     =   765
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   4210752
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "创建新存档"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Color_Border    =   8421504
      Can_Text_Move   =   0   'False
   End
   Begin 市井中孤傲的烟火.PButton PButton5 
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1296
      Color_Back      =   14737632
      Color_Back_Down =   4210752
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "返回"
      Font_Size       =   17
      Font_Bold       =   -1  'True
      Color_Border    =   8421504
      Can_Text_Move   =   0   'False
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth     As Long                                                '保存窗体的原始宽度
Private FormOldHeight     As Long                                               '保存窗体的原始高度
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '在程序装入时必须加入
    
    Form15.BackColor = &H80000
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, 50, LWA_ALPHA                           '   窗体透明
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    
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
    If Dir(App.Path & "\save\", vbDirectory) = "" Then                          '判断文件夹是否存在
        MkDir (App.Path & "\save\")                                             '创建文件夹
    End If
    If Dir(App.Path & "\save\save" & Form3.saveid & ".fsave") = "" Then
        Open App.Path & "\save\save" & Form3.saveid & ".fsave" For Output As #1
        Print #1, ""
        Close #1
        Dim write1 As Long
        If Form13.PSwitch1.Value = True Then write1 = WritePrivateProfileString("save", "mode", "2", App.Path & "\save\save" & Form3.saveid & ".fsave")
        If Form13.PSwitch1.Value = False Then write1 = WritePrivateProfileString("save", "mode", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "baoshidu", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "yinshui", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "tili", "1000", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("human", "money", "100", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "baoshidux", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "yinshuix", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "tilix", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
        write1 = WritePrivateProfileString("extrahuman", "moneyx", "0", App.Path & "\save\save" & Form3.saveid & ".fsave")
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
    
    Dim read_OK As Long
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
    Form3.baoshidux = CInt(rbaoshidux)
    Form3.yinshuix = CInt(ryinshuix)
    Form3.tilix = CInt(rtilix)
    Form3.moneyx = CInt(rmoneyx)
    Form3.daya = CInt(rdaya)
    Form1.houra = CInt(rhoura)
    Form1.mina = CInt(rmina)
    Form1.seca = CInt(rseca)
    Form20.rx = CInt(rx)
    Form20.ry = CInt(ry)
    Form1.Daodepingzhi = CInt(rdaodepingzhi)
    Form1.jingye = CInt(rjingye)
    Form1.haoxue = CInt(rhaoxue)
    Form1.dandang = CInt(rdandang)
    Form1.chengxin = CInt(rchengxin)
    If Replace(rpengren, Chr(0), "") = "False" Then Form1.Pengren = False
    If Replace(rpengren, Chr(0), "") = "True" Then Form1.Pengren = True
    If Replace(rdriven, Chr(0), "") = "False" Then Form1.Driven = False
    If Replace(rdriven, Chr(0), "") = "True" Then Form1.Driven = True
    If Replace(rjinrong, Chr(0), "") = "False" Then Form1.Jinrong = False
    If Replace(rjinrong, Chr(0), "") = "True" Then Form1.Jinrong = True
    '  'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "模式读取为" & rmoshi
    Form1.moshi = CInt(rmoshi)
    'Form32.Text1.Text = 'Form32.Text1.Text & vbCrLf & "模式为" & Form1.moshi
    If Form1.moshi = 2 Then Form1.mina = 20
    Form51.Show
End Sub

Private Sub PButton5_Click()
    Form7.Show
    Unload Form14
    Unload Form13
    Unload Form15
End Sub
