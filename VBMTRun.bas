Attribute VB_Name = "VBMTRun"
Option Explicit

'定义线程句柄
Public VBThreadHandle1 As Long, VBThreadHandle2 As Long
'定义线程ID
Public VBThreadID1 As Long, VBThreadID2 As Long

'************************************注意：VB6多线程必须以SUB MAIN为启动对象***************************************************
'***************************本示例中已经设置好了，自己使用时注意在工程——属性——启动对象中自行选择************************
Sub Main()
    If AvoidReentrant = False Then                                              '防止主线程重复运行
        AvoidReentrant = True
        If App.PrevInstance Then                                                '防止程序重复运行
            MessageBox ByVal 0&, StrPtr("程序正在运行或未完全退出"), StrPtr("重复运行"), vbCritical
            Exit Sub
        Else
            InitCommonControls                                                  '初始化通用控件
            GETVBHeader                                                         '获取VB数据头
            
            Form21.Show                                                         '在这里加载主窗体
        End If
    End If
End Sub

Public Sub Thread1()                                                            '子线程1
    '***********************************（重要！）VB6线程环境初始化*************************************************
    CreateIExprSrvObj 0&, 4&, 0&                                                'VB6运行库初始化
    CoInitializeEx ByVal 0&, ByVal (COINIT_MULTITHREADED Or COINIT_SPEED_OVER_MEMORY) 'COM组件初始化
    InitVBdll                                                                   '诱导VB6运行库内部其他部分的初始化
    '***********************************（重要！）VB6线程环境初始化*************************************************
    Dim StrTemp As String, n As Long
    Dim MyForm As Form19
    Set MyForm = New Form19                                                     '克隆FORM19的数据
    MyForm.Show
    '###########################################
    '#                                         #
    '#     是的你没看错，多线程我根本没用      #
    '#           因为这程序足够流畅            #
    '#                                         #
    '###########################################
    CoUninitialize                                                              '卸载COM组件（省掉也不会影响稳定性，但可能造成句柄或内存泄漏。为了养成好习惯，还是写上）
End Sub
