VERSION 5.00
Begin VB.UserControl PButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F2AF00&
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   MousePointer    =   1  'Arrow
   ScaleHeight     =   615
   ScaleWidth      =   1005
   ToolboxBitmap   =   "PButton.ctx":0000
   Begin 市井中孤傲的烟火.PUIMgr PM 
      Left            =   480
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox TPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Index           =   0
      Left            =   240
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      Begin VB.PictureBox TPic 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   1
         Left            =   0
         ScaleHeight     =   135
         ScaleWidth      =   135
         TabIndex        =   2
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "PButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'██████↓定义存储属性的变量↓███████████████████████████████████████████████████████
Dim C_Color_Back As OLE_COLOR                                                   '背景颜色
Dim C_Color_Back_Down As OLE_COLOR                                              '鼠标按下时的背景颜色
Dim C_Color_Begin As OLE_COLOR                                                  '渐变开始的颜色
Dim C_Color_End As OLE_COLOR                                                    '渐变结束的颜色
Dim C_Color_Text As OLE_COLOR                                                   '按钮文本的颜色
Dim C_Color_Text_MouseMoved As OLE_COLOR                                        '触碰后按钮文本的颜色
Dim C_Text As String                                                            '文本
Dim C_Font_Name As String                                                       '字体名字
Dim C_Font_Size As Integer                                                      '字体大小
Dim C_Font_Bold As Boolean                                                      '字体加粗
Dim C_Font_Italic As Boolean                                                    '字体斜体
Dim C_Font_Underline As Boolean                                                 '字体下划线
Dim C_Is_Enabled As Boolean                                                     '是否可用
Dim C_Style_Border As Border                                                    '边框形式
Dim C_Color_Border As OLE_COLOR                                                 '边框颜色
Dim C_Can_Text_Move As Boolean                                                  '鼠标按下文本会向右下角移动
Dim C_Color_Back_ChangeSpeed As Integer                                         '渐变速度
Dim C_Text_Deviate_X As Integer                                                 '偏离的X坐标
Dim C_Text_Deviate_Y As Integer                                                 '偏离的Y坐标
Dim C_Color_Back_TransparentDegree As Integer                                   '决定是否背景透明
Dim C_Is_Text_Transparent As Boolean                                            '决定是否文字透明
'██████↓定义使用中所需的变量↓███████████████████████████████████████████████████████
Dim State As Integer                                                            '0无任何 1触碰 2触碰+按下
'██████↓定义边框形式的枚举量↓███████████████████████████████████████████████████████
Public Enum Border
    None = 0
    Opposite = 1
    Custom = 2
End Enum
'██████↓定义事件↓███████████████████████████████████████████████████████
Public Event Click()                                                            '单击事件
Public Event DblClick()                                                         '双击事件
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标按下事件
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标触碰事件
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '鼠标弹起事件
'██████↓重绘（刷新）过程↓███████████████████████████████████████████████████████
Public Sub Refresh()
    Cls                                                                         '清空
    Dim i As Long
    If C_Color_Back_TransparentDegree = 100 Then
        If BeAbleToBeBackTransparent Then
            '            For i = 0 To UserControl.ParentControls.Count - 1
            '                If UserControl.ParentControls.Item(i).Name = UserControl.Extender.Name Then
            '                    UserControl.PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.ParentControls.Item(i).Left, UserControl.ParentControls.Item(i).Top, UserControl.Width, UserControl.Height
            UserControl.PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left, UserControl.Extender.Top, UserControl.Width, UserControl.Height
            '                    Exit For
            '                End If
            '            Next
        End If
        If Not C_Is_Text_Transparent Then DrawTextInUsercontrol
    Else
        If BeAbleToBeBackTransparent Then
            TPic(0).Cls
            TPic(1).Cls
            TPic(0).Width = UserControl.Width
            TPic(0).Height = UserControl.Height
            TPic(1).Width = UserControl.Width
            TPic(1).Height = UserControl.Height
            '            For i = 0 To UserControl.ParentControls.Count - 1
            '                If UserControl.ParentControls.Item(i).Name = UserControl.Extender.Name Then
            '                    TPic(0).PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.ParentControls.Item(i).Left, UserControl.ParentControls.Item(i).Top, UserControl.Width, UserControl.Height
            '                    Exit For
            '                End If
            '            Next
            TPic(0).PaintPicture UserControl.Extender.Container.Image, 0, 0, UserControl.Width, UserControl.Height, UserControl.Extender.Left, UserControl.Extender.Top, UserControl.Width, UserControl.Height
            TPic(1).BackColor = UserControl.BackColor
            TPic(1).FontName = C_Font_Name                                      '改变按钮的字体名字
            TPic(1).FontSize = C_Font_Size                                      '改变按钮的字体大小
            TPic(1).FontBold = C_Font_Bold                                      '改变按钮的字体加粗
            TPic(1).FontItalic = C_Font_Italic                                  '改变按钮的字体斜体
            TPic(1).FontUnderline = C_Font_Underline                            '改变按钮的字体下划线
            If State = 0 Then
                TPic(1).ForeColor = C_Color_Text
            Else
                TPic(1).ForeColor = C_Color_Text_MouseMoved
            End If
            If (State = 0) Or (State = 1) Then                                  '如果没有按下
                TPic(1).CurrentX = (UserControl.Width - Label1.Width) / 2 + C_Text_Deviate_X * 15
                TPic(1).CurrentY = (UserControl.Height - Label1.Height) / 2 + C_Text_Deviate_Y * 15 '打印到中心位置
            ElseIf (State = 2) Then
                If C_Can_Text_Move Then
                    TPic(1).CurrentX = (UserControl.Width - Label1.Width) / 2 + 30 + C_Text_Deviate_X * 15
                    TPic(1).CurrentY = (UserControl.Height - Label1.Height) / 2 + 30 + C_Text_Deviate_Y * 15 '打印到中心偏右下位置
                Else
                    TPic(1).CurrentX = (UserControl.Width - Label1.Width) / 2 + C_Text_Deviate_X * 15
                    TPic(1).CurrentY = (UserControl.Height - Label1.Height) / 2 + C_Text_Deviate_Y * 15 '打印到中心位置
                End If
            End If
            If C_Is_Text_Transparent Then
                TPic(1).Print Label1                                            '打印文本
                PM.ControlTransparent TPic(0), TPic(1), Int(C_Color_Back_TransparentDegree / 100 * 255)
            Else
                PM.ControlTransparent TPic(0), TPic(1), Int(C_Color_Back_TransparentDegree / 100 * 255)
                TPic(1).Print Label1                                            '打印文本
            End If
            Set UserControl.Picture = TPic(1).Image
        Else
            DrawTextInUsercontrol
        End If
    End If
End Sub

Private Function BeAbleToBeBackTransparent() As Boolean
    On Error GoTo Err
    '    Dim i As String
    Dim pppp As StdPicture
    '    i = UserControl.ParentControls.Item(0).Name
    Set pppp = UserControl.Extender.Container.Image
    BeAbleToBeBackTransparent = True
    Exit Function
Err:
    BeAbleToBeBackTransparent = False
End Function

Private Sub DrawTextInUsercontrol()
    UserControl.FontName = C_Font_Name                                          '改变按钮的字体名字
    UserControl.FontSize = C_Font_Size                                          '改变按钮的字体大小
    UserControl.FontBold = C_Font_Bold                                          '改变按钮的字体加粗
    UserControl.FontItalic = C_Font_Italic                                      '改变按钮的字体斜体
    UserControl.FontUnderline = C_Font_Underline                                '改变按钮的字体下划线
    If State = 0 Then
        UserControl.ForeColor = C_Color_Text
    Else
        UserControl.ForeColor = C_Color_Text_MouseMoved
    End If
    If (State = 0) Or (State = 1) Then                                          '如果没有按下
        UserControl.CurrentX = (UserControl.Width - Label1.Width) / 2 + C_Text_Deviate_X * 15
        UserControl.CurrentY = (UserControl.Height - Label1.Height) / 2 + C_Text_Deviate_Y * 15 '打印到中心位置
    ElseIf (State = 2) Then
        If C_Can_Text_Move Then
            UserControl.CurrentX = (UserControl.Width - Label1.Width) / 2 + 30 + C_Text_Deviate_X * 15
            UserControl.CurrentY = (UserControl.Height - Label1.Height) / 2 + 30 + C_Text_Deviate_Y * 15 '打印到中心偏右下位置
        Else
            UserControl.CurrentX = (UserControl.Width - Label1.Width) / 2 + C_Text_Deviate_X * 15
            UserControl.CurrentY = (UserControl.Height - Label1.Height) / 2 + C_Text_Deviate_Y * 15 '打印到中心位置
        End If
    End If
    UserControl.Print Label1                                                    '打印文本
End Sub
'██████↓各种属性↓███████████████████████████████████████████████████████
Public Property Get Is_Enabled() As Boolean                                     '是否有效
    Is_Enabled = C_Is_Enabled
End Property

Public Property Let Is_Enabled(ByVal vNewValue As Boolean)
    C_Is_Enabled = vNewValue
    PropertyChanged "Is_Enabled"
End Property

Public Property Get Font_Name() As String                                       '字体名字
    Font_Name = C_Font_Name
End Property

Public Property Let Font_Name(ByVal vNewValue As String)
    C_Font_Name = vNewValue
    Label1.FontName = vNewValue
    Refresh
    PropertyChanged "Font_Name"
End Property

Public Property Get Font_Size() As Integer                                      '字体大小
    Font_Size = C_Font_Size
End Property

Public Property Let Font_Size(ByVal vNewValue As Integer)
    C_Font_Size = vNewValue
    Label1.FontSize = vNewValue
    Refresh
    PropertyChanged "Font_Size"
End Property

Public Property Get Font_Bold() As Boolean                                      '字体加粗
    Font_Bold = C_Font_Bold
End Property

Public Property Let Font_Bold(ByVal vNewValue As Boolean)
    C_Font_Bold = vNewValue
    Label1.FontBold = vNewValue
    Refresh
    PropertyChanged "Font_Bold"
End Property

Public Property Get Font_Italic() As Boolean                                    '字体斜体
    Font_Italic = C_Font_Italic
End Property

Public Property Let Font_Italic(ByVal vNewValue As Boolean)
    C_Font_Italic = vNewValue
    Label1.FontItalic = vNewValue
    Refresh
    PropertyChanged "Font_Italic"
End Property

Public Property Get Font_Underline() As Boolean                                 '字体下划线
    Font_Underline = C_Font_Underline
End Property

Public Property Let Font_Underline(ByVal vNewValue As Boolean)
    C_Font_Underline = vNewValue
    Label1.FontUnderline = vNewValue
    Refresh
    PropertyChanged "Font_Underline"
End Property

Public Property Get Text() As String                                            '文本
    Text = C_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    Label1 = vNewValue
    Refresh
    PropertyChanged "Text"
End Property

Public Property Get Color_Back() As OLE_COLOR                                   '背景颜色
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    UserControl.BackColor = vNewValue
    C_Color_Begin = C_Color_Back
    PropertyChanged "Color_Begin"
    Refresh
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_Back_Down() As OLE_COLOR                              '鼠标按下时背景颜色
    Color_Back_Down = C_Color_Back_Down
End Property

Public Property Let Color_Back_Down(ByVal vNewValue As OLE_COLOR)
    C_Color_Back_Down = vNewValue
    PropertyChanged "Color_Back_Down"
End Property

Public Property Get Color_Begin() As OLE_COLOR                                  '渐变开始的颜色
    Color_Begin = C_Color_Begin
End Property

Public Property Let Color_Begin(ByVal vNewValue As OLE_COLOR)
    C_Color_Begin = C_Color_Back
    PropertyChanged "Color_Begin"
End Property

Public Property Get Color_End() As OLE_COLOR                                    '渐变结束的颜色
    Color_End = C_Color_End
End Property

Public Property Let Color_End(ByVal vNewValue As OLE_COLOR)
    C_Color_End = vNewValue
    PropertyChanged "Color_End"
End Property

Public Property Get Color_Text() As OLE_COLOR                                   '文本颜色
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    Refresh
    PropertyChanged "Color_Text"
End Property

Public Property Get Color_Text_MouseMoved() As OLE_COLOR                        '触碰后文本颜色
    Color_Text_MouseMoved = C_Color_Text_MouseMoved
End Property

Public Property Let Color_Text_MouseMoved(ByVal vNewValue As OLE_COLOR)
    C_Color_Text_MouseMoved = vNewValue
    Refresh
    PropertyChanged "Color_Text_MouseMoved"
End Property

Public Property Get Style_Border() As Border                                    '边框形式
    Style_Border = C_Style_Border
End Property

Public Property Let Style_Border(ByVal vNewValue As Border)
    C_Style_Border = vNewValue
    PropertyChanged "Style_Border"
End Property

Public Property Get Color_Border() As OLE_COLOR                                 '边框颜色
    Color_Border = C_Color_Border
End Property

Public Property Let Color_Border(ByVal vNewValue As OLE_COLOR)
    C_Color_Border = vNewValue
    PropertyChanged "Color_Border"
End Property

Public Property Get Can_Text_Move() As Boolean                                  '鼠标按下文本会向右下角移动
    Can_Text_Move = C_Can_Text_Move
End Property

Public Property Let Can_Text_Move(ByVal vNewValue As Boolean)
    C_Can_Text_Move = vNewValue
    PropertyChanged "Can_Text_Move"
End Property

Public Property Get Color_Back_ChangeSpeed() As Integer                         '颜色渐变快慢
    Color_Back_ChangeSpeed = C_Color_Back_ChangeSpeed
End Property

Public Property Let Color_Back_ChangeSpeed(ByVal vNewValue As Integer)
    C_Color_Back_ChangeSpeed = vNewValue
    If C_Color_Back_ChangeSpeed < 1 Then C_Color_Back_ChangeSpeed = 1
    If C_Color_Back_ChangeSpeed > 30 Then C_Color_Back_ChangeSpeed = 30
    PropertyChanged "Color_Back_ChangeSpeed"
End Property

Public Property Get Text_Deviate_X() As Integer                                 '偏离的X坐标
    Text_Deviate_X = C_Text_Deviate_X
End Property

Public Property Let Text_Deviate_X(ByVal vNewValue As Integer)
    C_Text_Deviate_X = vNewValue
    Refresh
    PropertyChanged "Text_Deviate_X"
End Property

Public Property Get Text_Deviate_Y() As Integer                                 '偏离的Y坐标
    Text_Deviate_Y = C_Text_Deviate_Y
End Property

Public Property Let Text_Deviate_Y(ByVal vNewValue As Integer)
    C_Text_Deviate_Y = vNewValue
    Refresh
    PropertyChanged "Text_Deviate_Y"
End Property

Public Property Get Color_Back_TransparentDegree() As Integer                   '决定是否背景透明
    Color_Back_TransparentDegree = C_Color_Back_TransparentDegree
End Property

Public Property Let Color_Back_TransparentDegree(ByVal vNewValue As Integer)
    C_Color_Back_TransparentDegree = vNewValue
    If C_Color_Back_TransparentDegree < 0 Then C_Color_Back_TransparentDegree = 0
    If C_Color_Back_TransparentDegree > 100 Then C_Color_Back_TransparentDegree = 100
    Refresh
    PropertyChanged "Color_Back_TransparentDegree"
End Property

Public Property Get Is_Text_Transparent() As Boolean                            '决定是否文本透明
    Is_Text_Transparent = C_Is_Text_Transparent
End Property

Public Property Let Is_Text_Transparent(ByVal vNewValue As Boolean)
    C_Is_Text_Transparent = vNewValue
    PropertyChanged "Is_Text_Transparent"
End Property

Private Sub PM_ColorSmlyIng(nColor As Long)
    UserControl.BackColor = nColor
    Refresh
End Sub

'██████↓各种事件↓███████████████████████████████████████████████████████
Private Sub Timer1_Timer()                                                      '计时器1的计时事件
    Dim E As Long                                                               '定义终止颜色
    '    Dim R1 As Integer, G1 As Integer, B1 As Integer '定义起始颜色的RGB值
    '    Dim R2 As Integer, G2 As Integer, B2 As Integer '定义终止颜色的RGB值
    If State = 2 Then
        E = C_Color_Back_Down                                                   '终止颜色是鼠标按下背景颜色
    ElseIf State = 1 Then                                                       '如果鼠标触碰
        E = C_Color_End                                                         '终止颜色是渐变结束颜色
    Else                                                                        '如果鼠标没有触碰
        E = C_Color_Back                                                        '终止颜色是渐变开始(控件背景)颜色
    End If
    PM.ColorSmly UserControl.BackColor, E, C_Color_Back_ChangeSpeed, 1
    Timer1.Enabled = False
    '    R1 = UserControl.BackColor Mod 256 '计算起始颜色的RGB值
    '    G1 = (UserControl.BackColor Mod 65536) \ 256
    '    B1 = UserControl.BackColor \ 65536
    '    R2 = E Mod 256 '计算终止颜色的RGB值
    '    G2 = (E Mod 65536) \ 256
    '    B2 = E \ 65536
    '    If R1 < R2 Then '更新颜色的RGB值,变化量为C_Color_Back_ChangeSpeed
    '        If R1 + C_Color_Back_ChangeSpeed < R2 Then
    '            R1 = R1 + C_Color_Back_ChangeSpeed
    '        Else
    '            R1 = R2
    '        End If
    '    End If
    '    If R1 > R2 Then
    '        If R1 - C_Color_Back_ChangeSpeed > R2 Then
    '            R1 = R1 - C_Color_Back_ChangeSpeed
    '        Else
    '            R1 = R2
    '        End If
    '    End If
    '    If G1 < G2 Then
    '        If G1 + C_Color_Back_ChangeSpeed < G2 Then
    '            G1 = G1 + C_Color_Back_ChangeSpeed
    '        Else
    '            G1 = G2
    '        End If
    '    End If
    '    If G1 > G2 Then
    '        If G1 - C_Color_Back_ChangeSpeed > G2 Then
    '            G1 = G1 - C_Color_Back_ChangeSpeed
    '        Else
    '            G1 = G2
    '        End If
    '    End If
    '    If B1 < B2 Then
    '        If B1 + C_Color_Back_ChangeSpeed < B2 Then
    '            B1 = B1 + C_Color_Back_ChangeSpeed
    '        Else
    '            B1 = B2
    '        End If
    '    End If
    '    If B1 > B2 Then
    '        If B1 - C_Color_Back_ChangeSpeed > B2 Then
    '            B1 = B1 - C_Color_Back_ChangeSpeed
    '        Else
    '            B1 = B2
    '        End If
    '    End If
    '    UserControl.BackColor = RGB(R1, G1, B1) '更新控件颜色
    '    If (R1 = R2) And (G1 = G2) And (B1 = B2) Then '如果颜色渐变完成(即起始颜色等于终止颜色)
    '        Timer1.Enabled = False '关闭计时器1
    '        If State = 2 Then
    '            UserControl.BackColor = C_Color_Back_Down '终止颜色是鼠标按下背景颜色
    '        ElseIf State = 1 Then '如果鼠标触碰过
    '            UserControl.BackColor = C_Color_End '控件颜色变为渐变结束颜色
    '        Else '如果没有触碰过
    '            UserControl.BackColor = C_Color_Back '控件颜色变为渐变开始(控件背景)颜色
    '        End If
    '    End If
    '    Refresh '刷新
End Sub

Private Sub Timer2_Timer()
    If (判断鼠标是否指向指定控件上(UserControl.hwnd) = False) And (State <> 2) Then
        State = 0
        Timer1.Enabled = True                                                   '开启计时器1
        Timer2.Enabled = False                                                  '关闭计时器2
        Shape1.Visible = False                                                  '隐藏边框
    End If
End Sub

Private Sub UserControl_Click()                                                 '控件的单击事件
    If Is_Enabled = True Then RaiseEvent Click                                  '触发单击事件
End Sub

Private Sub UserControl_DblClick()                                              '控件的双击事件
    If Is_Enabled = True Then RaiseEvent DblClick                               '触发双击事件
End Sub

Private Sub UserControl_Initialize()                                            '控件的加载事件
    C_Is_Enabled = True                                                         '定义每种属性的初始值
    C_Color_Back = &HF2AF00
    C_Color_Back_Down = &HF2AF00
    C_Color_Begin = &HF2AF00
    C_Color_End = &HFF7402
    C_Color_Text = &H0&
    C_Color_Text_MouseMoved = &HFFFFFF
    C_Text = "PButton"
    C_Font_Name = "微软雅黑"
    C_Font_Size = 11
    C_Font_Bold = False
    C_Font_Italic = False
    C_Font_Underline = False
    C_Style_Border = 1
    C_Color_Border = &H0&
    C_Color_Back_ChangeSpeed = 5
    C_Text_Deviate_X = 0
    C_Text_Deviate_Y = 0
    Label1 = "PButton"                                                          '配置每种属性
    Label1.FontName = "微软雅黑"
    Label1.FontSize = 11
    Label1.FontBold = False
    Label1.FontItalic = False
    Label1.FontUnderline = False
    C_Color_Back_TransparentDegree = 0
    C_Is_Text_Transparent = False
    State = 0
    Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)           '控件的键盘按下事件
    If KeyCode = 32 Or KeyCode = 13 Then                                        '如果按下的键是回车或空格
        If Is_Enabled = True Then                                               '如果控件有效
            State = 2                                                           '鼠标按下
            Timer1.Enabled = True                                               '打开计时器1
            Timer2.Enabled = True                                               '打开计时器2
            Refresh                                                             '刷新
        End If
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)             '控件的键盘弹起事件
    If KeyCode = 32 Or KeyCode = 13 Then                                        '如果按下的键是回车或空格
        If Is_Enabled = True Then                                               '如果控件有效
            State = 0
            Refresh                                                             '刷新
            RaiseEvent Click                                                    '触发单击事件
        End If
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) '控件的鼠标按下事件
    If Is_Enabled = True Then                                                   '如果控件有效
        State = 2                                                               '鼠标按下
        Timer1.Enabled = True                                                   '打开计时器1
        Timer2.Enabled = True                                                   '打开计时器2
        Refresh                                                                 '刷新
        RaiseEvent MouseDown(Button, Shift, X, Y)                               '触发鼠标按下事件
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) '控件的鼠标触碰事件
    If Is_Enabled = True Then                                                   '如果控件有效
        If State <> 2 Then State = 1                                            '鼠标触碰
        Timer1.Enabled = True                                                   '打开计时器1
        Timer2.Enabled = True                                                   '打开计时器2
        Shape1.Height = UserControl.Height                                      '使边框大小与控件大小一致
        Shape1.Width = UserControl.Width
        Select Case C_Style_Border                                              '分情况讨论边框形式
        Case 0                                                                  '无边框
            '
        Case 1                                                                  '渐变结束颜色的相反色
            Shape1.BorderColor = RGB(Abs(255 - C_Color_End Mod 256), Abs(255 - (C_Color_End Mod 65536) \ 256), Abs(255 - C_Color_End \ 65536))
            Shape1.Visible = True
        Case 2                                                                  '自定义的颜色
            Shape1.BorderColor = C_Color_Border
            Shape1.Visible = True
        End Select
        RaiseEvent MouseMove(Button, Shift, X, Y)                               '触发鼠标触碰事件
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '控件的鼠标弹起事件
    If Is_Enabled = True Then
        State = 1                                                               '鼠标没有按下
        Refresh                                                                 '刷新
        RaiseEvent MouseUp(Button, Shift, X, Y)                                 '触发鼠标弹起事件
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)                  '控件的读取属性事件
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)                 '读取各种属性和初始值
    C_Color_Back_Down = PropBag.ReadProperty("Color_Back_Down", &HF2AF00)
    C_Color_Begin = PropBag.ReadProperty("Color_Begin", &HF2AF00)
    C_Color_End = PropBag.ReadProperty("Color_End", &HFF7402)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    C_Color_Text_MouseMoved = PropBag.ReadProperty("Color_Text_MouseMoved", &HFFFFFF)
    C_Text = PropBag.ReadProperty("Text", "PButton")
    C_Font_Name = PropBag.ReadProperty("Font_Name", "微软雅黑")
    C_Font_Size = PropBag.ReadProperty("Font_Size", 11)
    C_Font_Bold = PropBag.ReadProperty("Font_Bold", False)
    C_Font_Italic = PropBag.ReadProperty("Font_Italic", False)
    C_Font_Underline = PropBag.ReadProperty("Font_Underline", False)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Style_Border = PropBag.ReadProperty("Style_Border", 1)
    C_Color_Border = PropBag.ReadProperty("Color_Border", &H0&)
    C_Can_Text_Move = PropBag.ReadProperty("Can_Text_Move", True)
    C_Color_Back_ChangeSpeed = PropBag.ReadProperty("Color_Back_ChangeSpeed", 5)
    C_Text_Deviate_X = PropBag.ReadProperty("Text_Deviate_X", 0)
    C_Text_Deviate_Y = PropBag.ReadProperty("Text_Deviate_Y", 0)
    C_Color_Back_TransparentDegree = PropBag.ReadProperty("Color_Back_TransparentDegree", 0)
    C_Is_Text_Transparent = PropBag.ReadProperty("Is_Text_Transparent", False)
    UserControl.BackColor = C_Color_Back                                        '配置各种属性
    Label1 = C_Text
    Label1.FontName = C_Font_Name
    Label1.FontSize = C_Font_Size
    Label1.FontBold = C_Font_Bold
    Label1.FontItalic = C_Font_Italic
    Label1.FontUnderline = C_Font_Underline
    Refresh                                                                     '刷新
End Sub

Private Sub UserControl_Resize()                                                '控件的大小改变事件
    Refresh                                                                     '刷新
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)                 '控件的写属性事件
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)            '写入各种属性和初始值
    Call PropBag.WriteProperty("Color_Back_Down", C_Color_Back_Down, &HF2AF00)
    Call PropBag.WriteProperty("Color_Begin", C_Color_Begin, &HF2AF00)
    Call PropBag.WriteProperty("Color_End", C_Color_End, &HFF7402)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Color_Text_MouseMoved", C_Color_Text_MouseMoved, &HFFFFFF)
    Call PropBag.WriteProperty("Text", C_Text, "PButton")
    Call PropBag.WriteProperty("Font_Name", C_Font_Name, "微软雅黑")
    Call PropBag.WriteProperty("Font_Size", C_Font_Size, 11)
    Call PropBag.WriteProperty("Font_Bold", C_Font_Bold, False)
    Call PropBag.WriteProperty("Font_Italic", C_Font_Italic, False)
    Call PropBag.WriteProperty("Font_Underline", C_Font_Underline, False)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Style_Border", C_Style_Border, 1)
    Call PropBag.WriteProperty("Color_Border", C_Color_Border, &H0&)
    Call PropBag.WriteProperty("Can_Text_Move", C_Can_Text_Move, True)
    Call PropBag.WriteProperty("Color_Back_ChangeSpeed", C_Color_Back_ChangeSpeed, 5)
    Call PropBag.WriteProperty("Text_Deviate_X", C_Text_Deviate_X, 0)
    Call PropBag.WriteProperty("Text_Deviate_Y", C_Text_Deviate_Y, 0)
    Call PropBag.WriteProperty("Color_Back_TransparentDegree", C_Color_Back_TransparentDegree, 0)
    Call PropBag.WriteProperty("Is_Text_Transparent", C_Is_Text_Transparent, 0)
End Sub
'↑↑↑↑↑↑↑The End↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑
