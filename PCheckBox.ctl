VERSION 5.00
Begin VB.UserControl PCheckBox 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ScaleHeight     =   375
   ScaleWidth      =   1455
   ToolboxBitmap   =   "PCheckBox.ctx":0000
   Begin 市井中孤傲的烟火.PButton B2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
   End
   Begin 市井中孤傲的烟火.PButton B1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Text            =   "×"
   End
End
Attribute VB_Name = "PCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'██████↓定义存储属性的变量↓███████████████████████████████████████████████████████
Dim C_Color_Back As OLE_COLOR                                                   '背景颜色
Dim C_Color_End As OLE_COLOR                                                    '渐变结束的颜色
Dim C_Color_Text As OLE_COLOR                                                   '按钮文本的颜色
Dim C_Text As String                                                            '文本
Dim C_Font_Name As String                                                       '字体名字
Dim C_Font_Size As Integer                                                      '字体大小
Dim C_Font_Bold As Boolean                                                      '字体加粗
Dim C_Font_Italic As Boolean                                                    '字体斜体
Dim C_Font_Underline As Boolean                                                 '字体下划线
Dim C_Is_Enabled As Boolean                                                     '是否可用
Dim C_Value As Boolean                                                          '值
'██████↓定义事件↓███████████████████████████████████████████████████████
Public Event ValueChange(nValue As Boolean)                                     '值改变事件
Public Event Click()                                                            '单击事件
Public Event DblClick()                                                         '双击事件
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Boolean) '鼠标按下事件,NValue为新值
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Boolean) '鼠标触碰事件,NValue为新值
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single, nValue As Boolean) '鼠标弹起事件,NValue为新值
'██████↓各种属性↓███████████████████████████████████████████████████████
Public Property Get Value() As Boolean                                          '值
    Value = C_Value
End Property

Public Property Let Value(ByVal vNewValue As Boolean)
    C_Value = vNewValue
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
    PropertyChanged "Value"
End Property

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
    B1.Font_Name = vNewValue
    B2.Font_Name = vNewValue
    PropertyChanged "Font_Name"
End Property

Public Property Get Font_Size() As Integer                                      '字体大小
    Font_Size = C_Font_Size
End Property

Public Property Let Font_Size(ByVal vNewValue As Integer)
    C_Font_Size = vNewValue
    B1.Font_Size = vNewValue
    B2.Font_Size = vNewValue
    PropertyChanged "Font_Size"
End Property

Public Property Get Font_Bold() As Boolean                                      '字体加粗
    Font_Bold = C_Font_Bold
End Property

Public Property Let Font_Bold(ByVal vNewValue As Boolean)
    C_Font_Bold = vNewValue
    B1.Font_Bold = vNewValue
    B2.Font_Bold = vNewValue
    PropertyChanged "Font_Bold"
End Property

Public Property Get Font_Italic() As Boolean                                    '字体斜体
    Font_Italic = C_Font_Italic
End Property

Public Property Let Font_Italic(ByVal vNewValue As Boolean)
    C_Font_Italic = vNewValue
    B1.Font_Italic = vNewValue
    B2.Font_Italic = vNewValue
    PropertyChanged "Font_Italic"
End Property

Public Property Get Font_Underline() As Boolean                                 '字体下划线
    Font_Underline = C_Font_Underline
End Property

Public Property Let Font_Underline(ByVal vNewValue As Boolean)
    C_Font_Underline = vNewValue
    B1.Font_Underline = vNewValue
    B2.Font_Underline = vNewValue
    PropertyChanged "Font_Underline"
End Property

Public Property Get Text() As String                                            '文本
    Text = C_Text
End Property

Public Property Let Text(ByVal vNewValue As String)
    C_Text = vNewValue
    B2.Text = vNewValue
    PropertyChanged "Text"
End Property

Public Property Get Color_Back() As OLE_COLOR                                   '背景颜色
    Color_Back = C_Color_Back
End Property

Public Property Let Color_Back(ByVal vNewValue As OLE_COLOR)
    C_Color_Back = vNewValue
    B1.Color_Back = vNewValue
    B2.Color_Back = vNewValue
    PropertyChanged "Color_Back"
End Property

Public Property Get Color_End() As OLE_COLOR                                    '渐变结束的颜色
    Color_End = C_Color_End
End Property

Public Property Let Color_End(ByVal vNewValue As OLE_COLOR)
    C_Color_End = vNewValue
    B1.Color_End = vNewValue
    B2.Color_End = vNewValue
    PropertyChanged "Color_End"
End Property

Public Property Get Color_Text() As OLE_COLOR                                   '文本颜色
    Color_Text = C_Color_Text
End Property

Public Property Let Color_Text(ByVal vNewValue As OLE_COLOR)
    C_Color_Text = vNewValue
    B1.Color_Text = vNewValue
    B2.Color_Text = vNewValue
    PropertyChanged "Color_Text"
End Property
'██████↓各种事件↓███████████████████████████████████████████████████████
Private Sub B1_Click()                                                          'B1的单击事件
    Value = Not (Value)                                                         '值取相反
    If C_Is_Enabled = True Then RaiseEvent Click                                '如果有效,触发单击事件
    If C_Is_Enabled = True Then RaiseEvent ValueChange(C_Value)                 '如果有效,触发值改变事件
End Sub

Private Sub B1_DblClick()                                                       'B1的双击事件
    If C_Is_Enabled = True Then RaiseEvent DblClick                             '如果有效,触发双击事件
End Sub

Private Sub B1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标按下事件
    If C_Is_Enabled = True Then RaiseEvent MouseDown(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标按下事件
End Sub

Private Sub B1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标触碰事件
    If C_Is_Enabled = True Then RaiseEvent MouseMove(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标触碰事件
End Sub

Private Sub B1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B1的鼠标弹起事件
    If C_Is_Enabled = True Then RaiseEvent MouseUp(Button, Shift, X, Y, C_Value) '如果有效,触发鼠标弹起事件
End Sub

Private Sub B2_Click()                                                          'B2的单击事件
    B1_Click                                                                    '调用B1的单击事件
End Sub

Private Sub B2_DblClick()                                                       'B2的双击事件
    B1_DblClick                                                                 '调用B1的双击事件
End Sub

Private Sub B2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标按下事件
    B1_MouseDown Button, Shift, X, Y                                            '调用B1的鼠标按下事件
End Sub

Private Sub B2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标触碰事件
    B1_MouseMove Button, Shift, X, Y                                            '调用B1的鼠标触碰事件
End Sub

Private Sub B2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'B2的鼠标弹起事件
    B1_MouseUp Button, Shift, X, Y                                              '调用B1的鼠标弹起事件
End Sub

Private Sub UserControl_Initialize()                                            '控件的加载事件
    C_Color_Back = &HF2AF00                                                     '定义每种属性的初始值
    C_Color_End = &HFF7402
    C_Color_Text = &H0&
    C_Text = "PCheckBox"
    C_Font_Name = "微软雅黑"
    C_Font_Size = 11
    C_Font_Bold = False
    C_Font_Italic = False
    C_Font_Underline = False
    C_Is_Enabled = True
    C_Value = False
    B1.Color_Text = C_Color_Text                                                '配置每种属性
    B2.Color_Text = C_Color_Text
    B1.Color_End = C_Color_End
    B2.Color_End = C_Color_End
    B1.Color_Back = C_Color_Back
    B2.Color_Back = C_Color_Back
    B2.Text = C_Text
    B1.Font_Underline = C_Font_Underline
    B2.Font_Underline = C_Font_Underline
    B1.Font_Italic = C_Font_Italic
    B2.Font_Italic = C_Font_Italic
    B1.Font_Bold = C_Font_Bold
    B2.Font_Bold = C_Font_Bold
    B1.Font_Size = C_Font_Size
    B2.Font_Size = C_Font_Size
    B1.Font_Name = C_Font_Name
    B2.Font_Name = C_Font_Name
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)                  '控件的读取属性事件
    C_Color_Back = PropBag.ReadProperty("Color_Back", &HF2AF00)                 '读取各种属性和初始值
    C_Color_End = PropBag.ReadProperty("Color_End", &HFF7402)
    C_Color_Text = PropBag.ReadProperty("Color_Text", &H0&)
    C_Text = PropBag.ReadProperty("Text", "PButton")
    C_Font_Name = PropBag.ReadProperty("Font_Name", "微软雅黑")
    C_Font_Size = PropBag.ReadProperty("Font_Size", 11)
    C_Font_Bold = PropBag.ReadProperty("Font_Bold", False)
    C_Font_Italic = PropBag.ReadProperty("Font_Italic", False)
    C_Font_Underline = PropBag.ReadProperty("Font_Underline", False)
    C_Is_Enabled = PropBag.ReadProperty("Is_Enabled", True)
    C_Value = PropBag.ReadProperty("Value", False)
    B1.Color_Text = C_Color_Text                                                '配置各种属性
    B2.Color_Text = C_Color_Text
    B1.Color_End = C_Color_End
    B2.Color_End = C_Color_End
    B1.Color_Back = C_Color_Back
    B2.Color_Back = C_Color_Back
    B2.Text = C_Text
    B1.Font_Underline = C_Font_Underline
    B2.Font_Underline = C_Font_Underline
    B1.Font_Italic = C_Font_Italic
    B2.Font_Italic = C_Font_Italic
    B1.Font_Bold = C_Font_Bold
    B2.Font_Bold = C_Font_Bold
    B1.Font_Size = C_Font_Size
    B2.Font_Size = C_Font_Size
    B1.Font_Name = C_Font_Name
    B2.Font_Name = C_Font_Name
    If C_Value = True Then
        B1.Text = "√"
    Else
        B1.Text = "×"
    End If
End Sub

Private Sub UserControl_Resize()                                                '控件的大小改变事件
    If UserControl.Width < UserControl.Height Then UserControl.Width = UserControl.Height '改变每个控件的大小和位置
    B1.Width = UserControl.Height
    B1.Height = UserControl.Height
    B2.Height = UserControl.Height
    B2.Width = UserControl.Width - B1.Width
    B2.Left = UserControl.Width - B2.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)                 '控件的写属性事件
    Call PropBag.WriteProperty("Color_Back", C_Color_Back, &HF2AF00)            '写入各种属性和初始值
    Call PropBag.WriteProperty("Color_End", C_Color_End, &HFF7402)
    Call PropBag.WriteProperty("Color_Text", C_Color_Text, &H0&)
    Call PropBag.WriteProperty("Text", C_Text, "PButton")
    Call PropBag.WriteProperty("Font_Name", C_Font_Name, "微软雅黑")
    Call PropBag.WriteProperty("Font_Size", C_Font_Size, 11)
    Call PropBag.WriteProperty("Font_Bold", C_Font_Bold, False)
    Call PropBag.WriteProperty("Font_Italic", C_Font_Italic, False)
    Call PropBag.WriteProperty("Font_Underline", C_Font_Underline, False)
    Call PropBag.WriteProperty("Is_Enabled", C_Is_Enabled, True)
    Call PropBag.WriteProperty("Value", C_Value, False)
End Sub
'↑↑↑↑↑↑↑The End↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

