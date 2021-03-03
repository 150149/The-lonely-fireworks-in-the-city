VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   0  'None
   Caption         =   "Form5"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form5"
   ScaleHeight     =   5475
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth     As Long     '保存窗体的原始宽度
Private FormOldHeight     As Long     '保存窗体的原始高度
Private Sub Form_Load()
Call ResizeInit(Me)     '在程序装入时必须加入
Form5.Height = Screen.Height
Form5.Width = Screen.Width
End Sub
Private Sub Form_Resize()
Call ResizeForm(Me)     '确保窗体改变时控件随之改变
End Sub




'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
      Dim Pos(4)     As Double
      Dim i     As Long, TempPos       As Long, StartPos       As Long
      Dim Obj     As Control
      Dim ScaleX     As Double, ScaleY       As Double
      ScaleX = FormName.ScaleWidth / FormOldWidth           '保存窗体宽度缩放比例
      ScaleY = FormName.ScaleHeight / FormOldHeight           '保存窗体高度缩放比例
      On Error Resume Next
      For Each Obj In FormName
          StartPos = 1
          For i = 0 To 4
          '读取控件的原始位置与大小
              TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
              If TempPos > 0 Then
                  Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                  StartPos = TempPos + 1
              Else
                  Pos(i) = 0
              End If
        '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
              Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
          Next i
      Next Obj
      On Error GoTo 0
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

