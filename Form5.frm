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
   StartUpPosition =   2  '��Ļ����
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
Private FormOldWidth     As Long     '���洰���ԭʼ���
Private FormOldHeight     As Long     '���洰���ԭʼ�߶�
Private Sub Form_Load()
Call ResizeInit(Me)     '�ڳ���װ��ʱ�������
Form5.Height = Screen.Height
Form5.Width = Screen.Width
End Sub
Private Sub Form_Resize()
Call ResizeForm(Me)     'ȷ������ı�ʱ�ؼ���֮�ı�
End Sub




'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
      Dim Pos(4)     As Double
      Dim i     As Long, TempPos       As Long, StartPos       As Long
      Dim Obj     As Control
      Dim ScaleX     As Double, ScaleY       As Double
      ScaleX = FormName.ScaleWidth / FormOldWidth           '���洰�������ű���
      ScaleY = FormName.ScaleHeight / FormOldHeight           '���洰��߶����ű���
      On Error Resume Next
      For Each Obj In FormName
          StartPos = 1
          For i = 0 To 4
          '��ȡ�ؼ���ԭʼλ�����С
              TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
              If TempPos > 0 Then
                  Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                  StartPos = TempPos + 1
              Else
                  Pos(i) = 0
              End If
        '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
              Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
          Next i
      Next Obj
      On Error GoTo 0
End Sub



'�ڵ���ResizeFormǰ�ȵ��ñ�����
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

