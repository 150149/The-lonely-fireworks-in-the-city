VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "��������3.frx":0000
   LinkTopic       =   "Form7"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002C2C2C&
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Text            =   "�µ�����"
      Top             =   2880
      Width           =   9135
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   0
   End
   Begin �о��й°����̻�.PButton PButton2 
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Back      =   12648447
      Color_Back_Down =   12648447
      Color_Begin     =   12648447
      Color_End       =   12648447
      Color_Text      =   16777215
      Text            =   "�����޾�ģʽ"
      Font_Size       =   15
      Font_Bold       =   -1  'True
      Color_Border    =   12632256
      Can_Text_Move   =   0   'False
   End
   Begin �о��й°����̻�.PSwitch PSwitch1 
      Height          =   735
      Left            =   6120
      TabIndex        =   2
      Top             =   4920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1296
      Color_Top       =   12632256
      Color_Back      =   8421504
      Color_Back_True =   14737632
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   735
      Left            =   10680
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
      Color_Back      =   12648447
      Color_Back_Down =   12648447
      Color_Begin     =   12648447
      Color_End       =   12648447
      Color_Text      =   16777215
      Text            =   "��������ģʽ"
      Font_Size       =   15
      Font_Bold       =   -1  'True
      Color_Border    =   12632256
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   3840
      TabIndex        =   7
      Top             =   7200
      Width           =   9135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   3720
      TabIndex        =   6
      Top             =   7080
      Width           =   9255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ�꣺������ʱ,Ŭ���������,������ʱ�䴴�����޿���"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10680
      TabIndex        =   5
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ�꣺��������,������ʱ,���ʱ��,Ŭ����ְ"
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   5880
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�浵����"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private pb1e As Double
Private pb2e As Double
Private pb3e As Double
Private pb4e As Double
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1


Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������
    'Form13.BackColor = &H80000
    'Dim rtn As Long
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, 0, 70, LWA_ALPHA '   ����͸��
    
    Me.BackColor = &HC0FFFF
    Dim rtn As Long
    Dim BorderStyler
    BorderStyler = 0
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, &HC0FFFF, 70, LWA_COLORKEY
    
    'Dim rtn As Long
    'rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    'rtn = rtn Or WS_EX_LAYERED
    'SetWindowLong hwnd, GWL_EXSTYLE, rtn
    'SetLayeredWindowAttributes hwnd, vbBlack, 70, LWA_COLORKEY
    
    Form13.Width = Screen.Width
    Form13.Height = Screen.Height
    Form14.Show
    Form13.Show
    Form15.Show
    Form15.Height = Form13.Label6.Height
    Form15.Width = Form13.Label6.Width
    Form15.Top = Form13.Label6.Top
    Form15.Left = Form13.Label6.Left
    
    
End Sub



'�������ı���ڸ�Ԫ���Ĵ�С���ڵ���ReSizeFormǰ�ȵ���ReSizeInit����
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4)     As Double
    Dim I     As Long, TempPos       As Long, StartPos       As Long
    Dim Obj     As Control
    Dim scaleX     As Double, scaleY       As Double
    scaleX = FormName.ScaleWidth / FormOldWidth                                 '���洰�������ű���
    scaleY = FormName.ScaleHeight / FormOldHeight                               '���洰��߶����ű���
    On Error Resume Next
    For Each Obj In FormName
        StartPos = 1
        For I = 0 To 4
            '��ȡ�ؼ���ԭʼλ�����С
            TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
            If TempPos > 0 Then
                Pos(I) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                StartPos = TempPos + 1
            Else
                Pos(I) = 0
            End If
            '���ݿؼ���ԭʼλ�ü�����ı��С�ı����Կؼ����¶�λ��ı��С
            Obj.Move Pos(0) * scaleX, Pos(1) * scaleY, Pos(2) * scaleX, Pos(3) * scaleY
        Next I
    Next Obj
    On Error GoTo 0
End Sub

Private Sub Form_Resize()
    Call ResizeForm(Me)                                                         'ȷ������ı�ʱ�ؼ���֮�ı�
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


