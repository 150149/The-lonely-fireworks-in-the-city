VERSION 5.00
Begin VB.Form Form36 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form36"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "���＼�ܵ�ѡ�����.frx":0000
   LinkTopic       =   "Form36"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '��Ļ����
   Begin �о��й°����̻�.PProgressBar PProgressBar1 
      Height          =   495
      Left            =   9960
      TabIndex        =   18
      Top             =   3360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin �о��й°����̻�.PButton PButton9 
      Height          =   615
      Left            =   7200
      TabIndex        =   12
      Top             =   6720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Text            =   "ȷ��"
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "0"
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "0"
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   9
      Text            =   "0"
      Top             =   4080
      Width           =   495
   End
   Begin �о��й°����̻�.PButton PButton2 
      Height          =   615
      Left            =   6840
      TabIndex        =   2
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   3360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "0"
      Top             =   3360
      Width           =   495
   End
   Begin �о��й°����̻�.PButton PButton3 
      Height          =   615
      Left            =   5520
      TabIndex        =   3
      Top             =   4080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   12632256
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton4 
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton5 
      Height          =   615
      Left            =   5520
      TabIndex        =   5
      Top             =   5520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   16777215
      Color_Back_Down =   8421504
      Color_Begin     =   16777215
      Color_End       =   8421504
      Color_Text      =   255
      Color_Text_MouseMoved=   128
      Text            =   "-"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton6 
      Height          =   615
      Left            =   6840
      TabIndex        =   6
      Top             =   4080
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton7 
      Height          =   615
      Left            =   6840
      TabIndex        =   7
      Top             =   4800
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PButton PButton8 
      Height          =   615
      Left            =   6840
      TabIndex        =   8
      Top             =   5520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text      =   49152
      Color_Text_MouseMoved=   16384
      Text            =   "+"
      Font_Size       =   20
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   70
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar2 
      Height          =   495
      Left            =   9960
      TabIndex        =   19
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar3 
      Height          =   495
      Left            =   9960
      TabIndex        =   20
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin �о��й°����̻�.PProgressBar PProgressBar4 
      Height          =   495
      Left            =   9960
      TabIndex        =   21
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      Color_Back      =   14737632
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   17
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "�۲���"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   16
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   15
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "�罻����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7560
      TabIndex        =   14
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "�ɷ���ļ��ܵ���"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   6240
      Width           =   2415
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Point1 As Integer
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Point1 = 20
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
    Label2.Caption = "�罻������" & Text1.Text
    Label3.Caption = "������" & Text2.Text
    Label4.Caption = "�۲�����" & Text3.Text
    Label5.Caption = "������" & Text4.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    
    If Screen.Width / Screen.TwipsPerPixelX = 1920 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1920.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1600 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1600.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1440 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1440.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1366 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1366.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1360 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1360.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1280 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1280.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1152 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1152.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1024 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1024.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 800 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\800.jpg")
    Else
        Me.Picture = LoadPicture(App.Path & "\picture\shape\1600.jpg")
    End If
End Sub

Private Sub PButton1_Click()
    If CInt(Text1.Text) <= 0 Then Exit Sub
    Text1.Text = CStr(CInt(Text1.Text) - 1)
    Point1 = Point1 + 1
    Label2.Caption = "�罻������" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton2_Click()
    If Point1 = 0 Then Exit Sub
    Text1.Text = CStr(CInt(Text1.Text) + 1)
    Point1 = Point1 - 1
    Label2.Caption = "�罻������" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton3_Click()
    If CInt(Text2.Text) <= 0 Then Exit Sub
    Text2.Text = CStr(CInt(Text2.Text) - 1)
    Point1 = Point1 + 1
    Label3.Caption = "������" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton4_Click()
    If CInt(Text3.Text) <= 0 Then Exit Sub
    Text3.Text = CStr(CInt(Text3.Text) - 1)
    Point1 = Point1 + 1
    Label4.Caption = "�۲�����" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton5_Click()
    If CInt(Text4.Text) <= 0 Then Exit Sub
    Text4.Text = CStr(CInt(Text4.Text) - 1)
    Point1 = Point1 + 1
    Label5.Caption = "������" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton6_Click()
    If Point1 = 0 Then Exit Sub
    Text2.Text = CStr(CInt(Text2.Text) + 1)
    Point1 = Point1 - 1
    Label3.Caption = "������" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton7_Click()
    If Point1 = 0 Then Exit Sub
    Text3.Text = CStr(CInt(Text3.Text) + 1)
    Point1 = Point1 - 1
    Label4.Caption = "�۲�����" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton8_Click()
    If Point1 = 0 Then Exit Sub
    Text4.Text = CStr(CInt(Text4.Text) + 1)
    Point1 = Point1 - 1
    Label5.Caption = "������" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub PButton9_Click()
    Form1.Shejiao = CInt(Text1.Text)
    Form1.Zhili = CInt(Text2.Text)
    Form1.Guanchali = CInt(Text3.Text)
    Form1.Naili = CInt(Text4.Text)
    If Form1.Shejiao < 6 Then Form1.Shejiao = 6
    If Form1.Guanchali < 2 Then Form1.Guanchali = 2
    If Form1.Naili < 2 Then Form1.Naili = 2
    Form20.Show
End Sub

Private Sub Text1_Change()
    If IsNumeric(Text1.Text) = False And Not Text1.Text = "" Then Text1.Text = "0"
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text = "" Then Text1.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text1.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label2.Caption = "�罻������" & Text1.Text
    PProgressBar1.Value = CSng(CInt(Text1.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub Text2_Change()
    If IsNumeric(Text2.Text) = False And Not Text2.Text = "" Then Text2.Text = "0"
End Sub

Private Sub Text3_Change()
    If IsNumeric(Text3.Text) = False And Not Text3.Text = "" Then Text3.Text = "0"
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4.Text) = False And Not Text4.Text = "" Then Text4.Text = "0"
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text = "" Then Text2.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text2.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label3.Caption = "������" & Text2.Text
    PProgressBar2.Value = CSng(CInt(Text2.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub Text3_LostFocus()
    If Text3.Text = "" Then Text3.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text3.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label4.Caption = "�۲�����" & Text3.Text
    PProgressBar3.Value = CSng(CInt(Text3.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
End Sub

Private Sub Text4_LostFocus()
    If Text4.Text = "" Then Text4.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    If Point1 < 0 Then Text4.Text = "0"
    Point1 = 20 - CInt(Text1.Text) - CInt(Text2.Text) - CInt(Text3.Text) - CInt(Text4.Text)
    Label5.Caption = "������" & Text4.Text
    PProgressBar4.Value = CSng(CInt(Text4.Text) / 20)
    Label1.Caption = "�ɷ���ļ��ܵ�����" & Point1
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
