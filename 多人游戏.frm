VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form9 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form9"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "������Ϸ.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   10380
   ScaleWidth      =   16995
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer Timer6 
      Left            =   120
      Top             =   2400
   End
   Begin �о��й°����̻�.PButton PButton4 
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   6960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   16777215
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "ע��"
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   50
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   7440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Timer Timer5 
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer4 
      Left            =   120
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   255
      ExtentX         =   450
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
      Location        =   ""
   End
   Begin �о��й°����̻�.PButton PButton3 
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   6360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   16777215
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "����"
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   50
   End
   Begin �о��й°����̻�.PButton PButton2 
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   4920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Color_Back      =   16777215
      Color_Back_Down =   8421504
      Color_Begin     =   16777215
      Color_End       =   14737632
      Color_Text      =   16777215
      Color_Text_MouseMoved=   14737632
      Text            =   "û���˺ţ���ע��"
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   100
   End
   Begin �о��й°����̻�.PButton PButton1 
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Color_Back      =   14737632
      Color_Back_Down =   12632256
      Color_Begin     =   14737632
      Color_End       =   12632256
      Text            =   "��¼"
      Color_Border    =   16777215
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   50
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�˺�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   4815
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   7335
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private page As Integer
Private szMyText, szMyText2 As String
Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To 256) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHostname As String) As Long
Private Const WS_VERSION_REQD = &H101
   
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Const WS_EX_LAYERED = &H80000
Const GWL_EXSTYLE = (-20)
Const LWA_ALPHA = &H2
'Const LWA_COLORKEY = &H1
Private ib As Integer
Private zh As Integer
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private FormOldWidth     As Long                                                '���洰���ԭʼ���
Private FormOldHeight     As Long                                               '���洰���ԭʼ�߶�
Private zc As Integer

Private Sub Form_Load()
    Call ResizeInit(Me)                                                         '�ڳ���װ��ʱ�������
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    If IsConnectedState Then
        WebBrowser1.Navigate "http://showtxt.cn/kbrt3"
        page = 1
    Else
    End If
    WebBrowser1.Silent = True
    
    Timer4.Enabled = False
    Timer5.Enabled = False
    
    
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
    PButton4.Visible = False
End Sub


Public Function IsConnectedState() As Boolean
    Dim udtWSAD As WSADATA
    Call WSAStartup(WS_VERSION_REQD, udtWSAD)
    IsConnectedState = CBool(gethostbyname("www.baidu.com"))
    Call WSACleanup
End Function





Private Sub PButton1_Click()
    If Len(Text1.Text) < 5 Then
        Tishim = Tishi("��ʾ", "���ֳ���С��5")
        Form3.Tishiback = 9
        Exit Sub
    End If
    If Len(Text1.Text) < 5 Then
        Tishim = Tishi("��ʾ", "���볤��С��5")
        Form3.Tishiback = 9
        Exit Sub
    End If
    If Len(Text1.Text) >= 5 And Len(Text2.Text) >= 5 Then
        If IsConnectedState Then
            Timer1.Enabled = False
            Timer3.Enabled = False
            WebBrowser1.Navigate "http://showtxt.cn/kbrt4"
            page = 3
        ElseIf IsConnectedState = False Then
            Tishim = Tishi("��ʾ", "����δ����")
            Form3.Tishiback = 9
        End If
    End If
End Sub

Private Sub PButton2_Click()
    If zc = 0 Then
        PButton4.Visible = True
        PButton4.Top = PButton1.Top
        PButton1.Visible = False
        zc = 1
        PButton2.Text = "�����˺ţ����¼"
    ElseIf zc = 1 Then
        PButton4.Visible = False
        PButton1.Visible = True
        PButton2.Text = "û���˺ţ���ע��"
        zc = 0
    End If
End Sub

Private Sub PButton3_Click()
    Form40.Show
End Sub

Private Sub PButton4_Click()
    If Len(Text1.Text) < 5 Then
        Tishim = Tishi("��ʾ", "���ֳ���С��5")
        Form3.Tishiback = 9
        Exit Sub
    End If
    If Len(Text1.Text) < 5 Then
        Tishim = Tishi("��ʾ", "���볤��С��5")
        Form3.Tishiback = 9
        Exit Sub
    End If
    If IsConnectedState Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        WebBrowser1.Navigate "http://showtxt.cn/kbrt4"
        page = 5
    Else
        MsgBox "����δ����"
    End If
End Sub

Private Sub Timer1_Timer()
    If WebBrowser1.busy Then
    Else
        Timer1.Enabled = False
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "��ʼ"                                                 '����Ҫ���ҵ��ַ�����ͷ
        szFindStrEnd = "����"                                                   '����Ҫ���ҵ��ַ�����β
        
        szText = WebBrowser1.Document.body.innerText                            '�õ��������֣���ʱ��ģ�壬ʵ��ʹ���л���ȥWebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '�ҿ�ͷ�ַ���
        If nBegin > 0 Then                                                      '���������ҵ���ͷ�˲ż���
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '�ҽ�β�ַ���
            If nEnd > nBegin Then                                               '��β����ȿ�ͷ��λ�ô�
                
                '�������ҵ��ַ���ģʽ��ע�͵������2��
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '������Ҫ��ȡ���ַ�������,���Ҫ�������ҵ��ַ�������1�У�ע������2��
                
                '���������ҵ��ַ���ģʽ
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '������������ҵ��ַ���������2��
                nBegin = nBegin + Len(szFindStrBegin)                           '������������ҵ��ַ���������2��
                
                szMyText = Mid(szText, nBegin, nLength)                         'ȡ����before then.���� "test" �м�Ķ���
                Debug.Print ("��ȡ���ݣ�" & szMyText)
                
                Dim jiesuan As String
                Dim mAES As New mAES                                            ''AES���������
                jiesuan = mAES.DecryptStr(szMyText, "150149AE|set", Bit128, Bit128, False)
                Debug.Print "���㣺" & jiesuan
                Label1.Caption = Mid(jiesuan, InStr(jiesuan, "����:"), InStr(jiesuan, ":����") - InStr(jiesuan, "����:"))
                jiesuan = ""
            End If
        End If
        
        
    End If
End Sub

Private Sub Timer2_Timer()
    If WebBrowser1.busy Then
    Else
        Timer2.Enabled = False
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "��ʼ"                                                 '����Ҫ���ҵ��ַ�����ͷ
        szFindStrEnd = "����"                                                   '����Ҫ���ҵ��ַ�����β
        
        szText = WebBrowser1.Document.body.innerText                            '�õ��������֣���ʱ��ģ�壬ʵ��ʹ���л���ȥWebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '�ҿ�ͷ�ַ���
        If nBegin > 0 Then                                                      '���������ҵ���ͷ�˲ż���
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '�ҽ�β�ַ���
            If nEnd > nBegin Then                                               '��β����ȿ�ͷ��λ�ô�
                
                '�������ҵ��ַ���ģʽ��ע�͵������2��
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '������Ҫ��ȡ���ַ�������,���Ҫ�������ҵ��ַ�������1�У�ע������2��
                
                '���������ҵ��ַ���ģʽ
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '������������ҵ��ַ���������2��
                nBegin = nBegin + Len(szFindStrBegin)                           '������������ҵ��ַ���������2��
                
                szMyText = Mid(szText, nBegin, nLength)                         'ȡ����before then.���� "test" �м�Ķ���
                Debug.Print ("��ȡ���ݣ�" & szMyText)
                
                '-------------------------------�˺�����ȶ�----------------------------------------
                Randomize
                Dim cv As String
                cv = Int(Rnd * (10000 - 1 + 1)) + 1
                Open App.Path & "\" & cv For Output As #1
                Print #1, szMyText
                Close #1
                '------------------------------д����ʱ�ļ�-----------------------------------------
                '------------------------------�����ַ����ȶԽ��-----------------------------------
                Dim read_OK As Long
                Dim max As String
                max = String(10, 0)
                read_OK = GetPrivateProfileString("S", "max", "1", max, 256, App.Path & "\" & cv)
                Dim m1 As Integer
                Dim m2 As String
                Dim m3 As Integer
                Dim m4 As Integer
                Dim m5 As String
                Dim m6 As String
                Dim m7 As String
                Dim mAES As New mAES                                            ''AES���������
                m2 = mAES.EncryptStr(Text1.Text, "150149", Bit128, Bit128, False)
                Debug.Print "�����˺ţ�" & m2
                m7 = String(256, 0)
                For m1 = 1 To CInt(max)
                    read_OK = GetPrivateProfileString("A", CStr(m1), "", m7, 256, App.Path & "\" & cv)
                    If Replace(m7, Chr(0), "") = m2 Then
                        m3 = m1
                        Debug.Print "�˶����ݿ�ID��" & m3
                        Debug.Print "�˶Լ��ܺ��˻���" & Replace(m7, Chr(0), "")
                    Else
                        m4 = m4 + 1
                        Debug.Print "�ȶԴ���m4��" & m4
                    End If
                Next
                If m4 = CInt(max) And CInt(max) > 1 Then
                    Tishim = Tishi("��ʾ", "���û�δע��")
                    Form3.Tishiback = 9
                    Kill App.Path & "\" & cv
                Else
                    m5 = String(256, 0)
                    read_OK = GetPrivateProfileString("B", CStr(m3), "", m5, 256, App.Path & "\" & cv)
                    Kill App.Path & "\" & cv
                    Debug.Print "��ȡ��Ӧ�����봮��" & m5
                    m6 = mAES.EncryptStr(Text2.Text, "150149AE|set", Bit128, Bit128, False)
                    Debug.Print "�������봮��" & m6
                    If m6 = Replace(m5, Chr(0), "") Then
                        MsgBox "��¼�ɹ�,��ӭ��" & Text1.Text
                        Form30.mingzi = Text1.Text
                        m6 = ""
                        m5 = ""
                        Form30.Show
                        Unload Form9
                    Else
                        Tishim = Tishi("��ʾ", "��¼ʧ�ܣ����벻��ȷ")
                        Form3.Tishiback = 9
                        m6 = ""
                        m5 = ""
                    End If
                    m4 = 0
                    m3 = 0
                    m2 = ""
                    Text2.Text = ""
                End If
                '----------------------------------------------------------------------------------
            End If
            
            
        End If
    End If
End Sub

Private Sub Timer3_Timer()
    If WebBrowser1.busy Then
    Else
        Timer3.Enabled = False
        Dim szText As String
        Dim szFindStrBegin As String
        Dim szFindStrEnd As String
        Dim nBegin As Long
        Dim nEnd As Long
        Dim nLength  As Long
        szFindStrBegin = "��ʼ"                                                 '����Ҫ���ҵ��ַ�����ͷ
        szFindStrEnd = "����"                                                   '����Ҫ���ҵ��ַ�����β
        
        szText = WebBrowser1.Document.body.innerText                            '�õ��������֣���ʱ��ģ�壬ʵ��ʹ���л���ȥWebBrowser1.Document.body.innerText
        
        nBegin = InStr(szText, szFindStrBegin)                                  '�ҿ�ͷ�ַ���
        If nBegin > 0 Then                                                      '���������ҵ���ͷ�˲ż���
            nEnd = InStr(nBegin, szText, szFindStrEnd)                          '�ҽ�β�ַ���
            If nEnd > nBegin Then                                               '��β����ȿ�ͷ��λ�ô�
                
                '�������ҵ��ַ���ģʽ��ע�͵������2��
                'nLength = nEnd - nBegin + Len(szFindStrEnd)   '������Ҫ��ȡ���ַ�������,���Ҫ�������ҵ��ַ�������1�У�ע������2��
                
                '���������ҵ��ַ���ģʽ
                nLength = nEnd - nBegin - Len(szFindStrBegin)                   '������������ҵ��ַ���������2��
                nBegin = nBegin + Len(szFindStrBegin)                           '������������ҵ��ַ���������2��
                
                szMyText = Mid(szText, nBegin, nLength)                         'ȡ����before then.���� "test" �м�Ķ���
                Debug.Print ("��ȡ���ݣ�" & szMyText)
                '------------------------------д����ʱ�ļ�----------------------------------------
                Randomize
                Dim cv As String
                cv = Int(Rnd * (10000 - 1 + 1)) + 1
                Open App.Path & "\" & cv For Output As #1
                Print #1, szMyText
                Close #1
                '------------------------------д����ʱ�ļ�-----------------------------------------
                Dim read_OK As Long
                Dim max As String
                max = String(10, 0)
                read_OK = GetPrivateProfileString("S", "max", "1", max, 256, App.Path & "\" & cv)
                '--------------------------------����Ƿ���ע��-------------------------------------
                Dim m1 As Integer
                Dim m7 As String
                Dim m2 As String
                Dim mAES As New mAES                                            ''AES���������
                m7 = String(256, 0)
                m2 = mAES.EncryptStr(Text1.Text, "150149", Bit128, Bit128, False)
                For m1 = 1 To CInt(max)
                    read_OK = GetPrivateProfileString("A", CStr(m1), "", m7, 256, App.Path & "\" & cv)
                    If Replace(m7, Chr(0), "") = m2 Then
                        Debug.Print "�˶����ݿ�ID��" & m1
                        Tishim = Tishi("��ʾ", "���û���ע��")
                        Form3.Tishiback = 9
                        Exit Sub
                    End If
                Next
                
                Dim max1 As Integer
                max1 = CInt(max) + 1
                
                Dim n2 As String
                Dim n3 As String
                n2 = mAES.EncryptStr(Text1.Text, "150149", Bit128, Bit128, False)
                n3 = mAES.EncryptStr(Text2.Text, "150149AE|set", Bit128, Bit128, False)
                Dim write1 As Long
                write1 = WritePrivateProfileString("S", "max", CStr(max1), App.Path & "\" & cv)
                write1 = WritePrivateProfileString("A", CStr(max1), n2, App.Path & "\" & cv)
                write1 = WritePrivateProfileString("B", CStr(max1), n3, App.Path & "\" & cv)
                n2 = ""
                n3 = ""
                
                
                If IsConnectedState Then                                        '�����������
                    '-------------------------------------------��������-------------------------------------------------
                    Dim fasongneirong As String
                    Open App.Path & "\" & cv For Binary As #1
                    fasongneirong = StrConv(InputB(LOF(1), 1), vbUnicode)
                    Close #1
                    Kill App.Path & "\" & cv
                    
                    Dim vDoc, VTag, mType As String, mTagName As String
                    Dim ia As Integer
                    Set vDoc = WebBrowser1.Document
                    For ia = 0 To vDoc.All.Length - 1
                        Select Case UCase(vDoc.All(ia).tagName)
                        Case "TEXTAREA"                                         '"TEXTAREA" ��ǩ,�ı������д
                            Set VTag = vDoc.All(ia)
                            VTag.Value = "��ʼ" & fasongneirong & "����"        '��Text1�е���������
                            Debug.Print ("�������ݣ�" & "��ʼ" & fasongneironng & "����")
                            Timer2.Enabled = True
                            Timer2.Interval = 300
                            MsgBox "ע��ɹ�,����Ϊ����¼"
                            Exit Sub
                            Timer3.Enabled = False
                        End Select
                        
                    Next ia
                ElseIf IsConnectedState = False Then                            '�����������
                    Tishim = Tishi("��ʾ", "����δ����")
                    Form3.Tishiback = 9
                End If
                
            End If
        End If
        
        
    End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim vDoc, X, VTag
    If (pDisp Is WebBrowser1.Object) And page = 1 Then                          '����ҳ�������ִ����������
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '������б�ǩ
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Value = "�ύ" Then
                    VTag.Click                                                  '����ύ�ˣ�һ�ж�OK��
                    page = 2
                    Timer1.Enabled = True
                    Timer1.Interval = 300
                End If
            End If
        Next X
    ElseIf page = 3 And URL = "http://showtxt.cn/kbrt4" Then
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '������б�ǩ
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Value = "�ύ" Then
                    VTag.Click                                                  '����ύ�ˣ�һ�ж�OK��
                    page = 4
                    Timer2.Enabled = True
                    Timer2.Interval = 300
                End If
            End If
        Next X
    ElseIf page = 5 And URL = "http://showtxt.cn/kbrt4" Then
        WebBrowser1.Document.getElementsByTagName("input")("submit_pw").Value = "189159"
        
        Set vDoc = WebBrowser1.Document
        For X = 0 To vDoc.All.Length - 1                                        '������б�ǩ
            If UCase(vDoc.All(X).tagName) = "INPUT" Then                        '�ҵ�input��ǩ
                Set VTag = vDoc.All(X)
                If VTag.Value = "�ύ" Then
                    VTag.Click                                                  '����ύ�ˣ�һ�ж�OK��
                    page = 6
                    Timer3.Enabled = True
                    Timer3.Interval = 300
                End If
            End If
        Next X
    End If
End Sub



Private Sub Timer4_Timer()
    ib = ib + 5
    SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, ib, LWA_ALPHA                        '150Ϊ͸����(0-255)
    If ib = 255 Then Timer4.Enabled = False
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
    Me.WebBrowser1.Left = 0
    Me.WebBrowser1.Top = 0
    Me.WebBrowser1.Width = 0
    Me.WebBrowser1.Height = 0
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

