VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "P�ؼ��� �汾5 ������ķ���ʵ����������Ч��"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   7455
   StartUpPosition =   2  '��Ļ����
   Begin P�ؼ���.PWin8Form PWin8Form1 
      Height          =   7800
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   13758
      Icon            =   "Form1.frx":000C
      Caption         =   "P�ؼ��� �汾5 ������ķ���ʵ����������Ч��"
      Color_Border    =   6181967
      Color_Frame     =   5326916
      Can_Move_Smoothly=   0   'False
      Has_MaxButton   =   0   'False
      Is_Resizable    =   0   'False
      Begin P�ؼ���.PTab PTab1 
         Height          =   7200
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   12700
         Color_Back      =   5326916
         Color_Text      =   14273737
         Distance_Transverse=   135
         Distance_Vertical=   120
         Color_Selector  =   8421504
         Texts           =   "��ӭ|��ť|��ѡ��|����|������|�������|��������|ͼƬ��|ƽ��ֱ������ϵ|��ʱͨѶ|�б��|��ѧ����|������Ļ|��ҳ��ǩ|����|����|����"
         Height_Selector =   45
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   6
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   53
            Tag             =   "��������"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PVScrollBar PVScrollBar1 
               Height          =   4815
               Left            =   3240
               TabIndex        =   54
               Top             =   480
               Width           =   375
               _ExtentX        =   661
               _ExtentY        =   8493
               Color_Top       =   6181967
               Color_Back      =   6118749
               Style_Border    =   2
               Color_Border    =   5326916
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   14
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   50
            Tag             =   "��Win8����"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PWeather PW 
               Left            =   3960
               Top             =   2160
               _ExtentX        =   847
               _ExtentY        =   847
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ȼ�ֿ������ˣ��ۣ�"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   5.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   5640
               TabIndex        =   52
               Top             =   5400
               Width           =   1050
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�����棡"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2880
               TabIndex        =   51
               Top             =   1920
               Width           =   1140
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   15
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   48
            Tag             =   "��Win8����"
            Top             =   960
            Width           =   6975
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��������أ�"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2640
               TabIndex        =   49
               Top             =   1920
               Width           =   1710
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   13
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   46
            Tag             =   "��ҳ��ǩ"
            Top             =   960
            Width           =   6975
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "û�����������Ǹ���"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2040
               TabIndex        =   47
               Top             =   1920
               Width           =   2850
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   12
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   43
            Tag             =   "������Ļ"
            Top             =   960
            Width           =   6975
            Begin VB.TextBox Text4 
               BorderStyle     =   0  'None
               Height          =   225
               Left            =   1920
               TabIndex        =   44
               Text            =   "P�ؼ���"
               Top             =   2760
               Width           =   3060
            End
            Begin P�ؼ���.PScreen PScreen1 
               Height          =   885
               Left            =   1920
               TabIndex        =   45
               Top             =   1800
               Width           =   3060
               _ExtentX        =   5398
               _ExtentY        =   1561
               Color_Text      =   255
               Color_Text_Back =   0
               Text            =   "P�ؼ���"
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   11
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   34
            Tag             =   "��������"
            Top             =   960
            Width           =   6975
            Begin VB.TextBox Text3 
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   1440
               TabIndex        =   37
               Top             =   3600
               Width           =   3015
            End
            Begin VB.TextBox Text2 
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   1440
               TabIndex        =   36
               Top             =   2040
               Width           =   4095
            End
            Begin VB.TextBox Text1 
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   1440
               TabIndex        =   35
               Top             =   1680
               Width           =   4095
            End
            Begin P�ؼ���.PMaths PMaths1 
               Left            =   4080
               Top             =   840
               _ExtentX        =   847
               _ExtentY        =   847
            End
            Begin P�ؼ���.PButton PButton6 
               Height          =   375
               Left            =   1440
               TabIndex        =   38
               Top             =   2400
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "+"
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton7 
               Height          =   375
               Left            =   2520
               TabIndex        =   39
               Top             =   2400
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "-"
               Font_Size       =   12
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton8 
               Height          =   375
               Left            =   3600
               TabIndex        =   40
               Top             =   2400
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "*"
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton9 
               Height          =   375
               Left            =   4680
               TabIndex        =   41
               Top             =   2400
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   661
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "/"
               Font_Size       =   9
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton10 
               Height          =   270
               Left            =   4680
               TabIndex        =   42
               Top             =   3600
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   476
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "����"
               Style_Border    =   2
               Color_Border    =   6181967
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   10
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   32
            Tag             =   "�б��"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PListBox PListBox1 
               Height          =   4200
               Left            =   1800
               TabIndex        =   33
               Top             =   720
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   7408
               Color_Back      =   5326916
               Color_Text      =   14273737
               Color_Top_ScrollBar=   6181967
               Color_Back_ScrollBar=   6118749
               Distance_Item   =   0
               Font_Size_Selected=   12
               Color_Back_Selected=   6181967
               Color_Back_Moved=   6118749
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   9
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   26
            Tag             =   "Winsock���ο���"
            Top             =   960
            Width           =   6975
            Begin MSComDlg.CommonDialog cd 
               Left            =   4800
               Top             =   840
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin P�ؼ���.PProgressBar PProgressBar2 
               Height          =   30
               Left            =   1080
               TabIndex        =   27
               Top             =   2760
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   53
               Color_Top       =   6181967
               Color_Back      =   6118749
            End
            Begin P�ؼ���.PButton PButton2 
               Height          =   495
               Left            =   1080
               TabIndex        =   28
               Top             =   2160
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "1������"
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PWinsock PWinsock2 
               Left            =   840
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
            End
            Begin P�ؼ���.PWinsock PWinsock1 
               Left            =   240
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
            End
            Begin P�ؼ���.PButton PButton3 
               Height          =   495
               Left            =   2280
               TabIndex        =   29
               Top             =   2160
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "2������"
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton4 
               Height          =   495
               Left            =   3480
               TabIndex        =   30
               Top             =   2160
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "��������"
               Style_Border    =   2
               Color_Border    =   6181967
            End
            Begin P�ؼ���.PButton PButton5 
               Height          =   495
               Left            =   4680
               TabIndex        =   31
               Top             =   2160
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   873
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Text            =   "�����ļ�"
               Style_Border    =   2
               Color_Border    =   6181967
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   8
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   23
            Tag             =   "ƽ��ֱ������ϵ"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PHScrollBar PHScrollBar4 
               Height          =   135
               Left            =   960
               TabIndex        =   24
               Top             =   5520
               Width           =   5175
               _ExtentX        =   9128
               _ExtentY        =   238
               Color_Top       =   6181967
               Color_Back      =   6118749
               Size            =   0.5
               Style_Border    =   2
               Color_Border    =   5326916
            End
            Begin P�ؼ���.PPRCS PPRCS1 
               Height          =   5175
               Left            =   960
               TabIndex        =   25
               Top             =   240
               Width           =   5175
               _ExtentX        =   9128
               _ExtentY        =   9128
               Color_Back      =   5326916
               Color_Grid      =   6118749
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   7
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   21
            Tag             =   "ͼƬ��"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PPictureBox PPictureBox1 
               Height          =   5655
               Left            =   120
               TabIndex        =   22
               Top             =   120
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   9975
               Picture         =   "Form1.frx":A93DE
               Color_Top       =   6181967
               Color_Back      =   6118749
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   5
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   19
            Tag             =   "�������"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PHScrollBar PHScrollBar3 
               Height          =   375
               Left            =   600
               TabIndex        =   20
               Top             =   2400
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   661
               Color_Top       =   6181967
               Color_Back      =   6118749
               Style_Border    =   2
               Color_Border    =   5326916
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   4
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   17
            Tag             =   "������"
            Top             =   960
            Width           =   6975
            Begin VB.Timer Timer1 
               Interval        =   1
               Left            =   1080
               Top             =   2520
            End
            Begin P�ؼ���.PProgressBar PProgressBar1 
               Height          =   375
               Left            =   600
               TabIndex        =   18
               Top             =   2400
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   661
               Color_Top       =   6181967
               Color_Back      =   6118749
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   3
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   15
            Tag             =   "����"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PSwitch PSwitch1 
               Height          =   1935
               Left            =   1560
               TabIndex        =   16
               Top             =   1800
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   3413
               Color_Top       =   6181967
               Color_Back      =   6118749
               Color_Back_True =   4210752
               Style_Border    =   2
               Color_Border    =   5326916
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   2
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   13
            Tag             =   "��ѡ��"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PCheckBox PCheckBox1 
               Height          =   1335
               Left            =   1320
               TabIndex        =   14
               Top             =   2160
               Width           =   4335
               _ExtentX        =   7646
               _ExtentY        =   2355
               Color_Back      =   6118749
               Color_End       =   6181967
               Text            =   "����D��"
               Font_Size       =   20
               Value           =   -1  'True
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   1
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   10
            Tag             =   "��ť"
            Top             =   960
            Width           =   6975
            Begin P�ؼ���.PHScrollBar PHScrollBar2 
               Height          =   135
               Left            =   960
               TabIndex        =   11
               Top             =   3840
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   238
               Color_Top       =   6181967
               Color_Back      =   6118749
               Value           =   0.1
               Size            =   0.5
               Style_Border    =   2
               Color_Border    =   5326916
            End
            Begin P�ؼ���.PButton PButton1 
               Height          =   2175
               Left            =   960
               TabIndex        =   12
               Top             =   1560
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   3836
               Color_Back      =   6118749
               Color_Back_Down =   6181967
               Color_Begin     =   6118749
               Color_End       =   5326916
               Font_Size       =   20
               Style_Border    =   2
               Color_Border    =   6181967
               Is_Text_Transparent=   -1  'True
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   0
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   7
            Tag             =   "��ӭ"
            Top             =   960
            Width           =   6975
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ķ���ʵ����������Ч��"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   18
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   840
               TabIndex        =   9
               Top             =   3960
               Width           =   5415
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "��ӭ����P�ؼ������磡"
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   36
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1890
               Left            =   1080
               TabIndex        =   8
               Top             =   1080
               Width           =   4875
            End
         End
         Begin VB.PictureBox P 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            Height          =   5895
            Index           =   16
            Left            =   120
            ScaleHeight     =   5835
            ScaleWidth      =   6915
            TabIndex        =   3
            Tag             =   "��ӭ"
            Top             =   960
            Width           =   6975
            Begin VB.TextBox Text5 
               BackColor       =   &H00514844&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3615
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   6
               Text            =   "Form1.frx":ADB99
               Top             =   120
               Width           =   6735
            End
            Begin VB.TextBox Text7 
               Alignment       =   2  'Center
               BackColor       =   &H00514844&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   5
               Text            =   "Form1.frx":ADE8F
               Top             =   4440
               Width           =   6735
            End
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               BackColor       =   &H00514844&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "΢���ź�"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   4
               Text            =   "Form1.frx":ADF14
               Top             =   3900
               Width           =   6735
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   55
            Top             =   6885
            Width           =   60
         End
         Begin VB.Image Image2 
            Height          =   135
            Left            =   600
            Top             =   6900
            Width           =   135
         End
         Begin VB.Image Image1 
            Height          =   135
            Left            =   120
            Top             =   6900
            Width           =   135
         End
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   225
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   397
      _cy             =   423
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    wmp.settings.Volume = 100
    wmp.URL = App.Path & "\����D��.mp3"
    Dim i As Integer
    For i = 0 To P.UBound
        PTab1.BeRelated P(i)
        P(i).BorderStyle = 0
        'P(i).PaintPicture PTab1.Picture, 0, 0, , , P(i).Left, P(i).Top
        P(i).BackColor = PTab1.Color_Back
    Next
    For i = 1 To 100
        PListBox1.AddItem "���ǵ�" & i & "�"
    Next
    PButton1.Refresh
'    PPRCS1.DrawAnyFunction "x^2+2x-3"
'    PPRCS1.DrawAnyFunction "-2x^2+12x-13"
'    PPRCS1.DrawAnyFunction "x+4"
'    PPRCS1.DrawAnyFunction "-3x-1"
'    PPRCS1.DrawLineY -3
'    PPRCS1.DrawLineX 2
    PPRCS1.DrawAnyFunction "sin(x)"
    PPRCS1.DrawAnyFunction "tan(x)"
    PPRCS1.DrawAnyFunction "cos(x)"
    Label4 = PW.GetWeather
    Dim Pic1 As StdPicture, Pic2 As StdPicture
    PW.GetPicture Pic1, Pic2
    Set Image1.Picture = Pic1
    Set Image2.Picture = Pic2
    Text5 = Left(Text5, Len(Text5) - 1)
    Text6 = Left(Text6, Len(Text6) - 1)
    Text7 = Left(Text7, Len(Text7) - 1)
'    PM.ColorSmly RGB(0, 0, 0), RGB(255, 255, 255), 10, 100
End Sub

Private Sub PButton10_Click()
    MsgBox PMaths1.VBCodetoNum(Text3)
End Sub

Private Sub PButton2_Click()
    PWinsock1.Listen
End Sub

Private Sub PButton3_Click()
    PWinsock2.Connect "127.0.0.1"
End Sub

Private Sub PButton4_Click()
    Randomize
    PWinsock1.SendData "������ɵ��������", Rnd
End Sub

Private Sub PButton5_Click()
    On Error GoTo Err
    cd.ShowOpen
    If cd.FileName <> "" Then PWinsock1.SendDocument cd.FileName
    cd.FileName = ""
    Exit Sub
Err:
    Exit Sub
End Sub

Private Sub PButton6_Click()
    MsgBox PMaths1.Add(Text1, Text2)
End Sub

Private Sub PButton7_Click()
    MsgBox PMaths1.Subtract(Text1, Text2)
End Sub

Private Sub PButton8_Click()
    MsgBox PMaths1.Multiply(Text1, Text2)
End Sub

Private Sub PButton9_Click()
    MsgBox PMaths1.Division(Text1, Text2)
End Sub

Private Sub PCheckBox1_ValueChange(NValue As Boolean)
    If NValue = True Then wmp.Controls.play
    If NValue = False Then wmp.Controls.Pause
End Sub

Private Sub PHScrollBar2_Scroll(NValue As Single)
    PButton1.Color_Back_TransparentDegree = Int(NValue * 100)
End Sub

'Private Sub PM_ColorSmlyComplete()
'    If Text5.ForeColor = RGB(0, 0, 0) Then
'        PM.ColorSmly RGB(0, 0, 0), RGB(255, 255, 255), 10, 100
'    Else
'        PM.ColorSmly RGB(255, 255, 255), RGB(0, 0, 0), 10, 100
'    End If
'End Sub

'Private Sub PM_ColorSmlyIng(nColor As Long)
'    Text5.ForeColor = nColor
'    Text6.ForeColor = nColor
'    Text7.ForeColor = nColor
'End Sub

Private Sub PHScrollBar4_Scroll(NValue As Single)
    PPRCS1.Resolution = Int(NValue * 99 + 1)
End Sub

Private Sub PWinsock1_ConnectSucceed()
    MsgBox "1�����ӳɹ�"
End Sub

Private Sub PWinsock1_DataSendSucceed()
    MsgBox "1�����ݷ��ͳɹ�"
End Sub

Private Sub PWinsock1_DocumentSending(Progress As Long, Total As Long)
    On Error Resume Next
    PProgressBar2.Value = Progress / Total
End Sub

Private Sub PWinsock1_DocumentSendSucceed()
    MsgBox "1���ļ����ͳɹ�"
End Sub

Private Sub PWinsock1_ListenBegin()
    MsgBox "������ʼ"
End Sub

Private Sub PWinsock2_ConnectBegin()
    MsgBox "���ӿ�ʼ"
End Sub

Private Sub PWinsock2_ConnectSucceed()
    MsgBox "2�����ӳɹ�"
End Sub

Private Sub PWinsock2_DataArrived(strTouwenjian As String, Shuju As Variant)
    MsgBox "2���յ�1�ŷ��������ݣ�" & strTouwenjian & Shuju
End Sub

Private Sub PWinsock2_DocumentComplete()
    MsgBox "2�Ž����ļ��ɹ�"
End Sub

Private Sub PWinsock2_DocumentQuest()
    On Error GoTo Err
    If MsgBox("1�����ļ������Ƿ���գ�", vbYesNo) = vbYes Then
        cd.ShowSave
        If cd.FileName <> "" Then PWinsock2.AcceptDocument cd.FileName
        cd.FileName = ""
    Else
        PWinsock2.RefuseDocument
    End If
    Exit Sub
Err:
    PWinsock2.RefuseDocument
    Exit Sub
End Sub

Private Sub Text4_Change()
    PScreen1.Text = Text4
End Sub

Private Sub Timer1_Timer()
    If PProgressBar1.Value <> 1 Then
        Randomize
        PProgressBar1.Value = PProgressBar1.Value + 0.005
    Else
        PProgressBar1.Value = 0
    End If
End Sub

Private Sub wmp_StatusChange()
    If wmp.playState = 1 Then wmp.Controls.play
End Sub
