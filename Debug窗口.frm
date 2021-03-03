VERSION 5.00
Begin VB.Form Form32 
   Caption         =   "Form32"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10395
   LinkTopic       =   "Form32"
   ScaleHeight     =   7650
   ScaleWidth      =   10395
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   7455
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      Text            =   "Debug窗口.frx":0000
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Label12"
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Label15"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Private Sub Form_Load()
    '    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
    
End Sub
