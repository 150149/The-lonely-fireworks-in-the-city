VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   0  'None
   Caption         =   "Form11"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   3120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1085
      Text            =   "确定"
      Can_Text_Move   =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "你贷款购买了家具，金钱-1000"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Private Sub Form_Unload(Cancel As Integer)
    Unload Form11
End Sub

Private Sub PButton1_Click()
    Form1.money = Form1.money - 1000
    Unload Form12
    Unload Form11
End Sub
