VERSION 5.00
Begin VB.Form Form52 
   BorderStyle     =   0  'None
   Caption         =   "Form52"
   ClientHeight    =   9900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19320
   Icon            =   "关于界面.frx":0000
   LinkTopic       =   "Form32"
   ScaleHeight     =   9900
   ScaleWidth      =   19320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin 市井中孤傲的烟火.PButton PButton1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      Color_Back      =   14737632
      Color_Back_Down =   8421504
      Color_Begin     =   14737632
      Color_End       =   8421504
      Color_Text_MouseMoved=   4210752
      Text            =   "返回"
      Style_Border    =   0
      Can_Text_Move   =   0   'False
      Color_Back_TransparentDegree=   90
   End
End
Attribute VB_Name = "Form52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Form52.Height = Screen.Height
    Form52.Width = Screen.Width
    
    If Screen.Width / Screen.TwipsPerPixelX = 1920 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1920.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1600 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1600.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1440 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1440.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1366 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1366.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1360 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1360.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1280 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1280.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1152 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1152.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 1024 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1024.jpg")
    ElseIf Screen.Width / Screen.TwipsPerPixelX = 800 Then
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\800.jpg")
    Else
        Me.Picture = LoadPicture(App.Path & "\picture\shape\about\1600.jpg")
    End If
    
End Sub

Private Sub PButton1_Click()
    Form54.Show
End Sub
