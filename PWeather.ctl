VERSION 5.00
Begin VB.UserControl PWeather 
   BackColor       =   &H00F2AF00&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "PWeather.ctx":0000
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "PWeather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Pica As Picture
Dim Picb As Picture

Public Function GetWeather(Optional Content As Integer = -1) As String
    Dim Weather As String
    Weather = GetWeatherInfo(List1, Pica, Picb)
    If Content = -1 Then
        GetWeather = Weather
    ElseIf Content > -1 And Content < 7 Then
        GetWeather = List1.List(Content)
    End If
End Function

Public Sub GetPicture(ByRef Pic1 As Picture, ByRef Pic2 As Picture)
    Dim Weather As String
    Weather = GetWeatherInfo(List1, Pica, Picb)
    Set Pic1 = Pica
    Set Pic2 = Picb
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 480
    UserControl.Height = 480
End Sub

Private Function GetWeatherInfo(ByRef lst As ListBox, ByRef Pic1 As Picture, ByRef Pic2 As Picture) As String
    On Error GoTo Err
    chengshi = ReadinteFile("http://61.4.185.48:81/g/")
    If InStr(chengshi, "var id=") <> 0 Then
        cityid = ""
        For q = InStr(chengshi, "var id=") + 7 To 100
            If Mid(chengshi, q, 1) <> ";" Then
                cityid = cityid & Mid(chengshi, q, 1)
            Else
                Exit For
            End If
        Next
    End If
    tianqi = Split(ReadinteFile("http://www.weather.com.cn/data/cityinfo/" & cityid & ".html"), Chr(34))
    lst.Clear
    For i = 0 To UBound(tianqi)
        If tianqi(i) <> cityid And tianqi(i) <> "," And tianqi(i) <> ":" And tianqi(i) <> ",:" And tianqi(i) <> "{" And tianqi(i) <> "}}" And tianqi(i) <> ":{" And tianqi(i) <> "weatherinfo" And tianqi(i) <> "city" And tianqi(i) <> "cityid" And tianqi(i) <> "temp1" And tianqi(i) <> "temp2" And tianqi(i) <> "weather" And tianqi(i) <> "img1" And tianqi(i) <> "img2" And tianqi(i) <> "ptime" Then
            lst.AddItem tianqi(i)
        End If
    Next
    Dim sSourceUrl As String
    Dim sLocalFile As String
    sSourceUrl = "http://m.weather.com.cn/img/" & lst.List(4)
    sLocalFile = "c:\PWeather" & lst.List(4)
    If DownloadFile(sSourceUrl, sLocalFile) Then
        Set Pic1 = LoadPicture("c:\PWeather" & lst.List(4))
        Kill "c:\PWeather" & lst.List(4)
    End If
    sSourceUrl = "http://m.weather.com.cn/img/" & lst.List(5)
    sLocalFile = "c:\PWeather" & lst.List(5)
    If DownloadFile(sSourceUrl, sLocalFile) Then
        Set Pic2 = LoadPicture("c:\PWeather" & lst.List(5))
        Kill "c:\PWeather" & lst.List(5)
    End If
    xingqi = Weekday(Now)
    Select Case xingqi
    Case "1"
        xingqi = "星期一"
    Case "2"
        xingqi = "星期二"
    Case "3"
        xingqi = "星期三"
    Case "4"
        xingqi = "星期四"
    Case "5"
        xingqi = "星期五"
    Case "6"
        xingqi = "星期六"
    Case "7"
        xingqi = "星期日"
    End Select
    GetWeatherInfo = Month(Now) & "月" & Day(Now) & "日" & " " & xingqi & " " & lst.List(0) & " " & lst.List(3) & " " & lst.List(2) & "~" & lst.List(1) & " 更新时间:" & lst.List(6)
    Exit Function
Err:
    GetWeatherInfo = "Error"
End Function
