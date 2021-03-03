Attribute VB_Name = "Functions"
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Const ERROR_SUCCESS As Long = 0
Const BINDF_GETNEWESTVERSION As Long = &H10
Const INTERNET_FLAG_RELOAD As Long = &H80000000
Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal HDCSRC As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

' ******************** 窗体透明 ********************
'***Module.bas
'****
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
 
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1
Public Const WS_EX_TRANSPARENT = &H20&
Public Const AC_SRC_OVER = &H0
'**************
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Tishim As Long



Public Function GetTaskbarHeight() As Integer
    Dim rectVal As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, rectVal, 0
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Public Function 判断鼠标是否指向指定控件上(hwn As Long) As Boolean
    Dim NowPOINT As POINTAPI
    Dim thWnd As Long
    GetCursorPos NowPOINT
    thWnd = WindowFromPoint(NowPOINT.x, NowPOINT.y)
    If hwn = thWnd Then 判断鼠标是否指向指定控件上 = True Else: 判断鼠标是否指向指定控件上 = False
End Function

Public Function DownloadFile(sSourceUrl As String, sLocalFile As String) As Boolean
    DownloadFile = URLDownloadToFile(0&, sSourceUrl, sLocalFile, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
End Function

Public Function ReadinteFile(ByVal sUrl As String) As String
    Dim xmlHTTP1 As Object
    Set xmlHTTP1 = CreateObject("Microsoft.XMLHTTP")
    xmlHTTP1.Open "get", sUrl, True
    xmlHTTP1.send
    While xmlHTTP1.ReadyState <> 4
        DoEvents
    Wend
    ReadinteFile = xmlHTTP1.responseText
    Set xmlHTTP1 = Nothing
End Function

'此代码可作为一种数字的加密手段
'如果要加密字符，可以将字符转为ASCII码再加密
Public Function Hex10_to_300(ByVal Num As Long) As String                       '加密
    Dim Neg As Boolean                                                          '负数标记
    Dim s As String                                                             '用来加/解密的汉字
    s = "叶古右占号叮可叵卟只叭史兄叽句叱台叹叼司叫叩叨叻另召吁吓吐吉吏吋吕同吊合吃向舌吒后各名吖吸吆吗呈吴吞呒呓杏呆吾吱呔吠呋呕吰否吥呖呃吨呀吵员呗呐呙吟含谷吩呛吽告听吹吻吪呜呌吝吭吣吢启吡吮君吚猫吧邑吼吲吷呿味咑哎咕呵咂呸咙咔呫咀呻呷咒呪呥鼠命呤呼咋知咊和咐呱咚咎鸣周咆咛呝咏呢咄咈亟呶咖呣咍呦咝哐哇咭哑哉哄哶哂咵咸咡咥咧咦哓哒咴哔呲呰虽品咽哕咣咶哙哈啕咮咻哌哗咱咿响哅咯哆哚咤咬哜哀咨咳咲咩咪哝哞哏哪哟唗唘唝哧唓哮哢唛唠哺哽哥唔唡唇哲哳成浩哯哨唢鸡哩哭唏唅唑哦唕唣唤唦唁哼唐哿唆唉唚唧啊唪啧啈唵喏唶喵啉啢啦啪啄啑唻啭啡喎啃啮唬唱啯啰念啥唫唾啝唯售啤啁死啖唿啵商啐唷唳啶啘唼兽啴所啷鹏啸唰狗喆喷喜喓喋嗒喃嗏"
    Neg = False                                                                 '负数标记
    If Num < 0 Then                                                             '判断正负
        Neg = True                                                              '标记为负数
        Num = -Num                                                              '按正数加密
    End If
    Do
        Hex10_to_300 = Mid(s, Num Mod 300 + 1, 1) & Hex10_to_300                '转换进制
        Num = Num \ 300
    Loop Until Num = 0
    If Neg = True Then Hex10_to_300 = "草" & Hex10_to_300                       '如果是负数就加前缀“草”
End Function

Public Function Hex300_to_10(ByVal Num As String) As Long                       '解密
    Dim Neg As Boolean                                                          '负数标记
    Dim i As Integer                                                            '解密的进度
    Dim s As String                                                             '用来加/解密的汉字
    s = "叶古右占号叮可叵卟只叭史兄叽句叱台叹叼司叫叩叨叻另召吁吓吐吉吏吋吕同吊合吃向舌吒后各名吖吸吆吗呈吴吞呒呓杏呆吾吱呔吠呋呕吰否吥呖呃吨呀吵员呗呐呙吟含谷吩呛吽告听吹吻吪呜呌吝吭吣吢启吡吮君吚猫吧邑吼吲吷呿味咑哎咕呵咂呸咙咔呫咀呻呷咒呪呥鼠命呤呼咋知咊和咐呱咚咎鸣周咆咛呝咏呢咄咈亟呶咖呣咍呦咝哐哇咭哑哉哄哶哂咵咸咡咥咧咦哓哒咴哔呲呰虽品咽哕咣咶哙哈啕咮咻哌哗咱咿响哅咯哆哚咤咬哜哀咨咳咲咩咪哝哞哏哪哟唗唘唝哧唓哮哢唛唠哺哽哥唔唡唇哲哳成浩哯哨唢鸡哩哭唏唅唑哦唕唣唤唦唁哼唐哿唆唉唚唧啊唪啧啈唵喏唶喵啉啢啦啪啄啑唻啭啡喎啃啮唬唱啯啰念啥唫唾啝唯售啤啁死啖唿啵商啐唷唳啶啘唼兽啴所啷鹏啸唰狗喆喷喜喓喋嗒喃嗏"
    Neg = False                                                                 '负数标记
    If Left(Num, 1) = "草" Then                                                 '判断正负
        Neg = True                                                              '标记为负数
        Num = Replace(Num, "草", "")                                            '按正数解密
    End If
    For i = 1 To Len(Num)                                                       '转换进制
        Hex300_to_10 = Hex300_to_10 * 300 + (InStr(s, Mid(Num, i, 1)) - 1)
    Next
    If Neg = True Then Hex300_to_10 = -Hex300_to_10                             '如果是负数就取结果的相反数
End Function

Public Function Encrypt(ByVal strText As String) As String
    If strText = "" Then Exit Function
    Encrypt = Hex10_to_300(Asc(Mid(strText, 1, 1)))
    Dim i As Long
    For i = 2 To Len(strText)
        Encrypt = Encrypt & "与" & Hex10_to_300(Asc(Mid(strText, i, 1)))
    Next
End Function

Public Function Declassified(ByVal strText As String) As String
    If strText = "" Then Exit Function
    Dim s() As String, i As Long
    s = Split(strText, "与")
    For i = 0 To UBound(s)
        Declassified = Declassified & Chr(Hex300_to_10(s(i)))
    Next
End Function

Public Function Tishi(ByVal Titou As String, Msg As String)
    If Titou = "" Then Exit Function
    If Msg = "" Then Exit Function
    
    If Tishiback = 5 Then
        Form5.Enabled = False
    ElseIf Tishiback = 7 Then
        Form7.Enabled = False
    ElseIf Tishiback = 8 Then
        Form8.Enabled = False
    ElseIf Tishiback = 9 Then
        Form9.Enabled = False
    ElseIf Tishiback = 23 Then
        Form23.Enabled = False
    ElseIf Tishiback = 24 Then
        Form24.Enabled = False
    ElseIf Tishiback = 25 Then
        Form25.Enabled = False
    ElseIf Tishiback = 26 Then
        Form26.Enabled = False
    ElseIf Tishiback = 27 Then
        Form27.Enabled = False
    ElseIf Tishiback = 28 Then
        Form28.Enabled = False
    ElseIf Tishiback = 29 Then
        Form29.Enabled = False
    ElseIf Tishiback = 44 Then
        Form44.Enabled = False
    End If
    Form34.Show
    Form34.Label2 = Msg
    Form34.Label1 = Titou
End Function

