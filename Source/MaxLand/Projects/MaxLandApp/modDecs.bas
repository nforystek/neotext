Attribute VB_Name = "modDecs"
#Const modDecs = -1
Option Explicit
'TOP DOWN


Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type MyScreen
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    clr As Long
    tu As Single
    tv As Single
End Type

Public Type MyVertex
    X As Single
    Y As Single
    Z As Single
    nx As Single
    ny As Single
    nz As Single
    tu As Single
    tv As Single
End Type

Public Type ImageDimensions
  Height As Long
  Width As Long
End Type

Public Const PixelsPerInchX As Single = 90 ' Logical pixels/inch in X
Public Const PixelsPerInchY As Single = 90 ' Logical pixels/inch in Y

Public Const PixelPerPointX As Single = 1.76630434782609
Public Const PixelPerPointY As Single = 1.7578125

Public Const PointsPerInchX As Single = 52.0861538461537 ' 1.625
Public Const PointsPerInchY As Single = 54.6133333333333 ' 1.6875

Public Const InchesPerPointX As Single = 0.017663043478261
Public Const InchesPerPointY As Single = 0.017578125

Public Const PointPerPixelX As Single = 0.566153846153845
Public Const PointPerPixelY As Single = 0.568888888888889

Public Const ColorCount8bitDepth As Long = 256
Public Const ColorCount16bitDepth As Long = 65536
Public Const ColorCount24bitDepth As Long = 16777216

Private Const COLOR_SCROLLBAR = 0           ' Scroll Bar
Private Const COLOR_BACKGROUND = 1          ' Windows desktop
Private Const COLOR_ACTIVECAPTION = 2       ' Caption of active window
Private Const COLOR_INACTIVECAPTION = 3     ' Caption of inactive window
Private Const COLOR_MENU = 4                ' Menu
Private Const COLOR_WINDOW = 5              ' Window background
Private Const COLOR_WINDOWFRAME = 6         ' Window frame
Private Const COLOR_MENUTEXT = 7            ' Menu text
Private Const COLOR_WINDOWTEXT = 8          ' Window text
Private Const COLOR_CAPTIONTEXT = 9         ' Text in window caption
Private Const COLOR_ACTIVEBORDER = 10       ' Border of active window
Private Const COLOR_INACTIVEBORDER = 11     ' Border of inactive window
Private Const COLOR_APPWORKSPACE = 12       ' Background of MDI desktop
Private Const COLOR_HIGHLIGHT = 13          ' Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14      ' Selected item text
Private Const COLOR_BTNFACE = 15            ' Button
Private Const COLOR_BTNSHADOW = 16          ' 3D shading of button
Private Const COLOR_GRAYTEXT = 17           ' Gray text, or zero if dithering is used
Private Const COLOR_BTNTEXT = 18            ' Button text
Private Const COLOR_INACTIVECAPTIONTEXT = 19    ' Text of inactive window
Private Const COLOR_BTNHIGHLIGHT = 20       ' 3D highlight of button

Public Const PM_NOREMOVE = &H0
Public Const PM_REMOVE = &H1
Public Const PM_NOYIELD = &H2

Public Const FVF_VERTEX_SIZE = 12
Public Const FVF_RENDER_SIZE = 32
Public Const FVF_SCREEN = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
Public Const FVF_RENDER = D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1
Public Const Transparent As Long = &HFFFF00FF
Public Const MouseSensitivity As Long = 0.6
Public Const MaxDisplacement As Single = 0.05
Public Const BeaconSpacing As Single = 40
Public Const BeaconRange As Single = 1000
Public Const FadeDistance As Single = 800
Public Const SpaceBoundary As Single = 3000
Public Const HoursInOneDay As Single = 24
Public Const LetterPerInch As Single = 10


Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
 
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_F1 = &H70

Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Declare Function timeGetTime Lib "winmm" () As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal HwndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Const SWW_HPARENT = -8


Public Function GetMonitorDPI(Optional ByVal LogicalPixelsX As Long = 88, Optional ByVal LogicalPixelsY As Long = 90) As ImageDimensions
    
    Dim hdc As Long
    Dim lngRetVal As Long
    
    hdc = GetDC(0)
    
    GetMonitorDPI.Width = GetDeviceCaps(hdc, LogicalPixelsX)
    GetMonitorDPI.Height = GetDeviceCaps(hdc, LogicalPixelsY)
    
    lngRetVal = ReleaseDC(0, hdc)

End Function

Public Function ImageDimensions(ByVal FileName As String, ByRef imgdim As ImageDimensions, Optional ByRef ext As String = "") As Boolean

    If PathExists(FileName, True) Then

        'declare vars
        Dim handle As Integer
        Dim byteArr(255) As Byte
        
        'open file and get 256 byte chunk
        handle = FreeFile
        On Error GoTo endFunction
        Open FileName For Binary Access Read As #handle
            Get handle, , byteArr
        Close #handle
        
        ImageDimensions = ImageDimensionsFromBytes(byteArr, imgdim, ext)

    
    Else
        Debug.Print "Invalid picture file [" & FileName & "]"
    End If
endFunction:


End Function

Public Function ImageDimensionsFromBytes(ByRef byteArr() As Byte, ByRef imgdim As ImageDimensions, Optional ByRef ext As String = "") As Boolean

    Dim isValidImage As Boolean
    Dim i As Integer
    
    'init vars
    isValidImage = False
    imgdim.Height = 0
    imgdim.Width = 0

    
    'check for jpg header (SOI): &HFF and &HD8
    ' contained in first 2 bytes
    If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
        isValidImage = True
    Else
        GoTo checkGIF
    End If
    
    'check for SOF marker: &HFF and &HC0 TO &HCF
    For i = 0 To UBound(byteArr) - 1
        If byteArr(i) = &HFF And byteArr(i + 1) >= &HC0 And byteArr(i + 1) <= &HCF Then
'            imgdim.Height = byteArr(I + 5) * 256 + byteArr(I + 6)
'            imgdim.Width = byteArr(I + 7) * 256 + byteArr(I + 8)
            imgdim.Height = byteArr(i + 2) * 256 + byteArr(i + 3)
            imgdim.Width = byteArr(i + 4) * 256 + byteArr(i + 5)
            Exit For
        End If
    Next i
    
    'get image type and exit
    ext = "jpg"
    GoTo endFunction
    
checkGIF:
    
    'check for GIF header
    If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 And byteArr(3) = &H38 Then
        imgdim.Width = byteArr(7) * 256 + byteArr(6)
        imgdim.Height = byteArr(9) * 256 + byteArr(8)
        isValidImage = True
    Else
        GoTo checkBMP
    End If
    
    'get image type and exit
    ext = "gif"
    GoTo endFunction
    
checkBMP:
    
    'check for BMP header
    If byteArr(0) = 66 And byteArr(1) = 77 Then
        isValidImage = True
    Else
        GoTo checkPNG
    End If
    
    'get record type info
    If byteArr(14) = 40 Then
    
        'get width and height of BMP
        imgdim.Width = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
        + byteArr(19) * 256 + byteArr(18)
        
        imgdim.Height = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
        + byteArr(23) * 256 + byteArr(22)
    
    'another kind of BMP
    ElseIf byteArr(17) = 12 Then
    
        'get width and height of BMP
        imgdim.Width = byteArr(19) * 256 + byteArr(18)
        imgdim.Height = byteArr(21) * 256 + byteArr(20)
        
    End If
    
    'get image type and exit
    ext = "bmp"
    GoTo endFunction
    
checkPNG:
    
    'check for PNG header
    If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E And byteArr(3) = &H47 Then
        imgdim.Width = byteArr(18) * 256 + byteArr(19)
        imgdim.Height = byteArr(22) * 256 + byteArr(23)
        isValidImage = True
    Else
        GoTo endFunction
    End If
    
    ext = "png"


endFunction:
    
    'return function's success status
    ImageDimensionsFromBytes = isValidImage

End Function


#If Not modCommon Then

Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim C2 As Integer
    Dim retVal As Boolean
    retVal = True
    If Not IsNumeric(Text) Then
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) = 46) Then
                C2 = C2 + 1
            ElseIf (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                retVal = False
                Exit For
            End If
        Next
    Else
        retVal = False
    End If
    Else
        retVal = True
    End If
    IsAlphaNumeric = retVal And (C2 <= 1)
End Function

#End If





Public Function ConvertColor(ByVal Color As Variant, Optional ByRef Red As Long, Optional ByRef Green As Long, Optional ByRef Blue As Long) As Long
On Error GoTo catch
    Dim lngColor As Long
    If InStr(CStr(Color), "#") > 0 Then
        GoTo HTMLorHexColor
    ElseIf InStr(CStr(Color), "&H") > 0 Then
        GoTo SysOrLongColor
    ElseIf IsAlphaNumeric(Color) Then
        If (Not (Len(Color) = 6)) And (Not Left(Color, 1) = "0") Then
            GoTo SysOrLongColor
        Else
            GoTo HTMLorHexColor2
        End If
    End If
SysOrLongColor:
    lngColor = CLng(Color)
    If Not (lngColor >= 0 And lngColor <= 16777215) Then 'if system colour
        Select Case lngColor
            Case SystemColorConstants.vbScrollBars
                lngColor = COLOR_SCROLLBAR          ' Scroll Bar
            Case SystemColorConstants.vbDesktop
                lngColor = COLOR_BACKGROUND        ' Windows desktop
            Case SystemColorConstants.vbActiveTitleBar
                lngColor = COLOR_ACTIVECAPTION       ' Caption of active window
            Case SystemColorConstants.vbInactiveTitleBar
                lngColor = COLOR_INACTIVECAPTION     ' Caption of inactive window
            Case SystemColorConstants.vbMenuBar
                lngColor = COLOR_MENU                 ' Menu
            Case SystemColorConstants.vbWindowBackground
                lngColor = COLOR_WINDOW               ' Window background
            Case SystemColorConstants.vbWindowFrame
                lngColor = COLOR_WINDOWFRAME         ' Window frame
            Case SystemColorConstants.vbMenuText
                lngColor = COLOR_MENUTEXT            ' Menu text
            Case SystemColorConstants.vbWindowText
                lngColor = COLOR_WINDOWTEXT           ' Window text
            Case SystemColorConstants.vbTitleBarText
                lngColor = COLOR_CAPTIONTEXT         ' Text in window caption
            Case SystemColorConstants.vbActiveBorder
                lngColor = COLOR_ACTIVEBORDER       ' Border of active window
            Case SystemColorConstants.vbInactiveBorder
                lngColor = COLOR_INACTIVEBORDER     ' Border of inactive window
            Case SystemColorConstants.vbApplicationWorkspace
                lngColor = COLOR_APPWORKSPACE       ' Background of MDI desktop
            Case SystemColorConstants.vbHighlight
                lngColor = COLOR_HIGHLIGHT          ' Selected item background
            Case SystemColorConstants.vbHighlightText
                lngColor = COLOR_HIGHLIGHTTEXT      ' Selected item text
            Case SystemColorConstants.vbButtonFace
                lngColor = COLOR_BTNFACE             ' Button
            Case SystemColorConstants.vbButtonShadow
                lngColor = COLOR_BTNSHADOW          ' 3D shading of button
            Case SystemColorConstants.vbGrayText
                lngColor = COLOR_GRAYTEXT           ' Gray text, or zero if dithering is used
            Case SystemColorConstants.vbButtonText
                lngColor = COLOR_BTNTEXT             ' Button text
            Case SystemColorConstants.vbInactiveCaptionText
                lngColor = COLOR_INACTIVECAPTIONTEXT     ' Text of inactive window
            Case SystemColorConstants.vb3DHighlight
                lngColor = COLOR_BTNHIGHLIGHT        ' 3D highlight of button
        End Select

'        lngColor = lngColor And Not &H80000000
        lngColor = GetSysColor(lngColor)
'
'    Else
'
    End If
HTMLorHexColor2:
    Color = Right("000000" & Hex(lngColor), 6)
HTMLorHexColor:
    Red = CByte("&h" & Mid(Color, 5, 2))
    Green = CByte("&h" & Mid(Color, 3, 2))
    Blue = CByte("&h" & Mid(Color, 1, 2))
    
    ConvertColor = RGB(Red, Green, Blue)
    If ConvertColor <> lngColor Then
        Err.Raise 8, "Exception."
    End If
    Exit Function

'    green = Val("&H" & Right(color, 2))
'    red = Val("&H" & Mid(color, 2, 2))
'    blue = Val("&H" & Mid(color, 4, 2))
'    ConvertColor = RGB(red, green, blue)
'    Exit Function
catch:
    Err.Clear
    ConvertColor = 0
End Function


Public Function Padding(ByVal Length As Long, ByVal Value As String, Optional ByVal PadWith As String = " ") As String
    Padding = String(Abs((Length * Len(PadWith)) - (Len(Value) \ Len(PadWith))), PadWith) & Value
End Function

Function FloatToDWord(F As Single) As Long
    Dim buf As D3DXBuffer
    Dim l As Long
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, F
    D3DX.BufferGetData buf, 0, 4, 1, l
    FloatToDWord = l
End Function

Public Function GetUserLoginName() As String

    Dim sBuffer As String
    Dim lSize As Long
    sBuffer = String(255, Chr(0))
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        sBuffer = Left(sBuffer, lSize)
    End If
    sBuffer = Replace(sBuffer, Chr(0), "")
    If sBuffer = "" Then sBuffer = "SHARED"
    GetUserLoginName = sBuffer
    
End Function

Public Function AppPath() As String

    Dim Ret As String
    Ret = IIf((Right(App.Path, 1) = "\"), App.Path, App.Path & "\")
#If VBIDE = -1 Then
    If Right(Ret, 9) = "Projects\" Then
        Ret = Left(Ret, Len(Ret) - 9) & "Binary\"
     End If
#End If
    AppPath = Ret
End Function

Public Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    Dim pos As Long
    cnt = 0
    pos = InStr(Text, Word)
    Do Until pos = 0
        cnt = cnt + 1
        pos = InStr(pos + Len(Word), Text, Word)
    Loop
    CountWord = cnt
End Function

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
        Else
            NextArg = Trim(TheParams)
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
        Else
            NextArg = TheParams
        End If
    End If
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator)))
        Else
            RemoveArg = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
        Else
            RemoveArg = ""
        End If
    End If
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function
Public Function NextQuotedArg(ByVal TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    NextQuotedArg = RemoveQuotedArg(TheParams, BeginQuote, EndQuote, Embeded, Compare)
End Function

Public Function RemoveLineArg(ByRef TheParams As Variant, Optional ByVal EndOfLine As String = " ") As String
    If (InStr(TheParams, EndOfLine) > 0) And (InStr(TheParams, vbCrLf) > 1) Then
        If InStr(TheParams, EndOfLine) < InStr(TheParams, vbCrLf) Then
            RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
        Else
            RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
        End If
        
    ElseIf (InStr(TheParams, vbCrLf) = 0) And (InStr(TheParams, EndOfLine) = 0) Then
        RemoveLineArg = TheParams
        TheParams = ""
    ElseIf (InStr(TheParams, vbCrLf) = 0) Then
        RemoveLineArg = Left(TheParams, InStr(TheParams, EndOfLine) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, EndOfLine))
    Else
        RemoveLineArg = Left(TheParams, InStr(TheParams, vbCrLf) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, vbCrLf) + Len(vbCrLf))
    End If
End Function

Public Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
    Dim retVal As String
    Dim X As Long
    X = InStr(1, TheParams, BeginQuote, Compare)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + Len(BeginQuote), TheParams, EndQuote, Compare) > 0) Then
            If (Not Embeded) Or (EndQuote = BeginQuote) Then
                retVal = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retVal, InStr(1, retVal, EndQuote, Compare) + Len(EndQuote))
                retVal = Left(retVal, InStr(1, retVal, EndQuote, Compare) - 1)
            Else
                Dim l As Long
                Dim Y As Long
                l = 1
                Y = X
                Do Until l = 0
                    If (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare) > 0) And (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare) < InStr(Y + Len(BeginQuote), TheParams, EndQuote, Compare)) Then
                        l = l + 1
                        Y = InStr(Y + Len(BeginQuote), TheParams, BeginQuote, Compare)
                    ElseIf (InStr(Y + Len(BeginQuote), TheParams, EndQuote, Compare) > 0) Then
                        l = l - 1
                        Y = InStr(Y + Len(EndQuote), TheParams, EndQuote, Compare)
                    Else
                        Y = Len(TheParams)
                        l = 0
                    End If
                Loop
                retVal = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retVal, (Y - X) + Len(EndQuote))
                retVal = Left(retVal, (Y - X) - 1)
            End If
        End If
    End If
    RemoveQuotedArg = retVal
End Function


Public Function RemoveNextArgNoTrim(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(TheParams, TheSeperator) > 0 Then
        RemoveNextArgNoTrim = Left(TheParams, InStr(TheParams, TheSeperator) - 1)
        TheParams = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator))
    Else
        RemoveNextArgNoTrim = TheParams
        TheParams = ""
    End If
End Function

Public Function PathExists(ByVal URL As String, Optional ByVal IsFile As Variant) As Boolean
    Dim fso As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    If IsMissing(IsFile) Then
        PathExists = fso.FileExists(URL) Or fso.FolderExists(URL)
    Else
        If IsFile Then
            PathExists = fso.FileExists(URL)
        Else
            PathExists = fso.FolderExists(URL)
        End If
    End If
    Set fso = Nothing
    
'    If (URL = vbNullString) Then
'        PathExists = False
'        Exit Function
'    ElseIf (Not IsMissing(IsFile)) Then
'        If ((GetFilePath(URL) = vbNullString) And IsFile And (Not (URL = vbNullString))) Or ((GetFileName(URL) = vbNullString) And (Not IsFile) And (Not (URL = vbNullString))) Then
'            PathExists = False
'            Exit Function
'        End If
'    End If
'    If (IsMissing(IsFile)) Then IsFile = False
'    If (Len(URL) = 2) And (Mid(URL, 2, 1) = ":") Then
'        URL = URL & "\"
'    End If
'    Dim Ret As Long
'    On Error Resume Next
'    Ret = GetAttr(URL)
'    If Err.number = 0 Then
'        PathExists = IIf(IsFile, Not CBool(Ret And vbDirectory), True)
'    Else
'        Err.Clear
'        PathExists = False
'    End If
'    On Error GoTo 0
End Function

Public Function ReadFile(ByVal Path As String) As String
    Dim num As Long
    Dim Text As String
    Dim timeout As Single
    
    num = FreeFile
    On Error Resume Next
    On Local Error Resume Next
    If PathExists(Path, True) Then
        Open Path For Append Shared As #num Len = 1 ' LenB(Chr(CByte(0)))
        Close #num
        Select Case Err.Number
            Case 54, 70, 75
                Err.Clear
                On Error GoTo tryagain
                On Local Error GoTo tryagain
                
                Open Path For Binary Access Read Lock Write As num Len = 1
                If timeout <> 0 Then
                    Open Path For Binary Shared As #num Len = 1
                End If
                Text = String(LOF(num), " ")
                Get #num, 1, Text
                Close #num
            Case Else
                On Error GoTo tryagain
                On Local Error GoTo tryagain
                
                Open Path For Binary Access Read As num Len = 1
                If timeout <> 0 Then
                    Open Path For Binary Shared As num Len = 1
                End If
                Text = String(LOF(num), " ")
                Get #num, 1, Text
                Close #num
        End Select
        If Err Then GoTo failit
        On Error GoTo 0
        On Local Error GoTo 0
    End If
    ReadFile = Text
    Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        On Error GoTo failit
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "ReadFile"
End Function

Public Function WriteFile(ByVal Path As String, ByRef Text As String) As Boolean

    If PathExists(Path, True) Then
        If (GetAttr(Path) And vbReadOnly) <> 0 Then Exit Function
    End If
    
    Dim timeout As Single
    Dim num As Integer
    
    On Error Resume Next
    On Local Error Resume Next
    
    num = FreeFile
    Open Path For Output Shared As #num Len = 1  'Len = LenB(Chr(CByte(0)))
    Close #num
    
    Select Case Err.Number

        Case 54, 70, 75
            Err.Clear
            On Error GoTo tryagain
            On Local Error GoTo tryagain
            
            Open Path For Binary Access Write Lock Read As #num Len = 1
            If timeout <> 0 Then
                Open Path For Binary Shared As #num Len = 1
            End If
            Put #num, 1, Text
            Close #num
            WriteFile = True
        Case 0
            On Error GoTo tryagain
            On Local Error GoTo tryagain
            
            Open Path For Binary Access Write As #num Len = 1
            If timeout <> 0 Then
                Open Path For Binary Shared As #num Len = 1
            End If
            Put #num, 1, Text
            Close #num
            WriteFile = True
    End Select

    If Err Then GoTo failit
    On Error GoTo 0
    On Local Error GoTo 0
    
    Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "WriteFile"
End Function

Public Function GetFilePath(ByVal URL As String) As String
    Dim nFolder As String
    If InStr(URL, "/") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "/") - 1)
        If nFolder = "" Then nFolder = "/"
    ElseIf InStr(URL, "\") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "\") - 1)
        If nFolder = "" Then nFolder = "\"
    Else
        nFolder = ""
    End If
    GetFilePath = nFolder
End Function

Public Function GetFileName(ByVal URL As String) As String
    If InStr(URL, "/") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "/") + 1)
    ElseIf InStr(URL, "\") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "\") + 1)
    Else
        GetFileName = URL
    End If
End Function

Public Function GetFileTitle(ByVal URL As String) As String
    URL = GetFileName(URL)
    If InStrRev(URL, ".") > 0 Then
        URL = Left(URL, InStrRev(URL, ".") - 1)
    End If
    GetFileTitle = URL
End Function

Public Function GetFileExt(ByVal URL As String, Optional ByVal LowerCase As Boolean = True, Optional ByVal RemoveDot As Boolean = False) As String
    If InStrRev(URL, ".") > 0 Then
        If LowerCase Then
            GetFileExt = Trim(LCase(Mid(URL, (InStrRev(URL, ".") + -CInt(RemoveDot)))))
        Else
            GetFileExt = Mid(URL, (InStrRev(URL, ".") + -CInt(RemoveDot)))
        End If
    Else
        GetFileExt = vbNullString
    End If
End Function
Public Sub Swap(ByRef Var1, ByRef Var2, Optional ByRef Var3, Optional ByRef Var4, Optional ByRef Var5, Optional ByRef Var6)
    Dim Var0
    If (VBA.IsObject(Var1) Or VBA.TypeName(Var1) = "Nothing") Or _
        (VBA.IsObject(Var2) Or VBA.TypeName(Var2) = "Nothing") Then
        
        If IsMissing(Var3) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var0
        ElseIf IsMissing(Var4) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var0
        ElseIf IsMissing(Var5) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var0
        ElseIf IsMissing(Var6) Then
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var5
            Set Var5 = Var0
        Else
            Set Var0 = Var1
            Set Var1 = Var2
            Set Var2 = Var3
            Set Var3 = Var4
            Set Var4 = Var5
            Set Var5 = Var6
            Set Var6 = Var0
        End If
    
    Else
        
        If IsMissing(Var3) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var0
        ElseIf IsMissing(Var4) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var0
        ElseIf IsMissing(Var5) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var0
        ElseIf IsMissing(Var6) Then
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var5
            Var5 = Var0
        Else
            Var0 = Var1
            Var1 = Var2
            Var2 = Var3
            Var3 = Var4
            Var4 = Var5
            Var5 = Var6
            Var6 = Var0
        End If
    End If
End Sub

Public Function FileSize(ByVal fName As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim F As Object
    Set F = fso.GetFile(fName)
    FileSize = F.Size
    Set F = Nothing
    Set fso = Nothing
End Function

'Public Function RandomPositive(LowerBound As Long, UpperBound As Long) As Single
'    Randomize
'    RandomPositive = CSng((UpperBound - LowerBound + 1) * Rnd + LowerBound)
'End Function
Function MakeScreen(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, Optional ByVal tu As Single = 0, Optional ByVal tv As Single = 0) As MyScreen
    MakeScreen.X = X
    MakeScreen.Y = Y
    MakeScreen.Z = Z
    MakeScreen.rhw = 1
    MakeScreen.clr = D3DColorARGB(255, 255, 255, 255)
    MakeScreen.tu = tu
    MakeScreen.tv = tv
End Function
'Public Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As D3DVECTOR
'    MakeVector.X = X
'    MakeVector.Y = Y
'    MakeVector.Z = Z
'End Function
'
'Public Function MakePoint(ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Point
'    Set MakePoint = New Point
'    MakePoint.X = X
'    MakePoint.Y = Y
'    MakePoint.Z = Z
'End Function
'
'Public Function MakeCoord(ByVal X As Single, ByVal Y As Single) As Coord
'    Set MakeCoord = New Coord
'    MakeCoord.X = X
'    MakeCoord.Y = Y
'End Function
'
'Public Function ToCoord(ByRef Vector As D3DVECTOR) As Coord
'    Set ToCoord = New Coord
'    ToCoord.X = Vector.X
'    ToCoord.Y = Vector.Y
'End Function
'
'Public Function ToVector(ByRef Point As Point) As D3DVECTOR
'    ToVector.X = Point.X
'    ToVector.Y = Point.Y
'    ToVector.Z = Point.Z
'End Function
'
'Public Function ToPoint(ByRef Vector As D3DVECTOR) As Point
'    Set ToPoint = New Point
'    ToPoint.X = Vector.X
'    ToPoint.Y = Vector.Y
'    ToPoint.Z = Vector.Z
'End Function
'
'Public Function SquareCenter(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR, ByRef V3 As D3DVECTOR) As D3DVECTOR
'
'    SquareCenter.X = (v0.X + V1.X + V2.X + V3.X) / 4
'    SquareCenter.Y = (v0.Y + V1.Y + V2.Y + V3.Y) / 4
'    SquareCenter.Z = (v0.Z + V1.Z + V2.Z + V3.Z) / 4
'
'End Function

'Public Function Distance(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
'    Distance = Sqr(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
'    'Distance = Sqr(((p1.X - p2.X) * (p1.X - p2.X)) + ((p1.Y - p2.Y) * (p1.Y - p2.Y)) + ((p1.z - p2.z) * (p1.z - p2.z)))
'End Function

'Public Function DistanceEx1(ByRef p1 As D3DVECTOR, ByRef p2 As Point) As Single
'    DistanceEx1 = Sqr(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
'    'Distance = Sqr(((p1.X - p2.X) * (p1.X - p2.X)) + ((p1.Y - p2.Y) * (p1.Y - p2.Y)) + ((p1.z - p2.z) * (p1.z - p2.z)))
'End Function
'Public Function DistanceEx2(ByRef p1 As Point, ByRef p2 As D3DVECTOR) As Single
'    DistanceEx2 = Sqr(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) + ((p1.Z - p2.Z) ^ 2))
'    'Distance = Sqr(((p1.X - p2.X) * (p1.X - p2.X)) + ((p1.Y - p2.Y) * (p1.Y - p2.Y)) + ((p1.z - p2.z) * (p1.z - p2.z)))
'End Function
'
'
'Public Function PointOfDistanceToPoint(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByVal Distance As Single) As D3DVECTOR
'
'      PointOfDistanceToPoint.z = -(Sqr(-(((p1.X - p2.X) ^ 2) + ((p1.Y - p2.Y) ^ 2) - (Distance ^ 2))) - p1.z)
'      PointOfDistanceToPoint.Y = -(Sqr(-(((p1.X - p2.X) ^ 2) + ((p1.z - p2.z) ^ 2) - (Distance ^ 2))) - p1.Y)
'      PointOfDistanceToPoint.X = -(Sqr(-(((p1.Y - p2.Y) ^ 2) + ((p1.z - p2.z) ^ 2) - (Distance ^ 2))) - p1.X)
'
'    PointOfDistanceToPoint.z = (Sqr(-((((p1.X / (p1.Y / p2.z)) ^ 2) + ((p1.Y / (p1.z / p2.X) ^ 2)) - (Distance ^ 2))) + p2.z)
'    PointOfDistanceToPoint.Y = (Sqr(-((((p1.X / (p1.Y / p2.z)) ^ 2) + ((p1.z / (p1.X / p2.Y) ^ 2)) - (Distance ^ 2))) + p2.Y)
'    PointOfDistanceToPoint.X = (Sqr(-((((p1.Y / (p1.z / p2.X)) ^ 2) + ((p1.z / (p1.X / p2.Y) ^ 2)) - (Distance ^ 2))) + p2.X)
'
'
'End Function
Public Function CreateVertex(X As Single, Y As Single, Z As Single, nx As Single, ny As Single, nz As Single, tu As Single, tv As Single) As MyVertex
    
    With CreateVertex
        .X = X: .Y = Y: .Z = Z
        .nx = nx: .ny = ny: .nz = nz
        .tu = tu: .tv = tv
    End With
    
End Function
Public Function ConvertVertexToVector(ByRef v As D3DVERTEX) As D3DVECTOR
    ConvertVertexToVector.X = v.X
    ConvertVertexToVector.Y = v.Y
    ConvertVertexToVector.Z = v.Z
End Function

Public Function VectorNormalize(ByRef v As D3DVECTOR) As D3DVECTOR
    Dim l As Single
    l = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z)
    If l = 0 Then l = 1
    VectorNormalize.X = (v.X / l)
    VectorNormalize.Y = (v.Y / l)
    VectorNormalize.Z = (v.Z / l)
End Function

Public Function VectorDotProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As Single
    VectorDotProduct = (u.X * v.X + u.Y * v.Y + u.Z * v.Z)
End Function

Public Function VectorCrossProduct(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorCrossProduct.X = v.Y * u.Z - v.Z * u.Y
    VectorCrossProduct.Y = v.Z * u.X - v.X * u.Z
    VectorCrossProduct.Z = v.X * u.Y - v.Y * u.X
End Function

Public Function VectorSubtract(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorSubtract.X = v.X - u.X
    VectorSubtract.Y = v.Y - u.Y
    VectorSubtract.Z = v.Z - u.Z
End Function
Public Function VectorAdd(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorAdd.X = v.X + u.X
    VectorAdd.Y = v.Y + u.Y
    VectorAdd.Z = v.Z + u.Z
End Function
Public Function VectorMultiply(ByRef v As D3DVECTOR, ByRef u As D3DVECTOR) As D3DVECTOR
    VectorMultiply.X = v.X * u.X
    VectorMultiply.Y = v.Y * u.Y
    VectorMultiply.Z = v.Z * u.Z
End Function

Public Function PointToPlane(ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR) As Single
    Dim p3 As D3DVECTOR
    p3 = VectorSubtract(p1, p2)
    PointToPlane = Sqr(VectorDotProduct(p3, p3))
End Function


'Public Function TriangleNormal(ByRef v0 As D3DVECTOR, ByRef v1 As D3DVECTOR, ByRef v2 As D3DVECTOR) As D3DVECTOR
'
'    TriangleNormal.x = Sqr(((v0.Y - v1.Y) * (v0.z - v2.z) - (v0.z - v1.z) * (v0.Y - v2.Y)) ^ 2)
'    TriangleNormal.Y = Sqr(((v0.z - v1.z) * (v0.x - v2.x) - (v0.x - v1.x) * (v0.z - v2.z)) ^ 2)
'    TriangleNormal.z = Sqr(((v0.x - v1.x) * (v0.Y - v2.Y) - (v0.Y - v1.Y) * (v0.x - v2.x)) ^ 2)
'
'End Function

Public Function TriangleNormal(p0 As D3DVECTOR, p1 As D3DVECTOR, p2 As D3DVECTOR) As D3DVECTOR

    Dim v01 As D3DVECTOR
    Dim v02 As D3DVECTOR
    Dim vNorm As D3DVECTOR

    D3DXVec3Subtract v01, p1, p0
    D3DXVec3Subtract v02, p2, p0

    D3DXVec3Cross vNorm, v01, v02

    D3DXVec3Normalize vNorm, vNorm

    TriangleNormal = vNorm

End Function

Public Function TriangleCenter(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR) As D3DVECTOR

    Dim vR As D3DVECTOR
 
    vR.X = (v0.X + V1.X + V2.X) / 3
    vR.Y = (v0.Y + V1.Y + V2.Y) / 3
    vR.Z = (v0.Z + V1.Z + V2.Z) / 3
    
    TriangleCenter = vR

End Function

Public Sub CreateSquare(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As D3DVECTOR, ByRef p2 As D3DVECTOR, ByRef p3 As D3DVECTOR, ByRef P4 As D3DVECTOR, Optional ByVal ScaleX As Single = 1, Optional ByVal ScaleY As Single = 1)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, ScaleX, ScaleY)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.Z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.Z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.Z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, 0, ScaleY)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, ScaleX, 0)
    Data(Index + 5) = CreateVertex(P4.X, P4.Y, P4.Z, 0, 0, 0, 0, 0)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.Z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.Z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.Z
    
End Sub

Public Sub CreateSquareEx(ByRef Data() As MyVertex, ByVal Index As Long, ByRef p1 As MyVertex, ByRef p2 As MyVertex, ByRef p3 As MyVertex, ByRef P4 As MyVertex)

    Dim vn As D3DVECTOR
    
    Data(Index + 0) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 1) = CreateVertex(p2.X, p2.Y, p2.Z, 0, 0, 0, p2.tu, p2.tv)
    Data(Index + 2) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, p3.tu, p3.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 0).X, Data(Index + 0).Y, Data(Index + 0).Z), _
                            MakeVector(Data(Index + 1).X, Data(Index + 1).Y, Data(Index + 1).Z), _
                            MakeVector(Data(Index + 2).X, Data(Index + 2).Y, Data(Index + 2).Z))
    Data(Index + 0).nx = vn.X: Data(Index + 0).ny = vn.Y: Data(Index + 0).nz = vn.Z
    Data(Index + 1).nx = vn.X: Data(Index + 1).ny = vn.Y: Data(Index + 1).nz = vn.Z
    Data(Index + 2).nx = vn.X: Data(Index + 2).ny = vn.Y: Data(Index + 2).nz = vn.Z
    
    Data(Index + 3) = CreateVertex(p1.X, p1.Y, p1.Z, 0, 0, 0, p1.tu, p1.tv)
    Data(Index + 4) = CreateVertex(p3.X, p3.Y, p3.Z, 0, 0, 0, p3.tu, p3.tv)
    Data(Index + 5) = CreateVertex(P4.X, P4.Y, P4.Z, 0, 0, 0, P4.tu, P4.tv)
    vn = TriangleNormal(MakeVector(Data(Index + 3).X, Data(Index + 3).Y, Data(Index + 3).Z), _
                            MakeVector(Data(Index + 4).X, Data(Index + 4).Y, Data(Index + 4).Z), _
                            MakeVector(Data(Index + 5).X, Data(Index + 5).Y, Data(Index + 5).Z))
    Data(Index + 3).nx = vn.X: Data(Index + 3).ny = vn.Y: Data(Index + 3).nz = vn.Z
    Data(Index + 4).nx = vn.X: Data(Index + 4).ny = vn.Y: Data(Index + 4).nz = vn.Z
    Data(Index + 5).nx = vn.X: Data(Index + 5).ny = vn.Y: Data(Index + 5).nz = vn.Z
    
End Sub

Public Function CreateMesh(ByRef Obj As Element, ByVal FileName As String, Mesh As D3DXMesh, Buffer As D3DXBuffer, MeshMaterials() As D3DMATERIAL8, MeshTextures() As Direct3DTexture8, MeshVerticies() As D3DVERTEX, MeshIndicies() As Integer, nMaterials As Long)
    Dim TextureName As String

    Set Mesh = D3DX.LoadMeshFromX(FileName, D3DXMESH_DYNAMIC, DDevice, Nothing, Buffer, nMaterials)

    If nMaterials > 0 Then
    
        ReDim MeshMaterials(0 To nMaterials - 1) As D3DMATERIAL8
        ReDim MeshTextures(0 To nMaterials - 1) As Direct3DTexture8
    
        Dim d As ImageDimensions
        
        Dim q As Integer
        For q = 0 To nMaterials - 1
    
            D3DX.BufferGetMaterial Buffer, q, MeshMaterials(q)
       
            TextureName = D3DX.BufferGetTextureName(Buffer, q)
            If (TextureName <> "") Then
                If ImageDimensions(AppPath & "Models\" & TextureName, d) Then
                    Set MeshTextures(q) = D3DX.CreateTextureFromFileEx(DDevice, AppPath & "Models\" & TextureName, d.Width, d.Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Transparent, ByVal 0, ByVal 0)
                Else
                    Debug.Print "IMAGE ERROR: ImageDimensions - " & AppPath & "Models\" & TextureName
                End If
            End If
            
        Next
    Else
        ReDim MeshTextures(0 To 0) As Direct3DTexture8
        ReDim MeshMaterials(0 To 0) As D3DMATERIAL8
    End If
    
    Dim vd As D3DVERTEXBUFFER_DESC
    Mesh.GetVertexBuffer.GetDesc vd

    ReDim MeshVerticies(0 To ((vd.Size \ FVF_VERTEX_SIZE) - 1)) As D3DVERTEX
    D3DVertexBuffer8GetData Mesh.GetVertexBuffer, 0, vd.Size, 0, MeshVerticies(0)

    Dim id As D3DINDEXBUFFER_DESC
    Mesh.GetIndexBuffer.GetDesc id

    ReDim MeshIndicies(0 To ((id.Size \ 2) - 1)) As Integer
    D3DIndexBuffer8GetData Mesh.GetIndexBuffer, 0, id.Size, 0, MeshIndicies(0)
    
    D3DX.ComputeNormals Mesh

    Dim i As Long
    
    Dim v As Long
    
    Dim avg As New Point
    If Obj.Displace Is Nothing Then Set Obj.Displace = New Point
    
    
    
    For i = 0 To ((id.Size \ 2) - 1)
        v = MeshIndicies(i)
            
        avg.X = avg.X + MeshVerticies(v).X
        avg.Y = avg.Y + MeshVerticies(v).Y
        avg.Z = avg.Z + MeshVerticies(v).Z
        If Abs(MeshVerticies(v).X - Obj.Origin.X) > Obj.Displace.X Then
            Obj.Displace.X = Abs(MeshVerticies(v).X - Obj.Origin.X)
        End If
        If Abs(MeshVerticies(v).Y - Obj.Origin.Y) > Obj.Displace.Y Then
            Obj.Displace.Y = Abs(MeshVerticies(v).Y - Obj.Origin.Y)
        End If
        If Abs(MeshVerticies(v).Z - Obj.Origin.Z) > Obj.Displace.Z Then
            Obj.Displace.Z = Abs(MeshVerticies(v).Z - Obj.Origin.Z)
        End If
    Next
    i = ((id.Size \ 2) - 1)
    avg.X = avg.X / i
    avg.Y = avg.Y / i
    avg.Z = avg.Z / i
    Set Obj.Centoid = avg
    
    

End Function

Public Function GetNow() As String

    GetNow = CStr(Now)

End Function

Public Function GetTimer() As String

    GetTimer = CStr(CDbl(Timer))
    
End Function

Public Sub SwapSingle(ByRef val1 As Single, ByRef val2 As Single)
    Dim tmp As Single
    tmp = val1
    val1 = val2
    val2 = tmp
End Sub

Public Function NoAngle() As Point
    Set NoAngle = MakePoint(360, 360, 360)
End Function
Public Function NoPoint() As Point
    Set NoPoint = MakePoint(0, 0, 0)
End Function
Public Function ClassifyPoint(ByRef v0 As D3DVECTOR, ByRef V1 As D3DVECTOR, ByRef V2 As D3DVECTOR, ByRef p As D3DVECTOR) As Single
    Dim dtp As Single
    Dim N As D3DVECTOR
    N = GetPlaneNormal(v0, V1, V2)
    dtp = VectorDotProduct(N, p) + -VectorDotProduct(N, v0)
    
    If dtp > Epsilon Then
        ClassifyPoint = 1 'front
    ElseIf dtp < -Epsilon Then
        ClassifyPoint = 2 'back
    Else
        ClassifyPoint = 0 'onplane
    End If
End Function



'Public Function PointBehindtriangle(ByRef Point As D3DVECTOR, ByRef center As D3DVECTOR, ByRef Lengths As D3DVECTOR) As Boolean
'
'' (GreatestX(v1, v2, v3) - LeastX(v1, v2, v3)) / 2, _
''                                (GreatestY(v1, v2, v3) - LeastY(v1, v2, v3)) / 2, _
''                                (GreatestZ(v1, v2, v3) - LeastZ(v1, v2, v3)) / 2, _
''                                Distance(v1, v2), Distance(v2, v3), Distance(v3, v1),
'
'    Dim v As D3DVECTOR
'    Dim u As D3DVECTOR
'    Dim n As D3DVECTOR
'
'    Dim d As Single
'
'    u = VectorCrossProduct(Point, center)
'
'    v = VectorSubtract(Point, center)
'    v = VectorSubtract(v, center)
'    v = VectorCrossProduct(v, center)
'
'    n.x = (((u.x + u.Y) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.x + v.Y) * (Lengths.x + Lengths.Y + Lengths.z)))
'    n.Y = (((u.x + u.z) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.x + v.z) * (Lengths.x + Lengths.Y + Lengths.z)))
'    n.x = (((u.Y + u.x) * (Lengths.x + Lengths.Y + Lengths.z)) - ((v.Y + v.x) * (Lengths.x + Lengths.Y + Lengths.z)))
'
'    d = Distance(u, v)
'
'    n.x = (n.x * u.x) / d
'    n.Y = (n.Y * u.Y) / d
'    n.z = (n.z * u.z) / d
'
'    'Debug.Print "n: " & n.x & "," & n.y & "," & n.z & "  u: " & u.x & "," & u.y & "," & u.z & "  v: " & v.x & "," & v.y & "," & v.z & "  d: " & d
'
'    PointBehindtriangle = n.x + n.Y + n.z > 0
'End Function
'
'Public Function TriangleIntersect(ByRef O1 As D3DVECTOR, ByRef O2 As D3DVECTOR, ByRef O3 As D3DVECTOR, ByRef Q1 As D3DVECTOR, ByRef Q2 As D3DVECTOR, ByRef Q3 As D3DVECTOR) As Long
'
'    Dim No As D3DVECTOR
'    Dim Nq As D3DVECTOR
'    Dim Co As D3DVECTOR
'    Dim Cq As D3DVECTOR
'    Dim Lc As Single
'    Dim Lo As Single
'    Dim Lq As Single
'    Dim Da As D3DVECTOR
'    Dim Oi As Single
'    Dim Qi As Single
'    Dim Jo As D3DVECTOR
'    Dim Jq As D3DVECTOR
'    Dim K As Integer
'
'    No = TriangleNormal(O1, O2, O3)
'    Nq = TriangleNormal(Q1, Q2, Q3)
'    Co = TriangleCenter(O1, O2, O3)
'    Cq = TriangleCenter(Q1, Q2, Q3)
'    Lc = Distance(Co, Cq)
'    Lo = Distance(O1, O2) + Distance(O2, O3) + Distance(O3, O1)
'    Lq = Distance(Q1, Q2) + Distance(Q2, Q3) + Distance(Q3, Q1)
'
'    Da.x = Sqr((((Lo + Lq) * (No.x + Co.x)) + (((Q1.Y + Q2.Y + Q3.Y - O1.Y + O2.Y + O3.Y) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.Y * No.z * Co.Y) + (O1.Y + O2.Y + O3.Y)))) ^ 2)
'    Da.Y = Sqr((((Lo + Lq) * (No.Y + Co.Y)) + (((Q1.x + Q2.x + Q3.x - O1.x + O2.x + O3.x) + (Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z)) * ((No.x * No.z * Co.x) + (O1.x + O2.x + O3.x)))) ^ 2)
'    Da.z = Sqr((((Lo + Lq) * (No.z + Co.z)) + (((Q1.z + Q2.z + Q3.z - O1.z + O2.z + O3.z) + (Q1.x + Q2.x + Q3.x - O1.x + O2.x + O3.x)) * ((No.x * No.x * Co.z) + (O1.z + O2.z + O3.z)))) ^ 2)
'
'    Oi = Sqr(((Da.x * Da.Y * Da.z) ^ 3 + ((Da.x + Da.Y + Da.z) * (Da.x + Da.Y + Da.z))))
'
'    Jo.x = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'    Jo.Y = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'    Jo.z = (((Q1.x + Q2.x + Q3.x) * (Q1.Y + Q2.Y + Q3.Y) * (Q1.z + Q2.z + Q3.z) * (Lc + Lc + Lo)) / Oi)
'
'    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jo.x) & " " & DebugPrint(Jo.y) & " " & DebugPrint(Jo.z) & " " & DebugPrint(Lo, 10) & " " & DebugPrint(Oi)
'
'    Da.x = Sqr((((Lq + Lo) * (Nq.x + Cq.x)) + (((O1.Y + O2.Y + O3.Y - Q1.Y + Q2.Y + Q3.Y) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.Y * Nq.z * Cq.Y) + (Q1.Y + Q2.Y + Q3.Y)))) ^ 2)
'    Da.Y = Sqr((((Lq + Lo) * (Nq.Y + Cq.Y)) + (((O1.x + O2.x + O3.x - Q1.x + Q2.x + Q3.x) + (O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z)) * ((Nq.x * Nq.z * Cq.x) + (Q1.x + Q2.x + Q3.x)))) ^ 2)
'    Da.z = Sqr((((Lq + Lo) * (Nq.z + Cq.z)) + (((O1.z + O2.z + O3.z - Q1.z + Q2.z + Q3.z) + (O1.x + O2.x + O3.x - Q1.x + Q2.x + Q3.x)) * ((Nq.x * Nq.x * Cq.z) + (Q1.z + Q2.z + Q3.z)))) ^ 2)
'
'    Qi = Sqr(((Da.x * Da.Y * Da.z) ^ 3 + ((Da.x + Da.Y + Da.z) * (Da.x + Da.Y + Da.z))))
'
'    Jq.x = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'    Jq.Y = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'    Jq.z = (((O1.x + O2.x + O3.x) * (O1.Y + O2.Y + O3.Y) * (O1.z + O2.z + O3.z) * (Lc + Lc + Lo)) / Qi)
'
'    'Debug.Print DebugPrint(Da.x) & " " & DebugPrint(Da.y) & " " & DebugPrint(Da.z) & " " & DebugPrint(Jq.x) & " " & DebugPrint(Jq.y) & " " & DebugPrint(Jq.z) & " " & DebugPrint(Lq, 10) & " " & DebugPrint(Qi)
'
'    K = (Sqr(Distance(Jo, Jq) * Sqr(((Oi / 2) + (Qi / 2)))))
'
'    'Debug.Print "Sect: " & CStr(K)
'
'    TriangleIntersect = K
'
'End Function
'



''procedure 3Dto2D (x, y, z, pan, centre, position)
''X = X + position.X
''Y = Y + position.Y
''Z = Z + position.Z
''new.x = x*cos(pan.x) - z*sin(pan.x)
''new.z = x*sin(pan.x) + z*cos(pan.x)
''new.y = y*cos(pan.y) - new.z*sin(pan.y)
''z = new.y*cos(pan.y) - new.z*sin(pan.y)
''x = new.x*cos(pan.z) - new.y*sin(pan.z)
''y = new.x*sin(pan.z) + new.y*cos(pan.z)
''If Z > 0 Then
''    Screen.X = X / Z * Zoom + centre.X
''    Screen.Y = Y / Z * Zoom + centre.Y
''End If

'Public Function ScreenVertex(ByVal VertexIndex As Long, ByRef ScreenX As Single, ByRef ScreenY As Single, ByRef ScreenZ As Single)
'    Dim r As D3DVECTOR 'VRP
'    Dim p As D3DVECTOR 'WCP
'    Dim n As D3DVECTOR
'    Dim up As D3DVECTOR
'    Dim v As D3DVECTOR
'    Dim u As D3DVECTOR
'    Dim dot As Single
'    Dim dif As D3DVECTOR
'
'
'    GetVertex VertexIndex, p
'    FR_Camera.GetPosition grid.Direct3DFrame, r
'    FR_Camera.GetOrientation grid.Direct3DFrame, n, up
'    DX_Main.VectorNormalize n
'
'
'    dot = DX_Main.VectorDotProduct(up, n)
'    v.X = up.X - dot * n.X
'    v.Y = up.Y - dot * n.Y
'    v.z = up.z - dot * n.z
'    DX_Main.VectorNormalize v
'
'
'    DX_Main.VectorCrossProduct u, n, v
'
'
'    DX_Main.VectorSubtract dif, p, r
'
'
'    ScreenX = DX_Main.VectorDotProduct(dif, u) + (ScreenWidth / 2)
'    ScreenY = DX_Main.VectorDotProduct(dif, v) + (ScreenHeight / 2)
'    ScreenZ = DX_Main.VectorDotProduct(dif, n)
'
'
'End Function
'
''Public Property Get ScreenX(ByVal VertexIndex As Integer) As Single
''    On Error Resume Next
''
''    Dim vert As D3DVECTOR
''    Dim vertA As D3DVECTOR
''    Dim vertB As D3DVECTOR
''
''    Dim cVert As D3DVECTOR
''    Dim cRot As D3DVECTOR
''    Dim cData As Single
''
''    FR_Camera.GetRotation FR_Root, cRot, cData
''    FR_Camera.GetPosition FR_Root, cVert
''
''    GetVertex VertexIndex, vert
''
''    'ScreenX = vert.X + (ScreenWidth / 2)
''
''    DX_Main.VectorAdd vertA, cVert, vert
''
''    DX_Main.VectorRotate vertB, vertA, cRot, cData
''
''
''    ScreenX = (vertB.X / vertB.Z * 1000) + (ScreenWidth / 2)    'Zoom + centre.X
''
''    'ScreenX = (vert.X / vert.Z * 100) + (ScreenWidth / 2)  'Zoom + centre.X
''End Property
''Public Property Get ScreenY(ByVal VertexIndex As Integer) As Single
''    On Error Resume Next
''
''    Dim vert As D3DVECTOR
''    Dim vertA As D3DVECTOR
''    Dim vertB As D3DVECTOR
''
''    Dim cVert As D3DVECTOR
''    Dim cRot As D3DVECTOR
''    Dim cData As Single
''
''    FR_Camera.GetRotation FR_Root, cRot, cData
''    FR_Camera.GetPosition FR_Root, cVert
''
''    GetVertex VertexIndex, vert
''
''    'ScreenY = vert.Y + (ScreenWidth / 2)
''
''    DX_Main.VectorAdd vertA, cVert, vert
''
''    DX_Main.VectorRotate vertB, vertA, cRot, cData
''
''
''    ScreenY = (vertB.Y / vertB.Z * 1000) + (ScreenHeight / 2)     'Zoom + centre.X
''
''    'ScreenY = (vert.Y / vert.Z * 100) + (ScreenHeight / 2)   'Zoom + centre.X
''End Property

