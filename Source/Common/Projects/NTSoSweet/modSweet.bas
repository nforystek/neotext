Attribute VB_Name = "modSweet"
#Const [True] = -1
#Const [False] = 0
#Const modSweet = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Private xColornerInc As Color
Private xIdentityInc As Long

Public Const StateErrors = 0
Public Const StateStages = 21
Public Const StateSyntax = 42

Public ParserMessages(0 To 62) As String

'Error Messages
Public Const ErrMsg_00 = 0
Public Const ErrMsg_01 = 1
Public Const ErrMsg_02 = 2
Public Const ErrMsg_03 = 3
Public Const ErrMsg_04 = 4
Public Const ErrMsg_05 = 5
Public Const ErrMsg_06 = 6
Public Const ErrMsg_07 = 7
Public Const ErrMsg_08 = 8
Public Const ErrMsg_09 = 9
Public Const ErrMsg_10 = 10
Public Const ErrMsg_11 = 11
Public Const ErrMsg_12 = 12
Public Const ErrMsg_13 = 13
Public Const ErrMsg_14 = 14
Public Const ErrMsg_15 = 15
Public Const ErrMsg_16 = 16
Public Const ErrMsg_17 = 17
Public Const ErrMsg_18 = 18
Public Const ErrMsg_19 = 19
Public Const ErrMsg_20 = 20

'Interpeter States
Public Const Stages_00 = 21
Public Const Stages_01 = 22
Public Const Stages_02 = 23
Public Const Stages_03 = 24
Public Const Stages_04 = 25
Public Const Stages_05 = 26
Public Const Stages_06 = 27
Public Const Stages_07 = 28
Public Const Stages_08 = 29
Public Const Stages_09 = 30
Public Const Stages_10 = 31
Public Const Stages_11 = 32
Public Const Stages_12 = 33
Public Const Stages_13 = 34
Public Const Stages_14 = 35
Public Const Stages_15 = 36
Public Const Stages_16 = 37
Public Const Stages_17 = 38
Public Const Stages_18 = 39
Public Const Stages_19 = 40
Public Const Stages_20 = 41

'Expected Syntax
Public Const Syntax_00 = 42
Public Const Syntax_01 = 43
Public Const Syntax_02 = 44
Public Const Syntax_03 = 45
Public Const Syntax_04 = 46
Public Const Syntax_05 = 47
Public Const Syntax_06 = 48
Public Const Syntax_07 = 49
Public Const Syntax_08 = 50
Public Const Syntax_09 = 51
Public Const Syntax_10 = 52
Public Const Syntax_11 = 53
Public Const Syntax_12 = 54
Public Const Syntax_13 = 55
Public Const Syntax_14 = 56
Public Const Syntax_15 = 57
Public Const Syntax_16 = 58
Public Const Syntax_17 = 59
Public Const Syntax_18 = 60
Public Const Syntax_19 = 61
Public Const Syntax_20 = 62


Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


Public Sub Main()
    soSweetPaserPresets
    
End Sub


Public Sub soSweetPaserPresets()
    'hoop and wrap are aliased for nesting in plurals of the same
    'Line and ream as aliased
    '
    ParserMessages(ParserCodes.ErrMsg_00) = "Expecting HOOPS statement above and before all."
    ParserMessages(ParserCodes.ErrMsg_01) = "Invalid COLOR color entity for HOOPS base."
    ParserMessages(ParserCodes.ErrMsg_02) = "Improper WRAP statement or HOOP arrangement."
    ParserMessages(ParserCodes.ErrMsg_03) = "Improper PUSH or MUSH statement in HOOP arrangement."
    ParserMessages(ParserCodes.ErrMsg_04) = "Invalid COLOR color entity for this HOOP line lace."
    ParserMessages(ParserCodes.ErrMsg_05) = "Invalid COLOR color entity for following HOOP marker."
    ParserMessages(ParserCodes.ErrMsg_06) = "Markers for the WORDS statement invoked HOOP incorrectly."
    ParserMessages(ParserCodes.ErrMsg_07) = "Markers for the LETTERS statement invoked HOOP incorrectly."
    ParserMessages(ParserCodes.ErrMsg_08) = "Markers for the DIGIT statement invoked HOOP incorrectly."
    ParserMessages(ParserCodes.ErrMsg_09) = "Expecting COLOR or other valid HOOPS continuations."
    ParserMessages(ParserCodes.ErrMsg_10) = "Expecting REAMS statement as the partition half."
    ParserMessages(ParserCodes.ErrMsg_11) = "Invalid COLOR color entity for REAMS partition half."
    ParserMessages(ParserCodes.ErrMsg_12) = "Improper RUSH or MUSH statement in REAM arrangement."
    ParserMessages(ParserCodes.ErrMsg_13) = "Improper PUSH or MUSH statement in REAM arrangement."
    ParserMessages(ParserCodes.ErrMsg_14) = "Invalid COLOR color entity for this REAM line lace."
    ParserMessages(ParserCodes.ErrMsg_15) = "Invalid COLOR color entity for following REAM marker."
    ParserMessages(ParserCodes.ErrMsg_16) = "Markers for the WORDS statement invoked REAM incorrectly."
    ParserMessages(ParserCodes.ErrMsg_17) = "Markers for the LETTERS statement invoked REAM incorrectly."
    ParserMessages(ParserCodes.ErrMsg_18) = "Markers for the DIGIT statement invoked REAM incorrectly."
    ParserMessages(ParserCodes.ErrMsg_19) = "Execution of the script preformed with syntax errors."
    ParserMessages(ParserCodes.ErrMsg_20) = "External errors prevented the script's execution."
    
    ParserMessages(ParserCodes.Stages_00) = "Parser recognized the HOOPS stage has statements."
    ParserMessages(ParserCodes.Stages_01) = "Parser recognized HOOPS level COLOR one statement."
    ParserMessages(ParserCodes.Stages_02) = "Parser traversal of HOOP lace WRAP statement."
    ParserMessages(ParserCodes.Stages_03) = "Parser traversal of HOOP lace PUSH or MUSH statement."
    ParserMessages(ParserCodes.Stages_04) = "Parser recognized HOOPS level COLOR two statement."
    ParserMessages(ParserCodes.Stages_05) = "Parser recognized HOOPS level COLOR three statement."
    ParserMessages(ParserCodes.Stages_06) = "Parser traversal at HOOPS level of WORDS markers."
    ParserMessages(ParserCodes.Stages_07) = "Parser traversal at HOOPS level of LETTERS markers."
    ParserMessages(ParserCodes.Stages_08) = "Parser traversal at HOOPS level of DIGIT markers."
    ParserMessages(ParserCodes.Stages_09) = "Parser recognized futher traversal of HOOPS level."
    ParserMessages(ParserCodes.Stages_10) = "Parser recognized the REAMS statement partition."
    ParserMessages(ParserCodes.Stages_11) = "Parser recognized REAMS level COLOR one statement."
    ParserMessages(ParserCodes.Stages_12) = "Parser traversal of REAM lace RUSH or MUSH statement."
    ParserMessages(ParserCodes.Stages_13) = "Parser traversal of REAM lace PUSH or MUSH statement."
    ParserMessages(ParserCodes.Stages_14) = "Parser recognized REAMS level COLOR two statement."
    ParserMessages(ParserCodes.Stages_15) = "Parser recognized REAMS level COLOR three statement."
    ParserMessages(ParserCodes.Stages_16) = "Parser traversal at REAMS level of WORDS markers."
    ParserMessages(ParserCodes.Stages_17) = "Parser traversal at REAMS level of LETTERS markers."
    ParserMessages(ParserCodes.Stages_18) = "Parser traversal at REAMS level of DIGIT markers."
    ParserMessages(ParserCodes.Stages_19) = "Parser recognized futher traversal of REAMS level."
    ParserMessages(ParserCodes.Stages_20) = "Parser has interpreted all phases of the script."
    
    ParserMessages(ParserCodes.Syntax_00) = "Syntax: hoops [color <color>]; Example: hoops"
    ParserMessages(ParserCodes.Syntax_01) = "Syntax: color <color>; Example: color #FF00FF"
    ParserMessages(ParserCodes.Syntax_02) = "Syntax: wrap ? hoop ? [color <color>]; Example: wrap "" hoop """
    ParserMessages(ParserCodes.Syntax_03) = "Syntax: push ? mush ? hoop ? [color <color>]; Example: push ' hoop '"
    ParserMessages(ParserCodes.Syntax_04) = "Syntax: color <color>; Example: ...hoop eol color background"
    ParserMessages(ParserCodes.Syntax_05) = "Syntax: color <color>; Example: color B2A4C6 words..."
    ParserMessages(ParserCodes.Syntax_06) = "Syntax: [color <color>] words ? ?...; word ?; Example: words con cat bit"
    ParserMessages(ParserCodes.Syntax_07) = "Syntax: [color <color>] letters ! !...; letter !; Example: letters & % # 1 a b"
    ParserMessages(ParserCodes.Syntax_08) = "Syntax: [color <color>] digits # ##...; digit ##; Example: digits 1 48 4800"
    ParserMessages(ParserCodes.Syntax_09) = "Syntax: ... hoop; Example: wrap ' mush '' hoop ' color #00FF00"
    ParserMessages(ParserCodes.Syntax_10) = "Syntax: reams [color <color>]; Example: reams color White"
    ParserMessages(ParserCodes.Syntax_11) = "Syntax: color <color>; Example: color &H124"
    ParserMessages(ParserCodes.Syntax_12) = "Syntax: rush ? [push ?] hoop ? [color <color>]; Example: rush /* hoop */"
    ParserMessages(ParserCodes.Syntax_13) = "Syntax: push ? [mush ?] hoop ? [color <color>]; Example: push * mush ** hoop *"
    ParserMessages(ParserCodes.Syntax_14) = "Syntax: color <color>; Example: ...ream color 0x052"
    ParserMessages(ParserCodes.Syntax_15) = "Syntax: color <color>; Example: color 2762 digit..."
    ParserMessages(ParserCodes.Syntax_16) = "Syntax: [color <color>] words ? ?...; word ?; Example: word isoneword"
    ParserMessages(ParserCodes.Syntax_17) = "Syntax: [color <color>] letters ! !...; letter !; Example: letter !"
    ParserMessages(ParserCodes.Syntax_18) = "Syntax: [color <color>] digits ## #...; digit #; Example: digit 0"
    ParserMessages(ParserCodes.Syntax_19) = "Syntax: ... ream; Example: rush [ push < push > ream ]"
    ParserMessages(ParserCodes.Syntax_20) = "Syntax: eol; eof; Example: words to express like eol"

End Sub

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
            GoTo HTMLorHexColor
        End If
    End If
SysOrLongColor:
    lngColor = CLng(Color)
    If Not (lngColor >= 0 And lngColor <= 16777215) Then 'if system colour
        lngColor = lngColor And Not &H80000000
        lngColor = GetSysColor(lngColor)
    End If
    Color = Right("000000" & Hex(lngColor), 6)
    Red = CByte("&h" & Mid(Color, 1, 2))
    Green = CByte("&h" & Mid(Color, 3, 2))
    Blue = CByte("&h" & Mid(Color, 5, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
HTMLorHexColor:
    Red = Val("&H" & Left(Color, 2))
    Green = Val("&H" & Mid(Color, 3, 2))
    Blue = Val("&H" & Right(Color, 2))
    ConvertColor = RGB(Red, Green, Blue)
    Exit Function
catch:
    Err.Clear
    ConvertColor = 0
End Function

Public Function GetNewIdentity() As Long
    xIdentityInc = xIdentityInc + 1
    GetNewIdentity = xIdentityInc
End Function

Public Function GetNewColorner(Optional ByVal Rawvalue As String = "", Optional ByVal Background As Boolean = False) As Object
    If xColornerInc Is Nothing Then
        Set GetNewColorner = New Color
        If (Rawvalue = "") Then
            GetNewColorner.Rawvalue = "#000000"
        Else
            GetNewColorner.Rawvalue = Rawvalue
        End If
    ElseIf (Not (Rawvalue = "")) Or (((Rawvalue <> xColornerInc.Rawvalue) And Background) Or Not Background) Then
        If Background Then
            Set xColornerInc = New Color
            xColornerInc.Rawvalue = Rawvalue
            Set GetNewColorner = xColornerInc
        Else
            Set GetNewColorner = New Color
            GetNewColorner.Rawvalue = Rawvalue
        End If
    End If
    
End Function

Public Function FileSize(ByVal fName As String) As Double
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim f As Object
    Set f = fso.GetFile(fName)
    FileSize = f.Size
    Set f = Nothing
    Set fso = Nothing
End Function

Public Function ReadFile(ByVal Path As String) As String
    Dim num As Long
    Dim Text As String
    num = FreeFile
    Open Path For Input Shared As #num Len = LenB(Chr(0))
    Close #num
    Open Path For Binary Shared As #num Len = LenB(Chr(0))
        Text = String(FileSize(Path), Chr(0))
        Get #num, 1, Text
    Close #num
    ReadFile = Text
End Function

Public Sub WriteFile(ByVal Path As String, ByRef Text As String)
    Dim num As Long
    num = FreeFile
    Open Path For Output Shared As #num Len = LenB(Chr(0))
    Close #num
    Kill Path
    Open Path For Binary Shared As #num Len = LenB(Chr(0))
        Put #num, 1, Text
    Close #num
End Sub

Public Function ClearCollection(ByRef ColList As VBA.Collection, Optional ByVal IsObjects As Boolean = False, Optional ByVal SetNothing As Boolean = False)
    If Not ColList Is Nothing Then
        If IsObjects Then
            Dim Obj As Object
            Do Until ColList.Count = 0
                Set Obj = ColList(1)
                ColList.Remove 1
                Set Obj = Nothing
            Loop
        Else
            Do Until ColList.Count = 0
                ColList.Remove 1
            Loop
        End If
        If SetNothing Then Set ColList = Nothing
    End If
End Function

Public Function TrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    TrimStrip = LTrimStrip(RTrimStrip(TheStr, TheChar), TheChar)
End Function
Public Function LTrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    Do While Left(TheStr, Len(TheChar)) = TheChar
        TheStr = Mid(TheStr, Len(TheChar) + 1)
    Loop
    LTrimStrip = TheStr
End Function
Public Function RTrimStrip(ByVal TheStr As String, ByVal TheChar As String) As String
    Do While Right(TheStr, Len(TheChar)) = TheChar
        TheStr = Left(TheStr, Len(TheStr) - Len(TheChar))
    Loop
    RTrimStrip = TheStr
End Function

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal NoTrim As Boolean = False) As String
    If NoTrim Then
        If InStr(TheParams, TheSeperator) > 0 Then
            NextArg = Left(TheParams, InStr(TheParams, TheSeperator) - 1)
        Else
            NextArg = TheParams
        End If
    Else
        If InStr(TheParams, TheSeperator) > 0 Then
            NextArg = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
        Else
            NextArg = Trim(TheParams)
        End If
    End If
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal NoTrim As Boolean = False) As String
    If NoTrim Then
        If InStr(1, TheParams, TheSeperator) > 0 Then
            RemoveArg = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
        Else
            RemoveArg = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator) > 0 Then
            RemoveArg = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator)))
        Else
            RemoveArg = ""
        End If
    End If
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal NoTrim As Boolean = False) As String
    If NoTrim Then
        If InStr(TheParams, TheSeperator) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(TheParams, TheSeperator) - 1)
            TheParams = Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    Else
        If InStr(TheParams, TheSeperator) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(TheParams, TheSeperator) - 1))
            TheParams = Trim(Mid(TheParams, InStr(TheParams, TheSeperator) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    End If
End Function
Public Function NextQuotedArg(ByVal TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False) As String
    NextQuotedArg = RemoveQuotedArg(TheParams, BeginQuote, EndQuote, Embeded)
End Function
Public Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False) As String
    Dim retVal As String
    Dim X As Long
    X = InStr(TheParams, BeginQuote)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + 1, TheParams, EndQuote) > 0) Then
            If (Not Embeded) Or (EndQuote = BeginQuote) Then

                retVal = Mid(TheParams, X + 1)
                TheParams = Left(TheParams, X - 1) & Mid(retVal, InStr(retVal, EndQuote) + 1)
                retVal = Left(retVal, InStr(retVal, EndQuote) - 1)

            Else
                Dim l As Long
                Dim Y As Long

                l = 1
                Y = X

                Do Until l = 0
                    If (InStr(Y + 1, TheParams, BeginQuote) > 0) And (InStr(Y + 1, TheParams, BeginQuote) < InStr(Y + 1, TheParams, EndQuote)) Then
                        l = l + 1
                        Y = InStr(Y + 1, TheParams, BeginQuote)
                    ElseIf (InStr(Y + 1, TheParams, EndQuote) > 0) Then
                        l = l - 1
                        Y = InStr(Y + 1, TheParams, EndQuote)
                    Else
                        Y = Len(TheParams)
                        l = 0
                    End If

                Loop

                retVal = Mid(TheParams, X + 1)
                TheParams = Left(TheParams, X - 1) & Mid(retVal, (Y - X) + 1)
                retVal = Left(retVal, (Y - X) - 1)
            End If

        End If
    End If
    RemoveQuotedArg = retVal
End Function

Public Function CountWord(ByVal Text As String, ByVal Word As String, Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt > 0 Then CountWord = cnt
End Function

Public Function WordCount(ByVal Text As String, Optional ByVal TheSeperator As String = " ", Optional ByVal Exact As Boolean = True) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, TheSeperator, , IIf(Exact, vbBinaryCompare, vbTextCompare)))
    If cnt >= 0 Then WordCount = cnt
End Function

Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
    Dim cnt As Integer
    Dim c2 As Integer
    Dim retVal As Boolean
    retVal = True
    If Len(Text) > 0 Then
        For cnt = 1 To Len(Text)
            If (Asc(LCase(Mid(Text, cnt, 1))) = 46) Then
                c2 = c2 + 1
            ElseIf (Not IsNumeric(Mid(Text, cnt, 1))) And (Not (Asc(LCase(Mid(Text, cnt, 1))) >= 97 And Asc(LCase(Mid(Text, cnt, 1))) <= 122)) Then
                retVal = False
                Exit For
            End If
        Next
    Else
        retVal = False
    End If
    IsAlphaNumeric = retVal And (c2 <= 1)
End Function
