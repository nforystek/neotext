VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pGreaterThan As NTStreams10.Stream
Private pLesserThan As NTStreams10.Stream
Private pLeftBracket As NTStreams10.Stream
Private pRightBracket As NTStreams10.Stream
Private pEqualSign As NTStreams10.Stream
Private pDoubleQuote As NTStreams10.Stream
Private pSingleQuote As NTStreams10.Stream
Private pColon As NTStreams10.Stream
Private pLeftBrace As NTStreams10.Stream
Private pRightBrace As NTStreams10.Stream
Private pSpaceChar As NTStreams10.Stream
Private pTabIndent As NTStreams10.Stream
Private pCr As NTStreams10.Stream
Private pLf As NTStreams10.Stream
Private pCrLf As NTStreams10.Stream

Public Property Get GreaterThan() As NTStreams10.Stream
    Set GreaterThan = pGreaterThan
End Property
Public Property Get LesserThan() As NTStreams10.Stream
    Set LesserThan = pLesserThan
End Property
Public Property Get LeftBracket() As NTStreams10.Stream
    Set LeftBracket = pLeftBracket
End Property
Public Property Get RightBracket() As NTStreams10.Stream
    Set RightBracket = pRightBracket
End Property
Public Property Get EqualSign() As NTStreams10.Stream
    Set EqualSign = pEqualSign
End Property
Public Property Get DoubleQuote() As NTStreams10.Stream
    Set DoubleQuote = pDoubleQuote
End Property
Public Property Get SingleQuote() As NTStreams10.Stream
    Set SingleQuote = pSingleQuote
End Property
Public Property Get Colon() As NTStreams10.Stream
    Set Colon = pColon
End Property
Public Property Get LeftBrace() As NTStreams10.Stream
    Set LeftBrace = pLeftBrace
End Property
Public Property Get RightBrace() As NTStreams10.Stream
    Set RightBrace = pRightBrace
End Property
Public Property Get SpaceChar() As NTStreams10.Stream
    Set SpaceChar = pSpaceChar
End Property
Public Property Get TabIndent() As NTStreams10.Stream
    Set TabIndent = pTabIndent
End Property
Public Property Get Cr() As NTStreams10.Stream
    Set Cr = pCr
End Property
Public Property Get Lf() As NTStreams10.Stream
    Set Lf = pLf
End Property
Public Property Get CrLf() As NTStreams10.Stream
    Set CrLf = pCrLf
End Property

Public Function ToString(ByRef Stream)
    If VBA.TypeName(Stream) = "Stream" Then
        ToString = StrConv(Stream.Partial, vbUnicode)
    ElseIf VBA.TypeName(Stream) = "String" Then
        ToString = Stream
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Public Function ToStream(ByRef str)
    If VBA.TypeName(str) = "String" Then
        Set ToStream = New Stream
        'Debug.Print StrConv(str, vbFromUnicode)
        ToStream = StrConv(str, vbFromUnicode)
        'ToStream.concat StrConv(str, vbFromUnicode)
    ElseIf VBA.TypeName(str) = "Stream" Then
        Set ToStream = str
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Private Function DetermineDataConcatenation(ByRef Obj, Optional ByRef Data) As Boolean
    If IsMissing(Data) Then Exit Function
    If VBA.TypeName(Obj) = "String" Then
        If VBA.TypeName(Data) = "String" Then
            Obj = Obj & Data
        ElseIf VBA.TypeName(Data) = "Stream" Then
            Dim str1 As Stream
            Set str1 = Data
            Obj = Obj & StrConv(str1.Partial, vbUnicode)
            Set str1 = Nothing
        End If
    ElseIf VBA.TypeName(Obj) = "Stream" Then
        Dim obj2 As Stream
        Set obj2 = Obj
        If VBA.TypeName(Data) = "String" Then
            obj2.ConcatBytes StrConv(Data, vbFromUnicode)
        ElseIf VBA.TypeName(Data) = "Stream" Then
            Dim str2 As Stream
            Set str2 = Data
            obj2.ConcatBytes str2.Partial
            Set str2 = Nothing
        End If
        Set obj2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
    DetermineDataConcatenation = True
End Function

Public Function ConcatenateEx(ByRef Data1, ByRef Data2, Optional ByRef Data3, Optional ByRef Data4, Optional ByRef Data5, _
    Optional ByRef Data6, Optional ByRef Data7, Optional ByRef Data8, Optional ByRef Data9, Optional ByRef Data10)
    If VBA.TypeName(Data1) = "Stream" Then
        Dim newObj As New Stream
        If Not DetermineDataConcatenation(newObj, Data1) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data2) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data3) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data4) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data5) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data6) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data7) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data8) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data9) Then GoTo doneobj
        DetermineDataConcatenation newObj, Data10
doneobj:
        Set ConcatenateEx = newObj
        Set newObj = Nothing
    ElseIf VBA.TypeName(Data1) = "String" Then
        Dim newStr As String
        If Not DetermineDataConcatenation(newStr, Data1) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data2) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data3) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data4) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data5) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data6) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data7) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data8) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data9) Then GoTo doneStr
        DetermineDataConcatenation newStr, Data10
doneStr:
        ConcatenateEx = newStr
        newStr = ""
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Public Sub ConcatenateTo(ByRef Obj, ByRef Data1, Optional ByRef Data2, Optional ByRef Data3, Optional ByRef Data4, _
    Optional ByRef Data5, Optional ByRef Data6, Optional ByRef Data7, Optional ByRef Data8, Optional ByRef Data9, Optional ByRef Data10)
    If VBA.TypeName(Obj) = "Stream" Then
        Dim newObj As New Stream
        If Not DetermineDataConcatenation(newObj, Data1) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data2) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data3) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data4) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data5) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data6) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data7) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data8) Then GoTo doneobj
        If Not DetermineDataConcatenation(newObj, Data9) Then GoTo doneobj
        DetermineDataConcatenation newObj, Data10
doneobj:
        Set Obj = newObj
        Set newObj = Nothing
    ElseIf VBA.TypeName(Obj) = "String" Then
        Dim newStr As String
        If Not DetermineDataConcatenation(newStr, Data1) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data2) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data3) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data4) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data5) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data6) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data7) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data8) Then GoTo doneStr
        If Not DetermineDataConcatenation(newStr, Data9) Then GoTo doneStr
        DetermineDataConcatenation newStr, Data10
doneStr:
        Obj = newStr
        newStr = ""
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Sub

                

Public Function InStr(StartPos, WithIn, Optional FindThis, Optional CaseSensitive)
    If VBA.TypeName(WithIn) = "String" Then
        If VBA.IsNumeric(StartPos) And VBA.TypeName(FindThis) = "String" Then
            If VBA.TypeName(CaseSensitive) = "Long" Then
                InStr = VBA.InStr(StartPos, WithIn, FindThis, CaseSensitive)
            ElseIf VBA.TypeName(CaseSensitive) = "Boolean" Then
                InStr = VBA.InStr(StartPos, WithIn, FindThis, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
            Else
                InStr = VBA.InStr(StartPos, WithIn, FindThis)
            End If
        ElseIf VBA.TypeName(StartPos) = "String" And VBA.TypeName(FindThis) = "String" And IsMissing(FindThis) Then
            InStr = VBA.InStr(1, WithIn, FindThis)
        Else
            Err.Raise 5, App.Title, "Invalid procedure call or argument"
        End If
    ElseIf VBA.TypeName(WithIn) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        If VBA.IsNumeric(StartPos) And VBA.TypeName(FindThis) = "Stream" Then
            Set arg1 = WithIn
            Set arg2 = FindThis
            If VBA.TypeName(CaseSensitive) = "Boolean" Then
                InStr = InStrP(StartPos, arg1, arg2, CaseSensitive)
            ElseIf VBA.TypeName(CaseSensitive) = "Long" Then
                InStr = InStrP(StartPos, arg1, arg2, (CaseSensitive = VbCompareMethod.vbBinaryCompare))
            Else
                InStr = InStrP(StartPos, arg1, arg2)
            End If
        ElseIf VBA.TypeName(StartPos) = "Stream" And VBA.TypeName(WithIn) = "Stream" And IsMissing(FindThis) Then
            Set arg1 = StartPos
            Set arg2 = WithIn
            InStr = InStrP(1, arg1, arg2)
        Else
            Err.Raise 5, App.Title, "Invalid procedure call or argument"
        End If
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function


Friend Function InStrP(ByVal StartPos As Long, ByRef WithIn As Stream, Optional ByRef FindThis As Stream, Optional ByVal CaseSensitive As Boolean = True)

    If WithIn Is Nothing Or FindThis Is Nothing Then
        Set InStrP = Nothing
        Exit Function
    ElseIf WithIn.Length = 0 Then
        InStrP = 0
        Exit Function
    ElseIf FindThis.Length = 0 Then
        InStrP = StartPos
        Exit Function
    End If
    
    Dim skp As Long
    skp = IIf(StartPos > 0, StartPos - 1, 0)
    
    If (FindThis.Length > (WithIn.Length - skp)) Then
        InStrP = 0
        Exit Function
    End If

    Dim pos As Long
    Dim cnt As Long
    Dim fnd As Long

    Dim ret() As Byte
    Dim tmp() As Byte

    Do
        cnt = 0
        fnd = 0
        ret = FindThis.Partial(0, 1)
        
        pos = WithIn.Poll(ret(LBound(ret)), 1, skp)
        If ((skp + pos) = WithIn.Length) And (Not CaseSensitive) Then
            If IsUpperCase(ret(LBound(ret))) Then
                pos = WithIn.Poll((ret(LBound(ret)) + 32), 1, skp)
            ElseIf IsLowerCase(ret(LBound(ret))) Then
                pos = WithIn.Poll((ret(LBound(ret)) - 32), 1, skp)
            End If
        End If
        
        If (skp + pos) < WithIn.Length Then
            fnd = (skp + pos)
            Do While cnt + 1 < FindThis.Length And (fnd - cnt) = (skp + pos)
                cnt = cnt + 1
                ret = FindThis.Partial(cnt, 1)
                tmp = WithIn.Partial((skp + pos) + cnt, 1)
                If ret(LBound(ret)) = tmp(LBound(ret)) Then
                    fnd = (skp + pos) + cnt
                ElseIf Not CaseSensitive Then
                    If IsUpperCase(ret(LBound(ret))) And IsLowerCase(tmp(LBound(ret))) Then
                        If ret(LBound(ret)) + 32 = tmp(LBound(ret)) Then
                            fnd = (skp + pos) + cnt
                        End If
                    ElseIf IsLowerCase(ret(LBound(ret))) And IsUpperCase(tmp(LBound(ret))) Then
                        If ret(LBound(ret)) - 32 = tmp(LBound(ret)) Then
                            fnd = (skp + pos) + cnt
                        End If
                    End If
                End If
            Loop
            If (fnd - cnt) = (skp + pos) Then
                InStrP = ((skp + pos) + 1)
                Exit Function
            End If
        End If
        skp = (skp + pos) + 1

    Loop While pos < WithIn.Length

End Function

Private Function IsUpperCase(ByVal Data As Byte) As Boolean
    IsUpperCase = Data >= 65 And Data <= 90
End Function
Private Function IsLowerCase(ByVal Data As Byte) As Boolean
    IsLowerCase = Data >= 97 And Data <= 122
End Function

Public Function Left(OfString, ByVal NumOfChars)
    If VBA.TypeName(OfString) = "String" Then
        Left = VBA.Left(OfString, NumOfChars)
    ElseIf VBA.TypeName(OfString) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OfString
        Set Left = LeftP(arg1, NumOfChars)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function
Friend Function LeftP(ByRef OfString As Stream, ByVal NumOfChars As Long) As Stream
    Set LeftP = New Stream
    If Not NumOfChars <= 0 Then
        LeftP.ConcatBytes OfString.Partial(0, NumOfChars)
    End If
End Function

Public Function Right(OfString, ByVal NumOfChars)
    If VBA.TypeName(OfString) = "String" Then
        Right = VBA.Right(OfString, NumOfChars)
    ElseIf VBA.TypeName(OfString) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OfString
        Set Right = RightP(arg1, NumOfChars)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function RightP(ByRef OfString As Stream, ByVal NumOfChars As Long) As Stream
    Set RightP = New Stream
    If Not NumOfChars <= 0 Then
        NumOfChars = (OfString.Length - NumOfChars)
        If NumOfChars > 0 Then
            RightP.ConcatBytes OfString.Partial(NumOfChars)
        Else
            RightP.ConcatBytes OfString.Partial
        End If
    End If
End Function


Public Function Mid(OfString, ByVal StartChar, Optional ByVal NumOfChars)
    If VBA.TypeName(OfString) = "String" Then
        If VBA.IsMissing(NumOfChars) Then
            Mid = VBA.Mid(OfString, StartChar)
        Else
            Mid = VBA.Mid(OfString, StartChar, NumOfChars)
        End If
    ElseIf VBA.TypeName(OfString) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OfString
        If VBA.IsMissing(NumOfChars) Then
            Set Mid = MidP(arg1, StartChar)
        Else
            Set Mid = MidP(arg1, StartChar, NumOfChars)
        End If
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function


Friend Function MidP(ByRef OfString As Stream, ByVal StartChar As Long, Optional ByVal NumOfChars As Variant) As Stream
    Set MidP = New Stream
    If Not VBA.IsMissing(NumOfChars) Then
        If NumOfChars <= 0 Then Exit Function
    End If
    If VBA.IsMissing(NumOfChars) Then NumOfChars = -1
    If (((StartChar - 1) >= 0) And (OfString.Length > 0)) And (Not (StartChar > OfString.Length)) Then
        MidP.ConcatBytes OfString.Partial(StartChar - 1, CLng(NumOfChars))
    End If
End Function

Public Function replace(OfString, FindText, replaceWith, Optional ByVal CaseSensitive)
    If VBA.TypeName(OfString) = "String" And VBA.TypeName(FindText) = "String" And VBA.TypeName(replaceWith) = "String" Then
        If VBA.TypeName(CaseSensitive) = "Long" Then
            replace = VBA.replace(OfString, FindText, replaceWith, CaseSensitive)
        ElseIf VBA.TypeName(CaseSensitive) = "Boolean" Then
            replace = VBA.replace(OfString, FindText, replaceWith, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
        Else
            replace = VBA.replace(OfString, FindText, replaceWith)
        End If
    ElseIf VBA.TypeName(OfString) = "Stream" And VBA.TypeName(FindText) = "Stream" And VBA.TypeName(replaceWith) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OfString
        Dim arg2 As Stream
        Set arg2 = FindText
        Dim arg3 As Stream
        Set arg3 = replaceWith
        If VBA.TypeName(CaseSensitive) = "Boolean" Then
            replace = ReplaceP(arg1, arg2, arg3, CaseSensitive)
        ElseIf VBA.TypeName(CaseSensitive) = "Long" Then
            replace = ReplaceP(arg1, arg2, arg3, (CaseSensitive = vbBinaryCompare))
        Else
            replace = ReplaceP(arg1, arg2, arg3)
        End If
        Set arg1 = Nothing
        Set arg2 = Nothing
        Set arg3 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function


Friend Function ReplaceP(ByRef OfString As Stream, ByRef FindText As Stream, ByRef replaceWith As Stream, Optional ByVal CaseSensitive As Boolean = True) As Stream
    Dim pos As Long
    Set ReplaceP = New Stream
    ReplaceP.ConcatBytes OfString.Partial
    If FindText.Length > 0 And ReplaceP.Length > 0 Then
        pos = 1
        Do
            pos = InStr(pos, ReplaceP, FindText, CaseSensitive)
            If pos > 0 And pos <= ReplaceP.Length Then
                If replaceWith.Length > 0 Then
                    ReplaceP.Pyramid replaceWith, (pos - 1), FindText.Length
                    pos = (pos + replaceWith.Length)
                Else
                    ReplaceP.Pinch (pos - 1), FindText.Length
                    pos = (pos + replaceWith.Length)
                End If
            End If
        Loop While pos > 0 And pos <= ReplaceP.Length
    End If
End Function

Public Function StrReverse(OnString)
    If VBA.TypeName(OnString) = "String" Then
        StrReverse = VBA.StrReverse(OnString)
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OnString
        Set StrReverse = StrReverseP(arg1)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function StrReverseP(ByRef OnString As Stream) As Stream
    Set StrReverseP = New Stream
    If OnString.Length > 0 Then
        Dim pos As Long
        For pos = OnString.Length - 1 To 0
            StrReverseP.Post OnString.Peek(pos)
        Next
    End If
End Function

Public Function UCase(OnString)
    If VBA.TypeName(OnString) = "String" Then
        UCase = VBA.UCase(OnString)
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As New Stream
        arg1 = OnString
        Set UCase = UCaseP(arg1)
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function UCaseP(ByRef OnString As Stream) As Stream
    Set UCaseP = New Stream
    UCaseP.ConcatBytes OnString.Partial
    If UCaseP.Length > 0 Then
        Dim pos As Long
        Dim Val As Byte
        For pos = 0 To UCaseP.Length - 1
            Val = UCaseP.Peek(pos)
            If IsLowerCase(Val) Then
                UCaseP.Plot pos, Val - 32
            End If
        Next
    End If
End Function

Public Function LCase(OnString)
    If VBA.TypeName(OnString) = "String" Then
        LCase = VBA.LCase(OnString)
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As New Stream
        arg1 = OnString
        Set LCase = LCaseP(arg1)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function LCaseP(ByRef OnString As Stream) As Stream
    Set LCaseP = New Stream
    LCaseP.ConcatBytes OnString.Partial
    If LCaseP.Length > 0 Then
        Dim pos As Long
        Dim Val As Byte
        For pos = 0 To LCaseP.Length - 1
            Val = LCaseP.Peek(pos)
            If IsUpperCase(Val) Then
                LCaseP.Plot pos, Val + 32
            End If
        Next
    End If
End Function

Public Function trim(OnString)
    If VBA.TypeName(OnString) = "String" Then
        trim = VBA.trim(OnString)
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As New Stream
        arg1 = OnString
        Set trim = TrimP(arg1)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function TrimP(ByRef OnString As Stream) As Stream
    Set TrimP = New Stream
    If OnString.Length > 0 Then
        TrimP.ConcatBytes OnString.Partial
        Dim pos As Long
        Do While TrimP.Length > 0
            If TrimP.Peek(0) = Asc(" ") Then
                If TrimP.Length = 1 Then
                    TrimP.Length = 0
                Else
                    TrimP.Length = TrimP.Length + 1
                    TrimP.Push 1
                    TrimP.Length = TrimP.Length - 1
                End If
            Else
                Exit Do
            End If
        Loop
        
        Do While TrimP.Length > 0
            If TrimP.Peek(TrimP.Length - 1) = Asc(" ") Then
                If TrimP.Length = 1 Then
                    TrimP.Length = 0
                Else
                    TrimP.Length = TrimP.Length - 1
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Public Function TrimStrip(ByRef OnString As Stream, ByRef StripOf As Stream, Optional ByVal CaseSensitive = True) As Stream
    Set TrimStrip = New Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    If OnString.Length > 0 Then
        TrimStrip.ConcatBytes OnString.Partial
        Dim pos As Long
        Do
            pos = Me.InStr(1, TrimStrip, StripOf, CaseSensitive)
            If pos = 1 Then
                If TrimStrip.Length = StripOf.Length Then
                    TrimStrip.Length = 0
                Else
                    TrimStrip.Length = TrimStrip.Length + StripOf.Length
                    TrimStrip.Push StripOf.Length
                    TrimStrip.Length = TrimStrip.Length - StripOf.Length
                End If
            Else
                Exit Do
            End If
        Loop While pos > 0
        
        Do
            pos = Me.InStr(((TrimStrip.Length - StripOf.Length) + 1), TrimStrip, StripOf, CaseSensitive)
            If pos = ((TrimStrip.Length - StripOf.Length) + 1) Then
                If TrimStrip.Length = StripOf.Length Then
                    TrimStrip.Length = 0
                Else
                    TrimStrip.Length = TrimStrip.Length - StripOf.Length
                End If
            Else
                Exit Do
            End If
        Loop While pos > 0
    End If
End Function

Public Function LTrimStrip(ByRef OnString, ByRef StripOf, Optional ByVal CaseSensitive = True)
    If VBA.TypeName(OnString) = "String" And VBA.TypeName(StripOf) = "String" Then
        LTrimStrip = modCommon.LTrimStrip(OnString, StripOf)
    ElseIf VBA.TypeName(OnString) = "Stream" And VBA.TypeName(StripOf) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Set arg1 = OnString
        Set arg2 = StripOf
        Set LTrimStrip = LTrimStripP(arg1, arg2, CaseSensitive)
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function LTrimStripP(ByRef OnString As Stream, ByRef StripOf As Stream, Optional ByVal CaseSensitive = True) As Stream
    Set LTrimStripP = New Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    If OnString.Length > 0 Then
        LTrimStripP.ConcatBytes OnString.Partial
        Dim pos As Long
        Do
            pos = Me.InStr(1, LTrimStripP, StripOf, CaseSensitive)
            If pos = 1 Then
                If LTrimStripP.Length = StripOf.Length Then
                    LTrimStripP.Length = 0
                Else
                    LTrimStripP.Length = LTrimStripP.Length + StripOf.Length
                    LTrimStripP.Push StripOf.Length
                    LTrimStripP.Length = LTrimStripP.Length - StripOf.Length
                End If
            Else
                Exit Do
            End If
        Loop While pos > 0
    End If
End Function

Public Function RTrimStrip(ByRef OnString, ByRef StripOf, Optional ByVal CaseSensitive = True)
    If VBA.TypeName(OnString) = "String" And VBA.TypeName(StripOf) = "String" Then
        RTrimStrip = modCommon.RTrimStrip(OnString, StripOf)
    ElseIf VBA.TypeName(OnString) = "Stream" And VBA.TypeName(StripOf) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Set arg1 = OnString
        Set arg2 = StripOf
        Set RTrimStrip = RTrimStripP(arg1, arg2, CaseSensitive)
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function RTrimStripP(ByRef OnString As Stream, ByRef StripOf As Stream, Optional ByVal CaseSensitive = True) As Stream
    Set RTrimStripP = New Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    If OnString.Length > 0 Then
        RTrimStripP.ConcatBytes OnString.Partial
        Dim pos As Long
        Do
            pos = InStrP(((RTrimStripP.Length - StripOf.Length) + 1), RTrimStripP, StripOf, CaseSensitive)
            If pos = ((RTrimStripP.Length - StripOf.Length) + 1) Then
                If RTrimStripP.Length = StripOf.Length Then
                    RTrimStripP.Length = 0
                Else
                    RTrimStripP.Length = RTrimStripP.Length - StripOf.Length
                End If
            Else
                Exit Do
            End If
        Loop While pos > 0
    End If
End Function

Public Function LTrim(ByRef OnString)
    If VBA.TypeName(OnString) = "String" Then
        LTrim = modCommon.LTrimStrip(OnString, " ")
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As New Stream
        arg1 = OnString
        Set LTrim = LTrimP(arg1)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function LTrimP(ByRef OnString As Stream) As Stream
    Set LTrimP = New Stream
    If OnString.Length > 0 Then
        LTrimP.ConcatBytes OnString.Partial
        Dim pos As Long
        Do While LTrimP.Length > 0
            If LTrimP.Peek(0) = Asc(" ") Then
                If LTrimP.Length = 1 Then
                    LTrimP.Length = 0
                Else
                    LTrimP.Length = LTrimP.Length + 1
                    LTrimP.Push 1
                    LTrimP.Length = LTrimP.Length - 1
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Public Function RTrim(ByRef OnString)
    If VBA.TypeName(OnString) = "String" Then
        RTrim = modCommon.RTrimStrip(OnString, " ")
    ElseIf VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As New Stream
        arg1 = OnString
        Set RTrim = RTrimP(arg1)
        Set arg1 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function RTrimP(ByRef OnString As Stream) As Stream
    Set RTrimP = New Stream
    If OnString.Length > 0 Then
        RTrimP.ConcatBytes OnString.Partial
        Dim pos As Long
        Do While RTrimP.Length > 0
            If RTrimP.Peek(RTrimP.Length - 1) = Asc(" ") Then
                If RTrimP.Length = 1 Then
                    RTrimP.Length = 0
                Else
                    RTrimP.Length = RTrimP.Length - 1
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Public Function IsNumeric(OnString) As Boolean
    If VBA.TypeName(OnString) = "Stream" Then
        Dim arg1 As Stream
        Set arg1 = OnString
        'IsNumeric = IsNumericP(arg1)
        IsNumeric = IsNumeric(StrConv(arg1.Partial, vbUnicode))
        Set arg1 = Nothing
    Else
        IsNumeric = VBA.IsNumeric(OnString)
    End If
End Function

'Private Function IsNumericP(ByRef OnString As Stream) As Boolean
'    If OnString.Length > 0 Then
'        IsNumericP = True
'        Dim pos As Long
'        Dim val As Byte
'
'        Dim has As Boolean
'
'        Dim dot As Boolean
'        Dim exp As Boolean
'        Dim min As Boolean
'        Dim dol As Boolean
'        Dim per As Boolean
'        Dim plu As Boolean
'
'        For pos = 0 To OnString.Length - 1
'            val = OnString.Peek(pos)
'
'            'flag determining 1 of any number exists somewhere
'            If ((Not has) And (val >= 48 And val <= 57)) Then has = True
'
'            'to refine the rulls commented below are not enforced:
'
'            'one dot anywhere in position is valid
'            If (val = 46) Then dot = True
'
'            'exponential must have numericals on both sides of it
'            If (val = 94) Then exp = True
'
'            'minus and plus muse be found at the beginning of the string
'            If (val = 45) Then min = True
'            If (val = 43) Then plu = True
'
'            'dollar and percent sign must be at the beginning or ending of the string
'            If (val = 36) Then dol = True
'            If (val = 38) Then per = True
'
'
'            IsNumericP = IsNumericP And ((val = 43) Or (val = 38) Or (val = 36) Or _
'                (val = 45) Or (val = 94) Or (val = 46) Or (val >= 48 And val <= 57))
'        Next
'        If IsNumericP And (min Or dot Or exp Or dol Or per Or plu) Then IsNumericP = has
'
'    End If
'End Function

Public Function NextArg(ByRef TheParams, ByRef TheSeperator, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True)
    If VBA.TypeName(TheParams) = "String" And VBA.TypeName(TheSeperator) = "String" Then
        NextArg = modCommon.NextArg(TheParams, TheSeperator, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare), TrimResult)
    ElseIf VBA.TypeName(TheParams) = "Stream" And VBA.TypeName(TheSeperator) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Set arg1 = TheParams
        Set arg2 = TheSeperator
        Set NextArg = NextArgP(arg1, arg2, CaseSensitive, TrimResult)
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function NextArgP(ByRef TheParams As Stream, ByRef TheSeperator As Stream, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True) As Stream
    Dim pos As Long
    Set NextArgP = New Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    If TrimResult Then
        pos = InStrP(1, TheParams, TheSeperator, CaseSensitive)
        If pos > 0 Then
            NextArgP.ConcatBytes TrimP(LeftP(TheParams, pos - 1)).Partial
        Else
            NextArgP.ConcatBytes TrimP(TheParams).Partial
        End If
    Else
        pos = InStrP(1, TheParams, TheSeperator, CaseSensitive)
        If pos > 0 Then
            NextArgP.ConcatBytes LeftP(TheParams, pos - 1).Partial
        Else
            NextArgP.ConcatBytes TheParams.Partial
        End If
    End If
End Function

Public Function RemoveArg(ByRef TheParams, ByRef TheSeperator, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True)
    If VBA.TypeName(TheParams) = "String" And VBA.TypeName(TheSeperator) = "String" Then
        RemoveArg = modCommon.RemoveArg(TheParams, TheSeperator, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare), TrimResult)
    ElseIf VBA.TypeName(TheParams) = "Stream" And VBA.TypeName(TheSeperator) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Set arg1 = TheParams
        Set arg2 = TheSeperator
        Set RemoveArg = RemoveArgP(arg1, arg2, CaseSensitive, TrimResult)
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function RemoveArgP(ByRef TheParams As Stream, ByRef TheSeperator As Stream, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True) As Stream
    Dim pos As Long
    Set RemoveArgP = New Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    If TrimResult Then
        pos = InStrP(1, TheParams, TheSeperator, CaseSensitive)
        If pos > 0 Then
            RemoveArgP.ConcatBytes TrimP(MidP(TheParams, pos + TheSeperator.Length)).Partial
        Else
            RemoveArgP.ConcatBytes TrimP(TheParams).Partial
        End If
    Else
        pos = InStrP(1, TheParams, TheSeperator, CaseSensitive)
        If pos > 0 Then
            RemoveArgP.ConcatBytes MidP(TheParams, pos + TheSeperator.Length, TheParams.Length - TheSeperator.Length).Partial
        Else
            RemoveArgP.ConcatBytes TheParams.Partial
        End If
    End If
End Function


Public Function RemoveNextArg(ByRef TheParams, ByRef TheSeperator, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True)
    If VBA.TypeName(TheParams) = "String" And VBA.TypeName(TheSeperator) = "String" Then
        Dim str1 As String
        str1 = TheParams
        RemoveNextArg = modCommon.RemoveNextArg(str1, TheSeperator, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare), TrimResult)
        TheParams = str1
    ElseIf VBA.TypeName(TheParams) = "Stream" And VBA.TypeName(TheSeperator) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Set arg1 = TheParams
        Set arg2 = TheSeperator
        Set RemoveNextArg = RemoveNextArgP(arg1, arg2, CaseSensitive, TrimResult)
        Set arg1 = Nothing
        Set arg2 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function

Friend Function RemoveNextArgP(ByRef TheParams As Stream, ByRef TheSeperator As Stream, Optional ByVal CaseSensitive = True, Optional ByVal TrimResult As Boolean = True) As Stream
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    Set RemoveNextArgP = NextArgP(TheParams, TheSeperator, CaseSensitive, TrimResult)
    Set TheParams = RemoveArgP(TheParams, TheSeperator, CaseSensitive, TrimResult)

End Function


Friend Function NextQuotedArgP(ByRef TheParams As Stream, Optional ByRef BeginQuote As Stream = Nothing, Optional ByRef EndQuote As Stream = Nothing, Optional ByVal Embeded As Boolean = False, Optional ByVal CaseSensitive = True) As Stream

    If BeginQuote Is Nothing Then
        Set BeginQuote = New Stream
        BeginQuote.Post Asc("""")
    End If
    If EndQuote Is Nothing Then
        Set EndQuote = New Stream
        EndQuote.Post Asc("""")
    End If
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    Dim newParams As New Stream
    If TheParams.Length > 0 Then newParams.ConcatBytes TheParams.Partial
    Set NextQuotedArgP = RemoveQuotedArgP(newParams, BeginQuote, EndQuote, Embeded, CaseSensitive)
    Set newParams = Nothing
    Set BeginQuote = Nothing
    Set EndQuote = Nothing
End Function

Public Function RemoveQuotedArg(ByRef TheParams, Optional ByRef BeginQuote, Optional ByRef EndQuote, Optional ByVal Embedded As Boolean = False, Optional ByVal CaseSensitive = True)
    If VBA.TypeName(TheParams) = "String" Then
        Dim str1 As String
        str1 = TheParams
        RemoveQuotedArg = modCommon.RemoveQuotedArg(str1, IIf(IsMissing(BeginQuote), """", BeginQuote), IIf(IsMissing(EndQuote), """", EndQuote), Embedded, IIf(CaseSensitive, vbBinaryCompare, vbTextCompare))
        TheParams = str1
    ElseIf VBA.TypeName(TheParams) = "Stream" Then
        Dim arg1 As Stream
        Dim arg2 As Stream
        Dim arg3 As Stream
        Set arg1 = TheParams
        Set arg2 = IIf(IsMissing(BeginQuote), Nothing, BeginQuote)
        Set arg3 = IIf(IsMissing(EndQuote), Nothing, EndQuote)
        Set RemoveQuotedArg = RemoveQuotedArgP(arg1, arg2, arg3, Embedded, CaseSensitive)
        Set TheParams = arg1
        Set arg1 = Nothing
        Set arg2 = Nothing
        Set arg3 = Nothing
    Else
        Err.Raise 5, App.Title, "Invalid procedure call or argument"
    End If
End Function



Friend Function RemoveQuotedArgP(ByRef TheParams As Stream, Optional ByRef BeginQuote As Stream = Nothing, Optional ByRef EndQuote As Stream = Nothing, Optional ByVal Embeded As Boolean = False, Optional ByVal CaseSensitive = True) As Stream

    If BeginQuote Is Nothing Then
        Set BeginQuote = New Stream
        BeginQuote.Post Asc("""")
    End If
    If EndQuote Is Nothing Then
        Set EndQuote = New Stream
        EndQuote.Post Asc("""")
    End If
    If VBA.TypeName(CaseSensitive) = "Long" Then
        CaseSensitive = (CaseSensitive = VbCompareMethod.vbBinaryCompare)
    End If
    
    Dim retVal As New Stream
    Dim newVal As New Stream
    Dim newParams As New Stream
    Dim X As Long
    X = InStrP(1, TheParams, BeginQuote, CaseSensitive)
    If (X > 0) And (X < TheParams.Length) Then
        If (InStrP(X + BeginQuote.Length, TheParams, EndQuote, CaseSensitive) > 0) Then
            If (Not Embeded) Or (EndQuote = BeginQuote) Then
                retVal.Reset
                retVal.ConcatBytes MidP(TheParams, X + BeginQuote.Length).Partial

                newParams.ConcatBytes LeftP(TheParams, X - 1).Partial
                newParams.ConcatBytes MidP(retVal, InStrP(1, retVal, EndQuote, CaseSensitive) + EndQuote.Length).Partial
                TheParams.Clone newParams
                newParams.Reset
                
                newParams.Reset
                newParams.ConcatBytes LeftP(retVal, InStrP(1, retVal, EndQuote, CaseSensitive) - 1).Partial
                retVal.Clone newParams
                newParams.Reset
            Else
            
                Dim L As Long
                Dim Y As Long
                L = 1
                Y = X
                Do Until L = 0
                    If (InStrP(Y + BeginQuote.Length, TheParams, BeginQuote, CaseSensitive) > 0) And (InStrP(Y + BeginQuote.Length, TheParams, BeginQuote, CaseSensitive) < InStrP(Y + BeginQuote.Length, TheParams, EndQuote, CaseSensitive)) Then
                        L = L + 1
                        Y = InStrP(Y + BeginQuote.Length, TheParams, BeginQuote, CaseSensitive)
                    ElseIf (InStrP(Y + BeginQuote.Length, TheParams, EndQuote, CaseSensitive) > 0) Then
                        L = L - 1
                        Y = InStrP(Y + EndQuote.Length, TheParams, EndQuote, CaseSensitive)
                    Else
                        Y = TheParams.Length
                        L = 0
                    End If
                Loop
                retVal.Reset
                retVal.ConcatBytes MidP(TheParams, X + BeginQuote.Length).Partial
                
                newParams.ConcatBytes LeftP(TheParams, X - 1).Partial
                newParams.ConcatBytes MidP(retVal, (Y - X) + EndQuote.Length).Partial
                TheParams.Clone newParams
                newParams.Reset

                newParams.ConcatBytes LeftP(retVal, (Y - X) - 1).Partial
                retVal.Clone newParams
                newParams.Reset
            End If
        End If
    End If
    Set RemoveQuotedArgP = retVal
    Set newParams = Nothing
    Set retVal = Nothing
End Function

Private Sub Class_Initialize()
    Set pGreaterThan = New NTStreams10.Stream
    Set pLesserThan = New NTStreams10.Stream
    Set pLeftBracket = New NTStreams10.Stream
    Set pRightBracket = New NTStreams10.Stream
    Set pEqualSign = New NTStreams10.Stream
    Set pDoubleQuote = New NTStreams10.Stream
    Set pSingleQuote = New NTStreams10.Stream
    Set pColon = New NTStreams10.Stream
    Set pLeftBrace = New NTStreams10.Stream
    Set pRightBrace = New NTStreams10.Stream
    Set pSpaceChar = New NTStreams10.Stream
    Set pTabIndent = New NTStreams10.Stream
    Set pCr = New NTStreams10.Stream
    Set pLf = New NTStreams10.Stream
    Set pCrLf = New NTStreams10.Stream
    
    pGreaterThan = ">"
    pLesserThan = "<"
    pLeftBracket = "["
    pRightBracket = "]"
    pEqualSign = "="
    pDoubleQuote = """"
    pSingleQuote = "'"
    pColon = ":"
    pLeftBrace = "{"
    pRightBrace = "}"
    pSpaceChar = " "
    pTabIndent = vbTab
    pCr = vbCr
    pLf = vbLf
    pCrLf = vbCrLf
End Sub

Private Sub Class_Terminate()
    Set pGreaterThan = Nothing
    Set pLesserThan = Nothing
    Set pLeftBracket = Nothing
    Set pRightBracket = Nothing
    Set pEqualSign = Nothing
    Set pDoubleQuote = Nothing
    Set pSingleQuote = Nothing
    Set pColon = Nothing
    Set pLeftBrace = Nothing
    Set pRightBrace = Nothing
    Set pSpaceChar = Nothing
    Set pTabIndent = Nothing
    Set pCr = Nothing
    Set pLf = Nothing
    Set pCrLf = Nothing
End Sub
