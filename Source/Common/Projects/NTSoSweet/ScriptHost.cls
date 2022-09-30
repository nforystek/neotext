VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Public Enum ParserCodes
    'Error Messages
    ErrMsg_00 = 0
    ErrMsg_01 = 1
    ErrMsg_02 = 2
    ErrMsg_03 = 3
    ErrMsg_04 = 4
    ErrMsg_05 = 5
    ErrMsg_06 = 6
    ErrMsg_07 = 7
    ErrMsg_08 = 8
    ErrMsg_09 = 9
    ErrMsg_10 = 10
    ErrMsg_11 = 11
    ErrMsg_12 = 12
    ErrMsg_13 = 13
    ErrMsg_14 = 14
    ErrMsg_15 = 15
    ErrMsg_16 = 16
    ErrMsg_17 = 17
    ErrMsg_18 = 18
    ErrMsg_19 = 19
    ErrMsg_20 = 20
    
    'Interpeter States
    Stages_00 = 21
    Stages_01 = 22
    Stages_02 = 23
    Stages_03 = 24
    Stages_04 = 25
    Stages_05 = 26
    Stages_06 = 27
    Stages_07 = 28
    Stages_08 = 29
    Stages_09 = 30
    Stages_10 = 31
    Stages_11 = 32
    Stages_12 = 33
    Stages_13 = 34
    Stages_14 = 35
    Stages_15 = 36
    Stages_16 = 37
    Stages_17 = 38
    Stages_18 = 39
    Stages_19 = 40
    Stages_20 = 41
    
    'Expected Syntax
    Syntax_00 = 42
    Syntax_01 = 43
    Syntax_02 = 44
    Syntax_03 = 45
    Syntax_04 = 46
    Syntax_05 = 47
    Syntax_06 = 48
    Syntax_07 = 49
    Syntax_08 = 50
    Syntax_09 = 51
    Syntax_10 = 52
    Syntax_11 = 53
    Syntax_12 = 54
    Syntax_13 = 55
    Syntax_14 = 56
    Syntax_15 = 57
    Syntax_16 = 58
    Syntax_17 = 59
    Syntax_18 = 60
    Syntax_19 = 61
    Syntax_20 = 62
End Enum

Private xColor As Colors

Private xHoops As Hoops
Private xReams As Reams

Private xWraps As Wraps
Private xLines As Lines

Private xOrderly() As Long

Private LineNum As Long

Private StateMajor As Integer
Private StateMinor As Integer
Private StateError As Boolean

Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)

Public Enum OrderlyInfo
    HostLinesTotal = 0
    HoopLinesTotal = -1
    ReamLinesTotal = -2
End Enum

Public Event HostUpdate()

Public Property Get Orderly(Optional ByVal Index As Long = 0)
    If ((UBound(xOrderly) > 0) And Index <= UBound(xOrderly)) Or (Index >= -2 And Index <= 0) Then
        If Index = 0 Then
            Orderly = UBound(xOrderly)
        ElseIf Index = -1 Then
            Orderly = xOrderly(0)
        ElseIf Index = -2 Then
            Orderly = UBound(xOrderly) - xOrderly(0)
        ElseIf Index > 0 And Index <= UBound(xOrderly) Then
            Dim lPtr As Long
            lPtr = xOrderly(Index)
        
            Dim lZero As Long
            Dim NewObj As Object
            RtlMoveMemory NewObj, lPtr, 4&
            Set Orderly = NewObj
            RtlMoveMemory NewObj, lZero, 4&
        Else
            Err.Raise 8, App.EXEName, "Invalid index bounds; use 0 for total count, -1 or count of Hoops, -2 for count of Reams."
        End If
    'Else
    '    Err.Raise 8, App.EXEName, "Script host not loaded; use 0 for total count, -1 or count of Hoops, -2 for count of Reams."
    End If
End Property
Public Property Get Colors() As Colors
    Set Colors = xColor
End Property

Public Property Get Hoops() As Hoops
    Set Hoops = xHoops
End Property

Public Property Get Reams() As Reams
    Set Reams = xReams
End Property

Private Property Get ErrorMessage(ByVal inLine As String) As String
    StateError = True
    ErrorMessage = ParserMessages(Val(CStr(IIf(StateMajor <> 1, "0", "1") & StateMinor + StateErrors))) & IIf(inLine <> "", " - [" & inLine & "]", "")
End Property
Public Property Get Helper(Optional ByVal CommandNumber As Long = -2) As String

    If CommandNumber < 0 Or CommandNumber > 20 Then
        If StateError Then
            Helper = ParserMessages(CLng(IIf(StateMajor <> 1, 0, 1) & StateMinor) + StateErrors)
            StateError = False
        Else
            Helper = ParserMessages(CLng(IIf(StateMajor <> 1, 0, 1) & StateMinor) + StateStages)
        End If
    Else
        Helper = ParserMessages(CommandNumber + StateSyntax)
    End If
    
End Property

Private Function StateFromLine(ByRef Items As Object, ByVal inLine As String) As Boolean
    If StateMajor = -1 Then
        If (LCase(NextArg(inLine, " ")) = "hoops" Or LCase(NextArg(inLine, " ")) = "wraps") And StateMinor = 0 Then
            StateMajor = 0
            StateMinor = 1
            RemoveNextArg inLine, " "

            Set Hoops.Color = ColorFromLine(inLine, True)
            Colors.Add Hoops.Color
            
            StateMinor = 2
            StateFromLine = True
        'Else
        '    Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
        End If
    ElseIf StateMajor = 0 Then
        If (LCase(NextArg(inLine, " ")) = "reams") Or (LCase(NextArg(inLine, " ")) = "lines") Then
            StateMajor = 1
            StateMinor = 1

            Set Reams.Color = ColorFromLine(RemoveArg(inLine, " "), True)
            Colors.Add Reams.Color
            
            StateMinor = 2
        ElseIf (StateMinor = 2) Then
            Set Items = New Hoops
            Items.Add WrapFromLine(inLine)
            Items.Add RushFromLine(inLine)
            Items.Add PushFromLine(inLine)
            
        End If
        StateFromLine = True
    ElseIf StateMajor = 1 Then
        If LCase(NextArg(inLine, " ")) = "eof" Then
            StateMajor = 2
        ElseIf (StateMinor = 2) Then
            Set Items = New Reams
            Items.Add LineFromLine(inLine)
            Items.Add RushFromLine(inLine)
            Items.Add PushFromLine(inLine)
            
        End If
        StateFromLine = True
    'Else
    '    Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
    End If
End Function

Private Function ColorFromLine(ByRef inLine As String, Optional ByVal Background As Boolean = False) As Color
    If LCase(NextArg(inLine, " ")) = "color" Then
        RemoveNextArg inLine, " "
        Set ColorFromLine = GetNewColorner(RemoveNextArg(inLine, " "), Background)
    End If
End Function

Private Function WrapFromLine(ByRef inLine As String) As Object
    If (LCase(NextArg(inLine, " ")) = "wrap") Or (LCase(NextArg(inLine, " ")) = "hoop") Then
        StateMinor = 2
        RemoveNextArg inLine, " "
        Dim W As New Wraps
        Dim Wrap As New Wrap
        Wrap.Rawvalue = RemoveNextArg(inLine, " ")
        W.Add Wrap
'        Wrap.Mushes.Add MushFromLine(inLine)
'        Wrap.Pushes.Add PushFromLine(inLine)
        W.Add MushFromLine(inLine)
        W.Add PushFromLine(inLine)
        W.Add MushFromLine(inLine)
        W.Add RushFromLine(inLine)
        W.Add MushFromLine(inLine)
        
        If (LCase(NextArg(inLine, " ")) = "wrap") Or (LCase(NextArg(inLine, " ")) = "hoop") Then
            RemoveNextArg inLine, " "
            Wrap.Endvalue = RemoveNextArg(inLine, " ")
        Else
            Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
        End If
        If Wrap.Rawvalue = "eol" Then Wrap.Rawvalue = Chr(13)
        If Wrap.Rawvalue = "bol" Then Wrap.Rawvalue = Chr(10)
        If Wrap.Endvalue = "eol" Then Wrap.Endvalue = Chr(13)
        If Wrap.Endvalue = "bol" Then Wrap.Endvalue = Chr(10)
        Set Wrap.Color = ColorFromLine(inLine)
        Colors.Add Wrap.Color
        Set WrapFromLine = W
        StateMinor = 4
    End If
End Function

Private Function LineFromLine(ByRef inLine As String) As Object

    If (LCase(NextArg(inLine, " ")) = "line") Or (LCase(NextArg(inLine, " ")) = "ream") Then
        StateMinor = 2
        RemoveNextArg inLine, " "
        Dim l As New Lines
        Dim Line As New Line
        Line.Rawvalue = RemoveNextArg(inLine, " ")
        l.Add Line
'        Line.Rushes.Add RushFromLine(inLine)
'        Line.Mushes.Add MushFromLine(inLine)
        l.Add MushFromLine(inLine)
        l.Add PushFromLine(inLine)
        l.Add MushFromLine(inLine)
        l.Add RushFromLine(inLine)
        l.Add MushFromLine(inLine)

        
        If (LCase(NextArg(inLine, " ")) = "line") Or (LCase(NextArg(inLine, " ")) = "ream") Then
            RemoveNextArg inLine, " "
            Line.Endvalue = RemoveNextArg(inLine, " ")
        Else
            Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
        End If
        If Line.Rawvalue = "eol" Then Line.Rawvalue = Chr(13)
        If Line.Rawvalue = "bol" Then Line.Rawvalue = Chr(10)
        If Line.Endvalue = "eol" Then Line.Endvalue = Chr(13)
        If Line.Endvalue = "bol" Then Line.Endvalue = Chr(10)
        Set Line.Color = ColorFromLine(inLine)
        Colors.Add Line.Color
        Set LineFromLine = l
        StateMinor = 4
    End If
End Function

Private Function PushFromLine(ByRef inLine As String) As Object
    If (LCase(NextArg(inLine, " ")) = "push") Then
        StateMinor = 3
        RemoveNextArg inLine, " "
        Dim P As New Pushes
        Dim Push As New Push
        P.Rawvalue = RemoveNextArg(inLine, " ")
        P.Add Push
'        Push.Mushes.Add MushFromLine(inLine)
'        Push.Rushes.Add RushFromLine(inLine)
        P.Add MushFromLine(inLine)
        P.Add RushFromLine(inLine)
        P.Add MushFromLine(inLine)
        
        Select Case Left(LCase(NextArg(inLine, " ")), 4)
            Case "ream", "hoop", "wrap", "line"
                RemoveNextArg inLine, " "
                Push.Endvalue = RemoveNextArg(inLine, " ")
            Case Else
                Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
        End Select
        If Push.Rawvalue = "eol" Then Push.Rawvalue = Chr(13)
        If Push.Rawvalue = "bol" Then Push.Rawvalue = Chr(10)
        If Push.Endvalue = "eol" Then Push.Endvalue = Chr(13)
        If Push.Endvalue = "bol" Then Push.Endvalue = Chr(10)
        Set Push.Color = ColorFromLine(inLine)
        Colors.Add Push.Color
        Set PushFromLine = P
        StateMinor = 4
    End If
End Function

Private Function RushFromLine(ByRef inLine As String) As Object
    If (LCase(NextArg(inLine, " ")) = "rush") Then
        StateMinor = 3
        RemoveNextArg inLine, " "
        Dim R As New Rushes
        Dim Rush As New Rush
        Rush.Rawvalue = RemoveNextArg(inLine, " ")
        R.Add Rush
'        Rush.Pushes.Add PushFromLine(inLine)
'        Rush.Mushes.Add MushFromLine(inLine)
        R.Add MushFromLine(inLine)
        R.Add PushFromLine(inLine)
        R.Add MushFromLine(inLine)
        
        Select Case Left(LCase(NextArg(inLine, " ")), 4)
            Case "ream", "hoop", "line", "ream"
                RemoveNextArg inLine, " "
                Rush.Endvalue = RemoveNextArg(inLine, " ")
            Case Else
                Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
        End Select
        If Rush.Rawvalue = "eol" Then Rush.Rawvalue = Chr(13)
        If Rush.Rawvalue = "bol" Then Rush.Rawvalue = Chr(10)
        If Rush.Endvalue = "eol" Then Rush.Endvalue = Chr(13)
        If Rush.Endvalue = "bol" Then Rush.Endvalue = Chr(10)
        Set Rush.Color = ColorFromLine(inLine)
        Colors.Add Rush.Color
        Set RushFromLine = R
        StateMinor = 4
    End If

End Function

Private Function MushFromLine(ByRef inLine As String) As Object
    If (LCase(NextArg(inLine, " ")) = "mush") Then
        StateMinor = 3
        RemoveNextArg inLine, " "
        Dim m As New Mushes
        Dim Mush As Mush
        Do
            Set Mush = New Mush
            Mush.Rawvalue = RemoveNextArg(inLine, " ")
            If Mush.Rawvalue = "eol" Then Mush.Rawvalue = Chr(13)
            If Mush.Rawvalue = "bol" Then Mush.Rawvalue = Chr(10)
            Set Mush.Color = Colors(Colors.Count)
            m.Add Mush
            Set Mush = Nothing
        Loop Until FinishStateLines(inLine) Or (inLine = "")
        Set MushFromLine = m
        StateMinor = 4
    End If

End Function

Private Function WordFromLine(ByRef inLine As String) As Boolean
    If (StrReverse(NextArg(StrReverse(inLine), " ")) = "words") Then
        Colors.Add ColorFromLine(inLine)
        If LCase(RemoveNextArg(inLine, " ")) = "words" Then
            WordFromLine = True
        End If
    End If
End Function

Private Function RainFromLine(ByRef inLine As String) As Boolean
    If (StrReverse(NextArg(StrReverse(inLine), " ")) = "rains") Then
        Colors.Add ColorFromLine(inLine)
        If LCase(RemoveNextArg(inLine, " ")) = "rains" Then
            RainFromLine = True
        End If
    End If
End Function

Private Function DigitFromLine(ByRef inLine As String) As Boolean
    If (StrReverse(NextArg(StrReverse(inLine), " ")) = "digit") Then
        Colors.Add ColorFromLine(inLine)
        If LCase(RemoveNextArg(inLine, " ")) = "digit" Then
            DigitFromLine = True
        End If
    End If
End Function

Private Sub DataFromLines(ByRef Items As Object, ByRef inLine As String)
    Dim conv As String
    Dim Temp As String
    
    If StateMinor = 6 Then
        Dim iobj1 As Word
        Do Until (Trim(inLine) = "")
            Set iobj1 = New Word
            If NextArg(inLine, " ", True) = "word" Then
                RemoveNextArg inLine, True
                iobj1.Rawvalue = inLine
                iobj1.Endvalue = iobj1.Rawvalue
            Else
                iobj1.Rawvalue = RemoveNextArg(inLine, " ", True)
                If NextArg(inLine, " ", True) = "word" Then
                    RemoveNextArg inLine, "word", True
                    iobj1.Endvalue = RemoveNextArg(inLine, " ", True)
                Else
                    iobj1.Endvalue = iobj1.Rawvalue
                End If
            End If
            Set iobj1.Color = Colors(Colors.Count)
            Items.Add iobj1
            Set iobj1 = Nothing
        Loop

    ElseIf StateMinor = 7 Then
        Dim iobj2 As Rain
        'many
        Do Until (Trim(inLine) = "")
            Set iobj2 = New Rain
            iobj2.Rawvalue = Left(inLine, 1)
            iobj2.Endvalue = iobj2.Rawvalue
            Set iobj2.Color = Colors(Colors.Count)
            inLine = Mid(inLine, 2)
            Items.Add iobj2
            Set iobj2 = Nothing
        Loop
    ElseIf StateMinor = 8 Then
        Dim iobj3 As Digit
        'many
        Do Until (Trim(inLine) = "")
            Set iobj3 = New Digit
            iobj3.Rawvalue = RemoveNextArg(inLine, " ")
            iobj3.Endvalue = iobj3.Rawvalue
            Set iobj3.Color = Colors(Colors.Count)
            Items.Add iobj3
            Set iobj3 = Nothing
        Loop
    End If
End Sub

Private Function FinishStateLines(ByVal inLine As String) As Boolean
    If (StateMinor >= 6 And StateMinor <= 8) Or (StateMinor = 2) Or (StateMinor = 3) Then
        Select Case LCase(RemoveNextArg(inLine, " "))
            Case "color", "reams", "ream", "hoop", "hoops", "eol", "eof", "words", "rains", "word", "rain", "digit", "wrap", "wraps", "lines", "line", "rush", "mush", "push", ""
                FinishStateLines = True
        End Select
    End If
End Function

Public Sub Interpreter(ByVal FileOrText As String)
    On Error Resume Next
    Dim tmp As Long
    tmp = FileSize(FileOrText)
    If tmp > 0 And Err.Number = 0 Then
        On Error GoTo 0
        LoadFromFile FileOrText
    Else
        Err.Clear
        On Error GoTo 0
        LoadFromText FileOrText
    End If
End Sub

Friend Sub LoadFromFile(ByVal FileName As String)
    LoadFromText ReadFile(FileName)
End Sub

Friend Sub LoadFromText(ByVal Text As String)
    Reset
    On Error GoTo ThrowException
    
    Dim Items As Object
    Dim Obj As Object
    Dim Rain As Rains
    Dim Digit As Digits
    Dim inLine As String
    Text = Replace(Replace(Text, vbCrLf, vbLf), vbCr, "")
    Do While InStr(Text, vbLf & vbLf) > 0
        Text = Replace(Text, vbLf & vbLf, vbLf)
    Loop
    Do Until (Text = "")
        inLine = RemoveNextArg(Text, vbLf)
        LineNum = LineNum + 1
        If (inLine <> "") Then
            If (StateMinor >= 6 And StateMinor <= 8) Or (StateMinor = 2) Then
                If FinishStateLines(inLine) Or inLine = "" Then
                    If StateMajor = 0 Then
                        If Not Items Is Nothing Then
                            If Items.Count > 0 Then
                                For Each Obj In Items
                                    ReDim Preserve xOrderly(0 To UBound(xOrderly) + 1) As Long
                                    xOrderly(UBound(xOrderly)) = ObjPtr(Obj)
                                    xOrderly(0) = UBound(xOrderly)
                                Next
                                Hoops.Add Items
                                Set Items = Nothing
                            End If
                        End If
                    ElseIf StateMajor = 1 Then
                        If Not Items Is Nothing Then
                            If Items.Count > 0 Then
                                For Each Obj In Items
                                    ReDim Preserve xOrderly(0 To UBound(xOrderly) + 1) As Long
                                    xOrderly(UBound(xOrderly)) = ObjPtr(Obj)
                                Next
                                Reams.Add Items
                                Set Items = Nothing
                            End If
                        End If
                    End If
                    StateMinor = 2
                Else
                    DataFromLines Items, inLine
                End If
            End If
            If Not StateFromLine(Items, inLine) Then Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
            If Not (StateMinor >= 6 And StateMinor <= 8) Then
                StateMinor = 5
                If WordFromLine(inLine) Then
                    Set Items = Nothing
                    Set Items = New Words
                    StateMinor = 6
                ElseIf RainFromLine(inLine) Then
                    Set Items = Nothing
                    Set Items = New Rains
                    StateMinor = 7
                ElseIf DigitFromLine(inLine) Then
                    Set Items = Nothing
                    Set Items = New Digits
                    StateMinor = 8
                Else
                    StateMinor = 2
                End If
            End If

        End If
    Loop
    If (StateMajor = 0) Then
ThrowException:
        Err.Raise 1002, , "Syntax Error" & IIf(LineNum > 0, " in row " & LineNum & vbCrLf, vbCrLf) & ErrorMessage(inLine)
    Else
        RaiseEvent HostUpdate
    End If

End Sub

'Public Function Generate() As String
'    StateMajor = -1
'    StateMinor = 0
'    StateError = False
'    Dim outText As String
'    Dim Obj As Object
'    Dim tmp As Long
'    Dim str As String
'    outText = outText & "hoops color " & Hoops.Color.Rawvalue & vbCrLf
'    Dim cnt As Long
'    cnt = 1
'    Do While cnt <= Hoops.Count
'        Set Obj = Hoops(cnt)
'        Select Case Typename(Obj)
'            Case "Word", "Words"
'                If InStr(Obj.Rawvalue, " ") > 0 Then
'                    outText = outText & "color " & Obj.Color & " words" & vbCrLf
'                Else
'                    outText = outText & "color " & Obj.Color & " words" & vbCrLf
'                    If Obj.Rawvalue <> Obj.Endvalue Then
'                        outText = outText & Obj.Rawvalue & " word " & Obj.Endvalue & vbCrLf
'                    Else
'                        outText = outText & Obj.Rawvalue & vbCrLf
'                    End If
'                End If
'            Case "Rain", "Rains"
'                If InStr(Obj.Rawvalue, " ") > 0 Then
'                    outText = outText & "color " & Obj.Color & " rains" & vbCrLf
'                Else
'                    outText = outText & "color " & Obj.Color & " rains" & vbCrLf
'                    If Obj.Rawvalue <> Obj.Endvalue Then
'                        outText = outText & Obj.Rawvalue & " rain " & Obj.Endvalue & vbCrLf
'                    Else
'                        outText = outText & Obj.Rawvalue & vbCrLf
'                    End If
'                End If
'            Case "Digit", "Digits"
'                If InStr(Obj.Rawvalue, " ") > 0 Then
'                    outText = outText & "color " & Obj.Color & " digits" & vbCrLf
'                Else
'                    outText = outText & "color " & Obj.Color & " digits" & vbCrLf
'                    If Obj.Rawvalue <> Obj.Endvalue Then
'                        outText = outText & Obj.Rawvalue & " digit " & Obj.Endvalue & vbCrLf
'                    Else
'                        outText = outText & Obj.Rawvalue & vbCrLf
'                    End If
'                End If
'            Case "Wrap", "Rush"
'                outText = outText & LCase(Typename(Obj)) & " " & Obj.Rawvalue & " "
'                If cnt < Hoops.Count Then
'                    tmp = cnt
'                    Do While (Typename(Reams(cnt)) = "Push" Or _
'                             Typename(Reams(cnt)) = "Rush" Or _
'                             Typename(Reams(cnt)) = "Mush") And _
'                              (cnt <= Hoops.Count)
'
'                             outText = outText & LCase(Typename(Hoops(cnt))) & " " & Obj.Rawvalue 'Hoops(cnt).rawlvaue
'                        cnt = cnt + 1
'                    Loop
'                    cnt = tmp
'                End If
'                outText = outText & "hoop " & Obj.Endvalue & " "
'                outText = outText & "color " & Obj.Color.Rawvalue & vbCrLf
'
'        End Select
'        cnt = cnt + 1
'    Loop
'    outText = outText & "reams color " & Reams.Color.Rawvalue & vbCrLf
'
'
'    cnt = 1
'    Do While cnt <= Reams.Count
'        Set Obj = Reams(cnt)
'        Select Case Typename(Obj)
''            Case "Word", "Rain", "Digit"
''                str = Typename(Obj)
''                If InStr(Obj.Rawvalue, " ") > 0 Then
''                    outText = outText & "color " & Obj.Color.Rawvalue & " " & LCase(str) & IIf(Typename(Obj) <> "Digit", "s", "") & vbCrLf
''                Else
''                    outText = outText & "color " & Obj.Color.Rawvalue & " " & LCase(str) & IIf(Typename(Obj) <> "Digit", "s", "") & vbCrLf
''                End If
''
''                cnt = cnt + 1
''
''                If cnt <= Reams.Count Then
''                    Do While (cnt <= Reams.Count)
''                        If (Typename(Reams(cnt)) <> str) Then
''                            cnt = cnt - 1
''                            Exit Do
''                        End If
''                        If InStr(Reams(cnt).Rawvalue, " ") > 0 Then
''                            outText = outText & LCase(str) & " " & Reams(cnt).Rawvalue & vbCrLf
''                        Else
''                            If Reams(cnt).Rawvalue <> Reams(cnt).Endvalue And Not Reams(cnt).Endvalue = "" Then
''                                outText = outText & Reams(cnt).Rawvalue & " " & LCase(str) & " " & Reams(cnt).Endvalue & vbCrLf
''                            Else
''                                outText = outText & Reams(cnt).Rawvalue & vbCrLf
''                            End If
''                        End If
''                        cnt = cnt + 1
''                    Loop
''                End If
'
'            Case "Line", "Rush", "Push"
'                outText = outText & LCase(Typename(Reams(cnt))) & " " & Obj.Rawvalue & " "
'                If cnt <= Reams.Count Then
'                    tmp = cnt
'                    Do While (Typename(Reams(cnt)) = "Push" Or _
'                             Typename(Reams(cnt)) = "Rush" Or _
'                             Typename(Reams(cnt)) = "Mush") And _
'                              (cnt < Reams.Count)
'                             outText = outText & LCase(Typename(Reams(cnt))) & " " & Reams(cnt).Rawvalue
'                        cnt = cnt + 1
'                    Loop
'                    cnt = tmp
'                End If
'                outText = outText & "ream " & Obj.Endvalue & " "
'                outText = outText & "color " & Obj.Color.Rawvalue & vbCrLf
'            Case Else
'                'Debug.Print Typename(Obj)
'                cnt = cnt + 1
'        End Select
'        cnt = cnt + 1
'    Loop
'    Generate = outText
'End Function

Public Sub Reset()
    StateMajor = -1
    StateMinor = 0
    StateError = False
    LineNum = 0
    xReams.Clear
    xHoops.Clear
    xColor.Clear
    xWraps.Clear
    xLines.Clear
    ReDim xOrderly(0 To 0) As Long
    xOrderly(0) = 0
End Sub
Private Sub Class_Initialize()
    StateMajor = -1
    StateMinor = 0
    StateError = False
    ReDim xOrderly(0 To 0) As Long
    Set xColor = New Colors
    Set xHoops = New Hoops
    Set xReams = New Reams
    Set xWraps = New Wraps
    Set xLines = New Lines
    LoadFromText "hoops color #000000" & vbCrLf & _
                "reams color #000000" & vbCrLf & _
                "eof" & vbCrLf
End Sub

Private Sub Class_Terminate()
    Set xReams = Nothing
    Set xHoops = Nothing
    Set xWraps = Nothing
    Set xLines = Nothing
    Set xColor = Nothing
    Erase xOrderly
End Sub



