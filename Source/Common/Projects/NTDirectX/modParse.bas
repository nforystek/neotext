Attribute VB_Name = "modParse"
Option Explicit

'#################################################################
'### These objects made public are added to ScriptControl also ###
'### are the only thing global exposed to modFactory.Execute   ###
'#################################################################

Public ScriptRoot As String
Public Include As New Include

Public All As New ntnodes10.Collection

Public Brilliants As New Brilliants
Public Molecules As New Molecules

Public Billboards As New Billboards
Public Planets As New Planets

Public Motions As New Motions

Public OnEvents As New ntnodes10.Collection
Public Bindings As New Bindings
Public Camera As New Camera

Public Function ParseQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal KeepDelim As Boolean = True) As String
    Dim retVal As String
    Dim X As Long
    X = InStr(1, TheParams, BeginQuote, Compare)
    If (X > 0) And (X < Len(TheParams)) Then
        If (InStr(X + Len(BeginQuote), TheParams, EndQuote, Compare) > 0) Then
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
            If KeepDelim Then
                retVal = Left(TheParams, X - 1) & Mid(TheParams, X)
                TheParams = Mid(retVal, (X - 1) + (Y - X) + Len(EndQuote) + Len(BeginQuote))
                retVal = Left(retVal, (X - 1) + (Y - X) + Len(BeginQuote))
            Else
                retVal = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retVal, (Y - X) + Len(EndQuote))
                retVal = Left(retVal, (Y - X) - 1)
            End If
        End If
    End If
    ParseQuotedArg = retVal
End Function
Private Function ParseReservedWord(ByVal txt As String) As Integer
    Dim inWord As String
    Do While IsAlphaNumeric(Left(txt, 1))
        inWord = inWord & Left(txt, 1)
        txt = Mid(txt, 2)
    Loop
    Select Case LCase(inWord)
        Case "oninrange", "onoutrange", "oncollide"
            ParseReservedWord = 1
        Case "molecule", "method", "brilliant", "serialize", "deserialize", "motion", "billboard", "planet"
            ParseReservedWord = 3
        Case "bindings", "camera"
            ParseReservedWord = -3
        Case Else
            ParseReservedWord = IIf(GetBindingIndex(inWord) > -1, 2, 0)
    End Select
End Function
Private Function ParseBracketOff(ByVal txt As String, ByVal StartDelim As String, ByVal StopDelim As String) As String
    txt = ParseWhiteSpace(txt)
    If Left(txt, 1) = StartDelim And Right(txt, 1) = StopDelim Then
        txt = Mid(txt, 2)
        txt = Left(txt, Len(txt) - 1)
    End If
    ParseBracketOff = ParseWhiteSpace(txt)
End Function
Public Function ParseWhiteSpace(ByVal txt As String) As String
    Static stack As Boolean
    If Not stack Then
        stack = True
        txt = StrReverse(txt)
        txt = ParseWhiteSpace(txt)
        txt = StrReverse(txt)
        stack = False
    End If
    Do While Left(txt, 1) = " " Or Left(txt, 1) = vbTab Or Left(txt, 1) = vbCr Or Left(txt, 1) = vbLf
        txt = Mid(txt, 2)
    Loop
    ParseWhiteSpace = txt
End Function
Private Function ParseInWith(ByVal inLine As String, Optional ByVal inWith As String = "") As String
    Dim pos As Long
    Dim inName As String
    pos = InStr(inLine, ".")
    Do While (pos > 0)
        If (pos = 1) Then
            If (inWith <> "") Then
                inLine = inWith & inLine
                pos = pos + Len(inWith)
            End If
        ElseIf (Not IsAlphaNumeric(Mid(inLine, pos - 1, 1))) Then
            If pos + 1 <= Len(inLine) Then
                If Not IsNumeric(Mid(inLine, pos + 1, 1)) Then
                    If (inWith <> "") Then
                        inLine = Left(inLine, pos - 1) & inWith & Mid(inLine, pos)
                        pos = pos + Len(inWith)
                    End If
                End If
            ElseIf (inWith <> "") Then
                inLine = Left(inLine, pos - 1) & inWith & Mid(inLine, pos)
                pos = pos + Len(inWith)
            End If
        End If
        pos = InStr(pos + 1, inLine, ".")
    Loop
    ParseInWith = Replace(inLine, "Debug.Print", "DebugPrint", , , vbTextCompare)
End Function

Private Function ParseDeserialize(ByRef nXML As String) As String

    Dim xml As New MSXML.DOMDocument
    xml.async = "false"
    xml.loadXML nXML
    Dim retVal As String
    Dim tmp As String
    Dim cnt As Long
    Dim cnt2 As Long
    Dim cnt3 As Long
    Dim inName As String
    retVal = "Sub Deserialize()" & vbCrLf
    If xml.parseerror.errorCode = 0 Then
        Dim child As MSXML.IXMLDOMNode
        Dim serial As MSXML.IXMLDOMNode
        For Each serial In xml.childNodes
            Select Case LCase(Include.SafeKey(serial.baseName))
            
                Case "serial"
                    For Each child In serial.childNodes
                        Select Case LCase(Include.SafeKey(child.baseName))
                            Case "datetime"
                                'temporary start new each time
                               ' If (FileDateTime(ScriptRoot & "\Index.vbx") <> Include.URLDecode(child.Text)) And (Not (InStr(1, LCase(Command), "/debug", vbTextCompare) > 0)) Then
                                    GoTo exitout:
                               'End If
                            Case "variables"
                                For cnt = 0 To child.childNodes.Length - 1
                                    tmp = Replace(Replace(Include.URLDecode(child.childNodes(cnt).Text), """", """"""), vbCrLf, """ & vbCrLf & """)
                                    If IsNumeric(tmp) Or LCase(tmp) = "false" Or LCase(tmp) = "true" Then
                                        retVal = retVal & child.childNodes(cnt).baseName & " = " & tmp & vbCrLf
                                    Else
                                        retVal = retVal & child.childNodes(cnt).baseName & " = """ & tmp & """" & vbCrLf
                                    End If
                                Next
                            Case "bindings"
                                For cnt = 0 To child.childNodes.Length - 1
                                    retVal = retVal & "Bindings(" & GetBindingIndex(child.childNodes(cnt).baseName) & ") = Include.URLDecode(""" & child.childNodes(cnt).Text & """)" & vbCrLf
                                    'Bindings(GetBindingIndex(child.childNodes(cnt).baseName)) = Include.URLDecode(child.childNodes(cnt).Text)
                                Next
                                retVal = retVal & "Bindings.Serialize = True" & vbCrLf
                                'Bindings.Serialize = True
                            Case "camera"
                                retVal = retVal & Include.URLDecode(child.childNodes(cnt).Text) & vbCrLf
                            Case "molecules", "brilliants", "planets", "billboards"
                                All.Add "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf & child.xml & vbCrLf & "</Serial>" & vbCrLf, Include.SafeKey(child.baseName)
                                retVal = retVal & "Set Include.Serialize = " & Include.SafeKey(child.baseName) & vbCrLf
                        End Select
                    Next
            End Select
        Next
    End If
exitout:
    retVal = retVal & "End Sub" & vbCrLf
    ParseDeserialize = retVal
    Set xml = Nothing
End Function
Private Function ParseSerialize(ByRef txt As String, Optional ByRef LineNum As Long = 0) As String
    On Error Resume Next
    Dim retVal As String
    retVal = "Function Serialize()" & vbCrLf
    
    retVal = retVal & "Serialize = ""<?xml version=""""1.0""""?>"" & vbCrLf" & vbCrLf
    retVal = retVal & "Serialize = Serialize & ""<Serial>"" & vbCrLf" & vbCrLf
    retVal = retVal & "Serialize = Serialize & Include.Serialize(""    "")" & vbCrLf
        
    retVal = retVal & "Serialize = Serialize & ""    <Variables>"" & vbCrLf" & vbCrLf
    txt = ParseWhiteSpace(txt)
    Do Until txt = ""
        retVal = retVal & "Serialize = Serialize & ""        <" & NextArg(txt, vbCrLf) & ">"" & Include.URLEncode(" & NextArg(txt, vbCrLf) & ") & ""        </" & NextArg(txt, vbCrLf) & ">"" & vbCrLf" & vbCrLf
        RemoveNextArg txt, vbCrLf
    Loop
    retVal = retVal & "Serialize = Serialize & ""    </Variables>"" & vbCrLf" & vbCrLf
    retVal = retVal & "Serialize = Replace(Serialize,""    <Variables>"" & vbCrLf & ""    </Variables>"" & vbCrLf ,"""")" & vbCrLf
    
    retVal = retVal & "If Bindings.Serialize Then" & vbCrLf
    retVal = retVal & "Serialize = Serialize & Bindings.ToString(""    "")" & vbCrLf
    retVal = retVal & "End If" & vbCrLf
    
    retVal = retVal & "Serialize = Serialize & ""</Serial>""" & vbCrLf
   
    retVal = retVal & "End Function" & vbCrLf
    
    ParseSerialize = retVal
    If Err.Number = 11 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 453, "Line " & (LineNum - CountWord(txt, vbCrLf)), "Unable to serialize."
    Else
        On Error GoTo 0
    End If
End Function
Private Sub ParseExecute(ByRef txt As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0)
    If GetFileExt(NextArg(txt, vbCrLf)) = ".vbx" Then
        If PathExists(NextArg(txt, vbCrLf), True) Then
            ParseScript RemoveNextArg(txt, vbCrLf), inWith, (LineNum - CountWord(txt, vbCrLf))
        Else
            frmMain.ExecuteStatement ParseInWith(RemoveNextArg(txt, vbCrLf), inWith)
        End If
    Else
        frmMain.ExecuteStatement ParseInWith(RemoveNextArg(txt, vbCrLf), inWith)
    End If
End Sub
Private Sub ParseSetting(ByRef inLine As String, ByRef inValue As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0)
    'Debug.Print "Setting: " & inLine & " B: " & inValue
    If Left(inLine, 8) = "variable" Then
        Dim inName As String
        inName = ParseQuotedArg(inLine, "<", ">", , , False)
        frmMain.AddCode "Dim " & inName & vbCrLf
        inLine = inName
    End If
    inValue = ParseWhiteSpace(inValue)
    If ((Left(inValue, 1) = "[") And (Right(inValue, 1) = "]")) Then
        frmMain.ExecuteStatement ParseInWith(inLine, inWith) & "=""" & Replace(Replace(inValue, """", """"""), vbCrLf, """ & vbcrlf & """) & """"
    Else
        frmMain.ExecuteStatement ParseInWith(inLine, inWith) & "=" & inValue
    End If
End Sub

Private Sub ParseBindings(ByRef inBind As String, ByRef inBlock As String, Optional ByRef LineNum As Long = 0)
    On Error Resume Next
    Dim BindIndex As Long
    Dim bindCode As String
    inBind = ParseWhiteSpace(inBind)
    inBlock = ParseWhiteSpace(inBlock)
    
    BindIndex = GetBindingIndex(RemoveNextArg(inBind, "="))
    If Left(inBlock, 1) = "[" And Right(inBlock, 1) = "]" Then
        bindCode = ParseQuotedArg(inBlock, "[", "]", , , False)
    Else
        bindCode = inBlock
    End If
    
    If BindIndex > -1 And bindCode <> "" Then
        Bindings(BindIndex) = bindCode
    End If

    If Err.Number = 11 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 453, "Line " & (LineNum - CountWord(inBind & inBlock, vbCrLf)), "The binding specified was not recognized."
    Else
        On Error GoTo 0
    End If
End Sub
Private Sub ParseEvent(ByRef inLine As String, ByRef inBlock As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0)
    'Debug.Print "Event: " & inLine & " B: " & inBlock
    Dim inEvent As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inEvent = inEvent & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    inBlock = ParseBracketOff(inBlock, "[", "]")
    Dim inName As String
    inName = ParseQuotedArg(inBlock, "<", ">", , , False)
    inBlock = ParseInWith(inBlock, inWith)
    inBlock = """" & Replace(Replace(inBlock, """", """"""), vbCrLf, """ & vbCrLf & """) & """"
    frmMain.ExecuteStatement "Set " & inWith & "." & inEvent & " = Nothing"
    If inName <> "" Then frmMain.ExecuteStatement inWith & "." & inEvent & ".ApplyTo = """ & inName & """"
    frmMain.ExecuteStatement inWith & "." & inEvent & ".Code = " & inBlock

End Sub
Private Function ParseObject(ByRef inLine As String, ByRef inBlock As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0) As String
    'Debug.Print "Object: " & inLine
    inBlock = ParseBracketOff(ParseBracketOff(inBlock, "{", "}"), "[", "]")
    Dim inName As String
    Dim Temporary As Object
    inName = ParseQuotedArg(inLine, "<", ">", , , False)
    inLine = ParseWhiteSpace(inLine)
    Dim inObj As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inObj = inObj & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    If LCase(inObj) = "serialize" Then
        frmMain.AddCode ParseSerialize(inBlock)
    ElseIf LCase(inObj) = "deserialize" Then
        ParseObject = ParseInWith(inBlock, inWith)
    ElseIf (LCase(inObj) = "method") Then
        frmMain.AddCode "Sub " & inName & "()" & vbCrLf & _
             ParseInWith(inBlock, inWith) & vbCrLf & "End Sub" & vbCrLf
    ElseIf ParseReservedWord(inObj) = -3 Then
        ParseScript inBlock, inObj, (LineNum - CountWord(inBlock, vbCrLf))
    ElseIf ParseReservedWord(inObj) = 3 Then

        inObj = Trim(UCase(Left(inObj, 1)) & LCase(Mid(inObj, 2)))
        Set Temporary = CreateObjectPrivate(inObj)

        If inName = "" Then inName = Include.Unnamed(All)

        If Not Include.Exists(inName) Then
            If inWith = "" Then frmMain.AddCode "Dim " & inName & vbCrLf
            All.Add Temporary, inName
        End If

        If inWith = "" Then
            frmMain.ExecuteStatement "Set " & inName & " =  All(""" & inName & """)"
        End If

        frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & "s.Add All(""" & inName & """), """ & inName & """"
        frmMain.ExecuteStatement "All(""" & inName & """).Key = """ & inName & """"

        ParseScript inBlock, IIf(inWith <> "", inWith & ".", "") & inObj & "s(""" & inName & """)", (LineNum - CountWord(inBlock, vbCrLf))
    End If
End Function

Public Function ParseScript(ByRef txt As String, Optional ByVal inWith As String = "", Optional ByRef LineNum As Long = 0) As String
    On Error GoTo parseerror
    If (InStr(Left(txt, 3), ":") > 0) Or (InStr(Left(txt, 3), "\") > 0) Then
        If PathExists(txt, True) Then
            ScriptRoot = GetFilePath(txt)
            txt = ReadFile(txt)
            If PathExists(ScriptRoot & "\Serial.xml", True) Then
                frmMain.AddCode ParseDeserialize(ReadFile(ScriptRoot & "\Serial.xml"))
            End If
        End If
    End If
    LineNum = LineNum + CountWord(txt, vbCrLf)
    Dim inBlock As String
    Dim inLine As String
    Dim reserve As Integer
    Do Until txt = ""
        txt = ParseWhiteSpace(txt)
        If Not Left(txt, 1) = "'" Then
            reserve = ParseReservedWord(txt)
            If reserve <> 0 Then
                If (((InStr(txt, "{") < InStr(txt, "[")) Or (InStr(txt, "[") = 0)) And (InStr(txt, "{") > 0)) Then
                    inLine = ParseQuotedArg(txt, "{", "}")
                    inBlock = "{" & RemoveArg(inLine, "{")
                    inLine = Replace(inLine, inBlock, "")
                ElseIf (((InStr(txt, "[") < InStr(txt, "{")) Or (InStr(txt, "{") = 0)) And (InStr(txt, "[") > 0)) Then
                    inLine = ParseQuotedArg(txt, "[", "]")
                    inBlock = "[" & RemoveArg(inLine, "[")
                    inLine = Replace(inLine, inBlock, "")
                End If
                inLine = ParseWhiteSpace(inLine)
                If reserve = 1 Then
                    ParseEvent inLine, inBlock, inWith, (LineNum - CountWord(txt, vbCrLf))
                ElseIf reserve = 2 Then
                    ParseBindings inLine, inBlock, (LineNum - CountWord(txt, vbCrLf))
                ElseIf Abs(reserve) > 2 Then
                    ParseScript = ParseScript & ParseObject(inLine, inBlock, inWith, (LineNum - CountWord(txt, vbCrLf)))
                End If
            ElseIf InStr(NextArg(txt, vbCrLf), "=") > 0 Then
                If ParseWhiteSpace(RemoveArg(NextArg(txt, "["), "=")) = "" Then
                    inLine = ParseQuotedArg(txt, "[", "]")
                    ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith, (LineNum - CountWord(txt, vbCrLf))
                ElseIf ParseWhiteSpace(RemoveArg(NextArg(txt, vbCrLf), "=")) <> "" Then
                    inLine = RemoveNextArg(txt, vbCrLf)
                    ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith, (LineNum - CountWord(txt, vbCrLf))
                Else
                    ParseExecute txt, inWith, (LineNum - CountWord(txt, vbCrLf))
                End If
            Else
                ParseExecute txt, inWith, (LineNum - CountWord(txt, vbCrLf))
            End If
        Else
            RemoveNextArg txt, vbCrLf
        End If
    Loop
    Exit Function
parseerror:
    Debug.Print "Error: "; Err.Number & " Line: " & (LineNum - CountWord(txt, vbCrLf)) & " Description: " & Err.Description
End Function

