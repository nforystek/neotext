Attribute VB_Name = "modParse"
Option Explicit

'#################################################################
'### These objects made public are added to ScriptControl also ###
'### are the only thing global exposed to modFactory.Execute   ###
'#################################################################

Global ScriptRoot As String
Global Include As New Include

Global All As New NTNodes10.Collection

Global Brilliants As New Brilliants
Global Molecules As New Molecules

Global Billboards As New Billboards
Global Planets As New Planets

Global Motions As New Motions

Global OnEvents As New NTNodes10.Collection
Global Bindings As New Bindings
Global Camera As New Camera

Public Function ParseQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal Embeded As Boolean = False, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal KeepDelim As Boolean = True) As String
    Dim retval As String
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
                retval = Left(TheParams, X - 1) & Mid(TheParams, X)
                TheParams = Mid(retval, (X - 1) + (Y - X) + Len(EndQuote) + Len(BeginQuote))
                retval = Left(retval, (X - 1) + (Y - X) + Len(BeginQuote))
            Else
                retval = Mid(TheParams, X + Len(BeginQuote))
                TheParams = Left(TheParams, X - 1) & Mid(retval, (Y - X) + Len(EndQuote))
                retval = Left(retval, (Y - X) - 1)
            End If
        End If
    End If
    ParseQuotedArg = retval
End Function
Public Function ParseNumerical(ByRef inValues As String) As Variant

    ParseNumerical = RemoveNextArg(inValues, ",")
    
    If IsNumeric(ParseNumerical) Then
        ParseNumerical = CSng(ParseNumerical)
    Else
        ParseNumerical = CSng(frmMain.Evaluate(ParseNumerical))
    End If
    

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
    Dim retval As String
    Dim tmp As String
    Dim cnt As Long
    Dim cnt2 As Long
    Dim cnt3 As Long
    Dim inName As String
    retval = "Sub Deserialize()" & vbCrLf
    If xml.parseerror.errorCode = 0 Then
        Dim child As MSXML.IXMLDOMNode
        Dim serial As MSXML.IXMLDOMNode
        For Each serial In xml.childNodes
            Select Case Include.SafeKey(serial.baseName)
            
                Case "serial"
                    For Each child In serial.childNodes
                        Select Case Include.SafeKey(child.baseName)
                            Case "datetime"
                                'If (FileDateTime(ScriptRoot & "\Index.vbx") <> Include.URLDecode(child.Text)) And (Not (InStr(1, LCase(Command), "/debug", vbTextCompare) > 0)) Then
                                    GoTo exitout:
                                'End If
                            Case "variables"
                                For cnt = 0 To child.childNodes.Length - 1
                                    tmp = Replace(Replace(Include.URLDecode(child.childNodes(cnt).Text), """", """"""), vbCrLf, """ & vbCrLf & """)
                                    If IsNumeric(tmp) Or LCase(tmp) = "false" Or LCase(tmp) = "true" Then
                                        retval = retval & child.childNodes(cnt).baseName & " = " & tmp & vbCrLf
                                    Else
                                        retval = retval & child.childNodes(cnt).baseName & " = """ & tmp & """" & vbCrLf
                                    End If
                                Next
                                
                            Case "molecules", "brilliants", "planets", "billboards", "camera", "bindings"
                                All.Add "<?xml version=""1.0""?>" & vbCrLf & "<Serial>" & vbCrLf & child.xml & vbCrLf & "</Serial>" & vbCrLf, Include.SafeKey(child.baseName)
                                retval = retval & "Set Include.Serialize = " & Include.SafeKey(child.baseName) & vbCrLf

                        End Select
                    Next
            End Select
        Next
    End If
exitout:
    retval = retval & "End Sub" & vbCrLf
    ParseDeserialize = retval
    Set xml = Nothing
End Function
Private Function ParseSerialize(ByVal inSection As Integer, Optional ByRef txt As String, Optional ByRef LineNum As Long = 0) As String
    'On Error Resume Next
    Static retval As String
    Select Case inSection
        Case 1
            retval = "Function Serialize()" & vbCrLf
            retval = retval & "Serialize = ""<Serial>"" & vbCrLf" & vbCrLf
        Case 2
            If retval <> "" Then
                retval = retval & "Serialize = Serialize & ""  <Variables>"" & vbCrLf" & vbCrLf
                txt = ParseWhiteSpace(txt)
                Do Until txt = ""
                    retval = retval & "Serialize = Serialize & ""        <" & NextArg(txt, vbCrLf) & ">"" & Include.URLEncode(" & NextArg(txt, vbCrLf) & ") & ""</" & NextArg(txt, vbCrLf) & ">"" & vbCrLf" & vbCrLf
                    RemoveNextArg txt, vbCrLf
                Loop
                retval = retval & "Serialize = Serialize & ""  </Variables>"" & vbCrLf" & vbCrLf
            End If
        Case 3
            If retval <> "" Then
                retval = retval & "Serialize = Serialize & Include.Serialize & vbCrLf" & vbCrLf
                
                retval = retval & "If Bindings.Serialize Then" & vbCrLf
                retval = retval & "Serialize = Serialize & Bindings.ToString()" & vbCrLf
                retval = retval & "End If" & vbCrLf
    
                retval = retval & "If Camera.Serialize Then" & vbCrLf
                retval = retval & "Serialize = Serialize & Camera.ToString()" & vbCrLf
                retval = retval & "End If" & vbCrLf
    
                retval = retval & "Serialize = Replace(Serialize,""  <Variables>"" & vbCrLf & ""  </Variables>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Bindings>"" & vbCrLf & "" </Bindings>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Billboards>"" & vbCrLf & ""  </Billboards>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Planets>"" & vbCrLf & ""  </Planets>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Brilliants>"" & vbCrLf & ""  </Brilliants>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Molecules>"" & vbCrLf & ""  </Molecules>"","""")" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize,""  <Camera>"" & vbCrLf & ""  </Camera>"","""")" & vbCrLf
                retval = retval & "Do While Instr(Serialize,vbCrLf & vbCrLf)>0" & vbCrLf
                retval = retval & "Serialize = Replace(Serialize, vbCrLf & vbCrLf, vbCrLf)" & vbCrLf
                retval = retval & "Loop" & vbCrLf
                retval = retval & "Serialize = Serialize & ""</Serial>""" & vbCrLf
                retval = retval & "End Function" & vbCrLf
                ParseSerialize = retval
                retval = ""
            End If
    End Select
    
    If Err.Number = 11 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 453, "Line " & (LineNum - CountWord(txt, vbCrLf)), "Unable to serialize."
    Else
        On Error GoTo 0
    End If
End Function

Private Function ParseExecute(ByRef txt As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0) As String
    If GetFileExt(NextArg(txt, vbCrLf)) = ".vbx" Then
        If PathExists(NextArg(txt, vbCrLf), True) Then
            ParseScript RemoveNextArg(txt, vbCrLf), inWith, (LineNum - CountWord(txt, vbCrLf))
        Else
            frmMain.ExecuteStatement ParseInWith(RemoveNextArg(txt, vbCrLf), inWith)
        End If
    Else
        frmMain.ExecuteStatement ParseInWith(RemoveNextArg(txt, vbCrLf), inWith)
    End If
End Function
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
    inName = ParseQuotedArg(inLine, "<", ">", , , False)
    inLine = ParseWhiteSpace(inLine)
    Dim inObj As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inObj = inObj & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    If LCase(inObj) = "serialize" Then
        ParseSerialize 2, inBlock
    ElseIf LCase(inObj) = "deserialize" Then
        ParseObject = ParseInWith(inBlock, inWith)
    ElseIf (LCase(inObj) = "method") Then
        frmMain.AddCode "Sub " & inName & "()" & vbCrLf & _
             ParseInWith(inBlock, inWith) & vbCrLf & "End Sub" & vbCrLf
    ElseIf ParseReservedWord(inObj) = -3 Then
        ParseScript inBlock, inObj, (LineNum - CountWord(inBlock, vbCrLf))
    ElseIf ParseReservedWord(inObj) = 3 Then

        Dim Temporary As Object
        inObj = Trim(UCase(Left(inObj, 1)) & LCase(Mid(inObj, 2)))
        Set Temporary = CreateObjectPrivate(inObj)
    
        ParseSetupObject Temporary, inObj, inWith, inName
        
        ParseScript inBlock, IIf(inWith <> "", inWith & ".", "") & inObj & "s(""" & inName & """)", (LineNum - CountWord(inBlock, vbCrLf))
    End If
End Function
Public Sub ParseSetupObject(ByRef Temporary As Object, ByRef inObj As String, Optional ByVal inWith As String = "", Optional ByRef inName As String)
    
    If inName = "" Then inName = Temporary.Key

    If Not All.Exists(inName) Then
        If inWith = "" Then frmMain.AddCode "Dim " & inName & vbCrLf
        All.Add Temporary, inName
    End If

    If inWith = "" Then
        frmMain.ExecuteStatement "Set " & inName & " =  All(""" & inName & """)"
    End If

    If frmMain.Evaluate(IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Exists(""" & inName & """)") Then
        frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Remove """ & inName & """"
    End If
    frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Add All(""" & inName & """), """ & inName & """"

    frmMain.ExecuteStatement "All(""" & inName & """).Key = """ & inName & """"
End Sub

Public Function ParseScript(ByRef txt As String, Optional ByVal inWith As String = "", Optional ByRef LineNum As Long = 0, Optional ByVal Deserialize As String = "Serial.xml") As String
    On Error GoTo parseerror
    Static serialLevel As Integer
    serialLevel = serialLevel + 1
    If (InStr(Left(txt, 3), ":") > 0) Or (InStr(Left(txt, 3), "\") > 0) Then
        If PathExists(txt, True) Then
            ScriptRoot = GetFilePath(txt)
            txt = ReadFile(txt)
            
            If PathExists(ScriptRoot & "\" & Deserialize, True) Then
                frmMain.AddCode ParseDeserialize(ReadFile(ScriptRoot & "\" & Deserialize))
            End If
            If serialLevel = 1 Then ParseSerialize 1
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
                    ParseObject inLine, inBlock, inWith, (LineNum - CountWord(txt, vbCrLf))
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
    serialLevel = serialLevel - 1
    If serialLevel = 0 Then
        frmMain.AddCode ParseSerialize(3)
        frmMain.Deserialize
    End If
    Exit Function
parseerror:
    serialLevel = 0
    Debug.Print "Error: "; Err.Number & " Line: " & (LineNum - CountWord(txt, vbCrLf)) & " Description: " & Err.Description
End Function

