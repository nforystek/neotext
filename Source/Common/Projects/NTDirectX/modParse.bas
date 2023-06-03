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

Global Frame As Boolean
Global Second As Single
Global Millis As Single

Private LastCall As String

'###########################################################
'### Commonly used functions among other parse functions ###
'###########################################################

Public Function ParseNumerical(ByRef inValues As String) As Variant
    'not used in this module, rather is for values tostring elsewhere
    
    ParseNumerical = RemoveNextArg(inValues, ",")
    If IsNumeric(ParseNumerical) Then
        ParseNumerical = CSng(ParseNumerical)
    Else
        ParseNumerical = CSng(frmMain.Evaluate(ParseNumerical))
    End If

End Function

Public Function ParseQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal NextBlock As Boolean = True) As String
    'passing true for NextBlock removes and returns an entire next section identified by beginquote and endquote with text before the block as well
    'passing false for NextBlock removes and retruns what is only inbetween the beginquote and endquote, stripping it and leaving before text
    'concatenated to any text after whwer ethe block was
    Dim retval As String
    Dim Compare As VbCompareMethod
    Compare = VbCompareMethod.vbBinaryCompare
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
            If NextBlock Then
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

Private Function ParseReservedWord(ByVal inLine As String, Optional ByVal inObj As String = "") As Integer
    'checks for the presence of a reserved word and returns the code for it
    Dim inWord As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inWord = inWord & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    Select Case LCase(inWord)
        Case "oninrange", "onoutrange", "oncollide"
            ParseReservedWord = 1
        Case "molecule", "method", "brilliant", "serialize", "deserialize", "motion", "billboard", "planet", "frame", "second", "millis"
            ParseReservedWord = 3
        Case "bindings", "camera"
            ParseReservedWord = -3
        Case Else
            ParseReservedWord = IIf(GetBindingIndex(inWord) > -1 Or LCase(Trim(inObj)) = "bindings", 2, 0)
    End Select
End Function
Private Function ParseBracketOff(ByVal inBlock As String, ByVal StartDelim As String, ByVal StopDelim As String) As String
    'similar to trimstrip with delimiters, and startdelim at the beginning is removed, any stopdelim at the end is removed
    inBlock = ParseWhiteSpace(inBlock)
    If Left(inBlock, 1) = StartDelim And Right(inBlock, 1) = StopDelim Then
        inBlock = Mid(inBlock, 2)
        inBlock = Left(inBlock, Len(inBlock) - 1)
    End If
    ParseBracketOff = ParseWhiteSpace(inBlock)
End Function
Public Function ParseWhiteSpace(ByVal inBlock As String) As String
    'similar to trimstrip with whitespaces, and sequence repeating white-
    'spaces at the bginning or end are removed from the result returned
    Static stack As Boolean
    If Not stack Then
        stack = True
        inBlock = StrReverse(inBlock)
        inBlock = ParseWhiteSpace(inBlock)
        inBlock = StrReverse(inBlock)
        stack = False
    End If
    Do While Left(inBlock, 1) = " " Or Left(inBlock, 1) = vbTab Or Left(inBlock, 1) = vbCr Or Left(inBlock, 1) = vbLf
        inBlock = Mid(inBlock, 2)
    Loop
    ParseWhiteSpace = inBlock
End Function
Private Function ParseInWith(ByVal inLine As String, Optional ByVal inWith As String = "") As String
    'builds the inline code with any "with" statement that starting with a "." extend
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

Private Function ParseName(ByVal inLine As String) As String
    ParseName = ParseQuotedArg(inLine, "<", ">", False)
End Function

Private Function ParseType(ByVal inLine As String) As String
    inLine = ParseWhiteSpace(inLine)
    Do While IsAlphaNumeric(Left(inLine, 1))
        ParseType = ParseType & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
End Function
'##################################################################
'### Deserialize and serialize functions executed in part parse ###
'##################################################################

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
                                If (FileDateTime(ScriptRoot & "\Index.vbx") <> Include.URLDecode(child.Text)) And (Not (InStr(1, LCase(Command), "/debug", vbTextCompare) > 0)) Then
                                    GoTo exitout:
                                End If
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
Private Function ParseSerialize(ByVal inSection As Integer, Optional ByRef txt As String) As String
    On Error Resume Next
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
    
    If Err.number = 11 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 453, "Serialize", "Unable to serialize."
    Else
        On Error GoTo 0
    End If
End Function

'#################################################################################
'### DFisposable parameter functions, anything passed is not changed to callee ###
'#################################################################################

Private Sub ParseSetting(ByVal inLine As String, ByVal inValue As String, ByVal inWith As String)
    'Debug.Print "Setting: " & inLine & " B: " & inValue
    LastCall = inLine & " = " & inValue
    Dim inName As String
    If Left(inLine, 8) = "variable" Then
        inName = ParseQuotedArg(inLine, "<", ">", False)
        inLine = inName
    End If
    inValue = ParseWhiteSpace(inValue)
    If ((Left(inValue, 1) = "[") And (Right(inValue, 1) = "]")) Then
        LastCall = ParseInWith(inLine, inWith) & "=""" & Replace(Replace(inValue, """", """"""), vbCrLf, """ & vbcrlf & """) & """"
        If inName <> "" Then frmMain.AddCode "Dim " & inName & vbCrLf
        frmMain.ExecuteStatement LastCall
    Else
        LastCall = ParseInWith(inLine, inWith) & "=" & inValue
        If inName <> "" Then frmMain.AddCode "Dim " & inName & vbCrLf
        frmMain.ExecuteStatement LastCall
    End If
End Sub

Private Sub ParseBindings(ByVal inBind As String, ByVal inBlock As String)
    LastCall = inBind & " = [" & inBlock & "]"

    On Error Resume Next
    Dim BindIndex As Long
    Dim bindCode As String
    inBind = ParseWhiteSpace(inBind)
    inBlock = ParseWhiteSpace(inBlock)
    
    BindIndex = GetBindingIndex(RemoveNextArg(inBind, "="))
    If Left(inBlock, 1) = "[" And Right(inBlock, 1) = "]" Then
        bindCode = ParseQuotedArg(inBlock, "[", "]", False)
    Else
        bindCode = inBlock
    End If
    If BindIndex > -1 And bindCode <> "" Then
        Bindings(BindIndex) = bindCode
    End If

    If Err.number = 11 Then
        Err.Clear
        On Error GoTo 0
        Err.Raise 453, "Bindings", "The binding specified was not recognized."
    Else
        On Error GoTo 0
    End If
End Sub

Private Sub ParseEvent(ByVal inLine As String, ByVal inBlock As String, ByVal inWith As String)
    'Debug.Print "Event: " & inLine & " B: " & inBlock
    LastCall = inLine & " = " & inBlock
    Dim inEvent As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inEvent = inEvent & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    inBlock = ParseBracketOff(inBlock, "[", "]")
    Dim inName As String
    inName = ParseQuotedArg(inBlock, "<", ">", False)
    inBlock = ParseInWith(inBlock, inWith)
    inBlock = """" & Replace(Replace(inBlock, """", """"""), vbCrLf, """ & vbCrLf & """) & """"
    frmMain.ExecuteStatement "Set " & inWith & "." & inEvent & " = Nothing"
    If inName <> "" Then frmMain.ExecuteStatement inWith & "." & inEvent & ".ApplyTo = """ & inName & """"
    frmMain.ExecuteStatement inWith & "." & inEvent & ".Code = " & inBlock
End Sub

Private Function ParseObject(ByVal inLine As String, ByVal inBlock As String, ByVal inWith As String) As String
    On Error GoTo scripterror:
    
    'Debug.Print "Object: " & inLine
    LastCall = inLine
    
    inBlock = ParseBracketOff(ParseBracketOff(inBlock, "{", "}"), "[", "]")
   
    Dim inName As String
    Dim inObj As String

    inName = ParseName(inLine)
    inObj = ParseType(inLine)
    
    LastCall = inObj & " <" & inName & ">"
        
    If LCase(inObj) = "serialize" Then
        ParseSerialize 2, inBlock
    ElseIf LCase(inObj) = "deserialize" Then
        ParseObject = ParseInWith(inBlock, inWith)
    ElseIf (LCase(inObj) = "method") Then
        frmMain.AddCode "Sub " & inName & "()" & vbCrLf & _
             ParseInWith(inBlock, inWith) & vbCrLf & "End Sub" & vbCrLf
    ElseIf (LCase(inObj) = "frame") Or (LCase(inObj) = "second") Or (LCase(inObj) = "millis") Then
        If (LCase(inObj) = "frame") Then Frame = True
        If (LCase(inObj) = "second") Then Second = Timer
        If (LCase(inObj) = "millis") Then Millis = Timer
    
        frmMain.AddCode "Sub " & inObj & "()" & vbCrLf & _
             inBlock & vbCrLf & "End Sub" & vbCrLf
    ElseIf ParseReservedWord(inObj) = 3 Then
        Dim Temporary As Object
        inObj = Trim(UCase(Left(inObj, 1)) & LCase(Mid(inObj, 2)))
        Set Temporary = CreateObjectPrivate(inObj)
        ParseSetupObject Temporary, inObj, inName, inWith
        ParseScript inBlock, IIf(inWith <> "", inWith & ".", "") & inObj & "s(""" & inName & """)"
    End If
scripterror:
        If Err.number <> 0 Then
            Dim num As Long
            Dim des As String
            Dim src As String
            num = Err.number
            src = Err.source
            des = Err.description
            Err.Clear
            On Error GoTo 0
            Err.Raise 51, GetFileTitle(src), "Subscript Error: " & num & " Source: " & src & " Description\n" & des
        End If
End Function

Public Sub ParseSetupObject(ByRef Temporary As Object, ByVal inObj As String, Optional ByRef inName As String = "", Optional ByVal inWith As String = "")
    
    If inName = "" Then inName = Temporary.Key
    If Not All.Exists(inName) Then
        If inWith = "" Then frmMain.AddCode "Dim " & inName & vbCrLf
        All.Add Temporary, inName
    End If
    If inWith = "" Then
        frmMain.ExecuteStatement "Set " & inName & " =  All(""" & inName & """)"
    End If
    
    If frmMain.Evaluate(IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & " Is Nothing") Then
        frmMain.ExecuteStatement "Set " & IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & " = CreateObjectPrivate(""" & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & """)"
    End If
    
    If frmMain.Evaluate(IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Exists(""" & inName & """)") Then
        frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Remove """ & inName & """"
    End If
    frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Add All(""" & inName & """), """ & inName & """"
    frmMain.ExecuteStatement "All(""" & inName & """).Key = """ & inName & """"
        
End Sub

Private Function ParseExecute(ByVal inLine As String, ByVal inWith As String) As String
    If GetFileExt(inLine) = ".vbx" Then
        If PathExists(inLine, True) Then
            On Error GoTo scripterror:
            Dim src As String
            src = RemoveNextArg(inLine, vbCrLf)
            modCommon.Swap src, ScriptRoot
            ParseScript src, inWith
            modCommon.Swap src, ScriptRoot
scripterror:
            If Err.number <> 0 Then
                Dim num As Long
                Dim des As String
                num = Err.number
                src = Err.source
                des = Err.description
                Err.Clear
                On Error GoTo 0
                Err.Raise 51, GetFileTitle(src), "Subscript Error: " & num & " Source: " & src & " Description\n" & des
            End If
        Else
            LastCall = ParseInWith(inLine, inWith)
            frmMain.ExecuteStatement LastCall
        End If
    Else
        LastCall = ParseInWith(inLine, inWith)
        frmMain.ExecuteStatement LastCall
    End If
End Function

'########################################
'### Main parsing of scripts function ###
'########################################

Public Function ParseScript(ByVal txt As String, Optional ByVal inWith As String = "", Optional ByRef LineNum As Long = 0, Optional ByVal Deserialize As String = "Serial.xml") As String
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
    Dim atLine As Long
    atLine = 1
    Do Until txt = ""
        txt = ParseWhiteSpace(txt)
        If Not Left(txt, 1) = "'" Then
            reserve = ParseReservedWord(txt)
            If reserve <> 0 Then
                If (((InStr(txt, "{") < InStr(txt, "[")) Or (InStr(txt, "[") = 0)) And (InStr(txt, "{") > 0)) Then
                    inLine = ParseQuotedArg(txt, "{", "}", True)
                    inBlock = "{" & RemoveArg(inLine, "{", vbBinaryCompare, False)
                    inLine = Replace(inLine, inBlock, "")
                ElseIf (((InStr(txt, "[") < InStr(txt, "{")) Or (InStr(txt, "{") = 0)) And (InStr(txt, "[") > 0)) Then
                    inLine = ParseQuotedArg(txt, "[", "]", True)
                    inBlock = "[" & RemoveArg(inLine, "[", vbBinaryCompare, False)
                    inLine = Replace(inLine, inBlock, "")
                End If
                inLine = ParseWhiteSpace(inLine)
                If reserve = 1 Then
                    ParseEvent inLine, inBlock, inWith
                ElseIf reserve = 2 Then
                    ParseBindings inLine, inBlock
                ElseIf ParseReservedWord(inLine) = -3 Then
                    inBlock = ParseBracketOff(ParseBracketOff(inBlock, "{", "}"), "[", "]")
                    ParseScript inBlock, ParseType(inLine), atLine
                ElseIf Abs(reserve) > 2 Then
                    ParseObject inLine, inBlock, inWith
                End If
            ElseIf InStr(NextArg(txt, vbCrLf), "=") > 0 Then
                If ParseWhiteSpace(RemoveArg(NextArg(txt, "["), "=")) = "" Then
                    inLine = ParseQuotedArg(txt, "[", "]", True)
                    ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith
                ElseIf ParseWhiteSpace(RemoveArg(NextArg(txt, vbCrLf), "=")) <> "" Then
                    inLine = RemoveNextArg(txt, vbCrLf)
                    ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith
                Else
                    ParseExecute RemoveNextArg(txt, vbCrLf), inWith
                End If
            Else
                ParseExecute RemoveNextArg(txt, vbCrLf), inWith
            End If
            atLine = (LineNum - CountWord(txt, vbCrLf))
        Else
            RemoveNextArg txt, vbCrLf
        End If
        LastCall = ""
    Loop
    serialLevel = serialLevel - 1
    If serialLevel = 0 Then
        frmMain.AddCode ParseSerialize(3)
        frmMain.Deserialize
    End If
    Exit Function
parseerror:
    serialLevel = 0
'    Debug.Print "Error: "; Err.number & " Line: " & (atLine - 1) & " Description: " & Err.description

    If Not ConsoleVisible Then
        ConsoleToggle
    End If
    Process "echo An error " & Err.number & " occurd in " & Err.source & " at line " & (atLine - 1) & "\n" & Err.description & "\n" & LastCall
    
    If frmMain.ScriptControl1.Error.number <> 0 Then frmMain.ScriptControl1.Error.Clear
    If Err.number <> 0 Then Err.Clear

End Function

