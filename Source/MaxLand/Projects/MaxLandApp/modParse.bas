Attribute VB_Name = "modParse"

Option Explicit

'#################################################################
'### These objects made public are added to ScriptControl also ###
'### are the only thing global exposed to modFactory.Execute   ###
'#################################################################

Global ScriptRoot As String
Global Include As New Include

Global All As ntnodes10.Collection
Global Beacons As ntnodes10.Collection
Global Bindings As Bindings
Global Boards As ntnodes10.Collection
Global Cameras As ntnodes10.Collection
Global Elements As ntnodes10.Collection
Global Lights As ntnodes10.Collection
Global Player As Player
Global Portals As ntnodes10.Collection
Global Screens As ntnodes10.Collection
Global Sounds As ntnodes10.Collection
Global Space As Space
Global Tracks As ntnodes10.Collection


Global Frame As Boolean
Global Second As Single
Global Millis As Single


Private Sourcing As String
Private LastCall As String

'###########################################################
'### Commonly used functions among other parse functions ###
'###########################################################

Public Function ParseNumerical(ByRef inValues As String) As Variant
    'parses the next numerical where seperated by commas, returns it
    'not used in this module, rather is for values tostring elsewhere
    If (InStr(inValues, ",") = 0 And InStr(inValues, " ") > 0) Then
        ParseNumerical = RemoveNextArg(inValues, " ")
    Else
        ParseNumerical = RemoveNextArg(inValues, ",")
    End If
    If IsNumeric(ParseNumerical) Then
        ParseNumerical = CSng(ParseNumerical)
    ElseIf Trim(ParseNumerical) = "" Then
        ParseNumerical = 0
    Else
        ParseNumerical = CSng(frmMain.Evaluate(ParseNumerical))
    End If

End Function

Public Function ParseQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """", Optional ByVal NextBlock As Boolean = True) As String
    'passing true for NextBlock removes and returns an entire next section identified by beginquote and endquote with text before the block as well
    'passing false for NextBlock removes and retruns what is only inbetween the beginquote and endquote, stripping it and leaving before text
    'concatenated to any text after whwer ethe block was
    Dim retVal As String
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

Private Function ParseReservedWord(ByVal inLine As String, Optional ByVal inObj As String = "") As Integer
    'checks for the presence of a reserved word and returns the code for it
    Dim inWord As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inWord = inWord & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    Select Case LCase(inWord)
        '#######################################################################################
        'event objects which have code brackets following the tag, which is to be executed later
        Case "oninrange", "onoutrange", "method", "script"
            ParseReservedWord = 1
        
        '#######################################################################################
        'objects whose tags repeat newly in creating, or special use case scenario functions
        
        Case "deserialize", "frame", "millis", "onidle", "second", "serialize", _
            "beacon", "board", "camera", "element", "light", "motion", "portal", "screen", "sound", "track"

            ParseReservedWord = 3
            
        '#######################################################################################
        'objects that are not to be new or multiple objects, they're singular already existing
        Case "bindings", "space", "player"
            ParseReservedWord = -3
        Case Else
            ParseReservedWord = IIf(GetBindingIndex(inWord) > -1, 2, 0)
    End Select
End Function
Public Function ParseEvalNow(ByVal inBlock As String) As Boolean
    'simply checks for eval now brackets {} or else assume eval later brackets []
    inBlock = ParseWhiteSpace(inBlock)
    ParseEvalNow = ParseWhiteSpace((Left(inBlock, 1) = "{" And Right(inBlock, 1) = "}"))
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
    'similar to trimstrip for tab, space, and crlf, and sequence repeating
    'white-spaces at the bginning or end are removed from the result return
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
    'builds the inline code with any "with" statement that's starting with a "." extend
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
    'if the statement contains < and > then the key in it is the name of the type
    'if there is no < and >, then the type, the first portion is the name also
    If (InStr(inLine, "<") > 0) And (InStr(inLine, ">") > InStr(inLine, "<")) Then
        ParseName = ParseQuotedArg(inLine, "<", ">", False)
    End If
End Function

Private Function ParseType(ByVal inLine As String) As String
    'returns the first portion of a statement,
    'is the type of object the statement is for
    inLine = ParseWhiteSpace(inLine)
    Do While IsAlphaNumeric(Left(inLine, 1))
        ParseType = ParseType & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
End Function
'##################################################################
'### Deserialize and serialize functions executed in part parse ###
'##################################################################

Private Function ParseDeserialize(ByRef nXML As String, Optional ByRef LineNum As Long = 0) As String
    On Error GoTo scripterror

    Dim retVal As String
    retVal = "Sub Deserialize()" & vbCrLf
    If nXML <> "" Then
        Dim tmp As String
        Dim cnt As Long
        Dim cnt2 As Long
        Dim cnt3 As Long
        Dim inName As String
        Dim xml As New MSXML.DOMDocument
        xml.async = "false"
        xml.loadXML nXML
        If xml.parseerror.errorCode = 0 Then
            Dim child As MSXML.IXMLDOMNode
            Dim serial As MSXML.IXMLDOMNode
            For Each serial In xml.childNodes
                Select Case Include.SafeKey(serial.baseName)
                
                    Case "serial"
                        For Each child In serial.childNodes
                            Select Case Include.SafeKey(child.baseName)
                                Case "datetime"
                                    If (FileDateTime(ScriptRoot & "Levels\" & CurrentLoadedLevel & ".vbx") <> Include.URLDecode(child.Text)) And (Not (InStr(1, LCase(Command), "/debug", vbTextCompare) > 0)) Then
                                        GoTo exitout:
                                    End If
                                Case "variables"
                                    For cnt = 0 To child.childNodes.Length - 1
                                        tmp = Replace(Replace(Include.URLDecode(child.childNodes(cnt).Text), """", """"""), vbCrLf, """ & vbCrLf & """)
                                        If IsNumeric(tmp) Or LCase(tmp) = "false" Or LCase(tmp) = "true" Then
                                            retVal = retVal & child.childNodes(cnt).baseName & " = " & tmp & vbCrLf
                                        Else
                                            retVal = retVal & child.childNodes(cnt).baseName & " = """ & tmp & """" & vbCrLf
                                        End If
                                    Next
                                Case "code"
                                    retVal = retVal & Include.URLDecode(child.Text) & vbCrLf
                            End Select
                        Next
                End Select
            Next
            
        Else
            AddMessage "There was an error parsing the serialized XML."
        End If
    End If
exitout:
    retVal = retVal & "End Sub" & vbCrLf
    ParseDeserialize = retVal
    Set xml = Nothing
scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured during serialization." & vbCrLf
        Err.Raise num, src, "Line: " & (LineNum + 1) & " Error: " & des
    End If
End Function
Private Function ParseSerialize(ByVal inSection As Integer, Optional ByRef txt As String, Optional ByRef LineNum As Long = 0) As String
    On Error Resume Next
    Static retVal As String
    Static deSer As String
    Dim tmp As String
    
    Select Case inSection
        Case 1
            retVal = "Function Serialize()" & vbCrLf
            retVal = retVal & "Serialize = Serialize & ""<?xml version=""""1.0""""?><Serial>""" & vbCrLf
        Case 2
            If retVal <> "" Then
                retVal = retVal & "Serialize = Serialize & ""  <Variables>"" & vbCrLf" & vbCrLf
                txt = ParseWhiteSpace(txt)
                Do Until txt = ""
                    tmp = Replace(Replace(NextArg(txt, vbCrLf), vbTab, ""), " ", "")
                    If tmp <> "" Then
                        retVal = retVal & "Serialize = Serialize & ""        <" & tmp & ">"" & " & tmp & " & ""</" & tmp & ">"" & vbCrLf" & vbCrLf
                    End If
                    RemoveNextArg txt, vbCrLf
                Loop
                retVal = retVal & "Serialize = Serialize & ""  </Variables>"" & vbCrLf" & vbCrLf
            End If
        Case 3
            deSer = txt
        Case 4
            If retVal <> "" Then
                retVal = retVal & "Serialize = Serialize & ""<DateTime>" & Include.URLEncode(FileDateTime(ScriptRoot & "Levels\" & CurrentLoadedLevel & ".vbx")) & "</DateTime>"" & vbCrLf" & vbCrLf
                
                retVal = retVal & "Serialize = Serialize & ""<Code>""" & vbCrLf
                retVal = retVal & "Serialize = Serialize & """ & Include.URLEncode(deSer) & """" & vbCrLf
                retVal = retVal & "Serialize = Serialize & ""</Code>""" & vbCrLf
                
                retVal = retVal & "Serialize = Serialize & Bindings.ToString()" & vbCrLf
                retVal = retVal & "Serialize = Serialize & Space.ToString()" & vbCrLf
                retVal = retVal & "Serialize = Serialize & ""</Serial>""" & vbCrLf

                retVal = retVal & "End Function" & vbCrLf
                 ParseSerialize = retVal
                retVal = ""
            End If
            Do Until deSer = ""
                frmMain.ExecuteStatement RemoveNextArg(deSer, vbCrLf)
            Loop
    End Select

scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured during serialization." & vbCrLf
        Err.Raise num, src, "Line: " & (LineNum + 1) & " Error: " & des
    End If
End Function

'#################################################################################
'### DFisposable parameter functions, anything passed is not changed to callee ###
'#################################################################################

Private Sub ParseSetting(ByVal inLine As String, ByVal inValue As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0)
    'Debug.Print "Setting: " & inLine & " B: " & inValue
    LastCall = inLine & " = " & inValue

    On Error GoTo scripterror
    Dim inName As String
    If Left(inLine, 8) = "variable" Then
        inName = ParseQuotedArg(inLine, "<", ">", False)
        inLine = inName
    End If
    inValue = ParseWhiteSpace(inValue)
    If ((Left(inValue, 1) = "[") And (Right(inValue, 1) = "]")) Then
        LastCall = ParseInWith(inLine, inWith) & "=""" & Replace(Replace(inValue, """", """"""), vbCrLf, """ & vbcrlf & """) & """"
        If inName <> "" Then frmMain.AddCode "Dim " & inName & vbCrLf
        frmMain.ExecuteStatement LastCall, , LineNum
    Else
        LastCall = ParseInWith(inLine, inWith) & "=" & inValue
        If inName <> "" Then frmMain.AddCode "Dim " & inName & vbCrLf, , LineNum
        frmMain.ExecuteStatement LastCall, , LineNum
    End If

scripterror:
    If Err.Number <> 0 And Err.Number <> 11 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while building a binding." & vbCrLf
        Err.Raise num, src, "Line: " & (LineNum + 1) & " Error: " & des
    End If
End Sub

Private Sub ParseBindings(ByVal inBind As String, ByVal inBlock As String, Optional ByRef LineNum As Long = 0)
    LastCall = inBind & " = [" & inBlock & "]"

    On Error GoTo scripterror
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
    LineNum = LineNum + CountWord(inBind & inBlock, vbCrLf)
    
scripterror:
    If Err.Number <> 0 And Err.Number <> 11 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while building a binding." & vbCrLf
        Err.Raise num, src, "Line: " & (LineNum + 1) & " Error: " & des
    End If
End Sub

Private Sub ParseEvent(ByVal inLine As String, ByVal inBlock As String, ByVal inWith As String, ByVal inName As String, Optional ByVal LineNum As Long = 0)
    On Error GoTo scripterror
    LineNum = LineNum + 1
    
    'Debug.Print "Event: " & inLine & " Block: " & inBlock;
    
    LastCall = inLine & " = " & inBlock
    Dim inEvent As String
    Do While IsAlphaNumeric(Left(inLine, 1))
        inEvent = inEvent & Left(inLine, 1)
        inLine = Mid(inLine, 2)
    Loop
    If inEvent = "method" Then
        inBlock = ParseBracketOff(inBlock, "[", "]")
        
        frmMain.AddCode "Sub " & inName & "()" & vbCrLf & _
             ParseInWith(inBlock, inWith) & vbCrLf & _
             "End Sub" & vbCrLf, LineNum
    Else
        inBlock = ParseBracketOff(inBlock, "[", "]")
        inBlock = ParseInWith(inBlock, inWith)
    
        Dim inApplyTo As String
        If Left(inBlock, 1) = "<" Then
            inApplyTo = ParseName(inBlock)
            inBlock = RemoveArg(inBlock, ">")
            Do While Left(inBlock, 2) = vbCrLf
                inBlock = Mid(inBlock, 3)
            Loop
        End If
        
        inName = "m" & Replace(modGuid.GUID(), "-", "")
        frmMain.AddCode "Sub " & inName & "()" & vbCrLf & inBlock & vbCrLf & "End Sub" & vbCrLf, , LineNum
        

        If LCase(inEvent) = "script" Then
            frmMain.ExecuteStatement inWith & "." & inEvent & " = """ & LineNum & ":" & inName & """"
        ElseIf LCase(inEvent) = "oninrange" Or LCase(inEvent) = "onoutrange" Then
            Dim oev As OnEvent
            Set oev = New OnEvent
            oev.RunMethod = inName
            oev.AppliesTo = inApplyTo
            oev.StartLine = LineNum
            inName = RemoveQuotedArg(inWith, """", """")
            Dim p As Portal
            Set p = Portals(inName)
            Select Case LCase(inEvent)
                Case "oninrange"
                    Set p.OnInRange = oev
                Case "onoutrange"
                    Set p.OnOutRange = oev
            End Select
            Set p = Nothing
            Set oev = Nothing
        End If
        
     

'        If LCase(inEvent) = "script" Then
'            frmMain.ExecuteStatement inWith & "." & inEvent & " = """ & LineNum & ":" & inName & """"
'        ElseIf frmMain.Evaluate(inWith & "." & inEvent & " Is Nothing") Then
'            frmMain.ExecuteStatement "Set " & inWith & "." & inEvent & " = NewObject(""OnEvent"")"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".RunMethod=""" & inName & """"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".AppliesTo=""" & inApplyTo & """"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".StartLine=" & LineNum
'        End If
        
            
'        If LCase(inEvent) = "script" Then
'
'            frmMain.ExecuteStatement inWith & "." & inEvent & " = """ & LineNum & ":" & inName & """"
'
'        ElseIf frmMain.Evaluate(inWith & "." & inEvent & " Is Nothing") Then
'
'
'            frmMain.ExecuteStatement "Set " & inWith & "." & inEvent & " = NewObject(""OnEvent"")"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".RunMethod=""" & inName & """"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".AppliesTo=""" & inApplyTo & """"
'            frmMain.ExecuteStatement inWith & "." & inEvent & ".StartLine=" & LineNum
'
'        Else
'            Err.Raise 8, , "The event " & inEvent & " for " & inWith & " already exists."
'        End If
        
        
        
        
    End If
scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while building an event." & vbCrLf
        Err.Raise num, src, "Line: " & LineNum & " Error: " & des
    End If
End Sub

Private Function ParseObject(ByVal inLine As String, ByVal inBlock As String, ByVal inWith As String, Optional ByVal LineNum As Long = 0) As String
    On Error GoTo scripterror:
    LineNum = LineNum + 1
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
        ParseSerialize 3, inBlock
    ElseIf (LCase(inObj) = "frame") Or (LCase(inObj) = "second") Or (LCase(inObj) = "millis") Then
        If (LCase(inObj) = "frame") Then
            If Frame Then
                Err.Raise 8, , "The 'frame' method already exists."
            Else
                Frame = True
            End If
        End If
        If (LCase(inObj) = "second") Then
            If Second = 0 Then
                Second = Timer
            Else
                Err.Raise 8, , "The 'second' method already exists."
            End If
        End If
        If (LCase(inObj) = "millis") Then
            If Millis = 0 Then
                Millis = Timer
            Else
                Err.Raise 8, , "The 'millis' method already exists."
            End If
        End If
        
        frmMain.AddCode "Sub " & inObj & "()" & vbCrLf & _
             inBlock & vbCrLf & "End Sub" & vbCrLf, LineNum
             
    ElseIf ParseReservedWord(inObj) = 3 Then
        Dim Temporary As Object
        inObj = Trim(UCase(Left(inObj, 1)) & LCase(Mid(inObj, 2)))
        Set Temporary = NewObject(inObj)

        ParseSetupObject Temporary, inObj, inName, inWith, (LineNum - CountWord(inLine, vbCrLf))
        ParseScript inBlock, IIf(inWith <> "", inWith & ".", "") & inObj & "s(""" & inName & """)", LineNum

    End If
scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while building an object." & vbCrLf
        Err.Raise num, src, "Line: " & LineNum & " Error: " & des
    End If
End Function

Public Sub ParseSetupObject(ByRef Temporary As Object, ByVal inObj As String, Optional ByRef inName As String = "", Optional ByVal inWith As String = "", Optional ByVal LineNum As Long = 0)
    On Error GoTo scripterror:
    
    If inName = "" Then inName = Temporary.Key
    If Not All.Exists(inName) Then
        If inWith = "" Then frmMain.AddCode "Dim " & inName & vbCrLf, , LineNum
        All.Add Temporary, inName
    End If
    
    If inWith = "" Then
        frmMain.ExecuteStatement "Set " & inName & " =  All(""" & inName & """)", , LineNum
    End If

    If frmMain.Evaluate(IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & " Is Nothing", , LineNum) Then
        frmMain.ExecuteStatement "Set " & IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & " = NewObject(""" & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & """)", , LineNum
    End If

    If frmMain.Evaluate(IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Exists(""" & inName & """)", , LineNum) Then
        frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Remove """ & inName & """", , LineNum
    End If

    frmMain.ExecuteStatement IIf(inWith <> "", inWith & ".", "") & inObj & IIf(Right(inObj, 1) <> "s", "s", "") & ".Add All(""" & inName & """), """ & inName & """", , LineNum
    frmMain.ExecuteStatement "All(""" & inName & """).Key = """ & inName & """", , (LineNum - CountWord(inObj, vbCrLf))
    
    Temporary.Key = inName
        
    
scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        Dim src As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while setting up an object." & vbCrLf
        Err.Raise num, src, "Line: " & (LineNum + 1) & " Error: " & des
    End If
End Sub

Private Function ParseExecute(ByVal inLine As String, ByVal inWith As String, Optional ByRef LineNum As Long = 0) As String
    On Error GoTo scripterror:
    If GetFileExt(inLine) = ".vbx" Then
        If PathExists(inLine, True) Then
            
            Dim src As String
            src = RemoveNextArg(inLine, vbCrLf)
            Swap src, ScriptRoot
            ParseScript ScriptRoot, inWith, LineNum
            Swap src, ScriptRoot

        Else
            LastCall = ParseInWith(inLine, inWith)
            frmMain.ExecuteStatement LastCall, , LineNum
        End If
    Else
        LastCall = ParseInWith(inLine, inWith)
        frmMain.ExecuteStatement LastCall, , LineNum
    End If
scripterror:
    If Err.Number <> 0 Then
        Dim num As Long
        Dim des As String
        num = Err.Number
        src = Err.source
        des = Err.Description
        Err.Clear
        On Error GoTo 0
        '"An error occured while executinng a statement." & vbCrLf
        Err.Raise num, src, "Line: " & LineNum & " Error: " & des
    End If
            
End Function

'########################################
'### Main parsing of scripts function ###
'########################################

Public Function ParseScript(ByVal inText As String, Optional ByVal inWith As String = "", Optional ByRef LineNum As Long = 0, Optional ByVal NoDeserialize As Boolean = False) As String
    On Error GoTo parseerror
    On Local Error GoTo parseerror
    Static serialLevel As Integer
    Dim deSer As String
    Dim atLine As Long
    
    
    atLine = LineNum - 1
    serialLevel = serialLevel + 1
    If (InStr(Left(inText, 3), ":") > 0) Or (InStr(Left(inText, 3), "\") > 0) Then
        If PathExists(inText, True) Then
            'the inText argument is a filename to a vbx script
            ScriptRoot = GetFilePath(GetFilePath(inText)) & "\"
            Sourcing = GetFileTitle(inText) 'source for error handling
            inText = ReadFile(inText)  'load it as if it was passed in inText
            If Right(inText, 2) <> vbCrLf Then inText = inText & vbCrLf
            'add on a trailing vbCrLf ffor line counting in case inText ends in text
            If serialLevel = 1 Then 'so no stacking occurs for serilse/deserialize
                'at loading of a script, check whether a serialized XML file is present with it
                If PathExists(ScriptRoot & "Levels\" & CurrentLoadedLevel & ".xml", True) Then
                    'build the deserialize function that will be executed at the end of loading the script
                    deSer = ParseDeserialize(ReadFile(ScriptRoot & "Levels\" & CurrentLoadedLevel & ".xml"))
                    frmMain.AddCode deSer ', , (LineNum - CountWord(inText, vbCrLf)) 'and add it as script code
                Else
                    frmMain.AddCode "Sub Deserialize()" & vbCrLf & "End Sub" & vbCrLf
                End If
                'make the header the serialize function
                deSer = ParseSerialize(1)
            End If
            
        End If
    End If
    LineNum = (LineNum + CountWord(inText, vbCrLf))
    
    
    Dim inBlock As String
    Dim inLine As String
    Dim inName As String
    Dim inType As String
    Dim reserve As Integer
    Do Until inText = ""
        inText = ParseWhiteSpace(inText)
        If Not Left(inText, 1) = "'" Then
            
            reserve = ParseReservedWord(inText, inWith)
            If reserve <> 0 Then
                If (((InStr(inText, "{") < InStr(inText, "[")) Or (InStr(inText, "[") = 0)) And (InStr(inText, "{") > 0)) Then
                    inLine = ParseQuotedArg(inText, "{", "}", True)
                    inBlock = "{" & RemoveArg(inLine, "{", vbBinaryCompare, False)
                    inLine = Replace(inLine, inBlock, "")
                ElseIf (((InStr(inText, "[") < InStr(inText, "{")) Or (InStr(inText, "{") = 0)) And (InStr(inText, "[") > 0)) Then
                    inLine = ParseQuotedArg(inText, "[", "]", True)
                    inBlock = "[" & RemoveArg(inLine, "[", vbBinaryCompare, False)
                    inLine = Replace(inLine, inBlock, "")
                End If
                inLine = ParseWhiteSpace(inLine)

                If reserve = 1 Then
                    inName = ParseName(inLine)
                    inLine = ParseType(inLine)
                    ParseEvent inLine, inBlock, inWith, inName, atLine
                ElseIf reserve = 2 Then
                    ParseBindings inLine, inBlock, atLine
                ElseIf reserve = -3 Then
                
                    inBlock = ParseBracketOff(inBlock, "{", "}")
                    inName = ParseName(inLine)
                    inType = ParseType(inLine)

                    ParseScript inBlock, inType, atLine
                    
                    Select Case inType
                        Case "player"
                            If inName <> "" Then
                                frmMain.ExecuteStatement inType & ".key=""" & inName & """"
                                frmMain.ExecuteStatement "All.Add " & inName & ", """ & inName & """"
                            End If
                    End Select

                ElseIf Abs(reserve) > 2 Then
                    ParseObject inLine, inBlock, inWith, atLine
                End If
            Else
                
                If InStr(NextArg(inText, vbCrLf), "=") > 0 Then
                    If ParseWhiteSpace(RemoveArg(NextArg(inText, "["), "=")) = "" Then
                        inLine = ParseQuotedArg(inText, "[", "]", True)

                        ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith, atLine
                    ElseIf ParseWhiteSpace(RemoveArg(NextArg(inText, vbCrLf), "=")) <> "" Then
                        inLine = RemoveNextArg(inText, vbCrLf)
                        ParseSetting NextArg(inLine, "="), RemoveArg(inLine, "="), inWith, atLine
                    Else
                        inLine = RemoveNextArg(inText, vbCrLf)
                        ParseExecute inLine, inWith, atLine
                    End If
                Else
                    
                    inLine = RemoveNextArg(inText, vbCrLf)
                    ParseExecute inLine, inWith, atLine
                End If
            End If
             
        Else
            RemoveNextArg inText, vbCrLf
        End If
        LastCall = ""
        atLine = (LineNum - CountWord(inText, vbCrLf)) + 1
        
    Loop
    serialLevel = serialLevel - 1
    If serialLevel = 0 Then
        'before we deserialize, built the serialize function
        'that will be executed when the project is closing
        
        deSer = ParseSerialize(4) 'finalize adding with the footer
        frmMain.AddCode deSer, , atLine  'add it to the code
        If Not NoDeserialize Then
            frmMain.Run "Deserialize" 'now run the deserialzation generated at the beginning
        End If
    End If
    Exit Function
parseerror:
    serialLevel = 0
    'Debug.Print "Error: "; Err.Number & " Line: " & (atLine - 1) & " Description: " & Err.Description
    'frmMain.Print "Error: "; Err.Number & " Line: " & (atLine - 1) & " Description: " & Err.Description
    DoEvents

    If Not ConsoleVisible Then
        ConsoleToggle
    End If
    Process "echo An error " & Err.Number & " occurd in " & Err.source & "\nError: " & Err.Description & "\n" & LastCall
    'frmMain.Print "echo An error " & Err.Number & " occurd in " & Err.Source & " at line " & (atLine - 1) & "\n" & Err.Description & "\n" & LastCall
    
    If frmMain.ScriptControl1.Error.Number <> 0 Then frmMain.ScriptControl1.Error.Clear
    If Err.Number <> 0 Then Err.Clear

End Function


