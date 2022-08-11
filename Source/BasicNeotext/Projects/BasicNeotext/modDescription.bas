Attribute VB_Name = "modDescription"

Option Explicit

Option Compare Text

Private Enum HeaderInfo
    Declared = 0
    Commented = 1
    Attributed = 2
End Enum

Private Type VBPInfo
    PrjType As String
    Name As String
    CondComp As String
    Includes As String
    Files As String
    Reserved As String
    Neotext As String
End Type
    
Public Sub BuildFileDescriptions(ByVal FileName As String, ByVal LoadElseSave As Boolean)
    On Error GoTo nochanges
    On Local Error GoTo nochanges
    
    Dim txt As String
    Dim out As String
    Dim user As String
    Dim desc As String
    Dim head As String
    Dim back As String


    Select Case GetFileExt(FileName, True, True)
        Case "vbp"
            out = ""
            back = ReadFile(FileName)
            txt = back
            Dim vbp As VBPInfo
            

            Do Until txt = ""
                head = RemoveNextArg(txt, vbCrLf)
                Select Case Trim(LCase(NextArg(head, "=")))
                    Case "name"
                        vbp.Name = Replace(RemoveArg(head, "="), """", "")
                        out = out & head & vbCrLf
                    Case "type"
                        vbp.PrjType = head & vbCrLf
                    Case "reference"
                        vbp.Includes = vbp.Includes & head & vbCrLf
                    Case "object"
                        vbp.Includes = head & vbCrLf & vbp.Includes
                    Case "form", "module", "class", "usercontrol", "relateddoc", "designer", "userdocument", "resfile32"
                        vbp.Files = vbp.Files & head & vbCrLf
                    Case "condcomp"
                        head = Replace(Replace(RemoveArg(head, "="), """", ":"), " ", "")
                        out = out & "CondComp=%condcomp%" & vbCrLf
                        vbp.CondComp = head
                    Case "compcond"
                        head = Replace(Replace(RemoveArg(head, "="), """", ":"), " ", "")
                        vbp.Neotext = head
                    Case "[neotext]"
                    Case Else

                        out = out & head & vbCrLf
                End Select
            Loop

            BuildCondComp vbp
            
            out = Replace(out, "%name%", vbp.Name)
            If InStr(out, "%condcomp%") > 0 Then
                out = Replace(out, "%condcomp%", Replace(Replace("""" & vbp.CondComp & """", """:", """"), ":""", """"))
            ElseIf InStr(out, vbCrLf & "CondComp=""") = 0 Then
                out = Replace(out, vbCrLf & "Name=""", vbCrLf & "CondComp=" & Replace(Replace("""" & vbp.CondComp & """", """:", """"), ":""", """") & vbCrLf & "Name=""")
            End If
            
            out = vbp.PrjType & vbp.Includes _
                    & vbp.Files & out & "[Neotext]" & vbCrLf _
                    & "CompCond=" & Replace(Replace("""" & vbp.Neotext & """", """:", """"), ":""", """") & vbCrLf

            If out <> back Then WriteFile FileName, out
        Case "bas", "ctl", "cls", "frm", "dob", "dsr"
            out = ""
            back = ReadFile(FileName)
            
            If GetFileExt(FileName, True, True) = "bas" Then

                Const Header = "#Const [True] = -1" & vbCrLf & "#Const [False] = 0" & vbCrLf
                If InStr(back, Replace(Replace(Header, "[", ""), "]", "")) > 0 Then 'Or InStr(back, Header) > 0 Then
                    back = Replace(back, Header, Replace(Replace(Header, "[", ""), "]", ""))
                    back = Replace(back, Replace(Replace(Header, "[", ""), "]", ""), Header)
                End If
            End If

            txt = vbCrLf & back & vbCrLf
            Do Until txt = ""
                out = out & FindNextHeader(txt, head)
                 If GetUserDefined(head) <> "" Then
                    Debug.Print
                    Debug.Print "FULL NEXT HEADER INFORMATION"
                    Debug.Print head
                    Debug.Print "DECLARE: " & GetDeclareLine(head)
                    Debug.Print "USERDEFINED FROM DECLARE: "; GetUserDefined(head, Commented); " USER DEFINED FROM ATTRIBUTE: " & GetUserDefined(head, Attributed)
                    Debug.Print "COMMENTED DESCRIPTION: "; GetDescription(head, Commented); " ATTRIBUTE DESCRIPTION: " & GetDescription(head, Attributed)
                    If LoadElseSave Then
                        If CountWord(head, vbCrLf) = 2 Then
                            head = head & "Attribute " & GetUserDefined(head, Declared) & ".VB_Description = """ & GetDescription(head, Commented) & """" & vbCrLf
                        End If
                        out = out & GetDeclareLine(head) & " ' _" & vbCrLf & GetDescription(head, Attributed) & vbCrLf & _
                            "Attribute " & GetUserDefined(head, Attributed) & ".VB_Description = """ & GetDescription(head, Attributed) & """" & vbCrLf
                    Else
                        out = out & GetDeclareLine(head) & " ' _" & vbCrLf & GetDescription(head, Commented) & vbCrLf & _
                            "Attribute " & GetUserDefined(head, Declared) & ".VB_Description = """ & GetDescription(head, Commented) & """" & vbCrLf
                    End If
                Else
                    out = out & head
                End If
            Loop
            If out <> back Then
                WriteFile FileName, Mid(out, 3, Len(out) - 4)
            End If
    End Select

    Exit Sub
nochanges:
    Err.Clear
End Sub

Private Sub BuildCondComp(ByRef vbp As VBPInfo)
    On Error GoTo nochanges
    On Local Error GoTo nochanges
    
    Dim Var As String
    Dim val As String
    Dim Ret As String
    Dim tmp As String

    vbp.CondComp = Replace(vbp.CondComp, ":" & vbp.Name & "=-1", "")
    vbp.CondComp = Replace(vbp.CondComp, ":VBIDE=-1", "")
    
    tmp = vbp.Neotext
    Do Until tmp = ""
        val = RemoveNextArg(tmp, vbCrLf)
        Var = RemoveNextArg(val, "=")
        If InStr(1, vbp.Files, "Module=" & Var & ";") > 0 Then
            vbp.Neotext = Replace(vbp.Neotext, ":" & Var & "=-1", "")
            If InStr(1, ":" & Var & "=", vbTextCompare) > 0 Then
                vbp.CondComp = vbp.CondComp & ":" & Var & "=-1"
            End If
        Else
            vbp.CondComp = Replace(vbp.CondComp, ":" & Var & "=-1", "")
            vbp.Neotext = Replace(vbp.Neotext, ":" & Var & "=-1", "")
        End If
    Loop

    vbp.Neotext = ""
    tmp = vbp.Files
    Do Until tmp = ""
        val = RemoveNextArg(tmp, vbCrLf)
        Var = RemoveNextArg(val, "=")
        val = NextArg(val, ";")
        Select Case LCase(Var)
            Case "module"
                If InStr(1, vbp.CondComp, ":" & val & "=", vbTextCompare) = 0 Then
                    vbp.CondComp = vbp.CondComp & ":" & val & "=-1"
                End If
                
                vbp.Neotext = vbp.Neotext & ":" & val & "=-1"

        End Select
    Loop
    If InStr(1, vbp.CondComp, ":VBIDE=", vbTextCompare) = 0 Then
        vbp.CondComp = ":VBIDE=-1:" & vbp.CondComp
    End If
    If InStr(1, vbp.CondComp, ":" & vbp.Name & "=", vbTextCompare) = 0 Then
        vbp.CondComp = ":" & vbp.Name & "=-1:" & vbp.CondComp
    End If
    
    vbp.CondComp = Replace(vbp.CondComp, "::", ":")
    vbp.Neotext = Replace(vbp.Neotext, "::", ":")
    
    Exit Sub
nochanges:
    Err.Clear
End Sub
Private Function FindNextHeader(ByRef txt As String, ByRef head As String) As String
    'returns finished, removes head as well from txt, and head is found <> ""
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    pos1 = InStr(1, txt, " ' _" & vbCrLf, vbTextCompare)
    pos2 = InStr(1, txt, vbCrLf & "Attribute ", vbTextCompare)
    Do Until pos2 = 0
        If NextArg(Mid(txt, pos2 + 2), vbCrLf) Like "*Attribute *.VB_Description*" Then Exit Do
        pos2 = InStr(pos2 + 1, txt, vbCrLf & "Attribute ", vbTextCompare)
    Loop
    
    If pos1 < pos2 And pos1 > 1 Then
        pos1 = InStrRev(txt, vbCrLf, pos1 - 1, vbTextCompare)
    ElseIf pos2 < pos1 And pos2 > 1 Then
        pos1 = InStrRev(txt, vbCrLf, pos2 - 1, vbTextCompare)
    ElseIf pos1 > 0 Then
        pos1 = InStrRev(txt, vbCrLf, pos1 - 1, vbTextCompare)
    
    End If
    If pos1 = 0 And pos2 = 0 Then
        FindNextHeader = txt
        txt = ""
        head = ""
    ElseIf pos1 = 0 Then
    
        If pos2 > 0 Then

            pos3 = InStr(pos2, txt, """" & vbCrLf, vbTextCompare)
            If pos3 > 0 Then
                pos1 = InStrRev(txt, vbCrLf, pos2, vbTextCompare)
                pos3 = pos3 + 3
            Else
                FindNextHeader = txt
                txt = ""
                head = ""
            End If
            If pos1 > 0 Then pos1 = pos1 + 1
        Else
            FindNextHeader = txt
            txt = ""
            head = ""
        
        End If
    Else
        pos1 = pos1 + 1
        pos2 = InStr(pos1, txt, " ' _" & vbCrLf, vbTextCompare)
        pos3 = InStr(pos1, txt, vbCrLf & "Attribute ", vbTextCompare)
        If (pos3 < pos2) And (pos2 = 0) And (pos3 > 0) Then

            pos3 = InStr(pos1, txt, """" & vbCrLf, vbTextCompare)
            If pos3 > 0 Then
                pos3 = pos3 + 3
            Else
                pos3 = pos2 + 6
            End If
        ElseIf pos3 > 0 Then
            pos3 = InStr(pos3, txt, """" & vbCrLf, vbTextCompare)
            If pos3 > 0 Then pos3 = pos3 + 3
        ElseIf pos2 > 0 Then
            pos2 = pos2 + 6
            pos3 = InStr(pos2, txt, vbCrLf, vbTextCompare) + 2
        Else
            pos3 = Len(txt)
        End If
    End If
    
    If pos1 > 0 And pos3 > pos1 Then
        head = Replace(LTrimStrip(LTrimStrip(LTrimStrip(Mid(txt, pos1, (pos3 - pos1)), vbLf), vbCr), vbCrLf), vbCrLf & vbCrLf, vbCrLf)
        If Len(head) <> 0 Then
            Select Case CountWord(head, vbCrLf)
                Case 0, 1
                    head = NextArg(head, vbCrLf) & vbCrLf
                Case 1, 2
                    head = NextArg(head, vbCrLf) & vbCrLf & _
                            NextArg(RemoveArg(head, vbCrLf), vbCrLf) & vbCrLf
                            
                Case 3
                    head = NextArg(head, vbCrLf) & vbCrLf & _
                            NextArg(RemoveArg(head, vbCrLf), vbCrLf) & vbCrLf & _
                            NextArg(RemoveArg(RemoveArg(head, vbCrLf), vbCrLf), vbCrLf) & vbCrLf
                Case Else
                    head = ""
            End Select
        End If
        If Len(head) = 0 Then
            FindNextHeader = txt
            txt = ""
        Else
            FindNextHeader = Left(txt, pos1)
            If Len(head) > 0 Then
                txt = Mid(txt, pos1 + (pos3 - pos1))
            Else
                FindNextHeader = FindNextHeader & txt
                txt = ""
            End If
        End If
    End If

End Function

Private Function Clear(ByVal txt As String, ByVal find As String) As String
    Clear = Replace(txt, find, "", , , vbTextCompare)
End Function
Private Function GetDeclareLine(ByVal head As String) As String
    If InStr(head, "' _" & vbCrLf) > 0 Then
        GetDeclareLine = RTrimStrip(NextArg(head, "' _" & vbCrLf))
    Else
        GetDeclareLine = NextArg(head, vbCrLf)
    End If
End Function
Private Function GetDescription(ByVal head As String, Optional ByVal From As HeaderInfo = 0) As String
    If From = Declared Or From = Commented Then
        If InStr(head, " ' _" & vbCrLf) > 0 Then
            GetDescription = NextArg(RemoveArg(head, " ' _" & vbCrLf), vbCrLf)
        End If
    Else
        If InStr(head, vbCrLf & "Attribute ") > 0 Then
            GetDescription = RemoveQuotedArg(head, ".VB_Description = """, """" & vbCrLf)
        End If
    End If
End Function

Private Function GetUserDefined(ByVal head As String, Optional ByVal From As HeaderInfo = 0) As String
    If From = Declared Or From = Commented Then
        Do While GetUserDefined = "" And (head <> "")
            Select Case NextArg(head, " ")
                Case "Public", "Private", "Global", "Friend", "Static", "WithEvents"
                Case "Dim", "Const", "Declare", "Event"
                    GetUserDefined = NextArg(NextArg(NextArg(RemoveArg(head, " "), " "), "("), ",")
                Case "Type", "Enum"
                    GetUserDefined = NextArg(NextArg(RemoveArg(head, " "), " "), "(")
                Case "Property"
                    GetUserDefined = NextArg(NextArg(RemoveArg(RemoveArg(head, " "), " "), " "), "(")
                Case "Function", "Sub"
                    GetUserDefined = NextArg(NextArg(RemoveArg(head, " "), " "), "(")
            End Select
            RemoveNextArg head, " "
        Loop
    Else
        If InStr(head, vbCrLf & "Attribute ") > 0 Then
            GetUserDefined = RemoveQuotedArg(head, vbCrLf & "Attribute ", ".VB_Description = """)
        End If
    End If
End Function
Private Function TrimStrip(ByVal TheStr As String, Optional ByVal TheChar As String = "") As String
    If TheChar = "" Then
        TrimStrip = TrimStrip(TrimStrip(TheStr, " "), vbTab)
    Else
        TrimStrip = LTrimStrip(RTrimStrip(TheStr, TheChar), TheChar)
    End If
End Function
Private Function LTrimStrip(ByVal TheStr As String, Optional ByVal TheChar As String = "") As String
    If TheChar = "" Then
        LTrimStrip = LTrimStrip(LTrimStrip(TheStr, " "), vbTab)
    Else
        Do While Left(TheStr, Len(TheChar)) = TheChar
            TheStr = Mid(TheStr, Len(TheChar) + 1)
        Loop
        LTrimStrip = TheStr
    End If
End Function
Private Function RTrimStrip(ByVal TheStr As String, Optional ByVal TheChar As String = "") As String
    If TheChar = "" Then
        RTrimStrip = RTrimStrip(RTrimStrip(TheStr, " "), vbTab)
    Else
        Do While Right(TheStr, Len(TheChar)) = TheChar
            TheStr = Left(TheStr, Len(TheStr) - Len(TheChar))
        Loop
        RTrimStrip = TheStr
    End If
End Function

Private Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) - 1)
    Else
        NextArg = TheParams
    End If
End Function

Private Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
    Else
        RemoveArg = ""
    End If
End Function

Private Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String) As String
    If InStr(1, TheParams, TheSeperator, vbTextCompare) > 0 Then
        RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) - 1)
        TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, vbTextCompare) + Len(TheSeperator))
    Else
        RemoveNextArg = TheParams
        TheParams = ""
    End If
End Function
Private Function NextQuotedArg(ByVal TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    NextQuotedArg = RemoveQuotedArg(TheParams, BeginQuote, EndQuote)
End Function
Private Function RemoveQuotedArg(ByRef TheParams As String, Optional ByVal BeginQuote As String = """", Optional ByVal EndQuote As String = """") As String
    Dim retVal As String
    Dim x As Long
    x = InStr(1, TheParams, BeginQuote, vbTextCompare)
    If (x > 0) And (x < Len(TheParams)) Then
        If (InStr(x + Len(BeginQuote), TheParams, EndQuote, vbTextCompare) > 0) Then
            If True Or (EndQuote = BeginQuote) Then
                retVal = Mid(TheParams, x + Len(BeginQuote))
                TheParams = Left(TheParams, x - 1) & Mid(retVal, InStr(1, retVal, EndQuote, vbTextCompare) + Len(EndQuote))
                retVal = Left(retVal, InStr(1, retVal, EndQuote, vbTextCompare) - 1)
            Else
                Dim l As Long
                Dim Y As Long
                l = 1
                Y = x
                Do Until l = 0
                    If (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare) > 0) And (InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare) < InStr(Y + Len(BeginQuote), TheParams, EndQuote, vbTextCompare)) Then
                        l = l + 1
                        Y = InStr(Y + Len(BeginQuote), TheParams, BeginQuote, vbTextCompare)
                    ElseIf (InStr(Y + Len(BeginQuote), TheParams, EndQuote, vbTextCompare) > 0) Then
                        l = l - 1
                        Y = InStr(Y + Len(EndQuote), TheParams, EndQuote, vbTextCompare)
                    Else
                        Y = Len(TheParams)
                        l = 0
                    End If
                Loop
                retVal = Mid(TheParams, x + Len(BeginQuote))
                TheParams = Left(TheParams, x - 1) & Mid(retVal, (Y - x) + Len(EndQuote))
                retVal = Left(retVal, (Y - x) - 1)
            End If
        End If
    End If
    RemoveQuotedArg = retVal
End Function

Private Function CountWord(ByVal Text As String, ByVal Word As String) As Long
    Dim cnt As Long
    cnt = UBound(Split(Text, Word, , vbTextCompare))
    If cnt > 0 Then CountWord = cnt
End Function
