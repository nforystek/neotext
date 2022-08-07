#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modScript"
#Const modScript = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module


Public Enumerators As New Collection
Public Procedures As New Collection
Public Function EnumeratorExists(ByVal EnumName As String) As Boolean
    On Error Resume Next
    Dim chk As clsEnumerator
    Set chk = Enumerators(EnumName)
    If Err = 0 Then
        EnumeratorExists = True
    Else
        Err.Clear
        EnumeratorExists = False
    End If
    Set chk = Nothing
End Function
Public Function ScriptToDisplay(ByVal ProcName As String) As String
    Dim retVal As String
    retVal = Replace(ProcName, "__", " - ")
    retVal = Replace(retVal, "_", " ")
    ScriptToDisplay = retVal
End Function
Public Function DisplayToScript(ByVal Displayname As String) As String
    Dim retVal As String
    retVal = Replace(Displayname, " - ", "__")
    retVal = Replace(retVal, " ", "_")
    DisplayToScript = retVal
End Function

Private Function AddEnumerator(ByVal EnumeratorName As String) As Integer
    Dim newEnum As New clsEnumerator
    newEnum.EnumeratorName = EnumeratorName
    Enumerators.Add newEnum, EnumeratorName
    AddEnumerator = Enumerators.Count
    Set newEnum = Nothing
End Function

Private Function ClearEnumerators()
    Do Until Enumerators.Count = 0
        Enumerators(1).ClearEnumValues
        Enumerators.Remove 1
    Loop
End Function

Private Function AddProcedure(ByVal ProcedureName As String) As Integer
    Dim newProc As New clsProcedure
    newProc.ProcedureName = ProcedureName
    Procedures.Add newProc, ProcedureName
    AddProcedure = Procedures.Count
    Set newProc = Nothing
End Function

Private Function ClearProcedures()
    Do Until Procedures.Count = 0
        Procedures(1).ClearParameters
        Procedures.Remove 1
    Loop
End Function
Public Function ClearRemindMeScript()
    ClearEnumerators
    ClearProcedures
End Function

Public Function GetRemindMeScript(ByVal Language As String, ByVal Code As String) As Long
    On Error GoTo catch
    
    Dim tCode As String
    Dim lineNum As Long
    Dim nextLine As String
    Dim inParam As String
    Dim inParamType As String
    Dim newProc As Integer

    lineNum = 0
    tCode = Code
    Do Until tCode = ""

        nextLine = TrimStrip(RemoveNextArg(tCode, vbCrLf), Chr(8))

        lineNum = lineNum + 1
        
        If (Left(nextLine, 1) = "'") And (Language = "VBScript") Then
            nextLine = Mid(nextLine, 2)
        ElseIf (Left(nextLine, 2) = "//") And (Language = "JScript") Then
            nextLine = Mid(nextLine, 3)
        Else
            nextLine = ""
        End If

        If Left(LCase(nextLine), 8) = "remindme" Then
            RemoveNextArg nextLine, ":"
        Else
            nextLine = ""
        End If

        If nextLine <> "" Then
            inParam = RemoveNextArg(nextLine, ":")
            Select Case LCase(inParam)
                Case "sub", "function"
                    inParam = RemoveNextArg(nextLine, "(")
                    nextLine = Replace(nextLine, ")", "")
                    newProc = AddProcedure(inParam)
                    If nextLine <> "" Then

                        Do Until nextLine = ""
                            inParam = RemoveNextArg(nextLine, ":")
                            inParamType = RemoveNextArg(nextLine, ",")
                            Procedures(newProc).AddParameter inParam, inParamType
                        Loop
                    End If
                Case "enum"
                    inParam = RemoveNextArg(nextLine, "(")
                    nextLine = Replace(nextLine, ")", "")
                    newProc = AddEnumerator(inParam)
                    If nextLine <> "" Then

                        Do Until nextLine = ""
                            inParam = RemoveNextArg(nextLine, ":")
                            inParamType = RemoveNextArg(nextLine, ",")
                            Enumerators(newProc).AddEnumValue inParam, inParamType
                        Loop
                    End If
            End Select

        End If
    Loop
    
    GetRemindMeScript = 0
    Exit Function
catch:
    Err.Clear
    GetRemindMeScript = lineNum
End Function

