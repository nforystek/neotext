Attribute VB_Name = "modDescription"

Option Explicit

Option Compare Text

Private Enum HeaderInfo
    Declared = 0
    Commented = 1
    Attributed = 2
End Enum

Public Enum BuildFunction
    AttributeToComments = 1
    CommentsToAttribute = 2
    InsertCommentDesc = 3
    DeleteCommentDesc = 4
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

Public Function GetModuleName(ByVal FileName As String) As String
    If PathExists(FileName, True) Then
        GetModuleName = NextArg(RemoveArg(ReadFile(FileName), "Attribute VB_Name="), """")
    End If
End Function
Public Function GetCodeModule2(ByRef vbcomp As VBComponent) As CodeModule
    On Error Resume Next
    Set GetCodeModule2 = vbcomp.CodeModule
    If Err.Number <> 0 Then Err.Clear
End Function
Public Function GetCodeModule(ByRef VBProjects As VBProjects, ByVal ProjectName As String, ByVal ModuleName As String) As CodeModule
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim Member As Member
    
    For Each vbproj In VBProjects
        If LCase(vbproj.Name) = LCase(ProjectName) Then
            For Each vbcomp In vbproj.VBComponents
                'If vbcomp.Name = ModuleName Then
                    Set GetCodeModule = GetCodeModule2(vbcomp)
                    Exit Function
                'End If
            Next
        End If
    Next
End Function

Public Sub DescriptionsStartup(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
                'If vbcomp.HasOpenDesigner Then
                    BuildComments InsertCommentDesc, GetCodeModule2(vbcomp)
                'End If
            Else
                'If vbcomp.HasOpenDesigner Then
                    BuildComments DeleteCommentDesc, GetCodeModule2(vbcomp)
                'End If
            End If
        Next
    Next
End Sub

Public Sub UpdateAttributeToCommentDescriptions(ByRef VBProjects As VBProjects)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        Dim vbproj As VBProject
        Dim vbcomp As VBComponent
        Dim Member As Member
        For Each vbproj In VBProjects
            For Each vbcomp In vbproj.VBComponents
                'If vbcomp.HasOpenDesigner Then
                    BuildComments AttributeToComments, GetCodeModule2(vbcomp)
                'End If
            Next
        Next
    End If
End Sub

Public Sub UpdateCommentToAttributeDescriptions(ByRef VBProjects As VBProjects)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        Dim vbproj As VBProject
        Dim vbcomp As VBComponent
        Dim Member As Member
        For Each vbproj In VBProjects
            For Each vbcomp In vbproj.VBComponents
                'If vbcomp.HasOpenDesigner Then
                    BuildComments CommentsToAttribute, GetCodeModule2(vbcomp)
                'End If
            Next
        Next
    End If
End Sub

Public Sub InsertDescriptions(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            'If vbcomp.HasOpenDesigner Then
                BuildComments InsertCommentDesc, GetCodeModule2(vbcomp)
            'End If
        Next
    Next
End Sub

Public Sub DeleteDescriptions(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            'If vbcomp.HasOpenDesigner Then
                BuildComments DeleteCommentDesc, GetCodeModule2(vbcomp)
            'End If
        Next
    Next
End Sub

Private Sub SetMemberDescription(ByRef Members As Members, ByVal ProcName As String, ByVal ProcDescription As String)
    Dim Member As Member
    For Each Member In Members
        If LCase(Member.Name) = LCase(ProcName) Then
            Member.Description = ProcDescription
            Exit Sub
        End If
    Next
End Sub

Public Sub BuildComments(ByVal BuildFunc As BuildFunction, ByRef CodeModule As CodeModule)
    On Error GoTo nochanges
    On Local Error GoTo nochanges
    
    If Not CodeModule Is Nothing Then

        Dim ProcName As String
        Dim ProcDeclare As String
        Dim ProcDescription As String
        
        Dim startrow As Long
        Dim startcol As Long
        Dim endrow As Long
        Dim endcol As Long
        
        CodeModule.CodePane.GetSelection startrow, startcol, endrow, endcol
    
        Dim Member As Member
        For Each Member In CodeModule.Members
            ProcName = Member.Name
            ProcDeclare = RTrimStrip(CodeModule.Lines(Member.CodeLocation, 1), " ")
            ProcDescription = Member.Description
    
            If (BuildFunc = AttributeToComments) Then
                If Right(ProcDeclare, 3) = "' _" Then
                    If ProcDescription = "" Then
                        CodeModule.ReplaceLine Member.CodeLocation, Left(ProcDeclare, Len(ProcDeclare) - 3)
                        CodeModule.DeleteLines Member.CodeLocation + 1
                    Else
                        CodeModule.ReplaceLine Member.CodeLocation + 1, ProcDescription
                    End If
                ElseIf ProcDescription <> "" Then
                    CodeModule.ReplaceLine Member.CodeLocation, ProcDeclare & " ' _" & vbCrLf & ProcDescription
                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
                End If
            ElseIf (BuildFunc = CommentsToAttribute) Then
                If Right(ProcDeclare, 3) = "' _" Then
                    ProcDescription = CodeModule.Lines(Member.CodeLocation + 1, 1)
                ElseIf ProcDescription <> "" Then
                    CodeModule.ReplaceLine Member.CodeLocation, ProcDeclare & " ' _" & vbCrLf & ProcDescription
                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
                End If
                SetMemberDescription CodeModule.Members, ProcName, ProcDescription
            Else
                ProcDescription = Member.Description
                If (BuildFunc = DeleteCommentDesc) Then
                    If Right(ProcDeclare, 3) = "' _" Then
                        If CodeModule.Lines(Member.CodeLocation + 1, 1) = ProcDescription Then
                            CodeModule.ReplaceLine Member.CodeLocation, Left(ProcDeclare, Len(ProcDeclare) - 3)
                            CodeModule.DeleteLines Member.CodeLocation + 1
                        Else
                            CodeModule.ReplaceLine Member.CodeLocation, Left(ProcDeclare, Len(ProcDeclare) - 3)
                            CodeModule.ReplaceLine Member.CodeLocation + 1, "'" & CodeModule.Lines(Member.CodeLocation, 1)
                        End If
                        SetMemberDescription CodeModule.Members, ProcName, ProcDescription
                    End If
                ElseIf (BuildFunc = InsertCommentDesc) Then
                    If Right(ProcDeclare, 3) <> "' _" And ProcDescription <> "" Then
                        CodeModule.ReplaceLine Member.CodeLocation, ProcDeclare & " ' _" & vbCrLf & ProcDescription
                        SetMemberDescription CodeModule.Members, ProcName, ProcDescription
                    End If
                End If
            End If
        Next
        
        CodeModule.CodePane.SetSelection startrow, startcol, endrow, endcol
        
    End If
    
    Exit Sub
nochanges:
    Err.Clear
End Sub

Public Sub BuildProject(ByVal FileName As String)
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
    Dim ret As String
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


