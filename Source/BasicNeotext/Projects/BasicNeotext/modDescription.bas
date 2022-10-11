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
Public Function GetCodeModule2(ByRef vbcomp As vbComponent) As CodeModule
    On Error Resume Next
    Set GetCodeModule2 = vbcomp.CodeModule
    If Err.Number <> 0 Then Err.Clear
End Function
Public Function GetCodeModule(ByRef VBProjects As VBProjects, ByVal ProjectName As String, ByVal ModuleName As String) As CodeModule
    Dim vbproj As VBProject
    Dim vbcomp As vbComponent
    Dim Member As Member
    
    For Each vbproj In VBProjects
        If LCase(vbproj.Name) = LCase(ProjectName) Then
            For Each vbcomp In vbproj.VBComponents
                Set GetCodeModule = GetCodeModule2(vbcomp)
                Exit Function
            Next
        End If
    Next
End Function

Public Sub DescriptionsStartup(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As vbComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                    BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), InsertCommentDesc, GetCodeModule2(vbcomp)
            End Select
            Else
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                    BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), DeleteCommentDesc, GetCodeModule2(vbcomp)
            End Select
            End If
        Next
    Next
End Sub

Public Sub UpdateAttributeToCommentDescriptions(ByRef VBProjects As VBProjects)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        Dim vbproj As VBProject
        Dim vbcomp As vbComponent
        Dim Member As Member
        For Each vbproj In VBProjects
            For Each vbcomp In vbproj.VBComponents
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), AttributeToComments, GetCodeModule2(vbcomp)
            End Select
            Next
        Next
    End If
End Sub

Public Sub UpdateCommentToAttributeDescriptions(ByRef VBProjects As VBProjects)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        Dim vbproj As VBProject
        Dim vbcomp As vbComponent
        Dim Member As Member
        For Each vbproj In VBProjects
            For Each vbcomp In vbproj.VBComponents
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), CommentsToAttribute, GetCodeModule2(vbcomp)
            End Select
            Next
        Next
    End If
End Sub

Public Sub InsertDescriptions(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As vbComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner
                BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), InsertCommentDesc, GetCodeModule2(vbcomp)
            End Select
            
        Next
    Next
End Sub

Public Sub DeleteDescriptions(ByRef VBProjects As VBProjects)
    Dim vbproj As VBProject
    Dim vbcomp As vbComponent
    Dim Member As Member
    For Each vbproj In VBProjects
        For Each vbcomp In vbproj.VBComponents
            Select Case vbcomp.Type
                Case vbext_ct_ClassModule, vbext_ct_UserControl, _
                vbext_ct_DocObject, vbext_ct_PropPage, vbext_ct_MSForm, _
                vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_ActiveXDesigner

                BuildComments GetProjectFileName(vbproj.Name, vbcomp.Name), DeleteCommentDesc, GetCodeModule2(vbcomp)

            End Select
        Next
    Next
End Sub
Public Function GetProjectFileName(ByVal ProjName As String, Optional ByVal ModuleName As String, Optional ByVal Projects As Project) As String
    If Projects Is Nothing Then Set Projects = Projs
    Dim proj As Project
    If LCase(Projects.Name) = LCase(ProjName) Then
        If ModuleName <> "" Then
            For Each proj In Projects.Includes
                GetProjectFileName = GetProjectFileName(ModuleName, , proj)
            Next
        Else
            GetProjectFileName = Projects.Location
            Exit Function
        End If
    Else
        
        For Each proj In Projects.Includes
            GetProjectFileName = GetProjectFileName(ProjName, ModuleName, proj)
        Next
    End If

End Function
Private Static Function GetMemberDescription(ByRef Members As Members, ByVal ProcName As String, ByRef LineNum As Long) As String
    Static index As Long
    Dim count As Long
    count = 0
    Do
        index = index + 1
        If index > Members.count Then index = 1
        If LCase(Members(index).Name) = LCase(ProcName) Then
            GetMemberDescription = Members(index).Description
            LineNum = Members(index).CodeLocation
            Exit Function
        End If
        count = count + 1
    Loop Until count > Members.count * 2
End Function

Private Static Sub SetMemberDescription(ByRef Members As Members, ByVal ProcName As String, ByVal ProcDescription As String)
    Static index As Long
    Dim count As Long
    count = 0
    Do
        index = index + 1
        If index > Members.count Then index = 1
        If LCase(Members(index).Name) = LCase(ProcName) Then
            Members(index).Description = ProcDescription
            Exit Sub
        End If
        count = count + 1
    Loop Until count > Members.count * 2
    
'    Dim Member As Member
'    For Each Member In Members
'        If LCase(Member.Name) = LCase(ProcName) Then
'            Member.Description = ProcDescription
'            Exit Sub
'        End If
'    Next
End Sub

Public Sub BuildComments(ByVal FileName As String, ByVal BuildFunc As BuildFunction, ByRef CodeModule As CodeModule)
   'On Error GoTo nochanges
   'On Local Error GoTo nochanges
    
    If Not CodeModule Is Nothing Then

        If CodeModule.CodePane.Window.Visible Then
        
            Debug.Print

            
            
'            Dim ProcName As String
'            Dim ProcDeclare As String
'            Dim ProcDescription As String
'
'            Dim startrow As Long
'            Dim startcol As Long
'            Dim endrow As Long
'            Dim endcol As Long
'            Dim ln As Long
'            Dim changed As Boolean
'
'            Dim txt As String
'            Dim head As String
'            Dim out As String
'            Dim back As String
'
'            CodeModule.CodePane.GetSelection startrow, startcol, endrow, endcol
'
'            Dim mend As Long
'            mend = 1
'            back = CodeModule.Lines(1, CodeModule.CountOfLines)
'
'            txt = vbCrLf & back & vbCrLf
'            Do Until txt = ""
'                out = out & FindNextHeader(txt, head)
'                ProcName = GetUserDefined(head)
'                If ProcName <> "" Then
'
'
'                    ProcDescription = GetMemberDescription(CodeModule.Members, ProcName, ln)
'
'                    ProcDeclare = GetDeclareLine(head)
'
'                    If InStr(" " & ProcDeclare, " Property Get ") > 0 Then
'                        ln = CodeModule.ProcStartLine(ProcName, vbext_ProcKind.vbext_pk_Get)
'                    ElseIf InStr(" " & ProcDeclare, " Property Let ") > 0 Then
'                        ln = CodeModule.ProcStartLine(ProcName, vbext_ProcKind.vbext_pk_Let)
'                    ElseIf InStr(" " & ProcDeclare, " Property Set ") > 0 Then
'                        ln = CodeModule.ProcStartLine(ProcName, vbext_ProcKind.vbext_pk_Set)
'                    ElseIf InStr(" " & ProcDeclare, " Sub ") > 0 Or InStr(ProcDeclare, "Function ") > 0 Then
'                        ln = CodeModule.ProcStartLine(ProcName, vbext_ProcKind.vbext_pk_Proc)
'                    End If
'
'                    ProcDeclare = CodeModule.Lines(ln, 1)
'
'
'                    If ln > 0 Then
'
'                        If (BuildFunc = AttributeToComments) Then
'                            If Right(ProcDeclare, 3) = "' _" Then
'                                If ProcDescription = "" And ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                                    CodeModule.ReplaceLine ln, Left(ProcDeclare, Len(ProcDeclare) - 3)
'                                    CodeModule.DeleteLines ln + 1
'                                    changed = True
'                                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                                ElseIf CodeModule.Lines(ln + 1, 1) <> ProcDescription Then
'                                    CodeModule.ReplaceLine ln + 1, ProcDescription
'                                    changed = True
'                                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                                End If
'                            ElseIf ProcDescription <> "" Then
'                                CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                                CodeModule.InsertLines ln + 1, ProcDescription
'                                changed = True
'                                SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                            End If
'
'                        ElseIf (BuildFunc = CommentsToAttribute) Then
'                            If Right(ProcDeclare, 3) = "' _" Then
'                                ProcDescription = CodeModule.Lines(ln + 1, 1)
'                            ElseIf ProcDescription <> "" And ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                                CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                                CodeModule.InsertLines ln + 1, ProcDescription
'                                changed = True
'                            End If
'                            SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                        Else
'                            If (BuildFunc = DeleteCommentDesc) Then
'                                If Right(ProcDeclare, 3) = "' _" Then
'                                    If ProcDescription = "" And CodeModule.Lines(ln + 1, 1) <> "" Then
'                                        ProcDescription = CodeModule.Lines(ln + 1, 1)
'                                    End If
'                                    CodeModule.ReplaceLine ln, Left(ProcDeclare, Len(ProcDeclare) - 3)
'                                    CodeModule.DeleteLines ln + 1
'                                    changed = True
'                                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                                End If
'                            ElseIf (BuildFunc = InsertCommentDesc) Then
'                                If Right(ProcDeclare, 3) <> "' _" Then
'                                    If ProcDescription <> "" Then
'                                        CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                                        CodeModule.InsertLines ln + 1, ProcDescription
'                                        changed = True
'                                        SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                                    End If
'                                ElseIf ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                                    If ProcDescription = "" Then
'                                     ProcDescription = IIf(ProcDescription = "", CodeModule.Lines(ln + 1, 1), ProcDescription)
'                                     CodeModule.DeleteLines ln + 1
'                                    ' CodeModule.InsertLines ln + 1, IIf(ProcDescription = "", CodeModule.Lines(ln + 1, 1), ProcDescription)
'
'                                     CodeModule.InsertLines ln + 1, ProcDescription
'                                    End If
'                                    changed = True
'                                    SetMemberDescription CodeModule.Members, ProcName, CodeModule.Lines(ln + 1, 1)
'                                End If
'
'                            End If
'                        End If
'                    End If
'                End If
'            Loop
'            If changed Then
'                CodeModule.CodePane.SetSelection startrow, startcol, endrow, endcol
'            End If
            
            
'            Dim txt As String
'            Dim out As String
'            Dim back As String
'            Dim line As String
'            Dim head As String
'
'            Dim DLHead As String
'            Dim UDHead As String
'            Dim DEHead As String
'
'            Dim UDComm As String
'            Dim UDAttr As String
'            Dim DEComm As String
'            Dim DEAttr As String
'            Dim endup As String
'
'
'            If CodeModule.CountOfLines = 0 Then Exit Sub
'
'            out = ""
'            back = CodeModule.Lines(1, CodeModule.CountOfLines)
'
'            txt = vbCrLf & back & vbCrLf
'            Do Until txt = ""
'                out = out & FindNextHeader(txt, head)
'                UDHead = GetUserDefined(head)
'                 If UDHead <> "" Then
'
'                    DLHead = GetDeclareLine(head)
'                    DEHead = GetUserDefined(head, Declared)
'                    UDComm = GetUserDefined(head, Commented)
'                    UDAttr = UDHead 'GetUserDefined(head, Attributed)
'                    DEComm = GetDescription(head, Commented)
'                    DEAttr = GetMemberDescription(CodeModule.Members, UDHead) 'GetDescription(head, Attributed)
'
'                    Debug.Print
'                    Debug.Print "FULL NEXT HEADER INFORMATION"
'                    Debug.Print head
'                    Debug.Print "DECLARE: " & DLHead
'                    Debug.Print "USERDEFINED FROM DECLARE: "; UDComm; " USER DEFINED FROM ATTRIBUTE: " & UDAttr
'                    Debug.Print "COMMENTED DESCRIPTION: "; DEComm; " ATTRIBUTE DESCRIPTION: " & DEAttr
'
'                    If (BuildFunc = AttributeToComments) Or (BuildFunc = InsertCommentDesc) Then
'                        If CountWord(head, vbCrLf) >= 2 Then 'for properties
'                            head = head & "Attribute " & DEHead & ".VB_Description = """ & IIf(DEAttr <> "", DEAttr, DEComm) & """" & vbCrLf
'                            DLHead = GetDeclareLine(head)
'                            DEAttr = GetDescription(head, Attributed)
'                            UDAttr = GetUserDefined(head, Attributed)
'                        End If
'
'                        line = DLHead & " ' _" & vbCrLf & IIf(DEAttr <> "", DEAttr, DEComm) & vbCrLf '& _
'                            IIf(InStr(DLHead, "Event ") = 0, "Attribute " & UDAttr & ".VB_Description = """ & IIf(DEAttr <> "", DEAttr, DEComm) & """" & vbCrLf, "")
'                        SetMemberDescription CodeModule.Members, UDHead, IIf(DEAttr <> "", DEAttr, DEComm)
'                    ElseIf (BuildFunc = CommentsToAttribute) Then
'                        line = DLHead & " ' _" & vbCrLf & IIf(DEComm <> "", DEComm, DEAttr) & vbCrLf '& _
'                            IIf(InStr(DLHead, "Event ") = 0, "Attribute " & DEHead & ".VB_Description = """ & IIf(DEComm <> "", DEComm, DEAttr) & """" & vbCrLf, "")
'
'                        SetMemberDescription CodeModule.Members, UDHead, IIf(DEComm <> "", DEComm, DEAttr)
'                    Else
'                        line = DLHead & vbCrLf '& _
'                            IIf(InStr(DLHead, "Event ") = 0, "Attribute " & UDAttr & ".VB_Description = """ & IIf(DEAttr = "", DEComm, DEAttr) & """" & vbCrLf, "")
'                        SetMemberDescription CodeModule.Members, UDHead, IIf(DEAttr = "", DEComm, DEAttr)
'                    End If
'                    If head <> line Then changed = True
'
'                    out = out & line
'
'                Else
'                    out = out & head
'                End If
'            Loop
'            out = out & endup
'            If Mid(out, 3, Len(out) - 4) <> back Then
'                CodeModule.DeleteLines 1, CodeModule.CountOfLines
'                CodeModule.InsertLines 1, Mid(out, 3, Len(out) - 4)
'                'CodeModule.AddFromString Mid(out, 3, Len(out) - 4)
'            End If

            
            
            
            
            
            
            
'            CodeModule.CodePane.GetSelection startrow, startcol, endrow, endcol
'            Dim mend As Long
'            mend = 1
'            Do While mend <= CodeModule.Members.count
'
'                ProcName = CodeModule.Members(mend).Name
'                ln = CodeModule.Members(mend).CodeLocation
'                Do
'                    ProcDeclare = RTrimStrip(CodeModule.Lines(ln, 1), " ")
'                    ln = ln + 1
'                Loop While ProcDeclare = ""
'                ln = ln - 1
'
'                ProcDescription = CodeModule.Members(mend).Description
'
'                If (BuildFunc = AttributeToComments) Then
'                    If Right(ProcDeclare, 3) = "' _" Then
'                        If ProcDescription = "" And ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                            CodeModule.ReplaceLine ln, Left(ProcDeclare, Len(ProcDeclare) - 3)
'                            CodeModule.DeleteLines ln + 1
'                            changed = True
'                            SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                        ElseIf CodeModule.Lines(ln + 1, 1) <> ProcDescription Then
'                            CodeModule.ReplaceLine ln + 1, ProcDescription
'                            changed = True
'                            SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                        End If
'                    ElseIf ProcDescription <> "" Then
'                        CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                        CodeModule.InsertLines ln + 1, ProcDescription
'                        changed = True
'                        SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                    End If
'
'                ElseIf (BuildFunc = CommentsToAttribute) Then
'                    If Right(ProcDeclare, 3) = "' _" Then
'                        ProcDescription = CodeModule.Lines(ln + 1, 1)
'                    ElseIf ProcDescription <> "" And ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                        CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                        CodeModule.InsertLines ln + 1, ProcDescription
'                        changed = True
'                    End If
'                    SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                Else
'                    If (BuildFunc = DeleteCommentDesc) Then
'                        If Right(ProcDeclare, 3) = "' _" Then
'                            If ProcDescription = "" And CodeModule.Lines(ln + 1, 1) <> "" Then
'                                ProcDescription = CodeModule.Lines(ln + 1, 1)
'                            End If
'                            CodeModule.ReplaceLine ln, Left(ProcDeclare, Len(ProcDeclare) - 3)
'                            CodeModule.DeleteLines ln + 1
'                            changed = True
'                            SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                        End If
'                    ElseIf (BuildFunc = InsertCommentDesc) Then
'                        If Right(ProcDeclare, 3) <> "' _" Then
'                            If ProcDescription <> "" Then
'                                CodeModule.ReplaceLine ln, ProcDeclare & " ' _"
'                                CodeModule.InsertLines ln + 1, ProcDescription
'                                changed = True
'                                SetMemberDescription CodeModule.Members, ProcName, ProcDescription
'                            End If
'                        ElseIf ProcDescription <> CodeModule.Lines(ln + 1, 1) Then
'                            CodeModule.InsertLines ln + 1, IIf(ProcDescription = "", CodeModule.Lines(ln + 1, 1), ProcDescription)
'                            changed = True
'                            SetMemberDescription CodeModule.Members, ProcName, CodeModule.Lines(ln + 1, 1)
'                        End If
'
'                    End If
'                End If
'                mend = mend + 1
'            Loop
'            If changed Then
'                CodeModule.CodePane.SetSelection startrow, startcol, endrow, endcol
'            End If
            
            

        End If
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
        Case "bas"
            out = ""
            back = ReadFile(FileName)

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

Private Function GetDeclareLine(ByVal head As String) As String
    If InStr(head, "' _" & vbCrLf) > 0 Then
        GetDeclareLine = RTrimStrip(NextArg(head, "' _" & vbCrLf), " ")
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
                Case "Dim", "Const", "Declare"
                '    GetUserDefined = NextArg(NextArg(NextArg(RemoveArg(head, " "), " "), "("), ",")
                Case "Type", "Enum"
                '    GetUserDefined = NextArg(NextArg(RemoveArg(head, " "), " "), "(")
                Case "Event"
                    GetUserDefined = NextArg(NextArg(NextArg(RemoveArg(head, " "), " "), "("), ",")
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

