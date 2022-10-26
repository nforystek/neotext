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

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

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
Public Function GetCodeModuleByCaption(ByRef VBInstance As VBE, ByVal Caption As String) As CodeModule
    Dim vbproj As VBProject
    Dim vbcomp As VBComponent
    Dim cm As CodeModule
    
    Dim Member As Member
    For Each vbproj In VBInstance.VBProjects
        For Each vbcomp In vbproj.VBComponents
            If InStr(Caption, " " & vbcomp.Name & " ") > 0 Then
                Set cm = GetCodeModule2(vbcomp)
                If Not cm Is Nothing Then
                    If cm.CodePane.Window.Caption = Caption Then
                        Set GetCodeModuleByCaption = cm
                        Set cm = Nothing
                        Exit Function
                    End If
                End If
                Set cm = Nothing
            End If
            
        Next
    Next
End Function

Public Function GetProjectNameByCodeModule(ByRef CodeModule As CodeModule) As String
    If Not CodeModule Is Nothing Then
        Dim vbproj As VBProject
        Dim vbcomp As VBComponent
        For Each vbproj In CodeModule.VBE.VBProjects
            For Each vbcomp In vbproj.VBComponents
                If ObjPtr(vbcomp) = ObjPtr(CodeModule.Parent) Or ObjPtr(vbcomp.CodeModule) = ObjPtr(CodeModule) Then
                    GetProjectNameByCodeModule = vbproj.Name
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Public Sub DescriptionsStartup(ByRef VBInstance As VBIDE.VBE)
    If Hooks.count > 0 Then
        Dim cnt As Long
        For cnt = 1 To Hooks.count
            If IsWindowVisible(Hooks(cnt).hWnd) Then
                If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
                    BuildComments VBInstance, InsertCommentDesc, Hooks(cnt)
                Else
                    BuildComments VBInstance, DeleteCommentDesc, Hooks(cnt)
                End If
            End If
        Next
    End If
End Sub

Public Sub UpdateAttributeToCommentDescriptions(ByRef VBInstance As VBIDE.VBE)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        If Hooks.count > 0 Then
            Dim cnt As Long
            For cnt = 1 To Hooks.count
                If IsWindowVisible(Hooks(cnt).hWnd) Then
                    BuildComments VBInstance, AttributeToComments, Hooks(cnt)
                End If
            Next
        End If

    End If
End Sub

Public Sub UpdateCommentToAttributeDescriptions(ByRef VBInstance As VBIDE.VBE)
    If CLng(GetSetting("BasicNeotext", "Options", "ProcedureDesc", 0)) = 1 Then
        If Hooks.count > 0 Then
            Dim cnt As Long
            For cnt = 1 To Hooks.count
                If IsWindowVisible(Hooks(cnt).hWnd) Then
                    BuildComments VBInstance, CommentsToAttribute, Hooks(cnt)
                End If
            Next
        End If

    End If
End Sub

Public Sub InsertDescriptions(ByRef VBInstance As VBIDE.VBE)
    If Hooks.count > 0 Then
        Dim cnt As Long
        For cnt = 1 To Hooks.count
            If IsWindowVisible(Hooks(cnt).hWnd) Then
                BuildComments VBInstance, InsertCommentDesc, Hooks(cnt)
            End If
        Next
    End If
End Sub

Public Sub DeleteDescriptions(ByRef VBInstance As VBIDE.VBE)
    If Hooks.count > 0 Then
        Dim cnt As Long
        For cnt = 1 To Hooks.count
            If IsWindowVisible(Hooks(cnt).hWnd) Then
                BuildComments VBInstance, DeleteCommentDesc, Hooks(cnt)
            End If
        Next
    End If
End Sub

Public Function GetProjectFileName(ByVal ProjName As String, Optional ByVal ModuleName As String, Optional ByVal Projects As Project) As String
    If Projects Is Nothing Then Set Projects = Projs
    Dim proj As Project
    If LCase(Projects.Name) = LCase(ProjName) Then
        If ModuleName <> "" Then
            For Each proj In Projects.Includes
                GetProjectFileName = GetProjectFileName(ModuleName, , proj)
                If GetProjectFileName <> "" Then Exit Function
            Next
        Else
            GetProjectFileName = Projects.Location
            Exit Function
        End If
    Else
        For Each proj In Projects.Includes
            GetProjectFileName = GetProjectFileName(ProjName, ModuleName, proj)
            If GetProjectFileName <> "" Then Exit Function
        Next
    End If
End Function

Private Static Function GetMemberDescription(ByRef Members As Members, ByVal ProcName As String, Optional ByRef LineNum As Long = 0) As String
    On Error Resume Next
    On Local Error Resume Next


    GetMemberDescription = Members(ProcName).Description
    If Err.Number = 0 Then
        LineNum = Members(ProcName).CodeLocation
    Else
        Err.Clear
        LineNum = ""
        GetMemberDescription = ""
    End If
    On Error GoTo -1
    On Local Error GoTo -1
'    Static Index As Long
'    Dim count As Long
'    count = 0
'    Do
'        Index = Index + 1
'        If Index > Members.count Then Index = 1
'        If LCase(Members(Index).Name) = LCase(ProcName) Then
'            GetMemberDescription = Replace(Replace(Members(Index).Description, vbCrLf, vbLf), vbLf, "")
'            LineNum = Members(Index).CodeLocation
'            Exit Function
'        End If
'        count = count + 1
'    Loop Until count > Members.count * 2
End Function

Private Static Sub SetMemberDescription(ByRef Members As Members, ByVal ProcName As String, ByVal ProcDescription As String)
    On Error Resume Next
    On Local Error Resume Next

    Members(ProcName).Description = ProcDescription
    If Err.Number <> 0 Then Err.Clear

    On Error GoTo -1
    On Local Error GoTo -1
    
'    Static Index As Long
'    Dim count As Long
'    count = 0
'    Do
'        Index = Index + 1
'        If Index > Members.count Then Index = 1
'        If LCase(Members(Index).Name) = LCase(ProcName) Then
'            Members(Index).Description = ProcDescription
'            Exit Sub
'        End If
'        count = count + 1
'    Loop Until count > Members.count * 2
End Sub

Public Function GetTemporaryFile() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetTempFileName(GetTemporaryFolder, App.Title, 0, winDir)
    If ret = 0 Then
        winDir = GetTemporaryFolder & "\" & Left(Left(App.Title, 3) & Hex(CLng(Mid(CStr(Rnd), 3))), 14) & ".tmp"
        ret = FreeFile
        Open winDir For Output As #ret
        Close #ret
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
    End If
    GetTemporaryFile = winDir
End Function

Private Sub SumPropertyHeaders(ByRef txt As String, ByRef head As String, Optional ByRef frm As FormHWnd = Nothing)
    'this goes ahead to get all a specific property defintions between
    'one, two or three and get, let and set and uses only the 1st as it
    'appears visual basic uses 1st with desc or else placing it on the 1st
    
    Dim back As String
    Dim head2 As String
    Dim head3 As String
    Dim user As String
    Dim out As String
    back = txt
    user = GetUserDefined(head)
    Do Until txt = ""
        out = out & FindNextHeader(txt, head2, user)
        If GetUserDefined(head2) = user Then
            out = out & GetDeclareLine(head2, True) & vbCrLf
            
            Do Until txt = ""
                out = out & FindNextHeader(txt, head3, user)
                If GetUserDefined(head3) = user Then
                    out = out & GetDeclareLine(head3, True) & vbCrLf
                    Exit Do
                Else
                    out = out & head3
                End If
            Loop
            Exit Do
        Else
            out = out & head2
        End If
    Loop

    Dim desc1 As String
    Dim desc2 As String
    Dim desc3 As String
    desc1 = GetDescription(head, Commented)
    desc2 = GetDescription(head2, Commented)
    desc3 = GetDescription(head3, Commented)

    If desc1 = "" Then desc1 = GetDescription(head, Attributed)
    If desc2 = "" Then desc2 = GetDescription(head2, Attributed)
    If desc3 = "" Then desc3 = GetDescription(head2, Attributed)
    
    If Not frm Is Nothing Then
        If desc1 = "" Then desc1 = GetMemberDescription(frm.CodeModule.Members, user)
        If desc2 = "" Then desc2 = GetMemberDescription(frm.CodeModule.Members, user)
        If desc3 = "" Then desc3 = GetMemberDescription(frm.CodeModule.Members, user)
    End If
            
    If (desc2 <> "") Then
        If InStr(desc1, desc2) = 0 Then
            desc1 = IIf(desc1 <> "", desc1 & ", " & desc2, desc2)
        End If
    End If

    If (desc3 <> "") Then
        If InStr(desc1, desc3) = 0 Then
            desc1 = IIf(desc1 <> "", desc1 & ", " & desc3, desc3)
        End If
    End If
    
    If desc1 <> "" Then
        head = GetDeclareLine(head, True) & " ' _" & vbCrLf & desc1 & vbCrLf & _
            "Attribute " & user & ".VB_Description = """ & desc1 & """" & vbCrLf
    Else
        head = GetDeclareLine(head, True) & vbCrLf
    End If
    
    txt = out & txt
                         
End Sub

Public Function BuildComments(ByRef VBInstance As VBIDE.VBE, ByVal BuildFunc As BuildFunction, ByRef frm As FormHWnd) As Boolean
  ' On Error GoTo nochanges
  ' On Local Error GoTo nochanges
   
    With frm
        
        If IsWindowVisible(frm.hWnd) Then
            If .CodeModule Is Nothing Then
                Set .CodeModule = GetCodeModuleByCaption(VBInstance, GetCaption(frm.hWnd))
            End If
            
            If .CodeModule Is Nothing Then Exit Function
        
            Dim vbcomp As VBComponent
            Set vbcomp = .CodeModule.Parent
            
            Dim startrow As Long
            Dim startcol As Long
            Dim endrow As Long
            Dim endcol As Long
            Dim changed As Boolean
    
            MSVBRedraw False
            
            .CodeModule.CodePane.GetSelection startrow, startcol, endrow, endcol
            
            Dim txt As String
            Dim out As String
            Dim user As String
            Dim desc As String
            Dim head As String
            Dim back As String
    
            Dim SavedCode As String
            Dim SavedFile As String
            Dim SavedFile2 As String
            Dim UnsavedCode As String
            Dim TempFile As String
            Dim propsDone As String
            
            Dim cnt As Long
            If .CodeModule.Parent.FileCount > 0 Then
    
                Select Case GetFileExt(.CodeModule.Parent.FileNames(1), True, True)
                    Case "bas", "cls", "ctl", "frm", "dsr", "pag", "dob"
                    
                        If PathExists(.CodeModule.Parent.FileNames(1), True) Then
                            SavedFile = .CodeModule.Parent.FileNames(1)
                        Else
                            SavedFile = GetProjectFileName(GetProjectNameByCodeModule(.CodeModule), .CodeModule.Parent.Name)
                        End If
                        If .CodeModule.Parent.FileCount = 2 Then SavedFile2 = GetFilePath(SavedFile) & "\" & GetFileName(.CodeModule.Parent.FileNames(2))
    
                        If PathExists(SavedFile, True) Then
                            SavedCode = ReadFile(SavedFile)
                        Else
                            SavedFile = ""
                            SavedFile2 = ""
                        End If
                        
                    Case Else
                        SavedFile = ""
                        SavedFile2 = ""
                End Select
    
            End If
            
            If SavedFile <> "" Then
                If .CodeModule.CountOfLines > 0 Then
                    Dim cap As String
                    cap = GetCaption(frm.hWnd)
                    
                    UnsavedCode = .CodeModule.Lines(1, .CodeModule.CountOfLines)
                    
                    TempFile = GetTemporaryFile
                    TempFile = GetFilePath(SavedFile) & "\" & GetFileTitle(TempFile) & GetFileExt(.CodeModule.Parent.FileNames(1))
                    
                    .CodeModule.Parent.SaveAs TempFile
                    
                    Set .CodeModule = GetCodeModuleByCaption(VBInstance, GetCaption(frm.hWnd))
   
                    out = ""
                    back = ReadFile(TempFile)
        
                    txt = vbCrLf & back & vbCrLf
    
                    Do Until txt = ""
                        out = out & FindNextHeader(txt, head)
                        user = GetUserDefined(head)
                        If user <> "" Then
                        
                            If IsPropertyHeader(head) Then
                                If InStr(propsDone, " " & user & " ") = 0 Then
                                    SumPropertyHeaders txt, head, frm
                                    propsDone = propsDone & " " & user & " "
                                Else
                                    user = ""
                                End If
                            End If
                            
                            If user <> "" Then
                                
'                                Debug.Print
'                                Debug.Print "FULL NEXT HEADER INFORMATION"
'                                Debug.Print head
'                                Debug.Print "DECLARE: " & GetDeclareLine(head, False)
'                                Debug.Print "USERDEFINED FROM DECLARE: "; GetUserDefined(head, Commented); " USER DEFINED FROM ATTRIBUTE: " & GetUserDefined(head, Attributed)
'                                Debug.Print "COMMENTED DESCRIPTION: "; GetDescription(head, Commented); " ATTRIBUTE DESCRIPTION: " & GetDescription(head, Attributed)
     
                                If (BuildFunc = AttributeToComments Or BuildFunc = InsertCommentDesc) Then
                                    
                                    If GetDescription(head, Attributed) = "" And GetDescription(head, Commented) <> "" Then
                                        desc = GetDescription(head, Commented)
                                    Else
                                        desc = GetDescription(head, Attributed)
                                    End If
    
                                    If desc = "" Then
                                        desc = GetMemberDescription(.CodeModule.Members, user)
                                    End If
                                    
                                    If desc <> "" Then
                                        out = out & GetDeclareLine(head, True) & " ' _" & vbCrLf & desc & vbCrLf & _
                                            "Attribute " & GetUserDefined(head, Attributed) & ".VB_Description = """ & desc & """" & vbCrLf
                                    Else
                                        out = out & GetDeclareLine(head, True) & vbCrLf
                                    End If
                                ElseIf BuildFunc = CommentsToAttribute Then
                                
                                    If GetDescription(head, Commented) = "" And GetDescription(head, Attributed) <> "" Then
                                        desc = GetDescription(head, Attributed)
                                    Else
                                        desc = GetDescription(head, Commented)
                                    End If
                                    If desc = "" Then
                                        desc = GetMemberDescription(.CodeModule.Members, user)
                                    End If
                                    
                                    If desc <> "" Then
                                        out = out & GetDeclareLine(head, True) & " ' _" & vbCrLf & desc & vbCrLf & _
                                            "Attribute " & GetUserDefined(head, Declared) & ".VB_Description = """ & desc & """" & vbCrLf
                                    Else
                                        out = out & GetDeclareLine(head, True) & vbCrLf
                                    End If
                                Else
                                    out = out & GetDeclareLine(head, True) & vbCrLf
                                End If
                            Else
                                out = out & head
                            End If
                            
                        Else
                            out = out & head
                        End If
                    Loop
        
                    If Mid(out, 3, Len(out) - 4) <> back Then
                                        
                        WriteFile TempFile, Mid(out, 3, Len(out) - 4)
                        changed = True

                    End If
    
                    Set vbcomp = .CodeModule.Parent
                    
                    .CodeModule.Parent.Reload

                    Set .CodeModule = vbcomp.CodeModule
                    
                    .CodeModule.Parent.SaveAs SavedFile
                    
                    .CodeModule.Parent.Reload
                        
                    WriteFile SavedFile, SavedCode
    
                    Kill TempFile
                    If PathExists(GetFilePath(SavedFile) & "\" & GetFileTitle(TempFile) & GetFileExt(SavedFile2), True) Then
                        Kill GetFilePath(SavedFile) & "\" & GetFileTitle(TempFile) & GetFileExt(SavedFile2)
                    End If
                                    
                End If
                Set .CodeModule = vbcomp.CodeModule
                
                BuildComments = True

                .CodeModule.CodePane.SetSelection startrow, startcol, endrow, endcol
    
            End If
            
            MSVBRedraw True
        End If
        
    End With
    
    Exit Function
nochanges:
    Err.Clear
    MSVBRedraw True
End Function

Public Sub BuildFileDescriptions(ByVal FileName As String, ByVal LoadElseSave As Boolean)
    'file only version of BuildComments, it does not access VBproject objects
    
    On Error GoTo nochanges
    On Local Error GoTo nochanges
    
    Dim txt As String
    Dim out As String
    Dim user As String
    Dim desc As String
    Dim head As String
    Dim back As String
    Dim propsDone As String
    
    Select Case GetFileExt(FileName, True, True)
        Case "cls", "ctl", "frm", "dsr", "pag", "dob"
            out = ""
            back = ReadFile(FileName)

            txt = vbCrLf & back & vbCrLf
            Do Until txt = ""
                out = out & FindNextHeader(txt, head)
                user = GetUserDefined(head)
                If user <> "" Then
                
                    If IsPropertyHeader(head) Then
                        If InStr(propsDone, " " & user & " ") = 0 Then
                            SumPropertyHeaders txt, head
                            propsDone = propsDone & " " & user & " "
                        Else
                            user = ""
                        End If
                    End If
                            
                    If user <> "" Then
'                        Debug.Print
'                        Debug.Print "FULL NEXT HEADER INFORMATION"
'                        Debug.Print head
'                        Debug.Print "DECLARE: " & GetDeclareLine(head, False)
'                        Debug.Print "USERDEFINED FROM DECLARE: "; GetUserDefined(head, Commented); " USER DEFINED FROM ATTRIBUTE: " & GetUserDefined(head, Attributed)
'                        Debug.Print "COMMENTED DESCRIPTION: "; GetDescription(head, Commented); " ATTRIBUTE DESCRIPTION: " & GetDescription(head, Attributed)

                        If LoadElseSave Then
                            desc = GetDescription(head, Attributed)
                            If desc = "" Then desc = GetDescription(head, Commented)
                            out = out & GetDeclareLine(head, True) & IIf(desc <> "", " ' _" & vbCrLf & desc & vbCrLf & _
                                "Attribute " & GetUserDefined(head, Attributed) & ".VB_Description = """ & desc & """" & vbCrLf, vbCrLf)
                        Else
                            desc = GetDescription(head, Commented)
                            If desc = "" Then desc = GetDescription(head, Attributed)
                            out = out & GetDeclareLine(head, True) & IIf(desc <> "", " ' _" & vbCrLf & desc & vbCrLf & _
                                "Attribute " & GetUserDefined(head, Declared) & ".VB_Description = """ & desc & """" & vbCrLf, vbCrLf)
                        End If
                    Else
                        out = out & head
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

Public Sub BuildProject(ByVal FileName As String)
    'this is for any project modifications we do
    'right now it is just the fixing of spaces in
    'compile conditions, which as allowed we setup
    'a [neotext] section to track differences we do
    
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
    'inserts const for each module into compile conditions and a VBIDE const as well
    'for debug and release in case it may be used, such as apppath function does here
    'to sort a binary and release style of implementation capable like that c++ does
    
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

Private Function SortText(ByVal Text As String, ByRef FindText1 As String, ByRef FindText2 As String, ByRef FindLoc1 As Long, ByRef FindLoc2 As Long) As Boolean
    'sorts two findtext strings and their findloc as order of found with in text and returns true if any of either is found
    FindLoc1 = InStr(Text, FindText1)
    FindLoc2 = InStr(Text, FindText2)
    
    If (((FindLoc1 > FindLoc2) Or (FindLoc1 = 0)) And (Not FindLoc2 = 0)) Then
        Swap FindText1, FindText2
        Swap FindLoc1, FindLoc2
    End If
    
    SortText = ((FindLoc1 <> 0) Or (FindLoc2 <> 0))
End Function

Private Function FindNextHeader(ByRef txt As String, ByRef head As String, Optional ByVal PropertyName As String = "") As String
    'searches txt for headers, returning any txt before the next header found
    'placing the header in head, altering txt to be all txt after a found header
    'if no header is found, the return is all of txt, and txt and head are blank
    '
    'a head is defined by any declaritive code that vb compiles hidden text
    'descriptions for, that are shown in the saved file when viewed as text
    'for our purpose, a comment trailing with underscore is also to be used
    'as description entered on the next line after the declaration, dually
    'then adaptive to the hidden text header with VB, to edit descriptions
    
    Dim pos As Long
    head = ""
    
    Do While txt <> ""
        pos = FindNextLine(txt, PropertyName)
        If pos > -1 Then
            FindNextHeader = FindNextHeader & Left(txt, pos - 1)
            txt = Mid(txt, pos)
            
            'remove any intro vbcrlf off the header
            Do While Left(txt, 2) = vbCrLf
                FindNextHeader = FindNextHeader & vbCrLf
                txt = Mid(txt, 3)
            Loop
            
            pos = FindLineEnd(txt, (PropertyName <> ""))
            If pos > 0 Then
                head = Left(txt, pos)
                txt = Mid(txt, pos + 1)
               ' Stop
                'comment descriptions and attribute descriptions may only be one line
                If Right(head, 5) = "' _" & vbCrLf Then
                    head = head & RTrimStrip(RemoveNextArg(txt, vbCrLf, , False), " ") & vbCrLf
                    If Left(LCase(txt), 10) = "attribute " Then
                        head = head & RTrimStrip(RemoveNextArg(txt, vbCrLf, , False), " ") & vbCrLf
                    End If
                Else
                    If Left(LCase(txt), 10) = "attribute " Then
                        head = head & RTrimStrip(RemoveNextArg(txt, vbCrLf, , False), " ") & vbCrLf
                    End If
                End If

                
                Exit Do
            Else
                FindNextHeader = FindNextHeader & txt
                txt = ""
            End If
        Else
            FindNextHeader = FindNextHeader & txt
            txt = ""
        End If
    Loop

End Function

Private Function FindLineStart(ByVal txt As String, ByVal pos As Long) As Long
    'accepts txt and pos with in txt where possible delcarative statements
    'are found, and then traces to the beginning of the line, returning pos
    'adbiding outside quotes and lines ending wtih underscares carrying over
    
    Do
        FindLineStart = pos
        pos = InStrRev(txt, vbCrLf, pos)
        If pos - 1 > 1 And InStrRev(txt, """", FindLineStart) < pos Then
            FindLineStart = pos
            If Mid(txt, pos - 1, 3) = "_" & vbCrLf Then
                pos = InStrRev(txt, vbCrLf, pos - 1)
                If InStrRev(txt, """", FindLineStart) < pos Then
                    FindLineStart = 0
                End If
            End If
        ElseIf pos = 1 Then
            FindLineStart = 1
        Else
            FindLineStart = InStrRev(txt, """", FindLineStart) + 1
        End If
    Loop While FindLineStart = 0
End Function

Private Function IsPropertyHeader(ByVal head As String) As Boolean
    IsPropertyHeader = Left(LCase(head), 8) = "property" Or InStr(head, " property ") > 0
End Function

Private Function GetFullLine(ByRef txt As String) As String

    'sew up any end of line carry overs
    'to find the end of the declarative
    Do While ((InStr(txt, "_" & vbCrLf) > 0) And (InStr(txt, "_" & vbCrLf) < InStr(txt, vbCrLf))) And (Not _
        ((InStr(txt, "' _" & vbCrLf) = (InStr(txt, "_" & vbCrLf) - 2)) And (InStr(txt, "_" & vbCrLf) > 0)))
        txt = NextArg(txt, "_" & vbCrLf, , False) & RemoveArg(txt, "_" & vbCrLf, , False)
    Loop

    
    GetFullLine = RemoveNextArg(txt, vbCrLf, , False) & vbCrLf
            
End Function

Private Function FindLineEnd(ByVal txt As String, Optional ByVal PropertyOnly As Boolean = False) As Long

    FindLineEnd = Len(txt)
    'check the header to the full line
    If ValidHeader(GetFullLine(txt), PropertyOnly) Then
        'return it's size
        FindLineEnd = FindLineEnd - Len(txt)
    Else
        FindLineEnd = 0
    End If
            
End Function

Private Function FindNextLine(ByVal txt As String, Optional ByVal PropertyName As String = "") As Long
    'searches txt for description supported line definitions
    'abiding outside of quotes, and returns a possible pos of
    
    Dim pos1 As Long
    Dim pos2 As Long
    Dim pos3 As Long
    Dim pos4 As Long
    Dim pos5 As Long

    Do
        If PropertyName = "" Then
            pos1 = InStr(pos5 + 1, LCase(txt), "property ")
            pos2 = InStr(pos5 + 1, LCase(txt), "event ")
            pos3 = InStr(pos5 + 1, LCase(txt), "function ")
            pos4 = InStr(pos5 + 1, LCase(txt), "sub ")
        Else
            pos1 = InStr(pos5 + 1, LCase(txt), "property get " & LCase(PropertyName) & "(")
            pos2 = InStr(pos5 + 1, LCase(txt), "property let " & LCase(PropertyName) & "(")
            pos3 = InStr(pos5 + 1, LCase(txt), "property set " & LCase(PropertyName) & "(")
            pos4 = 0
        End If
        If (pos1 > 0) And (pos1 < pos2 Or pos2 = 0) And (pos1 < pos3 Or pos3 = 0) And (pos1 < pos4 Or pos4 = 0) Then
            pos5 = pos1
            FindNextLine = FindLineStart(txt, pos1)
        ElseIf (pos2 > 0) And (pos2 < pos1 Or pos1 = 0) And (pos2 < pos3 Or pos3 = 0) And (pos2 < pos4 Or pos4 = 0) Then
            pos5 = pos2
            FindNextLine = FindLineStart(txt, pos2)
        ElseIf (pos3 > 0) And (pos3 < pos1 Or pos1 = 0) And (pos3 < pos2 Or pos2 = 0) And (pos3 < pos4 Or pos4 = 0) Then
            pos5 = pos3
            FindNextLine = FindLineStart(txt, pos3)
        ElseIf (pos4 > 0) And (pos4 < pos1 Or pos1 = 0) And (pos4 < pos3 Or pos3 = 0) And (pos4 < pos2 Or pos2 = 0) Then
            pos5 = pos4
            FindNextLine = FindLineStart(txt, pos4)
        Else
            FindNextLine = -1
        End If

    Loop Until FindNextLine <> 0

End Function
Private Function ValidHeader(ByVal head As String, Optional ByVal PropertyOnly As Boolean = False) As Boolean
    'accepts head information in a single line format and examines
    'for validity returning true or false whehter or not it's valid
    ValidHeader = True
    Do
        Select Case LCase(RemoveNextArg(head, " "))
            Case "public", "private", "global", "friend", "static", "withevents"
            'Case "dim", "const", "declare", "event", "type", "enum"
            Case "property"
                Exit Do
            Case "event", "function", "sub"
                If Not PropertyOnly Then Exit Do
            Case Else
                ValidHeader = False
                Exit Do
        End Select
    Loop While head <> ""
End Function


Private Function GetDeclareLine(ByVal head As String, ByVal rawform As Boolean) As String
    'gets only the declare line portion of a valid header

    If rawform Then
        'return it's size
        GetDeclareLine = head
        GetFullLine GetDeclareLine
        GetDeclareLine = Left(head, (Len(head) - Len(GetDeclareLine)))
    Else
        GetDeclareLine = GetFullLine(head)
    End If

    If InStr(GetDeclareLine, "' _") > 0 Then
        GetDeclareLine = RTrimStrip(NextArg(GetDeclareLine, "' _"), " ")
    Else
        GetDeclareLine = NextArg(GetDeclareLine, vbCrLf)
    End If

End Function
Private Function GetDescription(ByVal head As String, Optional ByVal From As HeaderInfo = 0) As String
    'gets the description portion of a valid header, with the option of returning the
    'description with in the comment, or the description with in the hidden attribute
    
    If From = Declared Or From = Commented Then
        If InStr(head, " ' _" & vbCrLf) > 0 Then
            GetDescription = NextArg(RemoveArg(head, " ' _" & vbCrLf), vbCrLf)
        End If
    Else
        If InStr(LCase(head), vbCrLf & "attribute ") > 0 Then
            GetDescription = RemoveQuotedArg(head, ".VB_Description = """, """" & vbCrLf, , vbTextCompare)
        End If
    End If
    GetDescription = Replace(Replace(GetDescription, vbCrLf, vbLf), vbLf, "")
End Function

Private Function GetUserDefined(ByVal head As String, Optional ByVal From As HeaderInfo = 0) As String
    'gets the "user defined name" portion of a valid header, with the option of returning the
    '"user defined name" from the declaration, or the "user defined name" from the hidden attribute
    If InStr(head, "' _" & vbCrLf) > 0 Then
        head = GetFullLine(NextArg(head, "' _" & vbCrLf, , True)) & "' _" & vbCrLf & RemoveArg(head, "' _" & vbCrLf, , True)
    ElseIf InStr(head, vbCrLf & "attribute ") > 0 Then
        head = GetFullLine(NextArg(head, vbCrLf & "attribute ", , True)) & vbCrLf & "attribute " & RemoveArg(head, vbCrLf & "attribute ", , True)
    Else
        head = GetFullLine(head)
    End If
    If From = Declared Or From = Commented Then
        
        Do While GetUserDefined = "" And (head <> "")
            Select Case LCase(NextArg(head, " "))
                Case "public", "private", "global", "friend", "static", "withevents"
                Case "dim", "const", "declare", "event"
                    GetUserDefined = NextArg(NextArg(NextArg(RemoveArg(head, " "), " "), "("), ",")
                Case "type", "enum"
                    GetUserDefined = NextArg(NextArg(RemoveArg(head, " "), " "), "(")
                Case "property"
                    GetUserDefined = NextArg(NextArg(RemoveArg(RemoveArg(head, " "), " "), " "), "(")
                Case "function", "sub"
                    GetUserDefined = NextArg(NextArg(RemoveArg(head, " "), " "), "(")
            End Select
            RemoveNextArg head, " "
        Loop
    Else
        If InStr(LCase(head), vbCrLf & "attribute ") > 0 Then
            GetUserDefined = RemoveQuotedArg(head, vbCrLf & "Attribute ", ".VB_Description = """, , vbTextCompare)
        End If
    End If
End Function

