Attribute VB_Name = "modMain"

#Const modMain = -1
Option Explicit

Public Enum SignTypes
    SignOnly = 0
    SignAndStamp = 1
    StampOnly = 2
End Enum

Public VBPID As Long

Public Projs As New Project
Public Execs As New VBA.Collection
Public Param As New VBA.Collection
Public Paths As New VBA.Collection
Public Hooks As New VBA.Collection
Public RegEd As New Registry

Public QuitCall As Boolean
Public QuitFail As Single

Public MainLoopElapse As Single

Public Sub Main()

    InitialSetup
    
    If App.StartMode = vbSModeStandalone Then

        ParseCommand

        If Execs.count = 0 Then
            RunProcessEx Paths("VisBasic"), "", , True
        Else
            Dim func As String
            Dim switch As Variant
            Set Projs = New Project
            
            For Each switch In Execs
                func = RemoveNextArg(switch, " ")
                Do Until (switch = "")
                    
                    Projs.Populate Paths(RemoveNextArg(switch, " "))

                    Select Case func
                        Case "copy"
                            DoCleanCopy
                        Case "install"
                            InstallSetup
                        Case "uninstall"
                            UninstallSetup
                        Case "runexit", "re"
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("/runexit """ & Projs.Location & """ " & Projs.CondComp & "  " & Projs.CmdLine), False, True
    
                        Case "run", "r"
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("/run """ & Projs.Location & """ " & Projs.CondComp & "  " & Projs.CmdLine), False, True
                        Case "make", "m"
                            If (Projs.Location <> "") Then
                                SetWorkingDir Projs
                                If (GetFileExt(Projs.Location, True, True) = "nsi") Then
                                    RunProcessEx Paths("MakeNSIS"), Trim("""" & Projs.Location & """ " & Projs.CondComp), , True
                                Else
                                    RunProcessEx Paths("VisBasic"), Trim("/make """ & Projs.Location & """ " & Projs.CondComp), , True
                                End If
                            End If
                        Case "signmake", "makesign", "sm", "ms"
                            
                            If (Projs.Location <> "") Then
                                SetWorkingDir Projs
                                If (GetFileExt(Projs.Location, True, True) = "nsi") Then
                                    RunProcessEx Paths("MakeNSIS"), Trim("""" & Projs.Location & """ " & Projs.CondComp), , True
                                Else
                                    RunProcessEx Paths("VisBasic"), Trim("/make """ & Projs.Location & """ " & Projs.CondComp), , True
                                End If
                            End If
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, SignAndStamp
                        Case "sign", "signonly", "s", "so"
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, SignOnly
                        Case "timestamp", "t", "to", "timeonly"
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, StampOnly
                        Case "?", "/?", "-?", "--?", "help", "/help", "-help", "--help"
                            frmHelp.Show
                        Case "d"
                        Case "open", "o"
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("""" & Projs.Location & """"), , True
                        Case Else

                            SetWorkingDir Projs
                            If func <> "" Then RunProcessEx Paths("VisBasic"), Command, , True

                    End Select
                    Projs.Cleanup
                Loop
            Next
            Set Projs = Nothing
        End If

    End If
End Sub

Private Sub SetWorkingDir(ByRef Projs As Project)
    If PathExists(GetFilePath(Projs.Compiled), False) Then ChDir GetFilePath(Projs.Compiled)
End Sub

Public Sub InitialSetup()
    Dim Path1 As String
    Dim Path2 As String
    Dim Path3 As String
    Path1 = GetSetting("BasicNeotext", "Options", "VisBasic")
    Path2 = GetSetting("BasicNeotext", "Options", "MakeNSIS")
    Path3 = GetSetting("BasicNeotext", "Options", "SignTool")
    
    If (Path1 = "") Or (Path2 = "") Or (Path3 = "") Then
        Dim Path4 As String
        
        If (PathExists(Left(AppPath, 2) & "\Program Files", False)) Then
            Path4 = Path4 & SearchPath("vb6.exe" & vbCrLf & "makensis.exe" & vbCrLf & "signtool.exe", , Left(AppPath, 2) & "\Program Files", MatchFlags.ExactMatch, , vbDirectory Or vbNormal)
        End If
        
        If (PathExists(Left(AppPath, 2) & "\Program Files (x86)", False)) Then
            Path4 = Path4 & SearchPath("vb6.exe" & vbCrLf & "makensis.exe" & vbCrLf & "signtool.exe", , Left(AppPath, 2) & "\Program Files (x86)", MatchFlags.ExactMatch, , vbDirectory Or vbNormal)
        End If
    
        Do Until (Path4 = "")
            If (InStr(1, LCase(NextArg(Path4, vbCrLf)), "vb6.exe", vbTextCompare) > 0) Then
                Path1 = NextArg(Path4, vbCrLf)
                SaveSetting "BasicNeotext", "Options", "VisBasic", Path1
            ElseIf (InStr(1, LCase(NextArg(Path4, vbCrLf)), "makensis.exe", vbTextCompare) > 0) Then
                Path2 = NextArg(Path4, vbCrLf)
            ElseIf (InStr(1, LCase(NextArg(Path4, vbCrLf)), "signtool.exe", vbTextCompare) > 0) Then
                Path3 = NextArg(Path4, vbCrLf)
            End If
            RemoveNextArg Path4, vbCrLf
        Loop

        SaveSetting "BasicNeotext", "Options", "VisBasic", IIf(Path1 = "", "(not found)", Path1)
        SaveSetting "BasicNeotext", "Options", "MakeNSIS", IIf(Path2 = "", "(not found)", Path2)
        SaveSetting "BasicNeotext", "Options", "SignTool", IIf(Path3 = "", "(not found)", Path3)

    End If
        
    Paths.Add IIf(Path1 = "(not found)", "", Path1), "VisBasic"
    Paths.Add IIf(Path2 = "(not found)", "", Path2), "MakeNSIS"
    Paths.Add IIf(Path3 = "(not found)", "", Path3), "SignTool"
End Sub

Private Sub InstallSetup()
    If (Paths("VisBasic") <> "") Then
        Dim entry As String
        entry = FolderQuoteName83(Replace(AppEXE(False, False), ".dll", ".exe"))
        If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "") <> entry & " /make ""%1"" " Then
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "", entry & " /make ""%1"" "
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\open\command", "", entry & " ""%1"""
            If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run\command", "", entry & " /run ""%1"" "
            ElseIf RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", entry & " /run ""%1"" "
            End If
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Make\command", "", entry & " /make ""%1"" "
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\open\command", "", entry & " ""%1"""
            If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", entry & " /run ""%1"" "
            ElseIf RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run ProjectGroup\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run ProjectGroup\command", "", entry & " /run ""%1"" "
            End If
        End If
    End If
End Sub

Private Sub UninstallSetup()
    If (Paths("VisBasic") <> "") Then
        Dim entry As String
        entry = FolderQuoteName83(Paths("VisBasic"))
        If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "") <> entry & " /make ""%1"" " Then
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Make\command", "", entry & " /make ""%1"" "
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\open\command", "", entry & " ""%1"""
            If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run\command", "", entry & " /run ""%1"" "
            ElseIf RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Run Project\command", "", entry & " /run ""%1"" "
            End If
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Make\command", "", entry & " /make ""%1"" "
            RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\open\command", "", entry & " ""%1"""
            If RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run\command", "", entry & " /run ""%1"" "
            ElseIf RegEd.GetValue(HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run ProjectGroup\command", "", "") <> "" Then
                RegEd.SetValue HKEY_CLASSES_ROOT, "VisualBasic.ProjectGroup\shell\Run ProjectGroup\command", "", entry & " /run ""%1"" "
            End If
        End If

        If PathExists(AppPath & "REG.BAK", True) Then
            Kill AppPath & "REG.BAK"
        End If
        
        RegEd.ExpellKey HKEY_CURRENT_USER, "Software\VB and VBA Program Settings\BasicNeotext"

    End If
End Sub

Public Sub SignTool(ByVal BinaryFile As String, Optional ByVal SignType As SignTypes)

    Dim cert As String
    Dim pwrd As String
    Dim turl As String
    Dim durl As String
    Dim auxc As String
    auxc = GetSetting("BasicNeotext", "Options", "Auxiliary", "")

    cert = GetSetting("BasicNeotext", "Options", "Certificate", "")
    pwrd = GetSetting("BasicNeotext", "Options", "Password", "")
    If (pwrd <> "") Then pwrd = DecryptString(pwrd, GetMachineName & "\\" & GetUserLoginName, True)
    turl = GetSetting("BasicNeotext", "Options", "TStampURL", "")
    durl = GetSetting("BasicNeotext", "Options", "DescURL", "")

    If (Not (cert = "")) And (Not (Paths("SignTool") = "")) And (BinaryFile <> "") Then
        If PathExists(cert, True) And PathExists(BinaryFile, True) And PathExists(Paths("SignTool"), True) Then
            Dim proceed As Boolean
            If (GetSetting("BasicNeotext", "Options", "RestrictOnly", 0) = 1) Then
                If (InStr(LCase(GetSetting("BasicNeotext", "Options", "RestrictList", "")), LCase(BinaryFile)) > 0) Then
                    proceed = True
                End If
            Else
                proceed = True
            End If
            If proceed Then
                If SignType = StampOnly Then
                    RunProcessEx Paths("SignTool"), "timestamp " & IIf(turl <> "", "/t " & turl & " ", "") & """" & BinaryFile & """", True, True
                ElseIf SignType = SignAndStamp Then
                    RunProcessEx Paths("SignTool"), "sign " & IIf(auxc <> "", auxc, "") & " /f """ & cert & """ " & IIf(pwrd <> "", "/p " & pwrd & " ", "") & IIf(turl <> "", "/t " & turl & " ", "") & IIf(durl <> "", "/du " & durl & " ", "") & """" & BinaryFile & """", True, True
                ElseIf SignType = SignOnly Then
                    RunProcessEx Paths("SignTool"), "sign " & IIf(auxc <> "", auxc, "") & " /f """ & cert & """ " & IIf(pwrd <> "", "/p " & pwrd & " ", "") & IIf(durl <> "", "/du " & durl & " ", "") & """" & BinaryFile & """", True, True
                End If
            End If
        End If
    End If
            
End Sub

Public Sub SetVBSettings()

    Dim regKey As String
    regKey = GetBNSettings

    SaveSettingByte HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Dock", StrConv(RemoveNextArg(regKey, vbCrLf), vbFromUnicode)
    SaveSettingByte HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Tool", StrConv(RemoveNextArg(regKey, vbCrLf), vbFromUnicode)
    SaveSettingByte HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "UI", StrConv(RemoveNextArg(regKey, vbCrLf), vbFromUnicode)

End Sub

Public Sub SetBNSettings()
    
    WriteFile GetFilePath(AppEXE(False)) & "\REG.BAK", GetVBSettings
    
End Sub

Public Function GetVBSettings() As String
    GetVBSettings = StrConv(GetSettingByte(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Dock"), vbUnicode) & vbCrLf & _
                    StrConv(GetSettingByte(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "Tool"), vbUnicode) & vbCrLf & _
                    StrConv(GetSettingByte(HKEY_CURRENT_USER, "Software\Microsoft\Visual Basic\6.0", "UI"), vbUnicode)
End Function

Public Function GetBNSettings() As String
    
    If PathExists(AppPath & "REG.BAK", True) Then
        GetBNSettings = ReadFile(AppPath & "REG.BAK")
    Else
        GetBNSettings = GetVBSettings
    End If

End Function

Public Sub RunProcessEx(ByVal path As String, ByVal Params As String, Optional ByVal Hide As Boolean = False, Optional ByVal Wait As Boolean = False)

    If Trim(path) <> "" Then
        Dim LastRun As String
        Dim LoopLatency As Single
        
        LastRun = Trim(Trim(path) & " " & Trim(Params))
        
        VBPID = Shell(LastRun, IIf(Hide, vbHide, vbNormalFocus))

        If (VBPID > 0) Then

            Do While ((IsProccessIDRunning(VBPID) Or QuitCall) And Wait) And (Not QuitFail = -1)
                LoopLatency = Timer
                DoLoop

                If QuitCall Then
                    If Not IsProccessIDRunning(VBPID) Then
                        VBPID = Shell(LastRun, vbNormalFocus)
                        QuitCall = False
                    End If
                End If

                If (Not QuitCall) Then
                    If (ProcessRunning("VB6.EXE") = 0) Then
                        If IsProccessIDRunning(VBPID) Then
                            QuitFail = -1
                            KillApp "VBN.EXE"
                        End If
                    End If
                End If
                MainLoopElapse = LoopLatency - Timer
                
            Loop
            

        End If
    End If
End Sub

Public Function GatherFileList(ByVal Location As String, Optional ByRef ProjList As String = "") As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim F As Object
    Dim fi As Object
    Dim sf As Object
    If PathExists(Location, True) Then
        Set F = fso.getfolder(GetFilePath(Location))
    Else
        Set F = fso.getfolder(Location)
    End If
    For Each sf In F.subfolders
        GatherFileList = GatherFileList & GatherFileList(sf.path, ProjList)
    Next
    For Each fi In F.Files
        Select Case GetFileExt(fi.Name, True, True)
            Case "bas", "cls"
                GatherFileList = GatherFileList & fi.path & vbCrLf
            Case "frm"
                GatherFileList = GatherFileList & fi.path & vbCrLf
                GatherFileList = GatherFileList & GetFilePath(fi.path) & "\" & GetFileTitle(fi.path) & ".frx" & vbCrLf
            Case "ctl"
                GatherFileList = GatherFileList & fi.path & vbCrLf
                GatherFileList = GatherFileList & GetFilePath(fi.path) & "\" & GetFileTitle(fi.path) & ".ctx" & vbCrLf
            Case "dob"
                GatherFileList = GatherFileList & fi.path & vbCrLf
                GatherFileList = GatherFileList & GetFilePath(fi.path) & "\" & GetFileTitle(fi.path) & ".vbd" & vbCrLf
            Case "dsr"
                GatherFileList = GatherFileList & fi.path & vbCrLf
                GatherFileList = GatherFileList & GetFilePath(fi.path) & "\" & GetFileTitle(fi.path) & ".dca" & vbCrLf
            Case "pag"
                GatherFileList = GatherFileList & fi.path & vbCrLf
                GatherFileList = GatherFileList & GetFilePath(fi.path) & "\" & GetFileTitle(fi.path) & ".pgx" & vbCrLf
            Case "vbp"
                ProjList = ProjList & fi.path & vbCrLf
        End Select
    Next
    Set fso = Nothing
End Function
Private Sub DoCleanUp(ByVal ProjList As String, ByRef FileListing As String)

    Dim Proj As String
    Dim p As Project
    Dim i As Project
    
    Do While ProjList <> ""
        Proj = RemoveNextArg(ProjList, vbCrLf)
        Set p = New Project
        p.Populate Proj

        For Each i In p.Includes
            Select Case GetFileExt(i.Location, True, True)
                Case "bas", "cls", "ctl", "ctx", "frm", "frx", "dsr", "dca", "dob", "vbd", "pag", "pgx"
                    FileListing = Replace(LCase(FileListing), LCase(i.Location & vbCrLf), "")
            End Select
        Next
        
    Loop
    
End Sub
Private Function ExcludePath(ByVal path As String, ByVal cancelExp As String) As Boolean

    Dim line As String
    
    Do While cancelExp <> ""
        line = RemoveNextArg(cancelExp, vbCrLf)
        If Trim(line) <> "" Then
            If InStr(LCase(path), LCase(line)) > 0 Then
                ExcludePath = True
                Exit Function
            End If
        End If
    Loop
    
End Function
Private Sub DoCopyProject(ByVal Source As String, ByVal Dest As String, ByVal ExcludeList As String, ByRef DeleteList As String, Optional ByVal ExcludeExperssions As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim F As Object
    Dim fi As Object
    Dim sf As Object

    If Not ExcludePath(Source, ExcludeExperssions) Then
    
        Set F = fso.getfolder(Source)
        
        Dim cancelExp As String
        If PathExists(Source & "\Exclude.txt", True) Then
            cancelExp = ReadFile(Source & "\Exclude.txt")
        End If

        MakeFolder Dest

        For Each sf In F.subfolders
            DoCopyProject Source & "\" & sf.Name, Dest & "\" & sf.Name, ExcludeList, DeleteList, ExcludeExperssions & vbCrLf & cancelExp
        Next
    
        For Each fi In F.Files
            If Not ExcludePath(Dest & "\" & fi.Name, ExcludeExperssions & vbCrLf & cancelExp & vbCrLf & "Exclude.txt") Then
                
                Select Case GetFileExt(fi.Name, True, True)
                    Case "bas", "cls", "frm", "ctl", "dob", "dsr"
                        
                        If InStr(LCase(ExcludeList), Source & "\" & fi.Name & vbCrLf) = 0 Then
                            DeleteList = Replace(LCase(DeleteList), LCase(Dest & "\" & fi.Name) & vbCrLf, "")
                            DoFileCopy fi.path, Dest & "\" & fi.Name
                        Else
                            DeleteList = DeleteList & Dest & "\" & fi.Name & vbCrLf
                        End If
                    Case Else
                        DeleteList = Replace(LCase(DeleteList), LCase(Dest & "\" & fi.Name & vbCrLf), "")
                        DoFileCopy fi.path, Dest & "\" & fi.Name
                End Select
            Else
                If PathExists(Dest & "\" & fi.Name, True) Then
                    SetAttr Dest & "\" & fi.Name, vbNormal
                    Kill Dest & "\" & fi.Name
                End If
            End If
        Next
        
    Else
        If PathExists(Dest, False) Then RemovePath Dest
    End If
    Set fso = Nothing
End Sub

Private Sub DoFileCopy(ByVal Source As String, ByVal Dest As String)
    If PathExists(Dest, True) Then
        If FileDateTime(Source) <> FileDateTime(Dest) Or _
            FileLen(Source) <> FileLen(Dest) Then
            Kill Dest
            FileCopy Source, Dest
        End If
    Else
        FileCopy Source, Dest
    End If
    'DoLoop
End Sub

Private Function DoCleanCopy() As Boolean
    
    Dim ProjList As String
    Dim FileList As String
    Dim UnusedList As String
    
    DoCleanCopy = True
    On Error Resume Next
    ProjList = Paths("copy1")
    FileList = Paths("copy2")
    If Err.Number = 0 Then
        
        If PathExists(Paths("copy1"), False) Then
            If PathExists(GetFilePath(Paths("copy2")), False) Then
                ProjList = ""
                FileList = ""
                
                
                FileList = GatherFileList(Paths("copy1"), ProjList)
                'project files are in projlist, and their files are in filelist
                UnusedList = FileList
                
                DoCleanUp ProjList, UnusedList
                'filelist is now only files not in any project in projlist
                
                'projlist now contains all source files in projects at the dest
                ProjList = ProjList & vbCrLf & FileList
                ProjList = Replace(ProjList, Paths("copy1"), Paths("copy2"))
                
                'FileList is passed to be excluded in copying the projlist files
                DoCopyProject Paths("copy1"), Paths("copy2"), UnusedList, ProjList

                'any file in the paths not copied is now in projlist

                'remove those in projlist as they are no longer in projects
                Do While ProjList <> ""
                    On Error Resume Next
                    UnusedList = RemoveNextArg(ProjList, vbCrLf)
                    SetAttr UnusedList, vbNormal
                    Kill UnusedList
                Loop
                
            Else
                DoCleanCopy = False
            End If
        Else
            DoCleanCopy = False
        End If
    Else
        Err.Clear
        DoCleanCopy = False
    End If
    If Not DoCleanCopy Then
        MsgBox "The /copy switch requires a source folder, followed by a destination folder." & vbCrLf & _
                "All projects under the source path will be copied to the destination folder" & vbCrLf & _
                "excluding any source file not with in projects, including everything else" & vbCrLf & _
                "and the destination folder may be cleaned of files not in the source path.", vbInformation
    End If
End Function

Public Static Sub DoLoop()
    DoTasks
    DoEvents
    modCommon.Sleep 1
                    
    Static elapse As Single
    Static latency As Single
    Static lastlat As Single
    Static multity As Long
    If elapse <> 0 Then
        elapse = Timer - elapse
        If elapse > latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then modCommon.Sleep lastlat
            End Select
            lastlat = elapse - latency
            multity = multity + 16
        ElseIf elapse < latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then modCommon.Sleep lastlat
            End Select
            lastlat = latency - elapse
            multity = multity + 4
        ElseIf lastlat <> 0 Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then modCommon.Sleep lastlat
            End Select
            If lastlat > 0 Then
                If Not multity = 0 Then
                    multity = multity \ 2
                Else
                    multity = multity + 2
                End If
            ElseIf lastlat < 0 Then
                If Not multity = 1024 Then
                    multity = multity * 2
                Else
                    multity = multity - 2
                End If
            End If
        ElseIf lastlat = 0 Then
            lastlat = 1
        End If
        latency = elapse
    End If
    elapse = Timer
End Sub

