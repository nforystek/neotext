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

Global MainLoopElapse As Single
Global TimerLoopElapse As Single

'Public CPUCurrent As Long
'Public CPUDirect As Long
'Public CPUPrior As Long
'Public Const CPUCores As Long = 4
'Public Const CPUUSage As Long = 25
'
'Public TargetPitch As Single
'Public TargetWhole As Single
'
'Public QueryStop As Boolean
'Public QueryObject As Object

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
                    
                    Select Case func
                        Case "copy"
                            DoCleanCopy
                            switch = ""
                        Case "install"
                            InstallSetup
                            switch = ""
                        Case "uninstall"
                            UninstallSetup
                            switch = ""
                        Case "runexit", "re"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("/runexit """ & Projs.Location & """ " & Projs.CondComp & "  " & Projs.CmdLine), False, True
                            Projs.Cleanup
                        Case "run", "r"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("/run """ & Projs.Location & """ " & Projs.CondComp & "  " & Projs.CmdLine), False, True
                            Projs.Cleanup
                        Case "make", "m"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            If (Projs.Location <> "") Then
                                SetWorkingDir Projs
                                If (GetFileExt(Projs.Location, True, True) = "nsi") Then
                                    RunProcessEx Paths("MakeNSIS"), Trim("""" & Projs.Location & """ " & Projs.CondComp), , True
                                Else
                                    RunProcessEx Paths("VisBasic"), Trim("/make """ & Projs.Location & """ " & Projs.CondComp), , True
                                End If
                            End If
                            Projs.Cleanup
                        Case "signmake", "makesign", "sm", "ms"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            If (Projs.Location <> "") Then
                                SetWorkingDir Projs
                                If (GetFileExt(Projs.Location, True, True) = "nsi") Then
                                    RunProcessEx Paths("MakeNSIS"), Trim("""" & Projs.Location & """ " & Projs.CondComp), , True
                                Else
                                    RunProcessEx Paths("VisBasic"), Trim("/make """ & Projs.Location & """ " & Projs.CondComp), , True
                                End If
                            End If
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, SignAndStamp
                            Projs.Cleanup
                        Case "sign", "signonly", "s", "so"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, SignOnly
                            Projs.Cleanup
                        Case "timestamp", "t", "to", "timeonly"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            If (Projs.Compiled <> "") Then SignTool Projs.Compiled, StampOnly
                            Projs.Cleanup
                        Case "?", "/?", "-?", "--?", "help", "/help", "-help", "--help"
                            frmHelp.Show
                            switch = ""
                        Case "d"
                            switch = ""
                        Case "open", "o"
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            SetWorkingDir Projs
                            If (Projs.Location <> "") Then RunProcessEx Paths("VisBasic"), Trim("""" & Projs.Location & """"), , True
                            Projs.Cleanup
                        Case Else
                            Projs.Populate Paths(RemoveNextArg(switch, " "))
                            SetWorkingDir Projs
                            If func <> "" Then RunProcessEx Paths("VisBasic"), Command, , True
                            Projs.Cleanup
                    End Select
                    
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
    
'        If IsWindows98 Then
'            Set QueryObject = New WinCPU
'        Else
'            Set QueryObject = New WinCPUNT
'        End If
'        QueryObject.Initialize
            
        Dim LastRun As String
        Dim LoopLatency As Single
        LoopLatency = Timer
        
        LastRun = """" & Trim(Trim(path) & """ " & Trim(Params))
        
        VBPID = Shell(LastRun, IIf(Hide, vbHide, vbNormalFocus))


        If (VBPID > 0) Then

            Do While ((IsProccessIDRunning(VBPID) Or QuitCall) And Wait) And (Not QuitFail = -1)
                
                MainLoopElapse = (Timer - LoopLatency) * 1000
                LoopLatency = Timer
                
                
                'SteadyService
                DoLoop
                
                If QuitCall Then
                    If (Not IsProccessIDRunning(VBPID)) Then
                        VBPID = Shell(LastRun, vbNormalFocus)
                        QuitCall = False
                    End If
                End If
                
                If (Not QuitCall) And InStr(LCase(LastRun), "vb6.exe") > 0 Then
                    If (ProcessRunning("VB6.EXE") = 0) Then
                        If IsProccessIDRunning(VBPID) Then
                            QuitFail = -1
                            KillApp "VBN.EXE"
                        End If
                    End If
                End If
            Loop

        End If
        
        'Set QueryObject = Nothing
        
    End If
End Sub

'Public Sub SteadyService()
'
'    TargetWhole = (100 / CPUCores)
'    TargetPitch = (TargetWhole * (CPUUSage / 100))
'
'    Dim lVal As Long, sVal As String
'    lVal = QueryObject.Query
'    If (lVal <> CPUPrior) And (lVal < 100) Then
'        sVal = Format$(lVal, "000")
'        lVal = CInt(sVal)
'        CPUDirect = CPUCurrent - lVal
'        CPUCurrent = lVal
'    End If
'    CPUPrior = lVal
'     'Debug.Print "CPUCurrent: " & CPUCurrent & ",  CPUDirect: " & CPUDirect & ", TargetWhole: " & TargetWhole & ",  TargetPitch: " & TargetPitch
'    If CPUDirect > TargetWhole Then TargetPitch = TargetPitch \ 2
'    If (CPUCurrent < TargetWhole) Then
'       If (TargetPitch > 0) Then Sleep TargetPitch
'    ElseIf (CPUCurrent > TargetWhole) Then
'        DoLoop
'    End If
'
'
'End Sub


Private Function DoCleanCopy() As Boolean

    DoCleanCopy = True
    On Error Resume Next
    Dim File As String
    File = Paths("copy1")
    File = Paths("copy2")
    If Err.Number = 0 Then

        If PathExists(Paths("copy1"), False) Then
            If PathExists(GetFilePath(Paths("copy2")), False) Then

                Dim ListAllSourceFiles As String
                ListAllSourceFiles = ApplyExclusionFiles(GatherFileList(Paths("copy1")))
                
                Dim ListAllDestFiles As String
                ListAllDestFiles = GatherFileList(Paths("copy2"))
                
                Dim ProjectSourceFileList As String
                
                ProjectSourceFileList = GatherProjectList(ListAllSourceFiles)

                Dim CodeSourceFileList As String
                CodeSourceFileList = GatherProjectCodeList(ProjectSourceFileList)
                
                Dim MasterCopyList As String
                Dim MasterDeleteList As String
                Dim IntermediateList As String

                MasterDeleteList = ApplyInclusionFilters(ListAllDestFiles, FilterListing(VisualBasicFilesFilter))
                
                MasterDeleteList = ApplyExclusionFilters(MasterDeleteList, FilterListing(ProjectSourceFileList & CodeSourceFileList))
                
                IntermediateList = ApplyExclusionFilters(ListAllDestFiles, FilterListing(VisualBasicFilesFilter))
                IntermediateList = ApplyExclusionFilters(IntermediateList, FilterListing(ListAllSourceFiles))
                
                MasterDeleteList = MasterDeleteList & IntermediateList
                
                MasterCopyList = ApplyExclusionFilters(ListAllSourceFiles, VisualBasicFilesFilter) & ProjectSourceFileList & CodeSourceFileList
                
                Dim Iterator As String
                Iterator = MasterDeleteList
                'delete dest files marked for delete
                Do While Iterator <> ""
                    File = RemoveNextArg(Iterator, vbCrLf)
                    If PathExists(File, True) Then 'ensure not folder
                        SetAttr File, vbNormal 'kill fails on hiddens
                        Kill File
                    End If
                Loop

                'copy files marked for copying
                Iterator = MasterCopyList
                Do While Iterator <> ""
                    File = RemoveNextArg(Iterator, vbCrLf)
                    If PathExists(File, True) Then 'ensure not folder
                        DoFileCopy Replace(File, Paths("copy2"), Paths("copy1"), , , vbTextCompare), _
                                Replace(File, Paths("copy1"), Paths("copy2"), , , vbTextCompare)
                    End If
                Loop

                'remove any empty dest folders
                Iterator = ListAllDestFiles
                Do While Iterator <> ""
                    File = RemoveNextArg(Iterator, vbCrLf)
                    If PathExists(File, False) Then 'ensure only folders
                        If Replace(Replace(SearchPath("*", True, File, FindAll), File & vbCrLf, ""), vbCrLf, "") = "" Then
                            RemovePath File
                        End If
                    End If
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

Public Function FilterListing(ByVal Listing As String) As String
    FilterListing = Replace(Replace(Replace(Replace(Replace(Listing, Paths("copy1"), Paths("copy2")), "[", "[[]"), "#", "[#]"), "(", "[(]"), ")", "[)]")
End Function
Public Function GatherFileList(ByVal Location As String) As String
    GatherFileList = SearchPath("*", True, Location, FindAll)
End Function

Public Function VisualBasicFilesFilter() As String
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.vbg" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.vbp" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.vbw" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.bas" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.cls" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.frm" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.frx" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.ctl" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.ctx" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.pag" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.pgx" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.dob" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.vbd" & vbCrLf
    
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.dsr" & vbCrLf
    VisualBasicFilesFilter = VisualBasicFilesFilter & "*.dca" & vbCrLf
End Function

Public Function ApplyExclusionFiles(ByVal FileListing As String) As String
    Dim line As String
    Dim exez As String
    Dim backedup As String
    Dim backedup2 As String
    Dim Exclusions As String
    backedup = FileListing
    Do Until FileListing = ""
        line = RemoveNextArg(FileListing, vbCrLf)
        If PathExists(line & "\exclude.txt", True) Then
            exez = "exclude.txt" & vbCrLf & ReadFile(line & "\exclude.txt")
            If Not Right(exez, 2) = vbCrLf Then exez = exez & vbCrLf
        End If
        If PathExists(line & "\.gitignore", True) Then
            exez = exez & ".gitignore" & vbCrLf & ReadFile(line & "\.gitignore")
        End If
        If PathExists(line, False) Then
            Do Until exez = ""
                Exclusions = Exclusions & line & "\*" & RemoveNextArg(exez, vbCrLf) & "*" & vbCrLf
            Loop
            Do While InStr(Exclusions, vbCrLf & vbCrLf) > 0
                Exclusions = Replace(Exclusions, vbCrLf & vbCrLf, vbCrLf)
            Loop
        End If
    Loop
    ApplyExclusionFiles = ApplyExclusionFilters(backedup, Exclusions)
End Function

Public Function ApplyExclusionFilters(ByVal FileListing As String, ByVal Exclusions As String) As String
    Dim line As String
    Dim exez As String
    Dim expr As String
    Dim fldr As String
    Dim test As String
    Dim File As String
    Do Until FileListing = ""
        line = RemoveNextArg(FileListing, vbCrLf)
        If line <> "" Then
            If InStr(Exclusions, line) = 0 Then
                exez = Exclusions
                Do Until exez = ""
                    expr = LCase(RemoveNextArg(exez, vbCrLf))
                    If (LCase(line) Like expr) Then
                        line = ""
                        Exit Do
                    End If
                    If line = "" Then Exit Do
                Loop
            Else
                line = ""
            End If
            If (line <> "") Then ApplyExclusionFilters = ApplyExclusionFilters & line & vbCrLf
        End If
    Loop
    Do While InStr(ApplyExclusionFilters, vbCrLf & vbCrLf) > 0
        ApplyExclusionFilters = Replace(ApplyExclusionFilters, vbCrLf & vbCrLf, vbCrLf)
    Loop
End Function


Public Function ApplyInclusionFilters(ByVal FileListing As String, ByVal Inclusions As String) As String
    Dim line As String
    Dim exez As String
    Dim expr As String
    Dim fldr As String
    Dim test As String
    Do Until FileListing = ""
        line = RemoveNextArg(FileListing, vbCrLf)
        If line <> "" Then
            If InStr(Inclusions, line) = 0 Then
                exez = Inclusions
                Do Until exez = ""
                    expr = LCase(RemoveNextArg(exez, vbCrLf))
                    If (LCase(line) Like expr) Then
                        ApplyInclusionFilters = ApplyInclusionFilters & line & vbCrLf
                        line = ""
                        Exit Do
                    End If
                    If line = "" Then Exit Do
                Loop
            Else
                ApplyInclusionFilters = ApplyInclusionFilters & line & vbCrLf
            End If
        End If
    Loop
    Do While InStr(ApplyInclusionFilters, vbCrLf & vbCrLf) > 0
        ApplyInclusionFilters = Replace(ApplyInclusionFilters, vbCrLf & vbCrLf, vbCrLf)
    Loop
End Function


Public Function GatherProjectList(ByVal FileListing As String) As String
    Dim line As String
    Do Until FileListing = ""
        line = RemoveNextArg(FileListing, vbCrLf)
        Select Case GetFileExt(line, True, True)
            Case "vbp", "vbg", "vbw"
                GatherProjectList = GatherProjectList & line & vbCrLf
        End Select
    Loop
End Function

Public Function GatherProjectCodeList(ByVal PrjectList As String) As String
    Dim proj As Project
    Dim inc As Project
    Dim line As String
    Do Until PrjectList = ""
        line = RemoveNextArg(PrjectList, vbCrLf)
        Select Case GetFileExt(line, True, True)
            Case "vbg", "vbp"
                Set proj = New Project
                proj.Populate line
                For Each inc In proj.Includes
                    Select Case GetFileExt(inc.Location, True, True)
                        Case "bas", "cls"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                            End If
                        Case "frm"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                                If PathExists(Left(inc.Location, Len(inc.Location) - 4) & ".frx", True) Then
                                    GatherProjectCodeList = GatherProjectCodeList & Left(inc.Location, Len(inc.Location) - 4) & ".frx" & vbCrLf
                                End If
                            End If
                        Case "ctl"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                                If PathExists(Left(inc.Location, Len(inc.Location) - 4) & ".ctx", True) Then
                                    GatherProjectCodeList = GatherProjectCodeList & Left(inc.Location, Len(inc.Location) - 4) & ".ctx" & vbCrLf
                                End If
                            End If
                        Case "pag"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                                If PathExists(Left(inc.Location, Len(inc.Location) - 4) & ".pgx", True) Then
                                    GatherProjectCodeList = GatherProjectCodeList & Left(inc.Location, Len(inc.Location) - 4) & ".pgx" & vbCrLf
                                End If
                            End If
                        Case "dob"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                                If PathExists(Left(inc.Location, Len(inc.Location) - 4) & ".vbd", True) Then
                                    GatherProjectCodeList = GatherProjectCodeList & Left(inc.Location, Len(inc.Location) - 4) & ".vbd" & vbCrLf
                                End If
                            End If
                        Case "dsr"
                            If InStr(GatherProjectCodeList, inc.Location & vbCrLf) = 0 Then
                                GatherProjectCodeList = GatherProjectCodeList & inc.Location & vbCrLf
                                If PathExists(Left(inc.Location, Len(inc.Location) - 4) & ".dca", True) Then
                                    GatherProjectCodeList = GatherProjectCodeList & Left(inc.Location, Len(inc.Location) - 4) & ".dca" & vbCrLf
                                End If
                            End If
                            
                    End Select
                Next
                Set proj = Nothing
        End Select
    Loop
End Function

Private Sub DoFileCopy(ByVal Source As String, ByVal Dest As String)
    
    If PathExists(Dest, True) Then
        If FileDateTime(Source) <> FileDateTime(Dest) Or _
            FileLen(Source) <> FileLen(Dest) Then
            SetAttr Dest, vbNormal
            Kill Dest
            FileCopy Source, Dest
        End If
    Else
        If Not PathExists(GetFilePath(Dest), False) Then
            MakeFolder GetFilePath(Dest)
        End If
        FileCopy Source, Dest
    End If
End Sub


Public Static Sub DoLoop()
    
    DoEvents
    modCommon.Sleep 1
    DoTasks

                    
'    Static elapse As Single
'    Static latency As Single
'    Static lastlat As Single
'    Static multity As Long
'    If elapse <> 0 Then
'        elapse = Timer - elapse
'        If elapse > latency Then
'            Select Case multity
'                Case 0, 4, 64, 1024
'                    DoEvents
'                    Debug.Print "DoEvents"
'                Case 1, 8, 128, 256
'                    DoTasks
'                    Debug.Print "DoTasks"
'                Case 2, 16, 32, 512
'                    If lastlat * 10000 < 1000 Then
'                        If lastlat * 10000 < 1 Then
'                            modCommon.Sleep 1
'                            Debug.Print "Sleep "; 1
'                        Else
'                            modCommon.Sleep lastlat * 10000
'                            Debug.Print "Sleep "; lastlat * 10000
'                        End If
'                    End If
'            End Select
'            lastlat = elapse - latency
'            multity = multity + 16
'        ElseIf elapse < latency Then
'            Select Case multity
'                Case 0, 4, 64, 1024
'                    DoEvents
'                    Debug.Print "DoEvents"
'                Case 1, 8, 128, 256
'                    DoTasks
'                    Debug.Print "DoTasks"
'                Case 2, 16, 32, 512
'                    If lastlat * 10000 < 1000 Then
'                        If lastlat * 10000 < 1 Then
'                            modCommon.Sleep 1
'                            Debug.Print "Sleep "; 1
'                        Else
'                            modCommon.Sleep lastlat * 10000
'                            Debug.Print "Sleep "; lastlat * 10000
'                        End If
'                    End If
'            End Select
'            lastlat = latency - elapse
'            multity = multity + 4
'        ElseIf lastlat <> 0 Then
'            Select Case multity
'                Case 0, 4, 64, 1024
'                    DoEvents
'                    Debug.Print "DoEvents"
'                Case 1, 8, 128, 256
'                    DoTasks
'                    Debug.Print "DoTasks"
'                Case 2, 16, 32, 512
'                    If lastlat * 10000 < 1000 Then
'                        If lastlat * 10000 < 1 Then
'                            modCommon.Sleep 1
'                            Debug.Print "Sleep "; 1
'                        Else
'                            modCommon.Sleep lastlat * 10000
'                            Debug.Print "Sleep "; lastlat * 10000
'                        End If
'                    End If
'            End Select
'            If lastlat > 0 Then
'                If Not multity = 0 Then
'                    multity = multity \ 2
'                Else
'                    multity = multity + 2
'                End If
'            ElseIf lastlat < 0 Then
'                If Not multity = 1024 Then
'                    multity = multity * 2
'                Else
'                    multity = multity - 2
'                End If
'            End If
'        ElseIf lastlat = 0 Then
'            lastlat = 1
'        End If
'        latency = elapse
'    End If
'    elapse = Timer
End Sub

