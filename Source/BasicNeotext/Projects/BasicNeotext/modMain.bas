Attribute VB_Name = "modMain"

#Const modMain = -1
Option Explicit

Public Enum SignTypes
    SignOnly = 0
    SignAndStamp = 1
    StampOnly = 2
End Enum

Public Projs As New Project
Public Execs As New VBA.Collection
Public Param As New VBA.Collection
Public Paths As New VBA.Collection
Public Hooks As New VBA.Collection
Public RegEd As New Registry

Public Sub Main()

    InitialSetup
    
    If App.StartMode = vbSModeStandalone Then

        ParseCommand

        If Execs.Count = 0 Then
            RunProcessEx Paths("VisBasic"), "", , False
        Else
            Dim func As String
            Dim switch As Variant
            Set Projs = New Project
            
            For Each switch In Execs
                func = RemoveNextArg(switch, " ")
                Do Until (switch = "")
                    
                    Projs.Populate Paths(RemoveNextArg(switch, " "))

                    Select Case func
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

Sub ShowDriveList()
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d In dc
        s = s & d.DriveLetter & " - "
        If d.DriveType = 3 Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
        s = s & n & vbCrLf
    Next
    MsgBox s
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
    Dim pId As Long
    If Trim(path) <> "" Then
       ' chkRegSet = GetVBSettings()
        'If GetBNSettings <> chkRegSet Then SetVBSettings
        pId = Shell(Trim(Trim(path) & " " & Trim(Params)), IIf(Hide, vbHide, vbNormalFocus))
        If (pId > 0) Then
            chkElapse = Now
            Dim latency As Single
            
            Do While IsProccessIDRunning(pId) And Wait
                DoTasks
                DoEvents
                modCommon.Sleep 1
                DoLoop
                If latency = 0 Or (Timer - latency) > 1 Then
                    ItterateDialogs
                    latency = Timer
                End If
                    
                If ProcessRunning("VB6.EXE") = 0 Then KillApp "VBN.EXE"
            Loop
        End If
    End If
End Sub


Private Static Sub DoLoop()
    Static Elapse As Single
    Static latency As Single
    Static lastlat As Single
    Static multity As Long
    If Elapse <> 0 Then
        Elapse = Timer - Elapse
        If Elapse > latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then modCommon.Sleep lastlat
            End Select
            lastlat = Elapse - latency
            multity = multity + 16
        ElseIf Elapse < latency Then
            Select Case multity
                Case 0, 4, 64, 1024
                    DoEvents
                Case 1, 8, 128, 256
                    DoTasks
                Case 2, 16, 32, 512
                    If lastlat < 1000 Then modCommon.Sleep lastlat
            End Select
            lastlat = latency - Elapse
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
        latency = Elapse
    End If
    Elapse = Timer
End Sub

