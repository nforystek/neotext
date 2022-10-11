Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
'TOP DOWN
'parse out all dll's from projects even api and generate nsi includes
'with the versioning and the shared references effort with other apps
'make two deploy includes per application one with system dll's one
'not and cab the compression of the files in prioritize defaults also
'use the reboot/cache for dll's

Private Const DLLTYPE = "DLL    $AlreadyInstalled REBOOT_PROTECTED"
Private Const REGDLLTYPE = "REGDLL $AlreadyInstalled REBOOT_PROTECTED"
Private Const REGEXETYPE = "REGEXE $AlreadyInstalled REBOOT_PROTECTED"
Private Const TLBTYPE = "TLB    $AlreadyInstalled REBOOT_PROTECTED"

Private fso As Scripting.FileSystemObject
Private FullManifest As String
Private Manifest As String
Private UnManifest As String
Private ManifestSys As String
Private UnManifestSys As String
Public Function GetFileEntry(ByVal File As String, ByVal LibType As String) As String
    Select Case GetFileExt(File)
        Case ".dll", ".ocx", ".oca"
            If GetFileTitle(File) = "maxlandlib" Then
                If Not LibType = DLLTYPE Then LibType = TLBTYPE
            Else
                If Not LibType = DLLTYPE Then LibType = REGDLLTYPE
            End If
        Case ".exe"
            If Not LibType = DLLTYPE Then LibType = REGEXETYPE
            File = ""
        Case ".tlb", ".olb"
            LibType = TLBTYPE
            Exit Function
            
    End Select
    Dim fvi As FILEVERINFO
    Dim verInfo As String
    verInfo = "0.0.0.0"
   
    If PathExists("C:\Development\Neotext\Common\Binary\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Common\Binary\" & GetFileName(File), fvi
        If Not LibType = DLLTYPE Then LibType = REGDLLTYPE
        If GetFileExt(File) <> ".vbp" Then GetFileEntry = _
            "!insertmacro InstallLib " & LibType & " ""${APPPATH}\Common\Binary\" & GetFileName(File) & """ ""$SYSDIR\" & GetFileName(File) & """ ""$SYSDIR""" & vbCrLf

    ElseIf PathExists("C:\Development\Neotext\Windows\ActiveX\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\ActiveX\" & GetFileName(File), fvi
        If Not LibType = DLLTYPE Then LibType = REGDLLTYPE
        If GetFileExt(File) <> ".vbp" Then GetFileEntry = _
            "!insertmacro InstallLib " & LibType & " ""${APPPATH}\Windows\ActiveX\" & GetFileName(File) & """ ""$SYSDIR\" & GetFileName(File) & """ ""$SYSDIR""" & vbCrLf
            
    ElseIf PathExists("C:\Development\Neotext\Windows\System\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\System\" & GetFileName(File), fvi
        LibType = DLLTYPE
        If GetFileExt(File) <> ".vbp" Then GetFileEntry = _
            "!insertmacro InstallLib " & LibType & " ""${APPPATH}\Windows\System\" & GetFileName(File) & """ ""$SYSDIR\" & GetFileName(File) & """ ""$SYSDIR""" & vbCrLf
        
    ElseIf PathExists("C:\Development\Neotext\Windows\Normal\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\Normal\" & GetFileName(File), fvi
        LibType = TLBTYPE
        If GetFileExt(File) <> ".vbp" Then GetFileEntry = _
            "!insertmacro InstallLib " & LibType & " ""${APPPATH}\Windows\Normal\" & GetFileName(File) & """ ""$SYSDIR\" & GetFileName(File) & """ ""$SYSDIR""" & vbCrLf

    ElseIf PathExists(SysPath & GetFileName(File), True) Then
        Debug.Print "C:\Development\Neotext\Windows\System\" & GetFileName(File)
        If Not PathExists("C:\Development\Neotext\Windows\System\" & GetFileName(File), True) Then
        
            'GetVersionInfo SysPath & GetFileName(File), fvi
            FileCopy SysPath & GetFileName(File), "C:\Development\Neotext\Windows\System\" & GetFileName(File)
            LibType = DLLTYPE
            If GetFileExt(File) <> ".vbp" Then GetFileEntry = _
            "!insertmacro InstallLib " & LibType & " ""${APPPATH}\Windows\System\" & GetFileName(File) & """ ""$SYSDIR\" & GetFileName(File) & """ ""$SYSDIR""" & vbCrLf

        End If

    End If
End Function
Public Function UnFileEntry(ByVal File As String) As String
    File = RemoveQuotedArg(File, """", """")
    Dim fvi As FILEVERINFO

    If PathExists("C:\Development\Neotext\Common\Binary\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Common\Binary\" & GetFileName(File), fvi
        UnFileEntry = "!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_NOTPROTECTED ""$SYSDIR\" & GetFileName(File) & """" & vbCrLf
    ElseIf PathExists("C:\Development\Neotext\Windows\ActiveX\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\ActiveX\" & GetFileName(File), fvi
        UnFileEntry = "!insertmacro UnInstallLib REGDLL SHARED NOREBOOT_PROTECTED ""$SYSDIR\" & GetFileName(File) & """" & vbCrLf
    ElseIf PathExists("C:\Development\Neotext\Windows\System\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\System\" & GetFileName(File), fvi
        UnFileEntry = "!insertmacro UnInstallLib DLL    SHARED REBOOT_PROTECTED ""$SYSDIR\" & GetFileName(File) & """" & vbCrLf
    ElseIf PathExists("C:\Development\Neotext\Windows\Normal\" & GetFileName(File), True) Then
        'GetVersionInfo "C:\Development\Neotext\Windows\Normal\" & GetFileName(File), fvi
        UnFileEntry = "!insertmacro UnInstallLib TLB    SHARED NOREBOOT_NOTPROTECTED ""$SYSDIR\" & GetFileName(File) & """" & vbCrLf
    ElseIf PathExists(SysPath & File, True) Then
        'GetVersionInfo SysPath & GetFileName(File), fvi
        UnFileEntry = "!insertmacro UnInstallLib DLL    SHARED NOREMOVE ""$SYSDIR\" & GetFileName(File) & """" & vbCrLf
    End If
End Function

Public Function CheckLibrary(ByVal Path As String, ByVal inVal As String, ByVal LibType As String) As Boolean
    
    If PathExists(MapFolder(GetFilePath(Path) & "..\Binary\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\Binary\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\Binary\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\Binary\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\Common\Binary\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    End If

    If PathExists(MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\..\Windows\System\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Windows\System\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\System\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\System\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\Windows\System\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    End If

    If PathExists(MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\..\Windows\ActiveX\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\Windows\ActiveX\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    End If


    If PathExists(MapFolder(GetFilePath(Path) & "..\..\..\Windows\Normal\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\Normal\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\..\Windows\Normal\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\..\Windows\Normal\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    ElseIf PathExists(MapFolder(GetFilePath(Path) & "..\..\Windows\Normal\", inVal), True) Then
        If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\Normal\", inVal), LibType))) = 0 Then
            Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path) & "..\..\Windows\Normal\", inVal), LibType)
            UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path) & "..\..\Windows\Normal\", inVal)), LibType))
        End If
        CheckLibrary = True
        Exit Function
    End If
    
    If InStr(LCase(Manifest), LCase(GetFileName(inVal))) = 0 Then
        If PathExists(MapFolder("C:\WINDOWS\SYSTEM32\", inVal), True) Then
            If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder("C:\WINDOWS\SYSTEM32\", inVal), LibType))) = 0 Then
                Manifest = Manifest & GetFileEntry(MapFolder("C:\WINDOWS\SYSTEM32\", inVal), LibType)
                UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder("C:\WINDOWS\SYSTEM32\", inVal)), LibType))
            End If
            CheckLibrary = True
            Exit Function
        ElseIf PathExists(MapFolder(GetFilePath(Path), inVal), True) Then
            If InStr(LCase(Manifest), LCase(GetFileEntry(MapFolder(GetFilePath(Path), inVal), LibType))) = 0 Then
                Manifest = Manifest & GetFileEntry(MapFolder(GetFilePath(Path), inVal), LibType)
                UnManifest = UnManifest & UnFileEntry(GetFileEntry(GetFileName(MapFolder(GetFilePath(Path), inVal)), LibType))
            End If
            CheckLibrary = True
            Exit Function
        End If
    End If
End Function

Public Function MapFolder(ByVal RootURL As String, ByVal vURL As String) As String
On Error GoTo exitthis

    'concatenates vURL to the RootURL properly by blind path specifications
    If InStr(RootURL, "\") > 0 Then
        RootURL = Replace(RootURL, "/", "\")
        vURL = Replace(vURL, "/", "\")
        If Left(vURL, 1) = "\" And Right(RootURL, 1) = "\" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "\" And Right(RootURL, 1) <> "\" Then
            vURL = RootURL & "\" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "\" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "\"
    ElseIf InStr(RootURL, "/") > 0 Then
        RootURL = Replace(RootURL, "\", "/")
        vURL = Replace(vURL, "\", "/")
        
        If Left(vURL, 1) = "/" And Right(RootURL, 1) = "/" Then
            vURL = RootURL & Mid(vURL, 2)
        ElseIf Left(vURL, 1) <> "/" And Right(RootURL, 1) <> "/" Then
            vURL = RootURL & "/" & vURL
        Else
            vURL = RootURL & vURL
        End If
        If Right(vURL, 1) = "/" Then vURL = Left(vURL, Len(vURL) - 1)
        If vURL = "" Then vURL = "/"

    End If
    vURL = Replace(vURL, "\*", "")
    
    Do While InStr(vURL, "..\") > 0
        If InStrRev(Left(vURL, InStr(vURL, "..\") - 5), "\") > 0 Then
            vURL = Left(vURL, InStrRev(Left(vURL, InStr(vURL, "..\") - 5), "\")) & Mid(vURL, InStr(vURL, "..\") + 3)
        End If
    Loop
    
    Do While InStr(vURL, "../") > 0
        If InStrRev(Left(vURL, InStr(vURL, "../") - 5), "/") > 0 Then
            vURL = Left(vURL, InStrRev(Left(vURL, InStr(vURL, "../") - 5), "/")) & Mid(vURL, InStr(vURL, "../") + 3)
        End If
    Loop
    
    MapFolder = vURL
    Exit Function
exitthis:
    Err.Clear
End Function

Public Function ParseSource(ByVal Path As String, ByVal Text As String) As String
    Dim backup As String
    backup = Text
    
    Dim inVal As String
    Do Until Text = ""
        RemoveNextArg Text, "Lib"
        If Text <> "" Then
        
            inVal = RemoveQuotedArg(Text, """", """")
            If inVal <> "" And InStr(LCase(inVal), ".dll") = 0 Then inVal = inVal & ".dll"
            CheckLibrary Path, inVal, DLLTYPE
        End If
    Loop
    
    Text = backup
    Do Until Text = ""
        RemoveNextArg Text, "CreateObject("
        If Text <> "" Then
        
            inVal = RemoveQuotedArg(Text, """", """")
            inVal = RemoveNextArg(inVal, ".")
            If inVal <> "" And InStr(LCase(inVal), ".dll") = 0 Then inVal = inVal & ".dll"
            If PathExists("C:\Development\Neotext\COmmon\Binary\" & inVal, True) Then
                CheckLibrary Path, inVal, REGDLLTYPE
            
            End If
        End If
    Loop
End Function

Public Function ParseProject(ByVal Path As String, ByVal Text As String) As String
    
    Dim inLine As String
    Dim inVar As String
    Dim inVal As String
    Do Until Text = ""
        inLine = RemoveNextArg(Text, vbCrLf)
        inVar = RemoveNextArg(inLine, "=")
        inVal = inLine
        Select Case LCase(inVar)
            Case "object", "reference"
                Do Until inVal = "" Or CheckLibrary(Path, inVal, REGDLLTYPE)
                    inVal = StrReverse(inVal)
                    inVar = RemoveNextArg(inVal, "#")
                    inVar = RemoveNextArg(inVar, ";")
                    inVar = StrReverse(inVar)
                    inVal = StrReverse(inVal)
                    CheckLibrary Path, inVar, REGDLLTYPE
                Loop
            Case "class", "module", "usercontrol", "form", "designer"
                If InStr(inVal, ";") > 0 Then RemoveNextArg inVal, ";"
                inVal = MapFolder(GetFilePath(Path), inVal)
                ParseSource inVal, ReadFile(inVal)
        End Select
    Loop

End Function

Public Function GetReferences(ByVal f As Folder) As String
    On Error GoTo catchexit
    
    Dim s As Folder
    Dim e As File
    
    For Each s In f.SubFolders
        If InStr(s.Path, "\Test") = 0 Then
            GetReferences s
        End If
    Next
    
    For Each e In f.Files
        Select Case GetFileExt(e.Path)
            Case ".bas", ".cls", ".frm", ".ctl"
                ParseSource e.Path, ReadFile(e.Path)
            Case ".vbp"
                ParseProject e.Path, ReadFile(e.Path)
        End Select
    Next
    Exit Function
catchexit:
    Err.Clear
    Resume Next
End Function

Public Function GenerateAppIncludes(ByVal f As Folder) As String
    Dim e As File
    
    Manifest = ""
    UnManifest = ""
    
    GetReferences f

    Dim tmp As String
    Dim tmp2 As String
    
    tmp = Manifest
    Do Until tmp = ""
        If InStr(NextArg(tmp, vbCrLf), "${APPPATH}\Common\Binary") > 0 Then
        
            tmp2 = RemoveQuotedArg(tmp, """", """")
            
            Set e = fso.GetFile("C:\Development\Neotext\Common\Projects\" & GetFileTitle(tmp2) & ".vbp")

            ParseProject e.Path, ReadFile(e.Path)
            
        End If
        RemoveNextArg tmp, vbCrLf
    Loop
    
    

'    tmp = Manifest
'    Manifest = ""
'
'    Do Until tmp = ""
'        If InStr(NextArg(tmp, vbCrLf), "InstallSystemLibrary") > 0 Then
'            ManifestSys = ManifestSys & RemoveNextArg(tmp, vbCrLf) & vbCrLf
'        Else
'            Manifest = Manifest & RemoveNextArg(tmp, vbCrLf) & vbCrLf
'        End If
'    Loop

'    tmp = UnManifest
'    UnManifest = ""
'
'    Do Until tmp = ""
'        If InStr(NextArg(tmp, vbCrLf), "InstallSystemLibrary") > 0 Then
'            UnManifestSys = UnManifestSys & RemoveNextArg(tmp, vbCrLf) & vbCrLf
'        Else
'            UnManifest = UnManifest & RemoveNextArg(tmp, vbCrLf) & vbCrLf
'        End If
'    Loop
'

    WriteFile "C:\Development\Neotext\Windows\Deploy\" & GetFileName(f.Path) & ".nsi", _
        "!macro InstallLibaries" & vbCrLf & _
        "IfFileExists ""$INSTDIR\*.exe"" 0 new_installation ;Replace MyApp.exe with your application filename" & vbCrLf & _
        "StrCpy $AlreadyInstalled 1" & vbCrLf & _
        "new_installation:" & vbCrLf & _
        "!insertmacro VB6RunTimeInstall C:\Development\Neotext\Windows\Runtime $AlreadyInstalled" & vbCrLf & _
        "SetOverwrite ifdiff" & vbCrLf & _
        Manifest & _
        "!macroend" & vbCrLf & _
        "!macro UninstallLibaries" & vbCrLf & _
        "!insertmacro VB6RunTimeUnInstall" & vbCrLf & _
        UnManifest & _
        "!macroend" & vbCrLf & vbCrLf
 
    Manifest = ""
    UnManifest = ""
    ManifestSys = ""
    UnManifestSys = ""
    
End Function

Public Function ResetLicense(ByVal f As Folder, ByVal FindText As String, ByVal SetText As String) As String
    Dim s As Folder
    Dim e As File
    
    For Each s In f.SubFolders
        ResetLicense s, FindText, SetText
    Next
    Dim txt As String
    
    For Each e In f.Files
        Select Case GetFileExt(e.Path)
            Case ".bas", ".cls", ".frm", ".ctl", ".dsr"
                If Trim(LCase(GetFilePath(e.Path))) <> LCase("C:\Development\Neotext\Source\Projects\Includes") And _
                    InStr(LCase(GetFilePath(e.Path)), "copy of") = 0 Then
                    On Error Resume Next
                    txt = ReadFile(e.Path)
                    If Not Err Then
                    
                        If InStr(txt, FindText) > 0 Then
                            txt = Replace(txt, FindText, SetText)
                            WriteFile e.Path, txt
                        End If
                    Else
                         Err.Clear
                    End If

                End If
        End Select
    Next
    
End Function

Public Sub Main()
    Set fso = New Scripting.FileSystemObject
        
    If Command = "/license on" Then
        ResetLicense fso.GetFolder("C:\Development\Neotext"), "'%LICENSE%", ReadFile("C:\Development\Neotext\Source\Binary\Controls.txt")
    ElseIf Command = "/license off" Then
        ResetLicense fso.GetFolder("C:\Development\Neotext"), ReadFile("C:\Development\Neotext\Source\Binary\Controls.txt"), "'%LICENSE%"
    Else
        CurDir "C:\Development\Neotext\Source\Binary"
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\BasicNeotext")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\Blacklawn")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\CrayonStill")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\Creata-Tree")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\HouseOfGlass")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\IdentAuth")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\Max-FTP")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\MaxLand")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\RemindMe")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\Sequencer")
        GenerateAppIncludes fso.GetFolder("C:\Development\Neotext\To-Doster")

    End If
    Set fso = Nothing
End Sub
