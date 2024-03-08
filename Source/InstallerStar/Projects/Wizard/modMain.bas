Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit
Option Compare Binary
Option Private Module

Private Const MultipleCabs = True

'TOP DOWN
Public Registry As New Registry
Public Program As New Program

Private Enum ManifestSection
    WizardDefaults = 1
    ProgramFiles = 2
    WindowsSystem32 = 3
    ExecuteWaits = 4
    ShortCuts = 5
    Excludes = 6
    Includes = 7
    FileTypes = 8
End Enum

Public Enum InstallMode
    Normal = 0
    ZeroUI = -1
    SheekG = 1
End Enum

Public Enum ResourceID
    LicenseAgreement = 101
    InstallerSEDFile = 102
    NSISInstallerFile = 103
End Enum

Public SimSilence As InstallMode
Public SimCommand As String
Private Signing As Boolean

Public Sub UninstallFirst()
    If (SearchPath("*.mdb", 1, GetCurrentAppDataFolder & "\" & Program.AppValue, FindAll) <> "") Or _
        (SearchPath("*.mdb", 1, GetProgramFilesFolder & "\" & Program.AppValue, FindAll) <> "") Or _
        PathExists(GetCurrentAppDataFolder & "\" & Program.AppValue & "\Uninstall.exe", True) Then
        Dim mdb As String
        Select Case SimSilence
            Case InstallMode.Normal, InstallMode.SheekG
                If MsgBox("Warning: a conflicting prior version needs to be uninstalled before" & vbCrLf & _
                        "installation.  Do you wish to still proceed with uninstalltion first?", vbQuestion + vbYesNo) = vbNo Then
                    Program.CloseAll
                    End
                End If
'                If PathExists(GetCurrentAppDataFolder & "\" & Program.AppValue & "\Uninstall.exe", True) Then
'                    RunProcess GetCurrentAppDataFolder & "\" & Program.AppValue & "\Uninstall.exe"
'                ElseIf PathExists(GetProgramFilesFolder & "\" & Program.AppValue & "\Uninstall.exe", True) Then
'                    RunProcess GetProgramFilesFolder & "\" & Program.AppValue & "\Uninstall.exe"
'                Else
                    mdb = SearchPath("*.mdb", 1, GetProgramFilesFolder & "\" & Program.AppValue, FirstOnly)
                    If mdb <> "" Then Kill NextArg(mdb, vbCrLf)
                    RemovePath GetCurrentAppDataFolder & "\" & Program.AppValue
'                End If
                
            Case Else
'                If PathExists(GetCurrentAppDataFolder & "\" & Program.AppValue & "\Uninstall.exe", True) Then
'                    RunProcess GetCurrentAppDataFolder & "\" & Program.AppValue & "\Uninstall.exe"
'                ElseIf PathExists(GetProgramFilesFolder & "\" & Program.AppValue & "\Uninstall.exe", True) Then
'                    RunProcess GetProgramFilesFolder & "\" & Program.AppValue & "\Uninstall.exe"
'                Else
                    mdb = SearchPath("*.mdb", 1, GetProgramFilesFolder & "\" & Program.AppValue, FirstOnly)
                    If mdb <> "" Then Kill NextArg(mdb, vbCrLf)
                    RemovePath GetCurrentAppDataFolder & "\" & Program.AppValue
'                End If
        End Select
    End If
End Sub

Public Function InstallLocation(Optional ItemType As String = "?") As String
    Dim retval As String

    If (Not Program.Legacy) Then

        Select Case ItemType
            Case "!", "-!"
                retval = GetCurrentAppDataFolder
            Case "?", "-?"
                retval = GetProgramFilesFolder
            Case "$", "-$"
                retval = GetAllUsersAppDataFolder
        End Select

    End If

    If retval = "" Then
        If PathExists(GetCurrentAppDataFolder & "\" & Program.AppValue & "\" & Program.Default, True) Then
            retval = GetCurrentAppDataFolder

        Else
            retval = GetProgramFilesFolder
        End If
    End If
    If Right(retval, 1) = "\" Then retval = Left(retval, Len(retval) - 1)

    UninstallFirst
    InstallLocation = retval
    
End Function

Public Function System32Location(Optional ByVal ItemName As String) As String
    Dim retval As String
    If (InStr(LCase(WinVerInfo), "server") > 0) Then
        If PathExists(GetSystem32Folder & "inetsrv", False) Then
            If ItemName <> "" Then
                If Not PathExists(GetSystem32Folder & ItemName, True) Then
                    retval = GetSystem32Folder & "inetsrv"
                End If
            End If
        End If
    End If
    If retval = "" Then retval = GetSystem32Folder
    If Right(retval, 1) = "\" Then retval = Left(retval, Len(retval) - 1)
    System32Location = retval & IIf(ItemName <> "", "\" & ItemName, "")
End Function
Public Sub Main()

    SimCommand = Command
    
    Dim ddfText As String
    Dim iniText As String
    Select Case UCase(NextArg(SimCommand, " "))
        Case "/S", "/Q0", "/QN0", "/QT", "/QNT", "/QUIET", "/SILENT"
            SimSilence = ZeroUI
            RemoveNextArg SimCommand, " "
        Case "/Q", "/Q1", "/QN1", "/SPLASH", "/SHEEK"
            SimSilence = SheekG
            RemoveNextArg SimCommand, " "
        Case Else
            SimSilence = Normal
    End Select

    If IsDebugger Or SimCommand = "" Then SimCommand = Left(AppPath, Len(AppPath) - 1)
    
    If IsDebugger Then
        Const NeotextAppFolder = "Max-FTP"
        Select Case MsgBox("Create test data using " & NeotextAppFolder & "?", vbYesNoCancel)
            Case vbYes
                SimCommand = "/compile " & NeotextAppFolder
            Case vbCancel
                End
        End Select
            
    End If

    If (NextArg(SimCommand, " ") = "/compile") Then

        Signing = PathExists("\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", True) And Not IsDebugger
        
        ChDir GetFilePath(GetFilePath(GetFilePath(AppPath))) & "\" & RemoveArg(SimCommand, " ")
        
        LoadManifest CurDir & "\Deploy\Manifest.ini"

        GenerateProgram CurDir & "\Binary", RemoveArg(SimCommand, " "), ddfText, iniText

        ddfText = ".Set DestinationDir=" & vbCrLf & ddfText & _
                ".Set DestinationDir=" & vbCrLf

        If MultipleCabs Then
            ddfText = ".Set CabinetNameTemplate=Inst000*.cab,Inst00*.cab,Inst0*.cab,Inst*.cab" & vbCrLf & _
                ".Set CabinetName1=Inst0001.cab" & vbCrLf & _
                ".Set CabinetName2=Inst0002.cab" & vbCrLf & _
                ".Set CabinetName3=Inst0003.cab" & vbCrLf & _
                ".Set CabinetName4=Inst0004.cab" & vbCrLf & _
                ".Set CabinetName5=Inst0005.cab" & vbCrLf & _
                ".Set CabinetName6=Inst0006.cab" & vbCrLf & _
                ".Set CabinetName7=Inst0007.cab" & vbCrLf & _
                ".Set CabinetName8=Inst0008.cab" & vbCrLf & _
                ".Set CabinetName9=Inst0009.cab" & vbCrLf & _
                ".Set CabinetName10=Inst0010.cab" & vbCrLf & _
                ".Set CabinetName11=Inst0011.cab" & vbCrLf & _
                ".Set CabinetName12=Inst0012.cab" & vbCrLf & _
                ".Set CabinetName13=Inst0013.cab" & vbCrLf & _
                ".Set CabinetName14=Inst0014.cab" & vbCrLf & _
                ".Set CabinetName15=Inst0015.cab" & vbCrLf & _
                ".Set CabinetName16=Inst0016.cab" & vbCrLf & _
                ".Set CabinetName17=Inst0017.cab" & vbCrLf & _
                ".Set CabinetName18=Inst0018.cab" & vbCrLf & _
                ".Set CabinetName19=Inst0019.cab" & vbCrLf & _
                ".Set CabinetName20=Inst0020.cab" & vbCrLf & ddfText
        Else
            ddfText = ".Set CabinetNameTemplate=InstStar.CAB" & vbCrLf & ddfText
        End If
        ddfText = ".Option Explicit" & vbCrLf & _
            ".Set ReservePerCabinetSize=6144" & vbCrLf & _
            ".Set Cabinet=on" & vbCrLf & _
            ".Set Compress=on" & vbCrLf & _
            ".Set CompressionType=MSZip" & vbCrLf & _
            ".Set CompressionLevel=7" & vbCrLf & _
            ".Set CompressionMemory=21" & vbCrLf & _
            ".Set MaxDiskSize=1.44M" & vbCrLf & _
            ".Set DiskDirectoryTemplate=" & vbCrLf & ddfText

                                        
        iniText = "[ProgramFiles]" & vbCrLf & iniText & vbCrLf & "[WindowsSystem32]" & vbCrLf
                
        GenerateSystem32 GetFilePath(CurDir), RemoveArg(SimCommand, " "), ddfText, iniText
        
        WriteFile GetFilePath(CurDir) & "\InstallerStar\Binary\Program.ddf", ddfText

        SaveManifest CurDir & "\Deploy\Manifest.ini", iniText
        SaveManifest GetFilePath(CurDir) & "\InstallerStar\Binary\Manifest.ini", iniText
        
        ChDir GetFilePath(CurDir) & "\InstallerStar\Binary"
        If Signing Then
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/signonly """ & CurDir & "\Wizard.exe""", vbHide, True
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/signonly """ & CurDir & "\Remove.exe""", vbHide, True
        End If
        KillFiles "C:\Development\Neotext\InstallerStar\Binary\Inst*.cab"
        
        RunProcess "makecab", "/F " & CurDir & "\Program.ddf", vbHide, True
        WriteFile CurDir & "\InstStar.nsi", Replace(Replace(StrConv(LoadResData(NSISInstallerFile, "CUSTOM"), vbUnicode), "%appvalue%", Program.AppValue), "%curdir%", CurDir)
        WriteFile CurDir & "\Installer.SED", AddCabFiles(Replace(Replace(Replace(StrConv(LoadResData(InstallerSEDFile, "CUSTOM"), vbUnicode), "%appvalue%", Program.AppValue), "%curdir%", CurDir), "HideExtractAnimation=0", "HideExtractAnimation=1"))
        
        If Signing Then
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/timeonly """ & CurDir & "\Wizard.exe""", vbHide, True
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/timeonly """ & CurDir & "\Remove.exe""", vbHide, True
        End If
        RunProcess "C:\Program Files\NSIS\makensis.exe", CurDir & "\InstStar.nsi", vbHide, True
        If Signing Then
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/sign """ & CurDir & "\InstStar.exe""", vbHide, True
        End If
        RunProcess GetSystem32Folder & "IExpress.exe", "/N " & IIf(IsDebugger, "/Q ", "") & "/M " & CurDir & "\Installer.SED", vbHide, True
        If Signing Then
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/sign """ & CurDir & "\Installer.exe""", vbHide, True
        End If
        FileCopy CurDir & "\Installer.exe", GetFilePath(GetFilePath(GetFilePath(AppPath))) & "\" & RemoveArg(SimCommand, " ") & "\Deploy\Installer.exe"

       If Not IsDebugger Then
        
            If PathExists(CurDir & "\Manifest.ini", True) Then Kill CurDir & "\Manifest.ini"
            If PathExists(CurDir & "\InstStar.nsi", True) Then Kill CurDir & "\InstStar.nsi"
            If PathExists(CurDir & "\InstStar.exe", True) Then Kill CurDir & "\InstStar.exe"
            
            If PathExists(CurDir & "\setup.inf", True) Then Kill CurDir & "\setup.inf"
            If PathExists(CurDir & "\setup.rpt", True) Then Kill CurDir & "\setup.rpt"

            If PathExists(CurDir & "\Program.ddf", True) Then Kill CurDir & "\Program.ddf"
            If PathExists(CurDir & "\System32.ddf", True) Then Kill CurDir & "\System32.ddf"
            If PathExists(CurDir & "\Installer.SED", True) Then Kill CurDir & "\Installer.SED"

            If PathExists(CurDir & "\Installer.exe", True) Then Kill CurDir & "\Installer.exe"
            
            DelCabFiles
            
        End If
        
        ChDir GetFilePath(GetFilePath(GetFilePath(AppPath))) & "\" & RemoveArg(SimCommand, " ") & "\Deploy"
            
        If Not PathExists(CurDir & "\Installer.exe", True) Then
            MsgBox "The package was not created.", vbCritical
        Else

            If PathExists(CurDir & "\" & Program.Package & ".exe", True) Then
                On Error Resume Next
                Kill CurDir & "\" & Program.Package & ".exe"
                If Err Then
                    Err.Clear
                    Do While MsgBox("Unable to write package """ & Program.Package & ".exe""", vbRetryCancel) = vbRetry
                        Kill CurDir & "\" & Program.Package & ".exe"
                        If Not Err Then
                            Exit Do
                        Else
                            Err.Clear
                        End If
                    Loop
                End If
            End If
            Name CurDir & "\Installer.exe" As CurDir & "\" & Program.Package & ".exe"
            If Signing Then
                RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/sign """ & CurDir & "\" & Program.Package & ".exe""", vbHide, True
            End If
        End If
            
        End

    End If
    
    If IsDebugger Then
        
        Select Case MsgBox("Do you want to preform the uninstallation?", vbYesNoCancel)
            Case vbYes
                Program.Installed = True
            Case vbNo
                
            Case vbCancel
                End
        End Select

    End If

    SimCommand = GetLongPath(SimCommand)
    
    ChDir Left(AppPath, Len(AppPath) - 1)
    
    LoadManifest AppPath & "Manifest.ini"

    If PathExists(SimCommand & "\Uninstall.exe", True) Or Program.Installed Then
        Program.Installed = True
        Select Case StrReverse(NextArg(StrReverse(Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Program.AppValue, "UninstallString", "/N")), " "))
            Case "/S"
                SimSilence = ZeroUI
            Case "/Q"
                SimSilence = SheekG
            Case "/N"
                SimSilence = Normal
        End Select
    End If
    
    Load frmMain
    
    If SimSilence = SheekG Then
        frmSplash.Show
    Else
        frmMain.StartWizard
    End If

End Sub

Private Function AddCabFiles(ByVal SEDText As String) As String
    Dim txt1 As String
    Dim txt2 As String
    Dim cnt As Long
    Dim nxt As String
    nxt = Dir(CurDir & "\Inst*.cab")
    If LCase(nxt) = "inststar.cab" Then
        SEDText = Replace(SEDText, "FILE1=" & vbCrLf, "FILE1=""InstStar.cab""" & vbCrLf)
        If Signing Then
            RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/sign """ & CurDir & "\InstStar.cab"""", vbHide, True"
        End If
    Else
        cnt = 1
        nxt = Dir(CurDir & "\Inst*" & Trim(CStr(cnt)) & ".cab")
        Do While nxt <> ""
            
            txt1 = txt1 & "FILE" & Trim(cnt) & "=""Inst" & Padding(4, cnt, "0") & ".cab""" & vbCrLf
            txt2 = txt2 & "%FILE" & Trim(cnt) & "%=" & vbCrLf
           ' Name CurDir & "\" & nxt As CurDir & "\Inst" & Padding(4, cnt, "0") & ".cab"
            If Signing Then
                RunProcess "C:\Program Files\Microsoft Visual Studio\VB98\VBN.EXE", "/sign """ & CurDir & "\Inst" & Padding(4, cnt, "0") & ".cab""", vbHide, True
            End If
            cnt = cnt + 1
            nxt = Dir(CurDir & "\Inst*" & Trim(CStr(cnt)) & ".cab")
        Loop

        
        SEDText = Replace(SEDText, "FILE1=" & vbCrLf, txt1)
        SEDText = Replace(SEDText, "%FILE1%=" & vbCrLf, txt2)
    End If
    AddCabFiles = SEDText
End Function

Private Sub DelCabFiles()
    Dim cnt As Long
    Dim nxt As String
    nxt = Dir(CurDir & "\Inst*.cab")

    Do While nxt <> ""
        Kill CurDir & "\" & nxt
        nxt = Dir(CurDir & "\Inst*.cab")
    Loop
End Sub

Private Sub SaveManifest(ByVal FilePath As String, ByVal iniText As String)
    Dim skText As String
    Dim cnt As Long
    Dim item As String
    If Program.ShortCuts.count > 0 Then
        skText = "[ShortCuts]" & vbCrLf

        For cnt = 1 To Program.ShortCuts.count
            item = Program.ShortCuts(cnt)
            Select Case RemoveNextArg(item, " ")
                Case "*"
                    skText = skText & "Folder = " & item & vbCrLf
                Case "?"
                    skText = skText & "File = " & item & vbCrLf
            End Select
        Next
        skText = skText & vbCrLf
    End If

    If Program.Excludes.count > 0 Then
        skText = skText & "[Excludes]" & vbCrLf

        For cnt = 1 To Program.Excludes.count
            item = Program.Excludes(cnt)
            Select Case RemoveNextArg(item, " ")
                Case "*"
                    skText = skText & "Wild = " & item & vbCrLf
                Case "?"
                    skText = skText & "Exact = " & item & vbCrLf
            End Select
        Next
        skText = skText & vbCrLf
    End If

    If Program.Includes.count > 0 Then
        skText = skText & "[Includes]" & vbCrLf

        For cnt = 1 To Program.Includes.count
            skText = skText & Program.Includes(cnt) & vbCrLf
        Next

        skText = skText & vbCrLf
    End If

    If Program.FileTypes.count > 0 Then
        skText = skText & "[FileTypes]" & vbCrLf

        For cnt = 1 To Program.FileTypes.count
            item = Program.FileTypes(cnt)
            skText = skText & item & vbCrLf
        Next
        skText = skText & vbCrLf
    End If
    
    iniText = "[WizardDefaults]" & vbCrLf & _
                "Display= " & Program.Display & vbCrLf & _
                "AppValue = " & Program.AppValue & vbCrLf & _
                "Package= " & Program.Package & vbCrLf & _
                "Default= " & Program.Default & vbCrLf & _
                "Author= " & Program.Author & vbCrLf & _
                "Website= " & Program.WebSite & vbCrLf & _
                "Contact= " & Program.Contact & vbCrLf & _
                "Restore= " & Program.Restore & vbCrLf & _
                "Legacy= " & Program.Legacy & vbCrLf & _
                vbCrLf & "[ExecuteWaits]" & vbCrLf & _
                "Backup=" & Program.Executes.backup & vbCrLf & _
                "Remove=" & Program.Executes.Remove & vbCrLf & _
                "Restore=" & Program.Executes.Restore & vbCrLf & _
                "Initial=" & Program.Executes.Initial & vbCrLf & _
                "Service=" & Program.Executes.Service & vbCrLf & _
                vbCrLf & skText & iniText & vbCrLf & vbCrLf

    WriteFile FilePath, iniText

End Sub

Private Sub LoadManifest(ByVal FilePath As String)
    Dim inText As String
    Dim inLine As String
    Dim section As ManifestSection
    inText = ReadFile(FilePath)
    Do Until inText = ""
        inLine = RemoveNextArg(inText, vbCrLf)
        inLine = Trim(RemoveNextArg(inLine, ";"))
        If inLine <> "" Then
            If InStr(inLine, "[") = 1 And InStr(inLine, "]") = Len(inLine) Then
                'section
                Select Case LCase(RemoveQuotedArg(inLine, "[", "]"))
                    Case "wizarddefaults"
                        section = WizardDefaults
                    Case "programfiles"
                        section = ProgramFiles
                    Case "windowssystem32"
                        section = WindowsSystem32
                    Case "executewaits"
                        section = ExecuteWaits
                    Case "shortcuts"
                        section = ShortCuts
                    Case "excludes"
                        section = Excludes
                    Case "includes"
                        section = Includes
                    Case "filetypes"
                        section = FileTypes
                        
                End Select
            Else
                'line
                Select Case section
                    Case WizardDefaults
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "display"
                                Program.Display = inLine
                            Case "appvalue"
                                Program.AppValue = inLine
                            Case "package"
                                Program.Package = inLine
                            Case "default"
                                Program.Default = inLine
                            Case "author"
                                Program.Author = inLine
                            Case "website"
                                Program.WebSite = inLine
                            Case "contact"
                                Program.Contact = inLine
                            Case "restore"
                                If Trim(LCase(inLine)) = "false" Or Trim(LCase(inLine)) = "true" Then
                                    Program.Restore = CBool(inLine)
                                End If
                            Case "legacy"
                                Program.Legacy = CBool(inLine)
                        End Select
                    Case ProgramFiles
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "folder"
                                Program.ProgramFiles.Add "* " & inLine
                            Case "file"
                                Program.ProgramFiles.Add "? " & inLine
                            Case "current", "custom"
                                Program.ProgramFiles.Add "-! " & inLine
                            Case "alluser"
                                Program.ProgramFiles.Add "-$ " & inLine
                        End Select
                    Case WindowsSystem32
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "shared"
                                Program.System32.Add "* " & inLine
                            Case "system"
                                Program.System32.Add "? " & inLine
                            Case "normal"
                                Program.System32.Add "! " & inLine
                        End Select
                    Case Includes
                        Program.Includes.Add inLine
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "folder"
                                Program.ProgramFiles.Add "=* " & inLine
                            Case "file"
                                Program.ProgramFiles.Add "=? " & inLine
                            Case "current", "custom"
                                Program.ProgramFiles.Add "=! " & inLine
                            Case "alluser"
                                Program.ProgramFiles.Add "=$ " & inLine
                            Case "shared"
                                Program.System32.Add "=* " & inLine
                            Case "system"
                                Program.System32.Add "=? " & inLine
                            Case "normal"
                                Program.System32.Add "=! " & inLine
                        End Select
                    Case ExecuteWaits
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "backup"
                                Program.Executes.backup = inLine
                            Case "remove"
                                Program.Executes.Remove = inLine
                            Case "restore"
                                Program.Executes.Restore = inLine
                            Case "initial"
                                Program.Executes.Initial = inLine
                            Case "service"
                                Program.Executes.Service = inLine
                        End Select
                    Case ShortCuts
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "folder"
                                Program.ShortCuts.Add "* " & inLine
                            Case "file"
                                Program.ShortCuts.Add "? " & inLine
                            Case "requireall"
                                Program.AllShortCuts = inLine
                        End Select
                    Case Excludes
                        Select Case LCase(RemoveNextArg(inLine, "="))
                            Case "wild"
                                Program.Excludes.Add "* " & inLine
                            Case "exact"
                                Program.Excludes.Add "? " & inLine
                        End Select
                    Case FileTypes
                        Program.FileTypes.Add inLine
                End Select
            End If
        End If
    Loop

    If PathExists(AppPath & "msvbvm60.dll", True) And Not Program.Installed Then
        Program.System32.Add "% msvbvm60.dll|" & FileLen(AppPath & "msvbvm60.dll") & "|" & GetFileDate(AppPath & "msvbvm60.dll") & "|" & GetFileVersion(AppPath & "msvbvm60.dll")
    End If
End Sub

Public Sub NetStart(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(Registry.GetValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 0) Then
            Registry.SetValue &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 1
            RunProcess LCase(Trim(exePath))
            timOut = Now
            Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)

                DoEvents
                Sleep 1
                
'                #If Not modCommon And Not modDoTasks Then
'                    DoEvents
'                #ElseIf modDoTasks Then
'                    modDoTasks.DoTasks
'                #ElseIf modCommon Then
'                    modCommon.DoTasks
'                #Else
'                    DoEvents
'                #End If
            Loop
        End If
    Else
        If EXEName = "" Then
            RunProcess SysPath & "net.exe", "start " & ServiceName, vbHide, True
        Else
            If ProcessRunning(EXEName) = 0 Then
                RunProcess SysPath & "net.exe", "start " & ServiceName, vbHide, True
                timOut = Now
                Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)
                    DoEvents
                    Sleep 1

'                    #If Not modCommon And Not modDoTasks Then
'                        DoEvents
'                    #ElseIf modDoTasks Then
'                        modDoTasks.DoTasks
'                    #ElseIf modCommon Then
'                        modCommon.DoTasks
'                    #Else
'                        DoEvents
'                    #End If
                Loop
'                If ProcessRunning(EXEName) = 0 Then
'                    StartNTService ServiceName
'                    If ProcessRunning(EXEName) = 0 Then
'                        If PathExists(AppPath & EXEName, True) Then
'                            RunProcess AppPath & EXEName
'                        Else
'                            RunFile EXEName
'                        End If
'                    End If
'                End If
            End If
        End If
    End If
End Sub

Public Sub NetStop(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(Registry.GetValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 1) Then
            KillApp LCase(Trim(exePath))
            timOut = Now
            Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
                DoEvents
                Sleep 1

'                #If Not modCommon And Not modDoTasks Then
'                    DoEvents
'                #ElseIf modDoTasks Then
'                    modDoTasks.DoTasks
'                #ElseIf modCommon Then
'                    modCommon.DoTasks
'                #Else
'                    DoEvents
'                #End If
            Loop
        End If
    Else
        If EXEName = "" Then
            RunProcess SysPath & "net.exe", "stop " & ServiceName, vbHide, True
        Else
            If ProcessRunning(EXEName) = 1 Then
                RunProcess SysPath & "net.exe", "stop " & ServiceName, vbHide, True
                timOut = Now
                Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
                    DoEvents
                    Sleep 1

'                    #If Not modCommon And Not modDoTasks Then
'                        DoEvents
'                    #ElseIf modDoTasks Then
'                        modDoTasks.DoTasks
'                    #ElseIf modCommon Then
'                        modCommon.DoTasks
'                    #Else
'                        DoEvents
'                    #End If
                Loop
'                If ProcessRunning(EXEName) = 1 Then
'                    StopNTService ServiceName
'                    If ProcessRunning(EXEName) = 1 Then
'                        KillApp EXEName
'                    End If
'                End If
            End If
        End If
    End If
End Sub

Public Sub NetContinue(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(Registry.GetValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 0) Then
            Registry.SetValue &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 1
            RunProcess LCase(Trim(exePath))
            timOut = Now
            Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)
                DoEvents
                Sleep 1

'                #If Not modCommon And Not modDoTasks Then
'                    DoEvents
'                #ElseIf modDoTasks Then
'                    modDoTasks.DoTasks
'                #ElseIf modCommon Then
'                    modCommon.DoTasks
'                #Else
'                    DoEvents
'                #End If
            Loop
        End If
    Else
        If EXEName = "" Then
            RunProcess SysPath & "net.exe", "continue " & ServiceName, vbHide, True
        Else
            If ProcessRunning(EXEName) = 0 Then
                RunProcess SysPath & "net.exe", "continue " & ServiceName, vbHide, True
                timOut = Now
                Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)
                    DoEvents
                    Sleep 1

'                    #If Not modCommon And Not modDoTasks Then
'                        DoEvents
'                    #ElseIf modDoTasks Then
'                        modDoTasks.DoTasks
'                    #ElseIf modCommon Then
'                        modCommon.DoTasks
'                    #Else
'                        DoEvents
'                    #End If
                Loop
            End If
        End If
    End If
End Sub

Public Sub NetPause(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(Registry.GetValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 1) Then
            KillApp LCase(Trim(exePath))
            timOut = Now
            Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
                DoEvents
                Sleep 1

'                #If Not modCommon And Not modDoTasks Then
'                    DoEvents
'                #ElseIf modDoTasks Then
'                    modDoTasks.DoTasks
'                #ElseIf modCommon Then
'                    modCommon.DoTasks
'                #Else
'                    DoEvents
'                #End If
            Loop
        End If
    Else
        If EXEName = "" Then
            RunProcess SysPath & "net.exe", "pause " & ServiceName, vbHide, True
        Else
            If ProcessRunning(EXEName) = 1 Then
                RunProcess SysPath & "net.exe", "pause " & ServiceName, vbHide, True
                timOut = Now
                Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
                    DoEvents
                    Sleep 1

'                    #If Not modCommon And Not modDoTasks Then
'                        DoEvents
'                    #ElseIf modDoTasks Then
'                        modDoTasks.DoTasks
'                    #ElseIf modCommon Then
'                        modCommon.DoTasks
'                    #Else
'                        DoEvents
'                    #End If
                Loop
            End If
        End If
    End If
End Sub

