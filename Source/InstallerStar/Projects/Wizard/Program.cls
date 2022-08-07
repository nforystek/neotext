VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Program"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Private Enum LockStates
    NoCurrentLock = 0
    InstallLock = 1
    RollBackLock = -1
    UninstallLock = 2
End Enum
    
Public Default As String
Public Contact As String
Public WebSite As String
Public Author As String
Public Restore As Boolean
Public Legacy As Boolean
Public AllShortCuts As Boolean

Public Display As String
Private pAppValue As String
Public Package As String

Private WithEvents CabFile As CabFile
Attribute CabFile.VB_VarHelpID = -1

Private LockState As LockStates

Private pExecutes As New Executes
Private pProgramFiles As New Collection
Private pSystem32 As New Collection
Private pRollback As New Collection
Private pShortCuts As New Collection
Private pExcludes As New Collection
Private pIncludes As New Collection
Private pFileTypes As New Collection

Public RebootFlag As Boolean
Public RemoveFlag As Boolean

Private pCurrentItem As Double
Public TotalOfItem As Double
Public ByteProgress As Double
Public TotalOfBytes As Double

Private CabXMLInfo As String

Private pInstalled As Boolean
Private pStartExecutes As Boolean

Public Property Get CurrentItem() As Double
    If pCurrentItem > TotalOfItem Then pCurrentItem = TotalOfItem
    CurrentItem = pCurrentItem
End Property
Public Property Let CurrentItem(ByVal RHS As Double)
    pCurrentItem = RHS
    If pCurrentItem > TotalOfItem Then pCurrentItem = TotalOfItem
End Property

Public Property Get FileTypes() As Collection
    Set FileTypes = pFileTypes
End Property

Public Property Get Includes() As Collection
    Set Includes = pIncludes
End Property


Public Property Get Excludes() As Collection
    Set Excludes = pExcludes
End Property

Public Property Get ShortCuts() As Collection
    Set ShortCuts = pShortCuts
End Property
Public Property Get Installed() As Boolean
    Installed = pInstalled
End Property
Friend Property Let Installed(ByVal newVal As Boolean)
    pInstalled = newVal
End Property

Public Function IsRunning() As Boolean
    If pProgramFiles.count > 0 Then
        Dim cnt As Long
        Dim fnum As Long
        Dim fname As String
        
        For cnt = 1 To pProgramFiles.count
            fname = NextArg(RemoveArg(pProgramFiles(cnt), " "), "|")
            
            Select Case GetFileExt(fname, True, True)
                Case "exe", "com"
                    If RunningProcessCount(fname) > 0 Then
                        IsRunning = True
                        Exit Function
                    End If
            End Select
        Next
        
    End If
End Function


Public Sub CloseAll()

    If RemoveArg(pExecutes.Service, "|") <> "" Then
        If ProcessRunning(InstallLocation & "\" & Program.AppValue & "\" & NextArg(pExecutes.Service, "|")) > 0 Then
            NetStop RemoveArg(pExecutes.Service, "|"), InstallLocation & "\" & Program.AppValue & "\" & NextArg(pExecutes.Service, "|")
        End If
    End If
    
    If pProgramFiles.count > 0 Then
        Dim cnt As Long
        Dim fnum As Long
        Dim fname As String
        
        For cnt = 1 To pProgramFiles.count
            fname = NextArg(RemoveArg(pProgramFiles(cnt), " "), "|")
            
            Select Case GetFileExt(fname, True, True)
                Case "exe", "com"
                    If Not GetFileName(fname) = GetFileName(AppEXE(False, True)) Then
                        If ProcessRunning(fname) > 0 Then
                            KillApp InstallLocation & "\" & Program.AppValue & "\" & fname
                        End If
                    End If
            End Select
        Next
    End If
End Sub


Private Function IncrementLibrary(ByVal libpath As String, ByVal libvers As String) As Boolean
    Dim libUse As String
    Dim libCnt As String
    libpath = LCase(libpath)
    libCnt = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", GetFileName(libpath), CStr(CInt(PathExists(libpath, True))))
    libUse = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), CStr(-CInt(Installed)))
    If libCnt = "-1" Then
        If CheckVersionEqualOrGreater(libvers, libpath) Then
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), "0"
            IncrementLibrary = True
        Else
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), "1"
        End If
    Else
        If libUse = "0" Then libCnt = CStr(CLng(CLng(libCnt) + 1))
        If Not PathExists(libpath, True) Then
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), "-1"
            IncrementLibrary = True
        ElseIf CheckVersionEqualOrGreater(libvers, libpath) Then
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), "0"
            IncrementLibrary = True
        Else
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), "1"
        End If
    End If
    If Not Installed Then Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", GetFileName(libpath), libCnt
    
End Function

Private Function DecrementLibrary(ByVal libpath As String, ByRef RemoveToo As Boolean) As Boolean
    Dim libUse As String
    Dim libCnt As String
    libpath = LCase(libpath)
    RemoveToo = False
    libCnt = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", GetFileName(libpath), CStr(-CInt(PathExists(libpath, True))))
    libUse = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath), -CLng(libCnt > 0))
    If Not libCnt = "-1" Then
        libCnt = CStr(CLng(CLng(libCnt) - 1))
        If libCnt = "0" Then
            Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", GetFileName(libpath)
            Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & pAppValue, GetFileName(libpath)
        ElseIf Not libCnt = "-1" Then
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", GetFileName(libpath), libCnt
        End If
    End If
    If libUse = "-1" Then RemoveToo = True
    If Not libUse = "1" Then
        DecrementLibrary = True
    End If
End Function


Public Property Get AppValue() As String
    AppValue = pAppValue
End Property
Public Property Let AppValue(ByVal RHS As String)
    pAppValue = RHS
End Property

Public Property Get Executes() As Executes
    Set Executes = pExecutes
End Property

Public Property Get ProgramFiles() As Collection
    Set ProgramFiles = pProgramFiles
End Property

Public Property Get System32() As Collection
    Set System32 = pSystem32
End Property

Public Property Get Summary() As String

    If pSystem32.count > 0 Then
        Dim cnt As Long
        cnt = 1
        Do
            Do While Left(pSystem32.item(cnt), 1) = "="
                pSystem32.Remove cnt
                If cnt > pSystem32.count Then Exit Do
            Loop
            cnt = cnt + 1
        Loop Until cnt > pSystem32.count
    End If
            
    Dim ret As String
    CurrentItem = 0
    ByteProgress = 0
    TotalOfBytes = 0
    TotalOfItem = 0
    Dim Inst As Boolean
    Inst = Installed
    
    ret = ret & Program.Display & " - " & IIf(Inst, "Uninstallation summary", "Installation summary") & vbCrLf
    If Program.Restore And PathExists(GetSystem32Folder & "wbem\wmic.exe", True) Then
        ret = ret & "         A restore point will be created;" & vbCrLf
    End If

    TotalOfItem = pProgramFiles.count + pSystem32.count
    Dim folderCount As Long
    
    If pProgramFiles.count > 0 Then
        For cnt = 1 To pProgramFiles.count
            If IsNumeric(NextArg(RemoveArg(pProgramFiles(cnt), "|"), "|")) Then
                TotalOfBytes = TotalOfBytes + CDbl(NextArg(RemoveArg(pProgramFiles(cnt), "|"), "|"))
            ElseIf NextArg(RemoveArg(pProgramFiles(cnt), "|"), "|") = "<DIR>" Then
                folderCount = folderCount + 1
            End If
        Next
    End If

    ret = ret & "         Main program requires;" & vbCrLf
    If pProgramFiles.count > 0 And folderCount > 0 Then
        ret = ret & "                   Files placed in, " & Program.AppValue & " folders:" & vbCrLf
        If pProgramFiles.count > 0 Then
            ret = ret & "                  " & pProgramFiles.count & " files totaling " & Trim(FormatFileSize(TotalOfBytes)) & vbCrLf
        End If
        If folderCount > 0 Then
            ret = ret & "                  " & folderCount & " folders will be created." & vbCrLf
        End If
    Else
        ret = ret & "                  No file is installed in this section." & vbCrLf
    End If
    folderCount = 0
    
    Dim tmp As Double
    If pSystem32.count > 0 Then
        For cnt = 1 To pSystem32.count
            If IsNumeric(NextArg(RemoveArg(pSystem32(cnt), "|"), "|")) Then
                tmp = tmp + CDbl(NextArg(RemoveArg(pSystem32(cnt), "|"), "|"))
            ElseIf NextArg(RemoveArg(pSystem32(cnt), "|"), "|") = "<DIR>" Then
                folderCount = folderCount + 1
            End If
        Next
    End If
    
    ret = ret & "         Shared libraries require;" & vbCrLf
    If pSystem32.count > 0 Then
        ret = ret & "                   Files placed in windows system folder:" & vbCrLf
        ret = ret & "                  " & pSystem32.count & " files totaling: " & Trim(FormatFileSize(tmp)) & vbCrLf
    Else
        ret = ret & "                  No file is installed in this section." & vbCrLf
    End If

    TotalOfBytes = TotalOfBytes + tmp
    
    tmp = 0
    
    If pShortCuts.count > 0 Then
        For cnt = 1 To pShortCuts.count
            If NextArg(pShortCuts(cnt), " ") = "?" Then
                tmp = tmp + 4096
            Else
                folderCount = folderCount + 1
            End If
        Next
    End If
    
    ret = ret & "         Menu shortcuts and file types require;" & vbCrLf
    If pShortCuts.count > 0 And folderCount > 0 Then
        ret = ret & "                   Files under the start menu programs:" & vbCrLf
        If pShortCuts.count > 0 Then
            ret = ret & "                  " & pShortCuts.count & " files approximately: " & Trim(FormatFileSize(tmp)) & vbCrLf
        End If
        If folderCount > 0 Then
            ret = ret & "                  " & folderCount & " folders will be created, and " & pFileTypes.count & " file types." & vbCrLf
        End If
    Else
        ret = ret & "                  No file is installed in this section." & vbCrLf
    End If
    
    TotalOfItem = pSystem32.count + pProgramFiles.count + pShortCuts.count
    
    TotalOfBytes = TotalOfBytes + tmp
    ret = ret & "Total of " & TotalOfItem & " files, at a total of " & Trim(FormatFileSize(TotalOfBytes)) & vbCrLf
    
    Summary = ret
End Property

Public Function LockInstalling() As Boolean
'starts the installing mode locking until completed
'no other lock can be started nor stopped, any unlock
'early where NextProgressive() is still true initiates
'a rollback which is similarly part of lock till done
    If LockState = NoCurrentLock Then
        LockState = InstallLock
        If Not PathExists(InstallLocation & "\" & Program.AppValue, False) Then
            MkDir InstallLocation & "\" & Program.AppValue
        End If
        Registry.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue
        Registry.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\" & Program.Author
        Registry.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\" & Program.Author & "\System"
        ChDir InstallLocation

        If PathExists(SimCommand & "\InstStar.exe", True) Then
            If PathExists(InstallLocation & "\" & Me.AppValue & "\Uninstall.exe", True) Then
                Kill InstallLocation & "\" & Me.AppValue & "\Uninstall.exe"
            End If
            FileCopy SimCommand & "\InstStar.exe", InstallLocation & "\" & Me.AppValue & "\Uninstall.exe"
            Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "UninstallString", """" & InstallLocation & "\" & Me.AppValue & "\Uninstall.exe""" & IIf(SimSilence = SheekG, " /Q", IIf(SimSilence = ZeroUI, " /S", " /N"))
        End If
        
        If Not Program.Installed Then
            'cab only needed on install
            FindNextCabFile
        End If
        
        LockInstalling = True
    End If
End Function
Private Sub FindNextCabFile()

    Dim cabNum As Long
    If Not CabFile Is Nothing Then
        cabNum = CLng(Mid(GetFileName(CabFile.CabName), 5, 4))
        Set CabFile = Nothing
    End If
    
    Set CabFile = New CabFile
    If PathExists(SimCommand & "\InstStar.cab", True) Then
        CabFile.CabName = SimCommand & "\InstStar.cab"
    Else
        cabNum = cabNum + 1
        Dim choice As Integer
        Do Until PathExists(SimCommand & "\Inst" & Padding(4, cabNum, "0") & ".cab", True)
            Select Case MsgBox("Please insert the disk labeled as number " & cabNum & " of the set.", vbRetryCancel, AppValue)
                Case vbRetry
                Case vbCancel
                    frmMain.Command1_Click 1
                    Exit Sub
            End Select
        Loop
        
        CabFile.CabName = SimCommand & "\Inst" & Padding(4, cabNum, "0") & ".cab"
    End If
    CabXMLInfo = CabFile.GetXML
    RemoveNextArg CabXMLInfo, "<FILES>"
    CabXMLInfo = Replace(RemoveNextArg(CabXMLInfo, "</FILES>"), vbCrLf, "")
End Sub

Private Sub RoundUpHistoryDLLS2(ByRef vals() As String)
    Dim cnt As Long
    Dim num As Long
    'only if the dll/ocx name does not already exist add it and
    'reguardless of activeX assume regsvr it will fail silently
    For cnt = LBound(vals) To UBound(vals)
        num = 0
        Select Case GetFileExt(vals(cnt), True, True)
            Case "dll", "ocx", "exe", "tlb"
                If pSystem32.count > 0 Then
                    num = 1
                    Do
                        If LCase(GetFileName(NextArg(RemoveArg(pSystem32(num), " "), "|"))) = LCase(vals(cnt)) Then
                            num = 0
                            Exit For
                        End If
                        num = num + 1
                    Loop While num <= pSystem32.count
                End If
        End Select
        If num > 0 Then
            If GetFileExt(vals(cnt), True, True) = "tlb" Then
                pSystem32.Add "! " & vals(cnt) & "|" & GetFileSize(System32Location(vals(cnt))) & "|" & GetFileDate(System32Location(vals(cnt))) & "|" & GetFileVersion(System32Location(vals(cnt)))
            Else
                pSystem32.Add "* " & vals(cnt) & "|" & GetFileSize(System32Location(vals(cnt))) & "|" & GetFileDate(System32Location(vals(cnt))) & "|" & GetFileVersion(System32Location(vals(cnt)))
            End If
        End If
    Next
End Sub
Private Sub RoundUpHistoryDLLS()
    
    Dim vals() As String 'start of uninstall get all dlls this appliaction had prior or otherwise in uninstall key
    If Registry.SubValues(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, vals()) > -1 Then
        RoundUpHistoryDLLS2 vals 'add them to uninstall collection
    End If
    Erase vals
    If Registry.SubKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", vals()) > 0 Then
        Dim cnt As Long 'next if there are no other Neotext applications with uninstall not ourself
        'uninstall then add the system shared enteries for complete uninstall of loose left over ends
        For cnt = LBound(vals) To UBound(vals)
            If vals(cnt) <> "" And vals(cnt) <> AppValue Then
                If Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & vals(cnt), "Publisher", "NULL") = Program.Author Then
                    Exit Sub 'found one not our own and is Neotext so exit the procedure, another will be last to unisntall and not this instance for shared entries
                End If
            End If
        Next
    End If
    Erase vals
    If Registry.SubValues(HKEY_LOCAL_MACHINE, "SOFTWARE\" & Program.Author & "\System", vals()) > -1 Then
        RoundUpHistoryDLLS2 vals 'add them to uninstall collection
    End If
    Dim nxt As String
    nxt = Dir(GetCommonFilesFolder & "\" & Me.Author & "\*.bak")
    Do While nxt <> ""
        If GetFileExt(nxt, True, True) = "bak" Then
            pSystem32.Add "& " & GetFileTitle(nxt) & "|" & GetFileSize(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt) & "|" & GetFileDate(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt) & "|" & GetFileVersion(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt)
        Else
            pSystem32.Add "& " & nxt & "|" & GetFileSize(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt) & "|" & GetFileDate(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt) & "|" & GetFileVersion(GetCommonFilesFolder & "\" & Me.Author & "\" & nxt)
        End If
        nxt = Dir
    Loop
    
End Sub
Public Function LockUninstalling() As Boolean
'starts the uninstalling mode locking until completed
'no other lock can be started nor stopped, any unlock
'early where NextProgressive() is still true fales with
'false until nexprogressive is false, then unlock true
    If LockState = NoCurrentLock Then
        LockState = UninstallLock
        RoundUpHistoryDLLS
        LockUninstalling = True
    End If
End Function

Public Function UnlockInstalling() As Boolean
'signals we want to stop installing (or are finished)
'when nextprogressive is false then this function call
'will return true and is requried to complete a install
'if this is called early then rollback is initiated on
'the current lock and similarly the unlock must complete
    If LockState = InstallLock Then
        If pProgramFiles.count = 0 And pSystem32.count = 0 Then
            CommitInstall
            LockState = NoCurrentLock
            UnlockInstalling = True
        Else
            LockState = RollBackLock
            UnlockInstalling = False
        End If
    ElseIf LockState = RollBackLock Then
        If pRollback.count = 0 Then
            CleanupInstall
            LockState = NoCurrentLock
            UnlockInstalling = True
        Else
            UnlockInstalling = False
        End If
    End If
End Function
Public Property Get Version() As String
    Dim fvi As FILEPROPERTIE
    fvi = GetFileInfo(InstallLocation & "\" & Me.AppValue & "\" & Default)
    Version = fvi.FileVersion
End Property
Private Sub CommitInstall()
    
    Set CabFile = Nothing

    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "DisplayName", Me.AppValue
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "Publisher", Me.Author
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "ProductVersion", Version
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "ModifyPath", """" & InstallLocation & "\" & Me.AppValue & "\Uninstall.exe"""
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "DisplayIcon", InstallLocation & "\" & Me.AppValue & "\" & Default
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "InstallUser", GetUserLoginName
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "InstallLocation", """" & InstallLocation & "\" & Me.AppValue & """"
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "HelpLink", Me.WebSite
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "HelpTelephone", Me.Contact
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "DisplayVersion", Me.Display
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "VersionMajor", NextArg(Version, ".")
    Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue, "VersionMinor", NextArg(RemoveArg(Version, "."), ".")

    TidySharedEntities
    
    If Not CurDir = InstallLocation & "\" & Program.AppValue Then ChDir InstallLocation & "\" & Program.AppValue
End Sub
Public Sub CleanupInstall()
    
    Set CabFile = Nothing

    RemoveFlag = True
    
    Registry.ExpellKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue
    
    TidySharedEntities True
End Sub

Private Sub TidySharedEntities(Optional Removing As Boolean = False)
    Dim progKeys() As String
    Dim cnt As Long

    'remove the common files path install or uninstall finished if it is empty
    If Not PathExists(GetCommonFilesFolder & "\Neotext\*.*", True) Then
        RemovePath GetCommonFilesFolder & "\" & Me.Author
    End If

    If Removing Then 'only on removing the application
        Registry.ExpellKey HKEY_LOCAL_MACHINE, "SOFTWARE\" & Program.Author & "\" & Program.AppValue 'old NSIS installation enteries
        
        'next check for presence of any Neotext application in the uninstall, if one exists exit out the function
        If Registry.SubKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", progKeys()) > -1 Then
            For cnt = LBound(progKeys) To UBound(progKeys)
                If progKeys(cnt) <> "" Then
                    If Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & progKeys(cnt), "Publisher", "NULL") = Program.Author Then
                        Exit Sub
                    End If
                End If
            Next
            Erase progKeys

            'if no Neotext applications installed, then remove shared entry
            Registry.ExpellKey HKEY_LOCAL_MACHINE, "SOFTWARE\" & Program.Author

        End If
    End If
End Sub
Public Function UnlockUninstalling() As Boolean
'signals we want to stop uninstalling (or are finished)
'calling this before nextprogresive is false will drop
'commitmet where it is when nextprogressive becomes false
'and then this returns true, canceling/pausing uninstall
    If LockState = UninstallLock Then
        CleanupInstall
        If pProgramFiles.count = 0 And pSystem32.count = 0 Then
            LockState = NoCurrentLock
            UnlockUninstalling = True
        End If
    End If
End Function
Private Sub DebugProgramFiles()
    Dim cnt As Long
    If pProgramFiles.count > 0 Then
        For cnt = 1 To pProgramFiles.count
            Debug.Print pProgramFiles.item(cnt)
        Next
    End If
End Sub
Private Sub DebugSystem23()
    Dim cnt As Long
    If pSystem32.count > 0 Then
        For cnt = 1 To pSystem32.count
            Debug.Print pSystem32.item(cnt)
        Next
    End If
End Sub

Private Static Function FindNextFileIndex() As Long

    Dim cnt As Long
    Dim fInfo As String
    
    If (pProgramFiles.count > 0) Then
        For cnt = 1 To pProgramFiles.count
            Select Case NextArg(pProgramFiles.item(cnt), " ")
                Case "*"
                    FindNextFileIndex = cnt
                    Exit Function
            End Select
        Next

    End If
    If (pSystem32.count > 0) Then
        For cnt = 1 To pSystem32.count
            Select Case NextArg(pSystem32.item(cnt), " ")
                Case "^", "%"
                    FindNextFileIndex = -cnt
                    Exit Function
            End Select
        Next

    End If

    If (CabXMLInfo = "") Then
        FindNextCabFile
    End If

    fInfo = RemoveQuotedArg(CabXMLInfo, "<FILE>", "</FILE>")
    fInfo = RemoveQuotedArg(fInfo, "<FULLNAME>", "</FULLNAME>")
    
    
    If pProgramFiles.count > 0 Then
        For cnt = 1 To pProgramFiles.count
            If LCase(fInfo) = Trim(LCase(NextArg(RemoveArg(pProgramFiles.item(cnt), " "), "|"))) Then
                FindNextFileIndex = cnt
                'Debug.Print fInfo & " " & pProgramFiles.item(cnt)
                Exit Function
            End If
        Next
    End If
    
    If pSystem32.count > 0 Then
        For cnt = 1 To pSystem32.count
            If LCase(fInfo) = Trim(LCase(NextArg(RemoveArg(pSystem32.item(cnt), " "), "|"))) Then
                FindNextFileIndex = -cnt
                'Debug.Print fInfo & " " & pSystem32.item(cnt)
                Exit Function
            End If
        Next
    End If
    
    
End Function

Public Function NextProgressive(ByRef InfoText As String) As Boolean
'continues the locked process progressive until
'there are non, at which time the functin returns
'false, otherwise it will return true and progress
    Dim colIndex As Long
    If Abs(CLng(LockState)) = CLng(InstallLock) Then
        If LockState = InstallLock Then
            If Not pStartExecutes Then
                pStartExecutes = (Not Executes.NextExecute(InfoText))
                NextProgressive = True
            Else
                If pProgramFiles.count > 0 Or pSystem32.count > 0 Then
                    colIndex = FindNextFileIndex()
                    If pProgramFiles.count > 0 And colIndex > 0 Then
                        InstallProgramFiles pProgramFiles(colIndex), InfoText, colIndex
                        pProgramFiles.Remove colIndex
                        NextProgressive = True
                    ElseIf pSystem32.count > 0 And colIndex < 0 Then
                        InstallSystem32 pSystem32(-colIndex), InfoText, -colIndex
                        pSystem32.Remove -colIndex
                        NextProgressive = True
                    Else
                        NextProgressive = True
                    End If
                ElseIf pShortCuts.count > 0 Then
                    If Not LCase(CurDir) = LCase(GetStartMenuProgramsFolder) Then ChDir GetStartMenuProgramsFolder
                    If SimSilence = Normal Or AllShortCuts Then
                        InstallShortCut pShortCuts(1), InfoText
                        pShortCuts.Remove 1
                    Else
                        InstallShortCut "? " & AppValue & "|" & InstallLocation & "\" & Program.AppValue & "\" & Default & "|", InfoText
                        Do Until pShortCuts.count = 0
                            pShortCuts.Remove 1
                        Loop
                    End If
                    NextProgressive = True
                ElseIf pFileTypes.count > 0 Then
                    InstallFileType pFileTypes(1), InfoText
                    pFileTypes.Remove 1
                    NextProgressive = True
                ElseIf pStartExecutes Then
                    NextProgressive = Executes.NextExecute(InfoText)
                End If
            End If
        ElseIf LockState = RollBackLock Then
            If pRollback.count > 0 Then
                RollbackEntry pRollback(pRollback.count), InfoText
                pRollback.Remove pRollback.count
                NextProgressive = True
            ElseIf pRollback.count = 0 And Registry.KeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue) Then
                Registry.DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & Me.AppValue
                NextProgressive = True
            Else
                NextProgressive = Executes.NextExecute(InfoText, True)
            End If
            If Not NextProgressive Then
                If PathExists(InstallLocation & "\" & Me.AppValue & "\Uninstall.exe", True) Then
                    Kill InstallLocation & "\" & Me.AppValue & "\Uninstall.exe"
                End If
                RemovePath InstallLocation & "\" & Me.AppValue
            End If
        End If
    ElseIf LockState = UninstallLock Then
        If Not pStartExecutes Then
            pStartExecutes = Not Executes.NextExecute(InfoText)
            NextProgressive = True
        Else
            If pProgramFiles.count > 0 Then
                UninstallProgramFiles pProgramFiles(1), InfoText, 1
                pProgramFiles.Remove 1
                NextProgressive = True
            ElseIf pSystem32.count > 0 Then
                'If Not LCase(CurDir) = LCase(Left(GetSystem32Folder, Len(GetSystem32Folder) - 1)) Then ChDir Left(GetSystem32Folder, Len(GetSystem32Folder) - 1)
                UninstallSystem32 pSystem32(1), InfoText, 1
                pSystem32.Remove 1

                NextProgressive = True
            ElseIf pShortCuts.count > 0 Then
                If Not LCase(CurDir) = LCase(GetStartMenuProgramsFolder) Then ChDir GetStartMenuProgramsFolder
                If SimSilence = Normal Then
                    UninstallShortCut pShortCuts(1), InfoText
                    pShortCuts.Remove 1
                Else
                    UninstallShortCut "? " & AppValue & "|" & InstallLocation & "\" & Program.AppValue & "\" & Default & "|", InfoText
                    Do Until pShortCuts.count = 0
                        pShortCuts.Remove 1
                    Loop
                End If
                NextProgressive = True
            ElseIf pFileTypes.count > 0 Then
                UninstallFileType pFileTypes(1), InfoText
                pFileTypes.Remove 1
                NextProgressive = True
            End If
        End If
    End If
End Function

Private Function InstallFileType(ByVal ItemData As String, ByRef InfoText As String) As Boolean
    On Error GoTo failitem

    Dim ItemType As String
    Dim ItemClass As String
    Dim ItemDesc As String
    Dim ItemAction As String
    Dim ItemProgam As String
    Dim ItemParams As String
    ItemType = RemoveNextArg(ItemData, "|")
    ItemClass = RemoveNextArg(ItemData, "|")
    ItemDesc = RemoveNextArg(ItemData, "|")
    
    Registry.CreateKey HKEY_CLASSES_ROOT, ItemType
    
    Registry.SetValue HKEY_CLASSES_ROOT, ItemType, "", ItemClass
    
    Registry.CreateKey HKEY_CLASSES_ROOT, ItemClass
    Registry.SetValue HKEY_CLASSES_ROOT, ItemClass, "", ItemDesc
    
    Registry.CreateKey HKEY_CLASSES_ROOT, ItemClass & "\DefaultIcon"
    Registry.SetValue HKEY_CLASSES_ROOT, ItemClass & "\DefaultIcon", "", InstallLocation & "\" & Program.AppValue & "\" & Program.Default & ",0"
    
    Registry.CreateKey HKEY_CLASSES_ROOT, ItemClass & "\shell"
    
    Do Until ItemData = ""
        ItemAction = RemoveNextArg(ItemData, "|")
        ItemProgam = RemoveNextArg(ItemData, "|")
        ItemParams = RemoveNextArg(ItemData, "|")
        
        Registry.CreateKey HKEY_CLASSES_ROOT, ItemClass & "\shell\" & ItemAction
        Registry.CreateKey HKEY_CLASSES_ROOT, ItemClass & "\shell\" & ItemAction & "\command"
        Registry.SetValue HKEY_CLASSES_ROOT, ItemClass & "\shell\" & ItemAction & "\command", "", """" & InstallLocation & "\" & Program.AppValue & "\" & IIf(ItemProgam = "", Program.Default, ItemProgam) & """ " & ItemParams
        
    Loop
    
    InfoText = "Installed file type: " & ItemDesc
    pRollback.Add "ft " & pFileTypes(1)
    CurrentItem = CurrentItem + 1
    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function

Private Function InstallShortCut(ByVal ItemData As String, ByRef InfoText As String) As Boolean
    On Error GoTo failitem
          
    Dim ItemType As String
    Dim ItemName As String
    Dim ItemLink As String
    Dim ItemArgs As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemLink = RemoveNextArg(ItemData, "|")
    ItemArgs = RemoveNextArg(ItemData, "|")
    
    Dim Target As String
    Dim DestFile As String
    DestFile = GetStartMenuProgramsFolder & "\" & ItemName
    
    Select Case ItemType

        Case "*" 'folder
            pRollback.Add "fe-= " & CurDir
            If Not PathExists(DestFile, False) Then
                InfoText = "Created folder: " & DestFile
                MkDir DestFile
                pRollback.Add "cf+- " & RemoveArg(pShortCuts(1), " ")
            Else
                InfoText = "Folder exists: " & DestFile
                pRollback.Add "cf+= " & RemoveArg(pShortCuts(1), " ")
            End If
            ChDir DestFile
            CurrentItem = CurrentItem + 1
        Case "?" 'file
            DestFile = DestFile & ".lnk"
            If Not PathExists(DestFile, True) Then

                InfoText = "Installed file: " & DestFile
                
                Target = ItemLink
                If InStr(Target, """") > 0 Then Target = RemoveQuotedArg(Target, """", """")
                Target = Replace(Replace(Replace(Target, "%systemroot%\", GetSystem32Folder, , , vbTextCompare), "%windir%\", GetWindowsFolder, , , vbTextCompare), "%approot%\", InstallLocation & "\" & AppValue & "\", , , vbTextCompare)
                If PathExists(Target, True) Then
                    CreateShortCut DestFile, Target, GetFileTitle(DestFile), ItemArgs
                Else
                    CreateShortCut DestFile, InstallLocation & "\" & AppValue & "\" & ItemLink, GetFileTitle(DestFile), ItemArgs
                End If
                pRollback.Add "pf+- " & RemoveArg(pShortCuts(1), " ")

            Else
                InfoText = "Skip existing file: " & DestFile
                pRollback.Add "pf-= " & RemoveArg(pShortCuts(1), " ")
            End If
            CurrentItem = CurrentItem + 1
        Case Else
            InfoText = "Unknown directive: """ & ItemData & """, ignoring..."
    End Select

    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function


Private Function InstallProgramFiles(ByVal ItemData As String, ByRef InfoText As String, ByVal colIndex As Long) As Boolean
    On Error GoTo failitem

    Dim ItemType As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    
    Dim DestFile As String
    DestFile = InstallLocation(ItemType) & "\" & Program.AppValue & "\" & ItemName
    
    If Not PathExists(GetFilePath(DestFile), False) Then
        CabXMLInfo = "<FILE><FULLNAME>" & ItemName & "</FULLNAME><NAME>" & GetFileName(ItemName) & "</NAME><DATE>" & ItemDate & "</DATE><SIZE>" & ItemSize & "</SIZE><PATH>" & GetFilePath(ItemName) & "</PATH></FILE>" & CabXMLInfo
        Do
            
            If Replace(DestFile, InstallLocation(ItemType) & "\", "", , , vbTextCompare) <> "" Then
                DestFile = GetFilePath(DestFile)
                ItemName = Replace(DestFile, InstallLocation(ItemType) & "\", "", , , vbTextCompare)


                pProgramFiles.Add "* " & ItemName & "|<DIR>|" & Now, , colIndex
            End If
        Loop Until PathExists(GetFilePath(DestFile), False) Or Replace(DestFile, InstallLocation(ItemType) & "\", "", , , vbTextCompare) = ""
                ItemType = "*"
                ItemSize = "<DIR>"
                ItemDate = Now
    Else
'        If Not CurDir = InstallLocation(ItemType) & "\" & Program.AppValue Then
'            ChDir InstallLocation(ItemType) & "\" & Program.AppValue
'        End If
        If Not CurDir = GetFilePath(DestFile) Then 'InstallLocation(ItemType) & "\" & Program.AppValue & "\" & GetFilePath(ItemName) Then
            ChDir GetFilePath(DestFile) 'InstallLocation(ItemType) & "\" & Program.AppValue & "\" & GetFilePath(ItemName)
        End If
    End If
    
    Select Case ItemType

        Case "*" 'folder
            pRollback.Add "fe-= " & ItemName
            If Not PathExists(DestFile, False) Then
                InfoText = "Created folder: " & DestFile
                MkDir DestFile
                pRollback.Add "cf+- " & ItemName & "|" & ItemSize & "|" & ItemDate
            Else
                InfoText = "Folder exists: " & DestFile
                pRollback.Add "cf+= " & ItemName & "|" & ItemSize & "|" & ItemDate
            End If
            CurrentItem = CurrentItem + 1
        Case "?", "-?"  'file
            If Not PathExists(DestFile, True) Then
                InfoText = "Installed file: " & DestFile
                Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                pRollback.Add "pf+- " & ItemName & "|" & ItemSize & "|" & ItemDate
            ElseIf ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                InfoText = "Overwrote file: " & DestFile
                Kill DestFile
                Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                pRollback.Add "pf+= " & ItemName & "|" & ItemSize & "|" & ItemDate
            Else
                InfoText = "Skip existing file: " & DestFile
                pRollback.Add "pf-= " & ItemName & "|" & ItemSize & "|" & ItemDate
            End If
            If InStr(LCase(DestFile), ".exe") > 0 Then
                If Not Registry.KeyExists(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers") Then
                    Registry.CreateKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
                End If
                If Not Registry.KeyExists(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers") Then
                    Registry.CreateKey HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
                End If
                Registry.SetValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile, "~ RUNASADMIN"
                Registry.SetValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile, "~ RUNASADMIN"
            End If
            CurrentItem = CurrentItem + 1
        Case "!", "-!"  'current
            If Not PathExists(DestFile, True) Then
                InfoText = "Installed file: " & DestFile
                Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                pRollback.Add "uf+- " & ItemName & "|" & ItemSize & "|" & ItemDate
            ElseIf ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                InfoText = "Skip current user file: " & DestFile
                pRollback.Add "uf-= " & ItemName & "|" & ItemSize & "|" & ItemDate
            Else
                InfoText = "Skip existing file: " & DestFile
                pRollback.Add "uf-= " & ItemName & "|" & ItemSize & "|" & ItemDate
            End If
            CurrentItem = CurrentItem + 1
        Case "$", "-$"  'alluser
            If Not PathExists(DestFile, True) Then
                InfoText = "Installed file: " & DestFile
                Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                pRollback.Add "af+- " & ItemName & "|" & ItemSize & "|" & ItemDate
            ElseIf ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                InfoText = "Overwrote file: " & DestFile
                Kill DestFile
                Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                pRollback.Add "af+= " & ItemName & "|" & ItemSize & "|" & ItemDate
            Else
                InfoText = "Skip existing file: " & DestFile
                pRollback.Add "af-= " & ItemName & "|" & ItemSize & "|" & ItemDate
            End If
            CurrentItem = CurrentItem + 1
        'Case "-?", "-!", "-$"
        Case Else
            InfoText = "Unknown directive: """ & ItemData & """, ignoring..."
    End Select
    If IsNumeric(ItemSize) Then ByteProgress = ByteProgress + CDbl(ItemSize)
    

    Exit Function
failitem:
Debug.Print "Error " & Err.Description & ": " & InfoText
    InfoText = "Error " & Err.Description & ": " & InfoText
    Err.Clear
    Resume Next
End Function


Private Function InstallSystem32(ByVal ItemData As String, ByRef InfoText As String, ByVal colIndex As Long) As Boolean
    On Error GoTo failitem

    Dim ItemType As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemVers As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemVers = RemoveNextArg(ItemData, "|")

    Dim DestFile As String
    DestFile = System32Location(ItemName)
    
    Dim skipInstead As Boolean
    Dim newerDate As Boolean
    If Not PathExists(GetCommonFilesFolder & "\" & Program.Author, False) Then
        MkDir GetCommonFilesFolder & "\" & Program.Author
    End If
                
    Select Case ItemType
        Case "%" 'virtualmachine
            If PathExists(DestFile, True) Then
                newerDate = (DateDiff("n", ItemDate, GetFileDate(DestFile)) < 0)
            Else
                newerDate = True
            End If
            skipInstead = True
            If IncrementLibrary(DestFile, ItemVers) Or newerDate Then
                If Not PathExists(DestFile, True) Then
                    InfoText = "Installed shared file: " & DestFile
                    RebootFlag = RebootFlag Or MoveLibrary(AppPath & ItemName, System32Location(ItemName), True, GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName & ".bak")
                    pRollback.Add "rf+- " & RemoveArg(pSystem32(colIndex), " ")
                    skipInstead = False
                ElseIf (ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile)) Or newerDate Then
                    If CheckVersionGreater(ItemVers, DestFile) Or newerDate Then
                        InfoText = "Overwrote file: " & DestFile
                        RebootFlag = RebootFlag Or MoveLibrary(AppPath & ItemName, System32Location(ItemName), True, GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName & ".bak")
                        pRollback.Add "rf+= " & RemoveArg(pSystem32(colIndex), " ")
                        skipInstead = False
                    End If
                End If
            End If
            If skipInstead Then
                InfoText = "Skip existing shared file: " & DestFile
                pRollback.Add "rf-= " & RemoveArg(pSystem32(colIndex), " ")
                
            End If
            If PathExists(DestFile, True) Then
                pSystem32.Add "^ " & RemoveArg(pSystem32(colIndex), " ")
            End If
            CurrentItem = CurrentItem + 1
        Case "^" 'register
            InfoText = "Registered shared file: " & DestFile
            If GetFileExt(DestFile, True, True) = "exe" Then
                RunProcess DestFile, "/regserver", vbHide, True
            Else
                RunProcess GetSystem32Folder & "regsvr32.exe", "/s """ & DestFile & """", vbHide, True
            End If
            pRollback.Add "rf+= " & RemoveArg(pSystem32(colIndex), " ")
            CurrentItem = CurrentItem + 1
        Case "*" 'shared
            If PathExists(DestFile, True) Then
                newerDate = (DateDiff("n", ItemDate, GetFileDate(DestFile)) < 0)
            Else
                newerDate = True
            End If

            skipInstead = True
            If IncrementLibrary(DestFile, ItemVers) Or newerDate Then
                If Not PathExists(DestFile, True) Then
                    InfoText = "Installed shared file: " & DestFile
                    If CurDir <> GetFilePath(DestFile) Then ChDir GetFilePath(DestFile)
                    Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                    pRollback.Add "rf+- " & RemoveArg(pSystem32(colIndex), " ")
                    skipInstead = False
                ElseIf (ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile)) Or newerDate Then
                    If CheckVersionGreater(ItemVers, DestFile) Or newerDate Then
                        InfoText = "Overwrote file: " & DestFile
                        If CurDir <> GetCommonFilesFolder & "\" & Me.Author Then ChDir GetCommonFilesFolder & "\" & Me.Author
        
                        Call CabFile.Extract(ItemName, GetCommonFilesFolder & "\" & Me.Author, GetFileName(ItemName))
                        RebootFlag = RebootFlag Or UpdateLibrary(ItemName, True)
                        pRollback.Add "rf+= " & RemoveArg(pSystem32(colIndex), " ")
                        skipInstead = False

                    End If
                End If
            End If
            If skipInstead Then
                InfoText = "Skip existing shared file: " & DestFile
                pRollback.Add "rf-= " & RemoveArg(pSystem32(colIndex), " ")
            End If
            
            If PathExists(DestFile, True) Then
                pSystem32.Add "^ " & RemoveArg(pSystem32(colIndex), " ")
            End If

            CurrentItem = CurrentItem + 1
        Case "?" 'system
            If PathExists(DestFile, True) Then
                newerDate = (DateDiff("n", ItemDate, GetFileDate(DestFile)) < 0)
            Else
                newerDate = True
            End If

            skipInstead = True
            If IncrementLibrary(DestFile, ItemVers) Or newerDate Then
                If Not PathExists(DestFile, True) Then
                    InfoText = "Installed system file: " & DestFile
                    If CurDir <> GetFilePath(DestFile) Then ChDir GetFilePath(DestFile)
                    Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                    pRollback.Add "sf+- " & RemoveArg(pSystem32(colIndex), " ")
                    skipInstead = False
                ElseIf (ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile)) Or newerDate Then
                    If CheckVersionGreater(ItemVers, DestFile) Or newerDate Then
                        InfoText = "Overwrote file: " & DestFile
                        If CurDir <> GetCommonFilesFolder & "\" & Me.Author Then ChDir GetCommonFilesFolder & "\" & Me.Author
                        Call CabFile.Extract(ItemName, GetCommonFilesFolder & "\" & Me.Author, GetFileName(ItemName))
                        RebootFlag = RebootFlag Or UpdateLibrary(ItemName, False)
                        pRollback.Add "sf+= " & RemoveArg(pSystem32(colIndex), " ")
                        skipInstead = False
                    End If
                End If
            End If
            If skipInstead Then
                InfoText = "Skip existing system file: " & DestFile
                pRollback.Add "sf-= " & RemoveArg(pSystem32(colIndex), " ")
            End If
            CurrentItem = CurrentItem + 1
        Case "!" 'normal
            If PathExists(DestFile, True) Then
                newerDate = (DateDiff("n", ItemDate, GetFileDate(DestFile)) < 0)
            Else
                newerDate = True
            End If

            skipInstead = True
            If IncrementLibrary(DestFile, ItemVers) Or newerDate Then
                If Not PathExists(DestFile, True) Then
                    InfoText = "Installed library file: " & DestFile
                    
                    If CurDir <> GetFilePath(DestFile) Then ChDir GetFilePath(DestFile)
                    Call CabFile.Extract(ItemName, GetFilePath(DestFile), GetFileName(ItemName))
                    pRollback.Add "sf+- " & RemoveArg(pSystem32(colIndex), " ")
                    skipInstead = False
                ElseIf (ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile)) Or newerDate Then
                    If CheckVersionGreater(ItemVers, DestFile) Or newerDate Then
                        InfoText = "Overwrote file: " & DestFile
                        If CurDir <> GetCommonFilesFolder & "\" & Me.Author Then ChDir GetCommonFilesFolder & "\" & Me.Author
                        Call CabFile.Extract(ItemName, GetCommonFilesFolder & "\" & Me.Author, GetFileName(ItemName))
                        RebootFlag = RebootFlag Or UpdateLibrary(ItemName, False)
                        pRollback.Add "sf+= " & RemoveArg(pSystem32(colIndex), " ")
                        skipInstead = False
                    End If
                End If
            End If
            If skipInstead Then
                InfoText = "Skip existing library file: " & DestFile
                pRollback.Add "sf-= " & RemoveArg(pSystem32(colIndex), " ")
            End If
            CurrentItem = CurrentItem + 1
        Case Else
            InfoText = "Unknown directive """ & ItemData & """, ignoring..."
    End Select
    If IsNumeric(ItemSize) Then ByteProgress = ByteProgress + CDbl(ItemSize)


    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function

Private Function RollbackEntry(ByVal ItemData As String, ByRef InfoText As String) As Boolean
    On Error GoTo failitem
    
    Dim rbFeat As String
    rbFeat = RemoveNextArg(ItemData, " ")
    Dim fileExisted As Boolean
    Dim fileSkipped As Boolean
    fileExisted = (Mid(ItemData, 4) = "=")
    fileSkipped = (Mid(ItemData, 3) = "-")
    
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemVers As String
    ItemName = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemVers = RemoveNextArg(ItemData, "|")
    Dim RemoveToo As Boolean
    Dim DestFile As String
    
    Select Case Left(rbFeat, 2)
        Case "rf" 'systemfiles
            DestFile = System32Location(ItemName)
            
            If DecrementLibrary(DestFile, RemoveToo) Then
                If RemoveToo Then
                    If PathExists(DestFile, True) Then
                        InfoText = "Removed shared file: " & ItemName
                        Kill DestFile
                    Else
                        InfoText = "File non existant: " & ItemName
                    End If
                ElseIf MakeRestoreLibrary(ItemName, DestFile, True) Then
                    InfoText = "Restored shared file: " & ItemName
                ElseIf PathExists(DestFile, True) Then
                    InfoText = "Keeping library file: " & ItemName
                Else
                    InfoText = "File non existant: " & ItemName
                End If
            ElseIf PathExists(DestFile, True) Then
                InfoText = "Keeping shared file: " & ItemName
            Else
                InfoText = "File non existant: " & ItemName
            End If
            
            CurrentItem = CurrentItem - 1
            
        Case "sf" 'systemfiles
            DestFile = System32Location(ItemName)
            If DecrementLibrary(DestFile, RemoveToo) Then
                If RemoveToo Then
                    If PathExists(DestFile, True) Then
                        InfoText = "Removed system file: " & ItemName
                        Kill DestFile
                    Else
                        InfoText = "File non existant: " & ItemName
                    End If
                ElseIf MakeRestoreLibrary(ItemName, DestFile, False) Then
                    InfoText = "Restored system file: " & ItemName
                ElseIf PathExists(DestFile, True) Then
                    InfoText = "Keeping library file: " & ItemName
                Else
                    InfoText = "File non existant: " & ItemName
                End If
            ElseIf PathExists(DestFile, True) Then
                InfoText = "Keeping system file: " & ItemName
            Else
                InfoText = "File non existant: " & ItemName
            End If
            
            CurrentItem = CurrentItem - 1
            
        Case "pf" 'programfiles
            DestFile = InstallLocation & "\" & Program.AppValue & "\" & ItemName
            
            InfoText = "Removed: " & DestFile
            If PathExists(DestFile, True) Then
                Kill DestFile
            End If

            CurrentItem = CurrentItem - 1
            
        Case "cf" 'createdfolder
            DestFile = InstallLocation & "\" & Program.AppValue & "\" & ItemName
            
            InfoText = "Deleted: " & DestFile

            RemovePath DestFile
            
            CurrentItem = CurrentItem - 1
        Case "fe" 'changedfolder
            DestFile = GetFilePath(InstallLocation & "\" & Program.AppValue & "\" & ItemName)
            If PathExists(DestFile, False) Then
            
            
                InfoText = "Changing: " & DestFile
                ChDir DestFile
            End If
            
            CurrentItem = CurrentItem - 1
        Case "uf" 'current user file
        
            DestFile = InstallLocation("!") & "\" & Program.AppValue & "\" & ItemName
            
            If ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                InfoText = "Keeping: " & DestFile
            Else
                InfoText = "Removed: " & ItemName
                If PathExists(DestFile, True) Then
                    Kill DestFile
                End If
            End If
            
            CurrentItem = CurrentItem - 1
        Case "af" 'current user file
        
            DestFile = InstallLocation("$") & "\" & Program.AppValue & "\" & ItemName
            
            If ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                InfoText = "Keeping: " & DestFile
            Else
                InfoText = "Removed: " & ItemName
                If PathExists(DestFile, True) Then
                    Kill DestFile
                End If
            End If
            
            CurrentItem = CurrentItem - 1
        Case "rf" 'register file
            DestFile = System32Location(ItemName)
            InfoText = "Unregistering shared file: " & DestFile
            If GetFileExt(DestFile, True, True) = "exe" Then
                RunProcess DestFile, "/unregserver", vbHide, True
            Else
                RunProcess GetSystem32Folder & "regsvr32.exe", "/s /u """ & DestFile & """", vbHide, True
            End If
        Case "ft"
            Registry.DeleteKey HKEY_CLASSES_ROOT, ItemName
            Registry.DeleteKey HKEY_CLASSES_ROOT, ItemSize
            InfoText = "Uninstalled file type: " & ItemSize
            CurrentItem = CurrentItem - 1
        Case Else
            InfoText = "Unknown directive: """ & ItemData & """, ignoring..."
            CurrentItem = CurrentItem + 1
    End Select

    If IsNumeric(ItemSize) Then ByteProgress = ByteProgress - CDbl(ItemSize)
    

    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function

Private Function UninstallFileType(ByVal ItemData As String, ByRef InfoText As String) As Boolean
    On Error GoTo failitem

    Dim ItemType As String
    Dim ItemClass As String
    Dim ItemDesc As String
    Dim ItemAction As String
    Dim ItemProgam As String
    Dim ItemParams As String
    ItemType = RemoveNextArg(ItemData, "|")
    ItemClass = RemoveNextArg(ItemData, "|")
    ItemDesc = RemoveNextArg(ItemData, "|")
    
   Registry.DeleteKey HKEY_CLASSES_ROOT, ItemType
   Registry.DeleteKey HKEY_CLASSES_ROOT, ItemClass
   Registry.DeleteKey HKEY_CLASSES_ROOT, ItemDesc
    
    InfoText = "Uninstalled file type: " & ItemDesc

    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function

Private Function UninstallShortCut(ByVal ItemData As String, ByRef InfoText As String) As Boolean
    On Error GoTo failitem
    
    Dim ItemType As String
    Dim ItemName As String
    Dim ItemLink As String
    Dim ItemArgs As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemLink = RemoveNextArg(ItemData, "|")
    ItemArgs = RemoveNextArg(ItemData, "|")
    Dim DestFile As String

    Select Case ItemType
        Case "^"
            If pShortCuts.count > 1 Then
                If InStr(1, pShortCuts(2), ItemName, vbTextCompare) > 0 Then
                    pShortCuts.Add pShortCuts(1)
                Else
                    If PathExists(ItemName, False) Then
                        InfoText = "Removed folder: " & ItemName
                        RmDir ItemName
                    Else
                        InfoText = "Folder non existant: " & ItemName
                    End If
                End If
            Else
                If PathExists(ItemName, False) Then
                    InfoText = "Removed folder: " & ItemName
                    RmDir ItemName
                Else
                    InfoText = "Folder non existant: " & ItemName
                End If
                
            End If
            CurrentItem = CurrentItem + 1
        Case "*" 'folder
            pShortCuts.Add "^ " & GetStartMenuProgramsFolder & "\" & ItemName
            CurrentItem = CurrentItem + 1
        Case "?" 'file
            DestFile = GetStartMenuProgramsFolder & "\" & ItemName & ".lnk"
            
            If PathExists(DestFile, True) Then
                InfoText = "Removed program file: " & DestFile
                Kill DestFile
            Else
                InfoText = "File non existant: " & DestFile
            End If
            
            If GetFileExt(ItemName, True, True) = "exe" Then
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", Replace(DestFile, "\", "\\")
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", StrConv(Replace(DestFile, "\", "\\"), vbUnicode)
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", Replace(DestFile, "\", "\\")
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", StrConv(Replace(DestFile, "\", "\\"), vbUnicode)
            End If
            CurrentItem = CurrentItem + 1
        Case Else
            InfoText = "Unknown directive: '" & ItemType & "', ignoring..."
            CurrentItem = CurrentItem + 1
    End Select

    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText

End Function


Private Function UninstallProgramFiles(ByVal ItemData As String, ByRef InfoText As String, ByVal colIndex As Long) As Boolean
    On Error GoTo failitem
    
    Dim ItemType As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemVers As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemVers = RemoveNextArg(ItemData, "|")
    Dim DestFile As String

    Select Case ItemType
        Case "^"
            If pProgramFiles.count > 1 Then
                If InStr(1, pProgramFiles(2), ItemName, vbTextCompare) > 0 Then
                    pProgramFiles.Add pProgramFiles(colIndex)
                Else
                    If PathExists(ItemName, False) Then
                        InfoText = "Removed folder: " & ItemName
                        RmDir ItemName
                    Else
                        InfoText = "Folder non existant: " & ItemName
                    End If
                End If
            Else
                If PathExists(ItemName, False) Then
                    InfoText = "Removed folder: " & ItemName
                    RmDir ItemName
                Else
                    InfoText = "Folder non existant: " & ItemName
                End If
                
            End If
            CurrentItem = CurrentItem + 1
        Case "*" 'folder
            pProgramFiles.Add "^ " & InstallLocation & "\" & Program.AppValue & "\" & ItemName
            CurrentItem = CurrentItem + 1
        Case "?" 'file
            DestFile = InstallLocation & "\" & Program.AppValue & "\" & ItemName
            
            If PathExists(DestFile, True) Then
                InfoText = "Removed program file: " & DestFile
                Kill DestFile
            Else
                InfoText = "File non existant: " & DestFile
            End If
            If GetFileExt(ItemName, True, True) = "exe" Then
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", Replace(DestFile, "\", "\\")
                Registry.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", StrConv(Replace(DestFile, "\", "\\"), vbUnicode)
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", DestFile
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", Replace(DestFile, "\", "\\")
                Registry.DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers", StrConv(Replace(DestFile, "\", "\\"), vbUnicode)
            End If
            CurrentItem = CurrentItem + 1
        Case "!" 'current
        
            DestFile = InstallLocation(ItemType) & "\" & Program.AppValue & "\" & ItemName
            If PathExists(DestFile, True) Then
                If ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                    InfoText = "Keeping user file: " & DestFile
                Else
                    InfoText = "Removed program file: " & ItemName
                    If PathExists(DestFile, True) Then
                        Kill DestFile
                    End If
                End If
            Else
                InfoText = "File non existant: " & ItemName
            End If
            
            CurrentItem = CurrentItem + 1
        Case "$", "-$" 'alluser
        
            DestFile = InstallLocation(ItemType) & "\" & Program.AppValue & "\" & ItemName
            If PathExists(DestFile, True) Then
                If ItemSize <> FileLen(DestFile) Or ItemDate <> GetFileDate(DestFile) Then
                    InfoText = "Keeping user file: " & DestFile
                Else
                    InfoText = "Removed program file: " & ItemName
                    If PathExists(DestFile, True) Then
                        Kill DestFile
                    End If
                End If
            Else
                InfoText = "File non existant: " & ItemName
            End If
            
            CurrentItem = CurrentItem + 1
        Case "-*", "-?", "-!", "-$"
        
        Case Else
            InfoText = "Unknown directive: '" & ItemType & "', ignoring..."
            CurrentItem = CurrentItem + 1
    End Select

    If IsNumeric(ItemSize) Then ByteProgress = ByteProgress + CDbl(ItemSize)
    
    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText

End Function



Private Function UninstallSystem32(ByVal ItemData As String, ByRef InfoText As String, ByVal colIndex As Long) As Boolean
    On Error GoTo failitem

    Dim ItemType As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemVers As String
    ItemType = RemoveNextArg(ItemData, " ")
    ItemName = RemoveNextArg(ItemData, "|")
    ItemSize = RemoveNextArg(ItemData, "|")
    ItemDate = RemoveNextArg(ItemData, "|")
    ItemVers = RemoveNextArg(ItemData, "|")

    Dim skipLib As Boolean
    Dim RemoveToo As Boolean
    Dim DestFile As String

    DestFile = System32Location(ItemName)

    If Not LCase(CurDir) = GetFilePath(DestFile) Then ChDir GetFilePath(DestFile)

    Select Case ItemType
        Case "*" 'shared
            If DecrementLibrary(DestFile, RemoveToo) Then
                        
                If RemoveToo Then
                    If PathExists(DestFile, True) Then
                        pSystem32.Add "^ " & ItemName & "|" & RemoveToo
                    Else
                        skipLib = True
                    End If
                ElseIf MakeRestoreLibrary(ItemName, DestFile, True) Then
                    pSystem32.Add "^ " & ItemName & "|" & RemoveToo
                ElseIf PathExists(DestFile, True) Then
                    skipLib = True
                    pSystem32.Add "^ " & ItemName & "|" & RemoveToo
                Else
                    skipLib = True
                    pSystem32.Add "^ " & ItemName & "|" & RemoveToo
                End If
                
                If Not skipLib Then
                    InfoText = "Unregistering shared file: " & DestFile
                    If GetFileExt(DestFile, True, True) = "exe" Then
                        RunProcess DestFile, "/unregserver", vbHide, True
                    Else
                        RunProcess GetSystem32Folder & "regsvr32.exe", "/s /u """ & DestFile & """", vbHide, True
                    End If
                End If
                
            ElseIf PathExists(DestFile, True) Then
                InfoText = "Keeping shared file: " & ItemName
            Else
                InfoText = "File non existant: " & ItemName
            End If
            
            CurrentItem = CurrentItem + 1
        Case "^"
            If CBool(ItemSize) Then
                If PathExists(DestFile, True) Then
                    InfoText = "Removed shared file: " & ItemName
                    Kill DestFile
                Else
                    InfoText = "File non existant: " & ItemName
                End If
            ElseIf MakeRestoreLibrary(ItemName, DestFile, True) Then
                InfoText = "Restored shared file: " & ItemName
                RunProcess GetSystem32Folder & "regsvr32.exe", "/s """ & DestFile & """", vbHide, True
            ElseIf PathExists(DestFile, True) Then
                InfoText = "Keeping library file: " & ItemName
            Else
                InfoText = "File non existant: " & ItemName
            End If
            CurrentItem = CurrentItem + 1
        Case "?", "!" 'normal 'system
            If DecrementLibrary(DestFile, RemoveToo) Then
                If RemoveToo Then
                    If PathExists(DestFile, True) Then
                        InfoText = "Removed shared file: " & ItemName
                        Kill DestFile
                    Else
                        InfoText = "File non existant: " & ItemName
                    End If
                ElseIf MakeRestoreLibrary(ItemName, DestFile, False) Then
                    InfoText = "Restored shared file: " & ItemName
                ElseIf PathExists(DestFile, True) Then
                    InfoText = "Keeping library file: " & ItemName
                Else
                    InfoText = "File non existant: " & ItemName
                End If
            ElseIf PathExists(DestFile, True) Then
                InfoText = "Keeping shared file: " & ItemName
            Else
                InfoText = "File non existant: " & ItemName
            End If

            CurrentItem = CurrentItem + 1
        Case "&" 'left over backup as last uninstaller
        
            MakeRestoreLibrary ItemName, DestFile, True

            If PathExists(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak", True) And PathExists(System32Location(ItemName), True) Then
                If CheckVersionEqual(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak", DestFile) Then
                    Kill GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak"
                End If
            End If

            InfoText = "Restored shared file: " & ItemName
            CurrentItem = CurrentItem + 1

        Case Else
            InfoText = "Unknown directive: """ & ItemData & """, ignoring..."
            CurrentItem = CurrentItem + 1
    End Select

    If IsNumeric(ItemSize) Then ByteProgress = ByteProgress + CDbl(ItemSize)


    Exit Function
failitem:
    InfoText = "Error " & Err.Description & ": " & InfoText
End Function


Public Function UpdateLibrary(ByVal ItemName As String, ByVal activeX As Boolean) As Boolean
    UpdateLibrary = MoveLibrary(GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName, System32Location(ItemName), activeX, GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName & ".bak")
End Function
Public Function RestoreLibrary(ByVal ItemName As String, ByVal activeX As Boolean) As Boolean
    RestoreLibrary = MoveLibrary(GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName & ".bak", System32Location(ItemName), activeX, GetCommonFilesFolder & "\" & Program.Author & "\" & ItemName)
End Function

Private Function MakeRestoreLibrary(ByVal ItemName As String, ByVal DestFile As String, ByVal activeX As Boolean, Optional ByVal Certianty As Boolean = False) As Boolean
    Dim libCnt As String
    libCnt = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\" & Me.Author & "\System", ItemName, "0")
    If (libCnt <> "-1") Or Certianty Then
        If (libCnt = "0") Or Certianty Then
            If PathExists(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak", True) Then
                Dim fvi As FILEPROPERTIE
                fvi = GetFileInfo(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak")
                If Not CheckVersionEqual(fvi.FileVersion, DestFile) Then
                    MakeRestoreLibrary = True
                End If
                RebootFlag = RebootFlag Or RestoreLibrary(ItemName, activeX)

                If PathExists(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak", True) And PathExists(System32Location(ItemName), True) Then
                    Kill GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName & ".bak"
                End If
                If PathExists(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName, True) Then
                    Kill GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName
                End If
                
            ElseIf PathExists(GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName, True) And PathExists(System32Location(ItemName), True) Then
                Kill GetCommonFilesFolder & "\" & Me.Author & "\" & ItemName
            End If
        End If
    End If

End Function

Private Sub CreateShortCut(ByVal FilePath As String, ByVal Target As String, ByVal Description As String, Optional ByVal Arguments As String)

    RunProcess """" & InstallLocation & "\" & AppValue & "\Uninstall.exe""", "/C " & FilePath & "|" & Target & "|" & Arguments & " |", vbHide, True

End Sub

Public Sub Finish()
     On Error Resume Next
     If Not IsDebugger Then
    
        If (Not RebootFlag) And PathExists(AppPath & "Manifest.ini", True) Then
            Kill AppPath & "Manifest.ini"
        ElseIf (Not PathExists(AppPath & "Manifest.ini", True)) And RebootFlag Then
            WriteFile AppPath & "Manifest.ini", ""
        End If
    End If
    If RemoveFlag Then

        If PathExists(AppPath & "Remove.exe", True) Then RunProcess AppPath & "Remove.exe", InstallLocation & "\" & AppValue & "\Uninstall.exe", vbHide, False

    End If
End Sub

