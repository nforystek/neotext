Attribute VB_Name = "modProcess"
#Const modProcess = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type


Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

#If Not modWindow Then

    Private Const GW_HWNDFIRST = 0
    Private Const GW_HWNDLAST = 1
    Private Const GW_HWNDNEXT = 2
    Private Const GW_HWNDPREV = 3
    Private Const GW_OWNER = 4
    Private Const GW_CHILD = 5
    Private Const GW_MAX = 5

    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

#End If

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

#If Not modShared Then

    Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
    
#End If

#If Not modCommon Then

    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    
    Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function RegGetValueEx Lib "advapi32" Alias "RegGetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Public Function IsWindows98() As Boolean
    Static isWin98 As Byte
    If isWin98 > 0 Then
        IsWindows98 = IIf(isWin98 = 1, True, False)
    Else
        Dim i As Long
        isWin98 = 2
        IsWindows98 = False
'
'        If GetProcAddress(GetModuleHandle("kernel32"), "RegisterServiceProcess") <> 0 Then
            On Error Resume Next
            i = RegisterServiceProcess(GetCurrentProcessId, 0)
        
            If Not (Err = 453) Then
                isWin98 = 1
                IsWindows98 = True
            Else
                Err.Clear
            End If
            On Error GoTo 0
      '  End If
    End If
End Function

#If Not modWindow Then

    Public Function WindowText(ByVal hwnd As Long) As String
        Dim sBuffer As String
        Dim lSize As Long
        sBuffer = Space$(255)
        lSize = Len(sBuffer)
        Call GetWindowText(hwnd, sBuffer, lSize)
        If lSize > 0 Then
            WindowText = Trim(Replace(Left$(sBuffer, lSize), Chr(0), ""))
        End If
    End Function
    
#End If

Public Function IsFileExecutable(ByVal FileName As String) As Boolean
    Select Case GetFileExt(FileName)
        Case ".exe", ".bat", ".com" ', ".msi"
            IsFileExecutable = True
        Case Else
            IsFileExecutable = False
    End Select
End Function

Public Function OpenWebsite(ByVal WebSite As String, Optional ByVal Silent As Boolean) As Boolean
    OpenWebsite = (RunFile(WebSite) <> 0)
End Function
Public Function RunFile(ByVal File As String, Optional ByVal Params As String = "", Optional ByVal FocusPID As Long = 1) As Long

#If modCommon Then
    
    If Not PathExists(File, True) Then
        RunFile = ShellExecute(0, "open", File, Params, 0&, FocusPID)
    Else
        RunFile = ShellExecute(0, "open", GetFileName(File), Params, GetFilePath(File), FocusPID)
    End If
#Else
    RunFile = ShellExecute(0, "open", File, Params, 0&, FocusPID)
#End If
End Function

Public Function RunProcess(ByVal path As String, Optional ByVal Params As String = "", Optional ByVal Focus As Integer = vbNormalFocus, Optional ByVal Wait As Boolean = False) As Long
    If Wait Then
        Dim PID As Double
        PID = Shell(Trim(Trim(path) & " " & Trim(Params)), Focus)
        Do While ProcessRunning(PID)
            modCommon.DoTasks
        Loop
        RunProcess = -CInt((PID > 0))
    Else
        RunProcess = Shell(Trim(path & " " & Params), Focus)
    End If
End Function
Public Function ProcessExeOrPIDBy(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Variant
    On Local Error GoTo catch

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, i - 1))

        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
                
                ProcessExeOrPIDBy = uProcess.th32ProcessID
                
            Exit Do
        ElseIf IsNumeric(EXEorPID) Then
            If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                ProcessExeOrPIDBy = szExename
                Exit Do
            End If
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    Call CloseHandle(hSnapshot)
    Exit Function
catch:
    Err.Clear
End Function

'Public Function RunProcess(ByVal Path As String, Optional ByVal Params As String = "", Optional ByVal Focus As Long = 1, Optional ByVal Wait As Variant = False) As Long
'    Dim RunCount As Long  'get count of process
'    If TypeName(Wait) = "Boolean" And CLng(Wait) = 0 Then
'        RunProcess = Shell(Trim(Trim(Path) & " " & Trim(Params)), Focus)
'    Else
'        If Focus = vbHidden Then Focus = vbHide 'use the correct one
'        'is running by exename to check for exiting
'
'        Focus = Shell(Trim(Trim(Path) & " " & Trim(Params)), Focus)
'        If (Focus > 0) And Wait Then
'            Dim lapse As Single
'            lapse = Timer
'            Do Until (lapse = -1) Or (ProcessRunning(Focus) = 0)
'                If IsNumeric(Wait) And (Not CStr(CInt(Wait)) = "-1") And (Not CStr(CInt(Wait)) = "0") Then
'                    If (Timer - lapse) >= Wait Then
'                        lapse = -1
'                    End If
'                End If
'                DoEvents
'                modCommon.Sleep 1
'            Loop
'        End If
'        RunProcess = ProcessRunning(GetFileName(Replace(Path, """", "")))
'        If RunProcess = 1 Then RunProcess = Focus
'    End If
'End Function



Public Function ProcessRunning(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Long
    On Local Error GoTo catch

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, i - 1))

        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
               
                cnt = cnt + 1
        ElseIf IsNumeric(EXEorPID) Then
            If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                cnt = cnt + 1
            End If
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    Call CloseHandle(hSnapshot)

    ProcessRunning = cnt
    Exit Function
catch:
    Err.Clear
End Function


'Public Function ProcessRunning(ByVal EXENameOrPID As Variant) As Variant
'
'    Dim uProcess As PROCESSENTRY32
'    Dim rProcessFound As Long
'    Dim hSnapshot As Long
'    Dim szExename As String
'    Dim i As Integer
'    Dim cnt As Long
'
'    Const TH32CS_SNAPPROCESS As Long = 2&
'
'    uProcess.dwSize = Len(uProcess)
'    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
'    rProcessFound = ProcessFirst(hSnapshot, uProcess)
'
'    cnt = 0
'
'    Do While rProcessFound
'        i = InStr(1, uProcess.szexeFile, Chr(0))
'        szExename = LCase(Left(uProcess.szexeFile, i - 1))
'        If ((szExename <> "") And (((Right$(szExename, Len(GetFileName(EXENameOrPID))) = LCase$(GetFileName(EXENameOrPID))) Or _
'            (Right$(EXENameOrPID, Len(GetFileName(szExename))) = LCase$(GetFileName(szExename)))) Or _
'            ((Right$(szExename, Len(EXENameOrPID)) = LCase$(EXENameOrPID)) Or _
'            (Right$(EXENameOrPID, Len(szExename)) = LCase$(szExename))))) Or _
'            ((CStr(uProcess.th32ProcessID) = CStr(EXENameOrPID)) And (CStr(EXENameOrPID) <> "0")) Then
'            cnt = cnt + 1
'            If IsNumeric(EXENameOrPID) Then
'                ProcessRunning = -1
'            Else
'                ProcessRunning = uProcess.th32ProcessID
'            End If
'        End If
'
'        rProcessFound = ProcessNext(hSnapshot, uProcess)
'
'    Loop
'
'    Call CloseHandle(hSnapshot)
'
'    If cnt > 0 And Not IsNumeric(ProcessRunning) Then
'        ProcessRunning = cnt
'    ElseIf Not (cnt = 1 And CStr(ProcessRunning) <> "0") Then
'        If ProcessRunning = -1 Then
'            EXENameOrPID = ProcessRunning
'            ProcessRunning = 1
'        Else
'            ProcessRunning = cnt
'        End If
'    ElseIf ProcessRunning = -1 Then
'        ProcessRunning = cnt
'    Else
'        EXENameOrPID = ProcessRunning
'        ProcessRunning = cnt
'    End If
'
'
'End Function

Public Function RunningProcessCount(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Long

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0

    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, i - 1))
        
        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
            
            RunningProcessCount = RunningProcessCount + 1


        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    
    Call CloseHandle(hSnapshot)

End Function

Public Function RunningEXEProcessByID(ByVal EXEPID As Long) As String

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0

    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, i - 1))
        If uProcess.th32ProcessID = EXEPID Then
            RunningEXEProcessByID = szExename
            Exit Do
        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    
    Call CloseHandle(hSnapshot)
   
End Function

Public Function IsProccessIDRunning(ByVal EXEPID As Long) As Boolean

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
'    toggle = Toggler(-Toggler(toggle))
    
    cnt = 0

    Do While rProcessFound

        If (uProcess.th32ProcessID = EXEPID) Then
            IsProccessIDRunning = True
            Exit Do
        End If
        
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    
    Call CloseHandle(hSnapshot)

End Function

Public Function IsProccessEXERunning(ByVal EXE As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim i As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0

    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, i - 1))
        
        If ProcessCheck(szExename, EXE, ExactMatch) Then
                
            IsProccessEXERunning = True

        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    
    Call CloseHandle(hSnapshot)

End Function


Public Function KillApp(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))

        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
                
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
            
        ElseIf IsNumeric(EXEorPID) Then
            If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                KillApp = True
                appCount = appCount + 1
                myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)
            End If
        End If
        
        
'        If (Right$(szExename, Len(myName)) = LCase$(myName) Or Right$(LCase(myName), Len(szExename)) = szExename) Or (Left(szExename, Len(myName)) = LCase$(myName) Or Left(LCase(myName), Len(szExename)) = szExename) Then
'            KillApp = True
'            appCount = appCount + 1
'            myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
'            AppKill = TerminateProcess(myProcess, exitCode)
'            Call CloseHandle(myProcess)
'        ElseIf InStr(szExename, "\") = 0 Then
'
'            If (Right$(szExename, Len(GetFileName(myName))) = LCase$(GetFileName(myName)) Or Right$(LCase(GetFileName(myName)), Len(szExename)) = szExename) Or _
'                (Left(szExename, Len(GetFileName(myName))) = LCase$(GetFileName(myName)) Or Left(LCase(GetFileName(myName)), Len(szExename)) = szExename) Then
'                KillApp = True
'                appCount = appCount + 1
'                myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
'                AppKill = TerminateProcess(myProcess, exitCode)
'                Call CloseHandle(myProcess)
'            End If
'        End If


        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop


    Call CloseHandle(hSnapshot)
Finish:
End Function
Private Function ColExists(ByRef col As Collection, ByVal Val As String) As Boolean
    If col.Count > 0 Then
        Dim i As Long
        For i = 1 To col.Count
            If col(i) = Val Then
                ColExists = True
                Exit Function
            End If
        Next
    End If
End Function
Public Function KillSubApps(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    Dim col As New Collection
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    Dim foundAdd As Boolean
    
    foundAdd = True
    
    Do While foundAdd
        foundAdd = False
        
        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
        Do While rProcessFound
            i = InStr(1, uProcess.szexeFile, Chr(0))
            szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
                If Not ColExists(col, uProcess.th32ProcessID) Then
                    col.Add CStr(uProcess.th32ProcessID), Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")
                    foundAdd = True
                End If
            ElseIf IsNumeric(EXEorPID) Then
                If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                    If Not ColExists(col, uProcess.th32ProcessID) Then
                        col.Add CStr(uProcess.th32ProcessID), Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")
                        foundAdd = True
                    End If
                ElseIf (uProcess.th32ParentProcessID = CLng(EXEorPID)) Or _
                    ColExists(col, uProcess.th32ParentProcessID) Then
                    If Not ColExists(col, uProcess.th32ProcessID) Then
                        col.Add CStr(uProcess.th32ProcessID)
                        foundAdd = True
                    End If
                End If
            ElseIf ColExists(col, uProcess.th32ParentProcessID) Then
                If Not ColExists(col, uProcess.th32ProcessID) Then
                    col.Add CStr(uProcess.th32ProcessID)
                    foundAdd = True
                End If
            End If
    
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        Loop
    
        Call CloseHandle(hSnapshot)
    
    Loop

    If col.Count > 0 Then
        For i = 1 To col.Count
                
            If col(i) <> col(Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")) Then
            
                KillSubApps = True
                appCount = appCount + 1
                myProcess = OpenProcess(1, False, col(i))
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)

            End If
            
        Next
    End If
    
Finish:
End Function


Private Function ProcessCheck(ByVal szExename As String, ByVal EXEorPID As Variant, ByVal ExactMatch As Boolean) As Boolean
    ProcessCheck = ( _
             ( _
               ( _
                 (Right(szExename, Len(EXEorPID)) = LCase(EXEorPID)) Or _
                 (Right(LCase(EXEorPID), Len(szExename)) = szExename) _
                ) _
                Or _
                ( _
                  (Left(szExename, Len(EXEorPID)) = LCase(EXEorPID)) Or _
                  (Left(LCase(EXEorPID), Len(szExename)) = szExename) _
                ) _
              ) _
              And (Not ExactMatch) _
            ) _
           Or _
           ( _
             ( _
                ( _
                  (LCase(szExename) = LCase(EXEorPID)) Or _
                  ((InStr(EXEorPID, "\") = 0 And InStr(EXEorPID, "/") = 0) And (LCase$(GetFileName(szExename)) = LCase(EXEorPID))) Or _
                  ((InStr(szExename, "\") = 0 And InStr(szExename, "/") = 0) And (LCase$(szExename) = LCase(GetFileName(EXEorPID)))) _
                ) _
              ) _
              And ExactMatch _
            )
End Function




