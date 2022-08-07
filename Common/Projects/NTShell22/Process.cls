VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private mCol As Collection

Public Function IsRunning(ByVal EXENameOrPID As Variant) As Variant
    IsRunning = ProcessRunning(EXENameOrPID)
End Function

Public Function Exec(ByVal File As String, Optional ByVal Params As String = "", Optional ByVal Focus As Long = 1) As Long
    Exec = modProcess.RunFile(File, Params, Focus)
End Function

Public Function Run(ByVal Path As String, Optional ByVal Params As String = "", Optional ByVal Focus As Long = 1, Optional ByVal WaitLatency As Single = False) As Long
    Run = modProcess.RunProcess(Path, Params, Focus, WaitLatency)
End Function

Public Function OpenProgram(ByVal pPath As String) As Long
    Dim lSuccess As Long
    Dim sNull As String
    Dim pInfo As PROCESS_INFORMATION
    Dim sInfo As STARTUPINFO
    
    sInfo.cb = Len(sInfo)
    lSuccess = CreateProcess(sNull, _
                                 pPath, _
                                 ByVal 0&, _
                                 ByVal 0&, _
                                 0&, _
                                 &H20, _
                                 ByVal 0&, _
                                 sNull, _
                                 sInfo, _
                                 pInfo)
    
    Dim pp As New Program
    
    pp.Process = pInfo.hProcess
    pp.ProcessId = pInfo.dwProcessId
    pp.Thread = pInfo.hThread
    pp.ThreadId = pInfo.dwThreadId
    pp.Path = pPath
    
    If lSuccess Then mCol.Add pp, pPath
    
    Set pp = Nothing
    OpenProgram = lSuccess
End Function
    
Public Function CloseProgram(ByVal pPath As String) As Long
    Dim lSuccess As Long
    Dim pp As Program
    Dim pInfo As PROCESS_INFORMATION
    Set pp = mCol.Item(pPath)
    
    pInfo.hProcess = pp.Process
    pInfo.dwProcessId = pp.ProcessId
    pInfo.hThread = pp.Thread
    pInfo.dwThreadId = pp.ThreadId
    
    lSuccess = CloseProgram2(pInfo)
    mCol.Remove pPath
    
    Set pp = Nothing
    
    CloseProgram = lSuccess
End Function

Private Function CloseProgram2(ByRef pInfo As PROCESS_INFORMATION) As Long
    Dim lSuccess As Long
     
    lSuccess = TerminateProcess(pInfo.hProcess, 0&)
    lSuccess = CloseHandle(pInfo.hThread)
    lSuccess = CloseHandle(pInfo.hProcess)

    CloseProgram2 = lSuccess
End Function

Public Property Get ProgramCount() As Long
    ProgramCount = mCol.count
End Property

Public Sub CloseAllPrograms()
    Dim pp As Program
    Dim pInfo As PROCESS_INFORMATION
    Do While mCol.count > 0
        Set pp = mCol(1)
        
        pInfo.hProcess = pp.Process
        pInfo.dwProcessId = pp.ProcessId
        pInfo.hThread = pp.Thread
        pInfo.dwThreadId = pp.ThreadId
    
        CloseProgram2 pInfo
        mCol.Remove 1
        
        Set pp = Nothing
    Loop
End Sub

Private Sub Class_Initialize()

    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    CloseAllPrograms
    Set mCol = Nothing
End Sub
