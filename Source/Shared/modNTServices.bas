#Const [True] = -1
#Const [False] = 0




Attribute VB_Name = "modNTServices"
#Const modNTServices = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
   dwProcessId As Long
   dwServiceFlags As Long
End Type

Private Const SC_MANAGER_CONNECT = &H1&
Private Const SERVICE_START = &H10&, SERVICE_STOP = &H20&

Private Enum SERVICE_CONTROL
   SERVICE_CONTROL_STOP = 1
   SERVICE_CONTROL_PAUSE = 2
   SERVICE_CONTROL_CONTINUE = 3
   SERVICE_CONTROL_INTERROGATE = 4
   SERVICE_CONTROL_SHUTDOWN = 5
End Enum

Private Enum SERVICE_STATE
   SERVICE_STOPPED = &H1
   SERVICE_START_PENDING = &H2
   SERVICE_STOP_PENDING = &H3
   SERVICE_RUNNING = &H4
   SERVICE_CONTINUE_PENDING = &H5
   SERVICE_PAUSE_PENDING = &H6
   SERVICE_PAUSED = &H7
End Enum

Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function OpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function StartService Lib "advapi32" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function ControlService Lib "advapi32" (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, lpServiceStatus As SERVICE_STATUS) As Long

Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type

Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias  As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private liOldIdleTime As LARGE_INTEGER
Private liOldSystemTime As LARGE_INTEGER

#If Not modProcess Then
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long


Public Function IsWindows98() As Boolean
'    Static isWin98 As Byte
'    If isWin98 > 0 Then
'        IsWindows98 = IIf(isWin98 = 1, True, False)
'    Else
        Dim i As Long
'        isWin98 = 2
'        IsWindows98 = False
'
'        If GetProcAddress(GetModuleHandle("kernel32"), "RegisterServiceProcess") <> 0 Then
            On Error Resume Next
            i = RegisterServiceProcess(GetCurrentProcessId, 0)
        
            If Not (Err = 453) Then
'                isWin98 = 1
                IsWindows98 = True
            Else
                Err.Clear
            End If
            On Error GoTo 0
'        End If
'    End If
End Function

#End If

Public Sub NetStart(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(GetSettingString(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 0) Then
            SaveSettingString &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 1
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
                If ProcessRunning(EXEName) = 0 Then
                    StartNTService ServiceName
                    If ProcessRunning(EXEName) = 0 Then
                        If PathExists(AppPath & EXEName, True) Then
                            RunProcess AppPath & EXEName
                        Else
                            RunFile EXEName
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub NetStop(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(GetSettingString(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
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
                If ProcessRunning(EXEName) = 1 Then
                    StopNTService ServiceName
                    If ProcessRunning(EXEName) = 1 Then
                        KillApp EXEName
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub NetContinue(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
    Dim timOut As String
    If IsWindows98 Then
        Dim exePath As String
        exePath = Replace(GetSettingString(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
        If EXEName = "" Then EXEName = exePath
        EXEName = GetFileName(EXEName)
        If Not (EXEName = "") And (ProcessRunning(EXEName) = 0) Then
            SaveSettingString &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 1
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
        exePath = Replace(GetSettingString(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, ""), """", "")
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

Public Function StartNTService(ServiceName As String) As Long

   ' This function starts service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_START)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If StartService(hService, 0, 0) = 0 Then
            StartNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         StartNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      StartNTService = Err.LastDllError
   End If

End Function

Public Function StopNTService(ServiceName As String) As Long

   ' This function stops service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_STOP)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
            StopNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         StopNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      StopNTService = Err.LastDllError
   End If

End Function


'Public Sub NetStart(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
'    Dim timOut As String
'    If IsWindows98 Then
'        Dim exePath As String
'        exePath = Registry.GetValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, "")
'        EXEName = GetFileName(exePath)
'        If Not (EXEName = "") And (ProcessRunning(EXEName) = 0) Then
'            SaveSettingLong &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\" & ServiceName, "NetStarting", 1
'            RunProcess LCase(Trim(exePath))
'            timOut = Now
'            Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)
'                DoTasks
'            Loop
'        End If
'    Else
'        If EXEName = "" Then
'            RunProcess SysPath & "net.exe", "start " & ServiceName, vbHide, True
'        Else
'            If ProcessRunning(EXEName) = 0 Then
'                RunProcess SysPath & "net.exe", "start " & ServiceName, vbHide, True
'                timOut = Now
'                Do Until (ProcessRunning(EXEName) = 1) Or (DateDiff("s", timOut, Now) > 10)
'                    DoTasks
'                Loop
'            End If
'        End If
'    End If
'End Sub
'
'Public Sub NetStop(ByVal ServiceName As String, Optional ByVal EXEName As String = "")
'    Dim timOut As String
'    If IsWindows98 Then
'        Dim exePath As String
'        exePath = Registry.GetValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", ServiceName, "")
'        EXEName = GetFileName(exePath)
'        If Not (EXEName = "") And (ProcessRunning(EXEName) = 1) Then
'            KillApp LCase(Trim(exePath))
'            timOut = Now
'            Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
'                DoTasks
'            Loop
'        End If
'    Else
'        If EXEName = "" Then
'            RunProcess SysPath & "net.exe", "stop " & ServiceName, vbHide, True
'        Else
'            If ProcessRunning(EXEName) = 1 Then
'                RunProcess SysPath & "net.exe", "stop " & ServiceName, vbHide, True
'                timOut = Now
'                Do Until (ProcessRunning(EXEName) = 0) Or (DateDiff("s", timOut, Now) > 10)
'                    DoTasks
'                Loop
'            End If
'        End If
'    End If
'End S