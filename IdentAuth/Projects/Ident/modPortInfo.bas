#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modPortInfo"
#Const modPortInfo = -1
Option Explicit
'TOP DOWN

Option Compare Binary

'For netstat
Private Const PROCESS_VM_READ           As Long = &H10
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_TERMINATE As Long = (&H1)

Private Const MIB_TCP_STATE_CLOSED      As Long = 1
Private Const MIB_TCP_STATE_LISTEN      As Long = 2
Private Const MIB_TCP_STATE_SYN_SENT    As Long = 3
Private Const MIB_TCP_STATE_SYN_RCVD    As Long = 4
Private Const MIB_TCP_STATE_ESTAB       As Long = 5
Private Const MIB_TCP_STATE_FIN_WAIT1   As Long = 6
Private Const MIB_TCP_STATE_FIN_WAIT2   As Long = 7
Private Const MIB_TCP_STATE_CLOSE_WAIT  As Long = 8
Private Const MIB_TCP_STATE_CLOSING     As Long = 9
Private Const MIB_TCP_STATE_LAST_ACK    As Long = 10
Private Const MIB_TCP_STATE_TIME_WAIT   As Long = 11
Private Const MIB_TCP_STATE_DELETE_TCB  As Long = 12
    
Private Type PMIB_UDPEXROW
    dwLocalAddr     As Long
    dwLocalPort     As Long
    dwProcessId     As Long
End Type

Private Type PMIB_TCPEXROW
    dwStats         As Long
    dwLocalAddr     As Long
    dwLocalPort     As Long
    dwRemoteAddr    As Long
    dwRemotePort    As Long
    dwProcessId     As Long
End Type

Private Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi" (ByRef pTcpTable As Any, ByRef bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long
Private Declare Function AllocateAndGetUdpExTableFromStack Lib "iphlpapi" (ByRef pTcpTable As Any, ByRef bOrder As Boolean, ByVal heap As Long, ByVal zero As Long, ByVal flags As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function EnumProcesses Lib "psapi" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "psapi" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long

Public mheap As Long

'to know all of processes in form4
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

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long

Public Function GetAllPortsByRemoteIP(ByVal remoteIP As String) As Collection

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long
    Dim Ret As Collection
    Set Ret = New Collection
    
    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If (remoteIP = GetIpString(TcpExTable(i).dwRemoteAddr)) Then
                    Ret.Add GetPortNumber(TcpExTable(i).dwLocalPort) & " , " & GetPortNumber(TcpExTable(i).dwRemotePort)
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

    Set GetAllPortsByRemoteIP = Ret
    Set Ret = Nothing
End Function

Public Function GetAllPortsByLocalIP(ByVal localIP As String) As Collection

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long
    Dim Ret As Collection
    Set Ret = New Collection
    
    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If (localIP = GetIpString(TcpExTable(i).dwLocalAddr)) Then
                    Ret.Add GetPortNumber(TcpExTable(i).dwLocalPort) & " , " & GetPortNumber(TcpExTable(i).dwRemotePort)
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

    Set GetAllPortsByLocalIP = Ret
    Set Ret = Nothing
End Function

Public Function GetRemoteIPByBothPorts(ByVal LocalPort As Long, ByVal RemotePort As Long) As String

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long
    Dim Ret As String
    
    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If (LocalPort = GetPortNumber(TcpExTable(i).dwLocalPort)) And _
                    (RemotePort = GetPortNumber(TcpExTable(i).dwRemotePort)) Then
                    Ret = GetIpString(TcpExTable(i).dwRemoteAddr)
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

    GetRemoteIPByBothPorts = Ret
End Function

Public Function GetLocalIPByBothPorts(ByVal LocalPort As Long, ByVal RemotePort As Long) As String

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long

    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If (LocalPort = GetPortNumber(TcpExTable(i).dwLocalPort)) And _
                    (RemotePort = GetPortNumber(TcpExTable(i).dwRemotePort)) Then
                    GetLocalIPByBothPorts = GetIpString(TcpExTable(i).dwLocalAddr)
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

End Function

Public Function GetProcessIDByBothPorts(ByVal lPort As Long, ByVal rPort As Long, Optional ByRef localIP As String = "") As Long

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long

    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If lPort = GetPortNumber(TcpExTable(i).dwLocalPort) And _
                 rPort = GetPortNumber(TcpExTable(i).dwRemotePort) Then
                    localIP = GetIpString(TcpExTable(i).dwLocalAddr)
                    GetProcessIDByBothPorts = TcpExTable(i).dwProcessId
                    Exit For
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

End Function

Public Function GetProcessIDByPort(ByVal Port As Long) As Long

    'to refresh list of all port process
    Dim TcpExTable() As PMIB_TCPEXROW
    Dim Distant      As String
    Dim Pointer      As Long
    Dim Number       As Long
    Dim Size         As Long
    Dim i            As Long

    mheap = GetProcessHeap()
    
    'for TCP
    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
        CopyMemory Number, ByVal Pointer, 4
        If Number Then
            ReDim TcpExTable(Number - 1)
            Size = Number * Len(TcpExTable(0))
            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
            For i = 0 To UBound(TcpExTable)
                If Port = GetPortNumber(TcpExTable(i).dwLocalPort) Then
                    GetProcessIDByPort = TcpExTable(i).dwProcessId
                End If
            Next
        End If
        HeapFree mheap, 0, ByVal Pointer
    End If

End Function

Public Function GetProcessName(ByVal ProcessID As Long) As String
    Dim strName     As String * 1024
    Dim hProcess    As Long
    Dim cbNeeded    As Long
    Dim hMod        As Long
    Select Case ProcessID
        Case 0:    GetProcessName = "Proccess Inactive"
        Case 4:    GetProcessName = "System"
        Case Else: GetProcessName = "Unknown"
    End Select
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessID)
    If hProcess Then
        If EnumProcessModules(hProcess, hMod, Len(hMod), cbNeeded) Then
            GetModuleBaseName hProcess, hMod, strName, Len(strName)
            GetProcessName = Left$(strName, lstrlen(strName))
        End If
        CloseHandle hProcess
    End If
End Function

Private Function GetIpString(ByVal Value As Long) As String
    Dim table(3) As Byte
    CopyMemory table(0), Value, 4
    GetIpString = table(0) & "." & table(1) & "." & table(2) & "." & table(3)
End Function

Private Function GetPortNumber(ByVal Value As Long) As Long
    GetPortNumber = (Value / 256) + (Value Mod 256) * 256
End Function

Public Function ProcessPathByPID(pID As Long) As String
    'to know process from its PID
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim Ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pID)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
        If Ret <> 0 Then
            ModuleName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            ProcessPathByPID = Left(ModuleName, Ret)
        End If
    End If
              
    Ret = CloseHandle(hProcess)
    If ProcessPathByPID = "" Then ProcessPathByPID = "SYSTEM"
End Function

'Private Function GetState(ByVal Value As Long) As String
'    Select Case Value
'        Case MIB_TCP_STATE_ESTAB: GetState = "ESTABLISH"
'        Case MIB_TCP_STATE_CLOSED: GetState = "CLOSED"
'        Case MIB_TCP_STATE_LISTEN: GetState = "LISTEN"
'        Case MIB_TCP_STATE_CLOSING: GetState = "CLOSING"
'        Case MIB_TCP_STATE_LAST_ACK: GetState = "LAST_ACK"
'        Case MIB_TCP_STATE_SYN_SENT: GetState = "SYN_SENT"
'        Case MIB_TCP_STATE_SYN_RCVD: GetState = "SYN_RCVD"
'        Case MIB_TCP_STATE_FIN_WAIT1: GetState = "FIN_WAIT1"
'        Case MIB_TCP_STATE_FIN_WAIT2: GetState = "FIN_WAIT2"
'        Case MIB_TCP_STATE_TIME_WAIT: GetState = "TIME_WAIT"
'        Case MIB_TCP_STATE_CLOSE_WAIT: GetState = "CLOSE_WAIT"
'        Case MIB_TCP_STATE_DELETE_TCB: GetState = "DELETE_TCB"
'    End Select
'End Function

'Public Function GetProcessIDByPort(ByVal Port As Long) As Long
'
'    'to refresh list of all port process
'    Dim TcpExTable() As PMIB_TCPEXROW
'    Dim UdpExTable() As PMIB_UDPEXROW
'    Dim Distant      As String
'    Dim Pointer      As Long
'    Dim Number       As Long
'    Dim Size         As Long
'    Dim i            As Long
'
'    mheap = GetProcessHeap()
'    'for TCP
'    If AllocateAndGetTcpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
'        CopyMemory Number, ByVal Pointer, 4
'        If Number Then
'            ReDim TcpExTable(Number - 1)
'            Size = Number * Len(TcpExTable(0))
'            CopyMemory TcpExTable(0), ByVal Pointer + 4, Size
'            For i = 0 To UBound(TcpExTable)
'
'                    'Debug.Print "TCP"
''                    Debug.Print GetIpString(TcpExTable(i).dwLocalAddr)
'                    If Port = GetPortNumber(TcpExTable(i).dwLocalPort) Then
'                        GetProcessIDByPort = TcpExTable(i).dwProcessId
'                    End If
''                    If Not (GetIpString(TcpExTable(i).dwRemoteAddr) = "0.0.0.0") Then
''                        Debug.Print GetIpString(TcpExTable(i).dwRemoteAddr)
''                        Debug.Print ResolveHostname(GetIpString(TcpExTable(i).dwRemoteAddr))
''                        Debug.Print GetPortNumber(TcpExTable(i).dwRemotePort)
''                    End If
''                    Debug.Print GetState(TcpExTable(i).dwStats)
''                    Debug.Print TcpExTable(i).dwProcessId
'                    'Debug.Print GetProcessName(TcpExTable(i).dwProcessId)
'                    'Debug.Print ProcessPathByPID(TcpExTable(i).dwProcessId)
'
'            Next
'        End If
'        HeapFree mheap, 0, ByVal Pointer
'    End If
'
'    'For UDP
'    If AllocateAndGetUdpExTableFromStack(Pointer, True, mheap, 2, 2) = 0 Then
'        CopyMemory Number, ByVal Pointer, 4
'        If Number Then
'            ReDim UdpExTable(Number - 1)
'            Size = Number * Len(UdpExTable(0))
'            CopyMemory UdpExTable(0), ByVal Pointer + 4, Size
'            For i = 0 To UBound(UdpExTable)
'
'                    'Debug.Print "UDP"
''                    Debug.Print GetIpString(UdpExTable(i).dwLocalAddr)
'                    'Debug.Print GetPortNumber(UdpExTable(i).dwLocalPort)
''                    Debug.Print UdpExTable(i).dwProcessId
'                    'Debug.Print GetProcessName(UdpExTable(i).dwProcessId)
'                    'Debug.Print ProcessPathByPID(UdpExTable(i).dwProcessId)
'                    If Port = GetPortNumber(UdpExTable(i).dwLocalPort) Then
'                        GetProcessIDByPort = TcpExTable(i).dwProcessId
'                    End If
'
'            Next
'        End If
'        HeapFree mheap, 0, ByVal Pointer
'    End If
'
'End Function
