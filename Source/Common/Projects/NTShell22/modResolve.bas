#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modResolve"
#Const modResolve = -1
Option Explicit
'TOP DOWN
Option Compare Binary


Option Private Module
Private Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Const hostent_size = 16

Private Const INADDR_NONE = &HFFFFFFFF
Private Const SOCKET_ERROR = -1

Private Declare Function GetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal host_name As String) As Long
Private Declare Function GetHostName Lib "wsock32" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Private Declare Function inet_addr Lib "wsock32" (ByVal cp As String) As Long

Private Declare Function WSAStartup Lib "wsock32" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Private Declare Function WSAAsyncSelect Lib "wsock32" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long
Private Declare Function WSAGetLastError Lib "wsock32" () As Long
    
Private Declare Sub CopyMemoryHost Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As HOSTENT, ByVal xSource As Long, ByVal nbytes As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByVal xSource As Long, ByVal nbytes As Long)

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    imaxudp As Integer
    lpszvenderinfo As Long
End Type

Private Function hiByte(ByVal wParam As Integer)
On Error GoTo catch

    hiByte = wParam \ &H100 And &HFF&

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Private Function lobyte(ByVal wParam As Integer)
On Error GoTo catch

    lobyte = wParam And &HFF&

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Private Function SocketsInitialize() As Long
On Error GoTo catch

    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String
    Dim sckOk As Long
    sckOk = 0
    
    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
    
    If iReturn <> 0 Then
        sckOk = 1
    Else
        If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
            WS_VERSION_MAJOR And hiByte(WSAD.wversion) < WS_VERSION_MINOR) Then
            sHighByte = Trim$(Str$(hiByte(WSAD.wversion)))
            sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
            sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
            sMsg = sMsg & " is not supported by winsock.dll "
            sckOk = 2
        End If
    End If
    SocketsInitialize = sckOk

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Private Function SocketsCleanup() As Long
On Error GoTo catch

    Dim lReturn As Long
    Dim sckOk As Long
    sckOk = 0
    lReturn = WSACleanup()
    
    If lReturn <> 0 Then
        sckOk = 1
    End If
    SocketsCleanup = sckOk

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Public Function ResolveHost(ByVal Host As String) As String

    Dim retVal As Long
    retVal = SocketsInitialize()
    If retVal = 0 Then
    
        Dim phe As Long
        Dim heDestHost As HOSTENT
        Dim addrList As Long
        Dim rc As Long
    
        Dim hostip_addr As Long
        
        Dim temp_ip_address() As Byte
        Dim i As Integer
        Dim ip_address As String
    
        rc = inet_addr(Host)
        If rc = SOCKET_ERROR Then
        
            phe = GetHostByName(Host)
            If phe <> 0 Then
            
                CopyMemoryHost heDestHost, phe, hostent_size
                CopyMemory addrList, heDestHost.h_addr_list, 4
                CopyMemory hostip_addr, addrList, heDestHost.h_length
                rc = hostip_addr
                
                ReDim temp_ip_address(1 To heDestHost.h_length)
                RtlMoveMemory temp_ip_address(1), hostip_addr, heDestHost.h_length
    
                For i = 1 To heDestHost.h_length
                    ip_address = ip_address & temp_ip_address(i) & "."
                Next
                ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    
                ResolveHost = ip_address
                
            Else
                rc = INADDR_NONE
            End If
            
        End If

        retVal = SocketsCleanup()


    End If
    
End Function

Public Function LocalHost() As String

    Dim retVal As Long
    retVal = SocketsInitialize()
    If retVal = 0 Then
    
        Dim buf As String
        Dim rc As Long
        
        buf = Space$(255)
        
        rc = GetHostName(buf, Len(buf))
        rc = InStr(buf, vbNullChar)
        
        If rc > 0 Then
            LocalHost = Left$(buf, rc - 1)
        Else
            LocalHost = ""
        End If
    
        retVal = SocketsCleanup()

    End If
End Function

