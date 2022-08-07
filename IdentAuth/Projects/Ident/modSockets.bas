#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modSockets"
#Const modSockets = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public Const ftpLocalSize = 1048576
Public Const ftpBufferSize = 65536
Public Const ftpPacketSize = 16384

Public Const WM_WINSOCK = 4025

Public Const INVALID_SOCKET = -1
Public Const SOCKET_ERROR = -1

Public Declare Function ioctlsocket Lib "wsock32" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long

Public Declare Function GetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal host_name As String) As Long
Public Declare Function GetHostName Lib "wsock32" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function GetServByName Lib "wsock32" Alias "getservbyname" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function htons Lib "wsock32" (ByVal hostshort As Long) As Integer
Public Declare Function inet_addr Lib "wsock32" (ByVal cp As String) As Long

Public Declare Function Bind Lib "wsock32" Alias "bind" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function socket Lib "wsock32" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function SocketAccept Lib "wsock32" Alias "accept" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
Public Declare Function SocketConnect Lib "wsock32" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function SocketListen Lib "wsock32" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function SocketSend Lib "wsock32" Alias "send" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function SocketRecv Lib "wsock32" Alias "recv" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function SocketClose Lib "wsock32" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function setsockopt Lib "wsock32" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "wsock32" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long

Public Declare Function WSAStartup Lib "wsock32" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Public Declare Function WSAAsyncSelect Lib "wsock32" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Public Declare Function WSACleanup Lib "wsock32" () As Long
Public Declare Function WSAGetLastError Lib "wsock32" () As Long
    
Public Declare Sub CopyMemoryHost Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As HOSTENT, ByVal xSource As Long, ByVal nbytes As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByVal xSource As Long, ByVal nbytes As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Public Const WSAEWOULDBLOCK = 10035

Public Const FD_READ = &H1
Public Const FD_WRITE = &H2
Public Const FD_ACCEPT = &H8
Public Const FD_CONNECT = &H10
Public Const FD_CLOSE = &H20
Public Const FD_OOB = &H4

Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Public Const SO_DEBUG = &H1&         ' Turn on debugging info recording
Public Const SO_ACCEPTCONN = &H2&    ' Socket has had listen() - READ-ONLY.
Public Const SO_REUSEADDR = &H4&     ' Allow local address reuse.
Public Const SO_KEEPALIVE = &H8&     ' Keep connections alive.
Public Const SO_DONTROUTE = &H10&    ' Just use interface addresses.
Public Const SO_BROADCAST = &H20&    ' Permit sending of broadcast msgs.
Public Const SO_USELOOPBACK = &H40&  ' Bypass hardware when possible.
Public Const SO_LINGER = &H80&       ' Linger on close if data present.
Public Const SO_OOBINLINE = &H100&   ' Leave received OOB data in line.
Public Const SO_CONDITIONAL_ACCEPT = &H3002

Public Const SO_DONTLINGER = Not SO_LINGER
Public Const SO_EXCLUSIVEADDRUSE = Not SO_REUSEADDR ' Disallow local address reuse.

Public Const SO_SNDTIMEO = &H1005
Public Const SO_RCVTIMEO = &H1006
Public Const SO_SNDBUF = &H1001&
Public Const SO_RCVBUF = &H1002&
Public Const SO_ERROR = &H1007&
Public Const SO_TYPE = &H1008&

Public Const SOL_SOCKET = 65535

Public Const TCP_NODELAY = &H1&      ' Turn off Nagel Algorithm.

Public Const INADDR_NONE = &HFFFFFFFF
Public Const INADDR_ANY = &H0

Public Const AF_UNSPEC = 0
Public Const AF_INET = 2
Public Const AF_IPX = 6
Public Const AF_APPLETALK = 16
Public Const AF_NETBIOS = 17
Public Const AF_INET6 = 23
Public Const AF_IRDA = 26
Public Const AF_BTH = 32

Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5

Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_IGMP = 2
Public Const BTHPROTO_RFCOMM = 13
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_ICMPV6 = 58
Public Const IPPROTO_RM = 113

Public Const FIONBIO = &H8004667E
Public Const FIONREAD = &H4004667F
Public Const SIOCATMARK = &H40047307

Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    imaxudp As Integer
    lpszvenderinfo As Long
End Type

Public Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16

Public Type HOSTENT
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16

Public Function WSAGetSelectEvent(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        WSAGetSelectEvent = (lParam And &HFFFF&) - &H10000
    Else
        WSAGetSelectEvent = lParam And &HFFFF&
    End If
End Function

Public Function WSAGetAsyncError(ByVal lParam As Long) As Integer
    WSAGetAsyncError = (lParam And &HFFFF0000) \ &H10000
End Function

Function HiByte(ByVal wParam As Integer)
On Error GoTo catch

    HiByte = wParam \ &H100 And &HFF&

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Function LoByte(ByVal wParam As Integer)
On Error GoTo catch

    LoByte = wParam And &HFF&

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Function SocketsInitialize() As Long
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
        If LoByte(WSAD.wversion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wversion) = _
            WS_VERSION_MAJOR And HiByte(WSAD.wversion) < WS_VERSION_MINOR) Then
            sHighByte = Trim$(str$(HiByte(WSAD.wversion)))
            sLowByte = Trim$(str$(LoByte(WSAD.wversion)))
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

Function SocketsCleanup() As Long
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
Public Function GetZeroNotation(ByVal IPAddr As String) As String

    Dim bt As String
    
    Do Until IPAddr = ""
    
        bt = RemoveNextArg(IPAddr, ".")
        
        GetZeroNotation = GetZeroNotation & Chr(CByte(CByte(bt) \ 256)) & Chr(CByte(CByte(bt) Mod 256))
        
    Loop

End Function
Public Function GetPortIP() As Collection
On Error GoTo catch

    Dim IPList As New Collection

    Dim init As Boolean
    Dim retVal As Long
    If Not WinsockControl Then
        init = True
        retVal = SocketsInitialize()
    End If
    If retVal = 0 Then
        Dim hostname As String * 256
        Dim hostent_addr As Long
        Dim host As HOSTENT
        Dim hostip_addr As Long
        Dim temp_ip_address() As Byte
        Dim i As Integer
        Dim ip_address As String
    
        If GetHostName(hostname, 256) = SOCKET_ERROR Then
            retVal = 1
        Else
            hostname = Trim$(hostname)
            hostent_addr = GetHostByName(hostname)

            If hostent_addr = 0 Then
                retVal = 2
            Else
                
                RtlMoveMemory host, hostent_addr, LenB(host)
                RtlMoveMemory hostip_addr, host.h_addr_list, 4
    
                Do
                    ReDim temp_ip_address(1 To host.h_length)
                    RtlMoveMemory temp_ip_address(1), hostip_addr, host.h_length
    
                    For i = 1 To host.h_length
                        ip_address = ip_address & temp_ip_address(i) & "."
                    Next
                    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    
                    IPList.Add ip_address
    
                    ip_address = ""
                    host.h_addr_list = host.h_addr_list + LenB(host.h_addr_list)
                    RtlMoveMemory hostip_addr, host.h_addr_list, 4
                Loop While (hostip_addr <> 0)
                
            End If

        End If
        
        If init Then
            retVal = SocketsCleanup()
        End If

    End If
    
    Set GetPortIP = IPList

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Public Function ResolveIP(ByVal host As String) As String
   
    Dim phe As Long
    Dim heDestHost As HOSTENT
    Dim addrList As Long
    Dim rc As Long

    Dim hostip_addr As Long
    
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String

    rc = inet_addr(host)
    If rc = SOCKET_ERROR Then
    
        phe = GetHostByName(host)
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

            ResolveIP = ip_address
            
        Else
            rc = INADDR_NONE
        End If
        
    End If
    
    
End Function

Public Function Resolve(ByVal host As String) As Long
   
    Dim phe As Long
    Dim heDestHost As HOSTENT
    Dim addrList As Long
    Dim rc As Long

    rc = inet_addr(host)
    If rc = SOCKET_ERROR Then
    
        phe = GetHostByName(host)
        If phe <> 0 Then
        
            CopyMemoryHost heDestHost, phe, hostent_size
            CopyMemory addrList, heDestHost.h_addr_list, 4
            CopyMemory rc, addrList, heDestHost.h_length
            
        Else
            rc = INADDR_NONE
        End If
        
    End If
    
    Resolve = rc
End Function


