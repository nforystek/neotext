Attribute VB_Name = "modSocks"

Option Private Module

Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, Addr As sockaddr, addrLen As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, Addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Public Declare Function connect Lib "ws2_32.dll" (ByVal s As Long, ByRef Addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (Addr As Long, addrLen As Long, addrType As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Long
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long  'removed ByVal from buf
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long
Public Declare Function send Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, toaddr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long


    Public Type HOSTENT
        h_name As Long          'Official name of the host (PC).
        h_aliases As Long       'Null-terminated array of alternate names.
        h_addrtype As Integer   'Type of address being returned.
        h_length As Integer     'Length of each address, in bytes.
        h_addr_list As Long     'Null-terminated list of addresses for the host.
    End Type

    Public Type protoent
        p_name As Long
        p_aliases As Long
        p_proto As Integer
    End Type

    Public Type servent
        s_name As Long
        s_aliases As Long
        s_port As Integer
        s_proto As Long
    End Type

    Public Type sockaddr
        sin_family As Integer
        sin_port As Integer
        sin_addr As Long
        sin_zero As String * 8
    End Type

    Public Type WSADATA                 'The members of the Windows Sockets WSADATA structure are:
        wVersion As Integer             'Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
        wHighVersion As Integer         'Highest version of the Windows Sockets specification that this .dll can support (also encoded as above).
        szDescription As String * 257   'Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows Sockets implementation.
        szSystemStatus As String * 129  'Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or configuration information.
        iMaxSockets As Integer          'Retained for backward compatibility, but should be ignored for Windows Sockets version 2 and later, as no single value can be appropriate for all underlying service providers.
        iMaxUdpDg As Integer            'Ignored for Windows Sockets version 2 and onward.
        lpVendorInfo As Long            'Ignored for Windows Sockets version 2 and onward.
    End Type


    'Constants

    Public Const AF_UNSPEC = 0      'unspecified
    'Although  AF_UNSPEC  is  defined for backwards compatibility, using
    'AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
    'DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
    'depends  on the actual address family chosen.  As environments grow
    'to  include  more  and  more  address families that use overlapping
    'protocol  values  there  is  more  and  more  chance of choosing an
    'undesired address family when AF_UNSPEC is used.
    Public Const AF_UNIX = 1        'local to host (pipes, portals)
    Public Const AF_INET = 2        'internetwork: UDP, TCP, etc.
    Public Const AF_IMPLINK = 3     'arpanet imp addresses
    Public Const AF_PUP = 4         'pup protocols: e.g. BSP
    Public Const AF_CHAOS = 5       'mit CHAOS protocols
    Public Const AF_NS = 6          'XEROX NS protocols
    Public Const AF_IPX = AF_NS     'IPX protocols: IPX, SPX, etc.
    Public Const AF_ISO = 7         'ISO protocols
    Public Const AF_OSI = AF_ISO    'OSI is ISO
    Public Const AF_ECMA = 8        'european computer manufacturers
    Public Const AF_DATAKIT = 9     'datakit protocols
    Public Const AF_CCITT = 10      'CCITT protocols, X.25 etc
    Public Const AF_SNA = 11        'IBM SNA
    Public Const AF_DECnet = 12     'DECnet
    Public Const AF_DLI = 13        'Direct data link interface
    Public Const AF_LAT = 14        'LAT
    Public Const AF_HYLINK = 15     'NSC Hyperchannel
    Public Const AF_APPLETALK = 16  'AppleTalk
    Public Const AF_NETBIOS = 17    'NetBios-style addresses
    Public Const AF_VOICEVIEW = 18  'VoiceView
    Public Const AF_FIREFOX = 19    'Protocols from Firefox
    Public Const AF_UNKNOWN1 = 20   'Somebody is using this!
    Public Const AF_BAN = 21        'Banyan
    Public Const AF_ATM = 22        'Native ATM Services
    Public Const AF_INET6 = 23      'Internetwork Version 6
    Public Const AF_CLUSTER = 24    'Microsoft Wolfpack
    Public Const AF_12844 = 25      'IEEE 1284.4 WG AF
    Public Const AF_IRDA = 26       'IrDA
    Public Const AF_NETDES = 28     'Network Designers OSI & gateway enabled protocols

    Public Const FD_READ_BIT = 0
    Public Const FD_READ = FD_READ_BIT
    Public Const FD_WRITE_BIT = 1
    Public Const FD_WRITE = FD_WRITE_BIT
    Public Const FD_OOB_BIT = 2
    Public Const FD_OOB = FD_OOB_BIT
    Public Const FD_ACCEPT_BIT = 3
    Public Const FD_ACCEPT = FD_ACCEPT_BIT
    Public Const FD_CONNECT_BIT = 4
    Public Const FD_CONNECT = FD_CONNECT_BIT
    Public Const FD_CLOSE_BIT = 5
    Public Const FD_CLOSE = FD_CLOSE_BIT
    Public Const FD_QOS_BIT = 6
    Public Const FD_QOS = FD_QOS_BIT
    Public Const FD_GROUP_QOS_BIT = 7
    Public Const FD_GROUP_QOS = FD_GROUP_QOS_BIT
    Public Const FD_ROUTING_INTERFACE_CHANGE_BIT = 8
    Public Const FD_ROUTING_INTERFACE_CHANGE = FD_ROUTING_INTERFACE_CHANGE_BIT
    Public Const FD_ADDRESS_LIST_CHANGE_BIT = 9
    Public Const FD_ADDRESS_LIST_CHANGE = FD_ADDRESS_LIST_CHANGE_BIT
    Public Const FD_MAX_EVENTS = 10
    Public Const FD_ALL_EVENTS = FD_MAX_EVENTS - 1

    Public Const INVALID_SOCKET = &HFFFF
    Public Const SOCKET_ERROR = -1

    Public Const IPPORT_ECHO = 7
    Public Const IPPORT_DISCARD = 9
    Public Const IPPORT_SYSTAT = 11
    Public Const IPPORT_DAYTIME = 13
    Public Const IPPORT_NETSTAT = 15
    Public Const IPPORT_FTP = 21
    Public Const IPPORT_TELNET = 23
    Public Const IPPORT_SMTP = 25
    Public Const IPPORT_TIMESERVER = 37
    Public Const IPPORT_NAMESERVER = 42
    Public Const IPPORT_WHOIS = 43
    Public Const IPPORT_MTP = 57

    Public Const IPPORT_TFTP = 69
    Public Const IPPORT_RJE = 77
    Public Const IPPORT_FINGER = 79
    Public Const IPPORT_TTYLINK = 87
    Public Const IPPORT_SUPDUP = 95

    Public Const IPPORT_EXECSERVER = 512
    Public Const IPPORT_LOGINSERVER = 513
    Public Const IPPORT_CMDSERVER = 514
    Public Const IPPORT_EFSSERVER = 520

    Public Const IPPORT_BIFFUDP = 512
    Public Const IPPORT_WHOSERVER = 513
    Public Const IPPORT_ROUTESERVER = 520   '521 also used
    Public Const IPPORT_RESERVED = 1024     'Ports < IPPORT_RESERVED are reserved for privileged processes (e.g. root).

    Public Const IPPROTO_IP = 0         'dummy for IP
    Public Const IPPROTO_ICMP = 1       'control message protocol
    Public Const IPPROTO_IGMP = 2       'internet group management protocol
    Public Const IPPROTO_GGP = 3        'gateway^2 (deprecated)
    Public Const IPPROTO_TCP = 6        'tcp
    Public Const IPPROTO_PUP = 12       'pup
    Public Const IPPROTO_UDP = 17       'user datagram protocol
    Public Const IPPROTO_IDP = 22       'xns idp
    Public Const IPPROTO_ND = 77        'UNOFFICIAL net disk proto
    Public Const IPPROTO_RAW = 255      'raw IP packet
    Public Const IPPROTO_MAX = 256

    Public Const MSG_OOB = &H1          'process out-of-band data
    Public Const MSG_PEEK = &H2         'peek at incoming message
    Public Const MSG_DONTROUTE = &H4    'send without using routing tables
    Public Const MSG_PARTIAL = &H8000   'partial send or recv for message xport
    
    Public Const SO_DEBUG = &H1             'turn on debugging info recording
    Public Const SO_ACCEPTCONN = &H2        'socket has had listen()
    Public Const SO_REUSEADDR = &H4         'allow local address reuse
    Public Const SO_KEEPALIVE = &H8         'keep connections alive
    Public Const SO_DONTROUTE = &H10        'just use interface addresses
    Public Const SO_BROADCAST = &H20        'permit sending of broadcast msgs
    Public Const SO_USELOOPBACK = &H40      'bypass hardware when possible
    Public Const SO_LINGER = &H80           'linger on close if data present
    Public Const SO_OOBINLINE = &H100       'leave received OOB data in line

    Public Const SO_SNDBUF = &H1001        'send buffer size
    Public Const SO_RCVBUF = &H1002        'receive buffer size
    Public Const SO_SNDLOWAT = &H1003      'send low-water mark
    Public Const SO_RCVLOWAT = &H1004      'receive low-water mark
    Public Const SO_SNDTIMEO = &H1005      'send timeout
    Public Const SO_RCVTIMEO = &H1006      'receive timeout
    Public Const SO_ERROR = &H1007         'get error status and clear
    Public Const SO_TYPE = &H1008          'get socket type

    Public Const SO_GROUP_ID = &H2001           'ID of a socket group
    Public Const SO_GROUP_PRIORITY = &H2002     'the relative priority within a group
    Public Const SO_MAX_MSG_SIZE = &H2003       'maximum message size
    Public Const SO_PROTOCOL_INFOA = &H2004     'WSAPROTOCOL_INFOA structure
    Public Const SO_PROTOCOL_INFOW = &H2005     'WSAPROTOCOL_INFOW structure
    Public Const PVD_CONFIG = &H3001            'configuration info for service provider
    Public Const SO_CONDITIONAL_ACCEPT = &H3002 'enable true conditional accept connection is not ack-ed to the other side until conditional function returns CF_ACCEPT
                                       
    Public Const SOCK_STREAM = 1        'stream socket
    Public Const SOCK_DGRAM = 2         'datagram socket
    Public Const SOCK_RAW = 3           'raw-protocol interface
    Public Const SOCK_RDM = 4           'reliably-delivered message
    Public Const SOCK_SEQPACKET = 5     'sequenced packet stream

    Public Const SOL_SOCKET = &HFFFF

    Public Const WSA_DESCRIPTION_LEN = 256 'Upto 256 char
    Public Const WSA_SYS_STATUS_LEN = 128
    
    
Public Function GetHostByIP(strIP As String) As String
    If Len(strIP) < 1 Then Exit Function 'Must contain text
    
    Dim Host As HOSTENT 'Cannot use HOSTENT
    Dim lngIP As Long
    Dim strHost As String * 255
    
    lngIP = inet_addr(strIP & Chr$(0))
    
    'Copy mem
    MoveMemory Host, apierror, Len(Host)
    MoveMemory ByVal strHost, Host.h_name, 255
    
    strIP = strHost 'I think you can use strHost
    
    'Pull from beginning to null
    If InStr(strIP, Chr$(0)) <> 0 Then
        strIP = Left$(strIP, InStr(strIP, Chr$(0)) - 1)
    End If
    
    strIP = Trim$(strIP)
    
    GetHostByIP = strIP 'Send back out
End Function

Public Sub GetIPByHost(strHost As String, aryIP() As String, lngIP As Long)
    If Len(strHost) < 1 Then Exit Sub      'Must contain text
    
    Dim Host As HOSTENT 'Cannot use HOSTENT
    Dim lngHostIp As Long
    Dim strIP As String
    Dim tmpIP() As Byte '1/4 ip = byte
    Dim tmpInt As Integer
    
    apierror = gethostbyname(strHost & Chr$(0))
    If apierror = 0 Then
        Failed "gethostbyname"
        Exit Sub
    End If

    'Copy mem
    MoveMemory Host, apierror, LenB(Host)
    MoveMemory lngHostIp, Host.h_addr_list, 4 'Copy 4 parts of ip

    Do
        ReDim tmpIP(1 To Host.h_length) 'Resize
        MoveMemory tmpIP(1), lngHostIp, Host.h_length 'Copy mem
        
        strIP = ""
        For tmpInt = 1 To Host.h_length 'Cyle through all parts
            strIP = strIP & tmpIP(tmpInt) & "." 'Add . in between
        Next
        strIP = Left$(strIP, Len(strIP) - 1) 'Remove extra .
        
        lngIP = lngIP + 1
        ReDim Preserve aryIP(lngIP)
        aryIP(lngIP) = strIP
        
        Host.h_addr_list = Host.h_addr_list + LenB(Host.h_addr_list)
        MoveMemory lngHostIp, Host.h_addr_list, 4
    Loop While lngHostIp <> 0
End Sub

Public Sub WinSockError(apiFunction As String)
    If WSAGetLastError > 0 Then
        Errors.Errors WSAGetLastError, apiFunction
    Else
        Failed apiFunction
    End If
End Sub

Public Sub WinSockStart()
    'Winsock 2.2 startup
    Dim WSADATA As WSADATA
    
    With WSADATA
        .szDescription = Space$(256)
        .szSystemStatus = Space$(128)
        .wHighVersion = 2
        .wVersion = 2
    End With
    Dim apierror As Long
    
    apierror = WSAStartup(&H202, WSADATA)
    If apierror <> 0 Then
        'Errors.Errors apierror, "WSAStartup"
    Else
        'Propriatary
        With WinsockData
            .Description = WSADATA.szDescription
            .SystemStatus = WSADATA.szSystemStatus
        End With
    End If
End Sub

Public Sub WinSockEnd()
    apierror = WSACleanup
    If apierror <> 0 Then 'Errors.Errors apierror, "WSACleanup"
End Sub



