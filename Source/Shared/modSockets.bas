Attribute VB_Name = "modSockets"

#Const modSockets = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Public Const WM_USER = &H400
Public Const SOCKET_MESSAGE = 32769
Public Const WM_WINSOCK = 4025
Public Const WINSOCK_MESSAGE = WM_USER + &H401
'
Public Const FIONBIO = &H8004667E
Public Const FIONREAD = &H4004667F

Public Const MSG_OOB = &H1          'process out-of-band data
Public Const MSG_PEEK = &H2         'peek at incoming message
Public Const MSG_DONTROUTE = &H4    'send without using routing tables
Public Const MSG_PARTIAL = &H8000   'partial send or recv for message xport
Public Const MSG_WAITALL = &H8

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17

Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_Name As String, ByVal proto As String) As Long
'Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
'Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_Name As String) As Long
'Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As Any, ByRef namelen As Long) As Long
Public Declare Function SocketAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, Addr As Any, addrlen As Long) As Long
Public Declare Function Bind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, Addr As sockaddr, ByVal namelen As Long) As Long
Public Declare Function socketclose Lib "ws2_32.dll" Alias "closesocket" (ByVal s As Long) As Long
Public Declare Function SocketConnect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef Addr As Any, ByVal namelen As Long) As Long
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (Addr As Long, addrlen As Long, addrtype As Long) As Long
Public Declare Function GetHostByName Lib "ws2_32.dll" Alias "gethostbyname" (ByVal host_name As String) As Long
Public Declare Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function SocketListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function SocketRecvLngPtr Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByVal Buf As Long, ByVal buflen As Long, ByVal Flags As Long) As Long 'removed ByVal from buf
'Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, ByRef fromaddr As sockaddr, ByRef fromlen As Long) As Long
Public Declare Function SocketSendLngPtr Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByVal Buf As Long, ByVal buflen As Long, ByVal Flags As Long) As Long
'Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal flags As Long, toaddr As sockaddr, ByVal tolen As Long) As Long
Public Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, ByRef optval As Long, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, ByVal optval As Long, optlen As Long) As Long
'Public Declare Function shutdown Lib "ws2_32.dll" (ByVal s As Long, ByVal how As Long) As Long
Public Declare Function Socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function SocketSendString Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByVal Buf As String, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function SocketRecvString Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByVal Buf As String, ByVal buflen As Long, ByVal Flags As Long) As Long



'Reciving and sending data on winsock functions
Public Declare Function WSARecv Lib "ws2_32.dll" Alias "recv" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function WSASend Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Winsock API functions to create a listening server
Public Declare Function WSABind Lib "ws2_32.dll" Alias "bind" (ByVal s As Long, ByRef name As sockaddr, ByRef namelen As Long) As Long
Public Declare Function WSAListen Lib "ws2_32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function WSAAccept Lib "ws2_32.dll" Alias "accept" (ByVal s As Long, ByRef Addr As sockaddr, ByRef addrlen As Long) As Long

Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long



Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal laddr As Long) As Long
Public Declare Function inet_pton Lib "ws2_32.dll" (ByVal af As Long, ByVal pszAddrString As String, ByRef pAddrBuf As Any) As Long
'Private Declare Function inet_pton Lib "ws2_32.dll" (ByVal af As Long, ByVal pszAddrString As String, ByRef pAddrBuf As Any) As Long
Public Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
'Private Declare Function inet_ntop Lib "ws2_32.dll" (ByVal af As Long, ByRef ppAddr As Any, ByRef pStringBuf As Any, ByVal StringBufSize As Long) As Long
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Long) As Integer
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Long) As Integer

Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal s As Long, ByVal cmd As Long, argp As Long) As Long '

''Public Declare Function gethostbyName Lib "wsock32.dll" (ByVal host_Name As String) As Long
'Public Declare Function gethostname Lib "wsock32.dll" Alias "gethostName" (ByVal host_name As String, ByVal namelen As Long) As Long
''Public Declare Function getservbyName Lib "wsock32.dll" (ByVal serv_Name As String, ByVal proto As String) As Long
''
''
'Public Declare Function Bind Lib "wsock32.dll" Alias "bind" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
'Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
'Public Declare Function SocketAccept Lib "wsock32.dll" Alias "accept" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
'Public Declare Function SocketConnect Lib "wsock32.dll" Alias "connect" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
'Public Declare Function SocketListen Lib "wsock32.dll" Alias "listen" (ByVal s As Long, ByVal backlog As Long) As Long
''Public Declare Function SocketSendString Lib "wsock32.dll" Alias "send" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long) As Long
'Public Declare Function SocketSendAny Lib "wsock32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
''Public Declare Function SocketSendLngPtr Lib "wsock32.dll" Alias "send" (ByVal s As Long, ByVal buf As Long, ByVal buflen As Long, ByVal flags As Long) As Long
''Public Declare Function SocketRecvString Lib "wsock32.dll" Alias "recv" (ByVal s As Long, ByVal buf As String, ByVal buflen As Long, ByVal flags As Long) As Long
'Public Declare Function SocketRecvAny Lib "wsock32.dll" Alias "recv" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
''Public Declare Function SocketRecvLngPtr Lib "wsock32.dll" Alias "recv" (ByVal s As Long, ByVal buf As Long, ByVal buflen As Long, ByVal flags As Long) As Long
'Public Declare Function socketclose Lib "wsock32.dll" Alias "closesocket" (ByVal s As Long) As Long
'Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
'Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
''
'Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADATA) As Long
'Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
'Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
'Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
'

Public Declare Sub CopyMemoryHost Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As hostent, ByVal xSource As Long, ByVal nbytes As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByVal xSource As Long, ByVal nbytes As Long)
Public Declare Sub RtlMoveMemory3 Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Public Declare Sub RtlMoveMemory2 Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, Source As Any, ByVal cbCopy As Long)

Public Const FD_READ       As Long = &H1
Public Const FD_WRITE      As Long = &H2
Public Const FD_OOB        As Long = &H4
Public Const FD_ACCEPT     As Long = &H8
Public Const FD_CONNECT    As Long = &H10
Public Const FD_CLOSE      As Long = &H20

'Public Const FD_READ = 0
'Public Const FD_WRITE = 1
'Public Const FD_OOB = 2
'Public Const FD_ACCEPT = 3
'Public Const FD_CONNECT = 4
'Public Const FD_CLOSE = 5
'Public Const FD_QOS = 6
'Public Const FD_GROUP_QOS = 7
'Public Const FD_ROUTING_INTERFACE_CHANGE = 8
'Public Const FD_ADDRESS_LIST_CHANGE = 9
'Public Const FD_MAX_EVENTS = 10

'Public Const FD_EVENTS As Long = (FD_READ Or FD_WRITE Or FD_OOB Or FD_ACCEPT Or FD_CONNECT Or FD_CLOSE)
'Public Const FD_LISTEN As Long = (FD_ACCEPT Or FD_CLOSE)

Public Const FD_EVENTS As Long = (FD_READ Or FD_WRITE Or FD_OOB Or FD_CONNECT Or FD_CLOSE)
Public Const FD_LISTEN As Long = (FD_ACCEPT Or FD_CLOSE)

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

'
Public Const WS_VERSION_REQD = &H101 '&H202
Public Const WS_SSL3_VERSION = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Public Const SO_SNDTIMEO = &H1005      'send timeout
Public Const SO_RCVTIMEO = &H1006      'receive timeout
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
'
Public Const SO_DONTLINGER = Not SO_LINGER
Public Const SO_EXCLUSIVEADDRUSE = Not SO_REUSEADDR ' Disallow local address reuse.

Public Const SO_SNDBUF = &H1001&     ' Send buffer size.
Public Const SO_RCVBUF = &H1002&     ' Receive buffer size.
Public Const SO_ERROR = &H1007&      ' Get error status and clear.
Public Const SO_TYPE = &H1008&       ' Get socket type - READ-ONLY.

Public Const TCP_NODELAY = &H1&      ' Turn off Nagel Algorithm.
'
Public Const INVALID_SOCKET = -1
Public Const Socket_ERROR = -1

Public Const SOCK_STREAM = 1

Public Const INADDR_NONE = &HFFFFFFFF
Public Const INADDR_ANY = &H0

Public Const AF_INET = 2

Public Const SOL_SOCKET = 65535

''
'' All Windows Sockets error constants are biased by WSABASEERR from
'' the "normal"
''
Public Const WSABASEERR = 10000
'
''
'' Windows Sockets definitions of regular Microsoft C error constants
'
Public Const WSAEINTR = (WSABASEERR + 4)
Public Const WSAEBADF = (WSABASEERR + 9)
Public Const WSAEACCES = (WSABASEERR + 13)
Public Const WSAEFAULT = (WSABASEERR + 14)
Public Const WSAEINVAL = (WSABASEERR + 22)
Public Const WSAEMFILE = (WSABASEERR + 24)

'
' Windows Sockets definitions of regular Berkeley error constants
'
Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)
Public Const WSAEINPROGRESS = (WSABASEERR + 36)
Public Const WSAEALREADY = (WSABASEERR + 37)
Public Const WSAENOTSOCK = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)
Public Const WSAEMSGSIZE = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)
Public Const WSAEADDRINUSE = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)
Public Const WSAENETDOWN = (WSABASEERR + 50)
Public Const WSAENETUNREACH = (WSABASEERR + 51)
Public Const WSAENETRESET = (WSABASEERR + 52)
Public Const WSAECONNABORTED = (WSABASEERR + 53)
Public Const WSAECONNRESET = (WSABASEERR + 54)
Public Const WSAENOBUFS = (WSABASEERR + 55)
Public Const WSAEISCONN = (WSABASEERR + 56)
Public Const WSAENOTCONN = (WSABASEERR + 57)
Public Const WSAESHUTDOWN = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)
Public Const WSAETIMEDOUT = (WSABASEERR + 60)
Public Const WSAECONNREFUSED = (WSABASEERR + 61)
Public Const WSAELOOP = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)
Public Const WSAENOTEMPTY = (WSABASEERR + 66)
Public Const WSAEPROCLIM = (WSABASEERR + 67)
Public Const WSAEUSERS = (WSABASEERR + 68)
Public Const WSAEDQUOT = (WSABASEERR + 69)
Public Const WSAESTALE = (WSABASEERR + 70)
Public Const WSAEREMOTE = (WSABASEERR + 71)

''
'' Extended Windows Sockets error constant definitions

Public Const WSASYSNOTREADY = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)
Public Const WSANOTINITIALISED = (WSABASEERR + 93)
Public Const WSAEDISCON = (WSABASEERR + 101)
Public Const WSAENOMORE = (WSABASEERR + 102)
Public Const WSAECANCELLED = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED = (WSABASEERR + 111)
Public Const WSAEREFUSED = (WSABASEERR + 112)

Public Const WSAHOST_NOT_FOUND = 11001
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const WSATRY_AGAIN = 11002
Public Const WSANO_RECOVERY = 11003
Public Const WSANO_DATA = 11004
'
Public Const SO_ON As Byte = &H1&
Public Const SO_OFF As Byte = &H0&
'
Public Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Public Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(1 To WSADESCRIPTION_LEN) As Byte
    szSystemstatus(1 To WSASYS_STATUS_LEN) As Byte
    iMaxSockets As Integer
    imaxudp As Integer
    lpszvenderinfo As Long
End Type

Private Type in_addr
   s_addr   As Long
End Type

Public Type sockaddr
    sin_family          As Integer  '2 bytes
    sin_port            As Integer  '2 bytes
    sin_addr            As Long  '4 bytes
    sin_zero(1 To 8)    As Byte     '8 bytes
End Type                            'Total 16 bytes

Public Const sockaddr_size = 16

Public Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Public Const hostent_size = 16


Public Type saddr
    sa_family           As Integer  '2 bytes
    sa_data(0 To 25)    As Byte     '26 bytes
End Type

'Public Type fd_set
'   fd_count As Long
'   fd_array(FD_SETSIZE) As Long
'End Type
'
'Public Type timeval
'   tv_sec As Long
'   tv_usec As Long
'End Type
'
'    Public Type hostent
'        h_name As Long          'Official Name of the host (PC).
'        h_aliases As Long       'Null-terminated array of alternate Names.
'        h_addrtype As Integer   'Type of address being returned.
'        h_length As Integer     'Length of each address, in bytes.
'        h_addr_list As Long     'Null-terminated list of addresses for the host.
'    End Type

'    Public Type protoent
'        p_Name As Long
'        p_aliases As Long
'        p_proto As Integer
'    End Type
'
'    Public Type servent
'        s_Name As Long
'        s_aliases As Long
'        s_port As Integer
'        s_proto As Long
'    End Type
'
'
'    Public Type WSAData                 'The members of the Windows Sockets WSADATA structure are:
'        wVersion As Integer             'Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
'        wHighVersion As Integer         'Highest version of the Windows Sockets specification that this .dll can support (also encoded as above).
'        szDescription As String * 257   'Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows Sockets implementation.
'        szSystemStatus As String * 129  'Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or configuration information.
'        iMaxSockets As Integer          'Retained for backward compatibility, but should be ignored for Windows Sockets version 2 and later, as no single value can be appropriate for all underlying service providers.
'        iMaxUdpDg As Integer            'Ignored for Windows Sockets version 2 and onward.
'        lpVendorInfo As Long            'Ignored for Windows Sockets version 2 and onward.
'    End Type
'
'
'    'Constants
'
'    Public Const AF_UNSPEC = 0      'unspecified
'    'Although  AF_UNSPEC  is  defined for backwards compatibility, using
'    'AF_UNSPEC for the "af" parameter when creating a socket is STRONGLY
'    'DISCOURAGED.    The  interpretation  of  the  "protocol"  parameter
'    'depends  on the actual address family chosen.  As environments grow
'    'to  include  more  and  more  address families that use overlapping
'    'protocol  values  there  is  more  and  more  chance of choosing an
'    'undesired address family when AF_UNSPEC is used.
'    Public Const AF_UNIX = 1        'local to host (pipes, portals)
'    Public Const AF_INET = 2        'internetwork: UDP, TCP, etc.
'    Public Const AF_IMPLINK = 3     'arpanet imp addresses
'    Public Const AF_PUP = 4         'pup protocols: e.g. BSP
'    Public Const AF_CHAOS = 5       'mit CHAOS protocols
'    Public Const AF_NS = 6          'XEROX NS protocols
'    Public Const AF_IPX = AF_NS     'IPX protocols: IPX, SPX, etc.
'    Public Const AF_ISO = 7         'ISO protocols
'    Public Const AF_OSI = AF_ISO    'OSI is ISO
'    Public Const AF_ECMA = 8        'european computer manufacturers
'    Public Const AF_DATAKIT = 9     'datakit protocols
'    Public Const AF_CCITT = 10      'CCITT protocols, X.25 etc
'    Public Const AF_SNA = 11        'IBM SNA
'    Public Const AF_DECnet = 12     'DECnet
'    Public Const AF_DLI = 13        'Direct data link interface
'    Public Const AF_LAT = 14        'LAT
'    Public Const AF_HYLINK = 15     'NSC Hyperchannel
'    Public Const AF_APPLETALK = 16  'AppleTalk
'    Public Const AF_NETBIOS = 17    'NetBios-style addresses
'    Public Const AF_VOICEVIEW = 18  'VoiceView
'    Public Const AF_FIREFOX = 19    'Protocols from Firefox
'    Public Const AF_UNKNOWN1 = 20   'Somebody is using this!
'    Public Const AF_BAN = 21        'Banyan
'    Public Const AF_ATM = 22        'Native ATM Services
'    Public Const AF_INET6 = 23      'Internetwork Version 6
'    Public Const AF_CLUSTER = 24    'Microsoft Wolfpack
'    Public Const AF_12844 = 25      'IEEE 1284.4 WG AF
'    Public Const AF_IRDA = 26       'IrDA
'    Public Const AF_NETDES = 28     'Network Designers OSI & gateway enabled protocols
'
'    Public Const FD_READ == 0
'    Public Const FD_READ = FD_READ_BIT
'    Public Const FD_WRITE == 1
'    Public Const FD_WRITE = FD_WRITE_BIT
'    Public Const FD_OOB == 2
'    Public Const FD_OOB = FD_OOB_BIT
'    Public Const FD_ACCEPT == 3
'    Public Const FD_ACCEPT = FD_ACCEPT_BIT
'    Public Const FD_CONNECT == 4
'    Public Const FD_CONNECT = FD_CONNECT_BIT
'    Public Const FD_CLOSE == 5
'    Public Const FD_CLOSE = FD_CLOSE_BIT
'    Public Const FD_QOS == 6
'    Public Const FD_QOS = FD_QOS_BIT
'    Public Const FD_GROUP_QOS == 7
'    Public Const FD_GROUP_QOS = FD_GROUP_QOS_BIT
'    Public Const FD_ROUTING_INTERFACE_CHANGE == 8
'    Public Const FD_ROUTING_INTERFACE_CHANGE = FD_ROUTING_INTERFACE_CHANGE_BIT
'    Public Const FD_ADDRESS_LIST_CHANGE == 9
'    Public Const FD_ADDRESS_LIST_CHANGE = FD_ADDRESS_LIST_CHANGE_BIT
'    Public Const FD_MAX_EVENTS = 10
'    Public Const FD_ALL_EVENTS = FD_MAX_EVENTS - 1
'
'    Public Const INVALID_SOCKET = &HFFFF
'    Public Const SOCKET_ERROR = -1
'
'    Public Const IPPORT_ECHO = 7
'    Public Const IPPORT_DISCARD = 9
'    Public Const IPPORT_SYSTAT = 11
'    Public Const IPPORT_DAYTIME = 13
'    Public Const IPPORT_NETSTAT = 15
'    Public Const IPPORT_FTP = 21
'    Public Const IPPORT_TELNET = 23
'    Public Const IPPORT_SMTP = 25
'    Public Const IPPORT_TIMESERVER = 37
'    Public Const IPPORT_NAMESERVER = 42
'    Public Const IPPORT_WHOIS = 43
'    Public Const IPPORT_MTP = 57
'
'    Public Const IPPORT_TFTP = 69
'    Public Const IPPORT_RJE = 77
'    Public Const IPPORT_FINGER = 79
'    Public Const IPPORT_TTYLINK = 87
'    Public Const IPPORT_SUPDUP = 95
'
'    Public Const IPPORT_EXECSERVER = 512
'    Public Const IPPORT_LOGINSERVER = 513
'    Public Const IPPORT_CMDSERVER = 514
'    Public Const IPPORT_EFSSERVER = 520
'
'    Public Const IPPORT_BIFFUDP = 512
'    Public Const IPPORT_WHOSERVER = 513
'    Public Const IPPORT_ROUTESERVER = 520   '521 also used
'    Public Const IPPORT_RESERVED = 1024     'Ports < IPPORT_RESERVED are reserved for privileged processes (e.g. root).
'
'    Public Const IPPROTO_IP = 0         'dummy for IP
'    Public Const IPPROTO_ICMP = 1       'control message protocol
'    Public Const IPPROTO_IGMP = 2       'internet group management protocol
'    Public Const IPPROTO_GGP = 3        'gateway^2 (deprecated)
'    Public Const IPPROTO_TCP = 6        'tcp
'    Public Const IPPROTO_PUP = 12       'pup
'    Public Const IPPROTO_UDP = 17       'user datagram protocol
'    Public Const IPPROTO_IDP = 22       'xns idp
'    Public Const IPPROTO_ND = 77        'UNOFFICIAL net disk proto
'    Public Const IPPROTO_RAW = 255      'raw IP packet
'    Public Const IPPROTO_MAX = 256
'
'    Public Const MSG_OOB = &H1          'process out-of-band data
'    Public Const MSG_PEEK = &H2         'peek at incoming message
'    Public Const MSG_DONTROUTE = &H4    'send without using routing tables
'    Public Const MSG_PARTIAL = &H8000   'partial send or recv for message xport
'
'    Public Const SO_DEBUG = &H1             'turn on debugging info recording
'    Public Const SO_ACCEPTCONN = &H2        'socket has had listen()
'    Public Const SO_REUSEADDR = &H4         'allow local address reuse
'    Public Const SO_KEEPALIVE = &H8         'keep connections alive
'    Public Const SO_DONTROUTE = &H10        'just use interface addresses
'    Public Const SO_BROADCAST = &H20        'permit sending of broadcast msgs
'    Public Const SO_USELOOPBACK = &H40      'bypass hardware when possible
'    Public Const SO_LINGER = &H80           'linger on close if data present
'    Public Const SO_OOBINLINE = &H100       'leave received OOB data in line

'    Public Const SO_SNDBUF = &H1001        'send buffer size
    'Public Const SO_RCVBUF = &H1002        'receive buffer size
'    Public Const SO_SNDLOWAT = &H1003      'send low-water mark
'    Public Const SO_RCVLOWAT = &H1004      'receive low-water mark
'    Public Const SO_SNDTIMEO = &H1005      'send timeout
'    Public Const SO_RCVTIMEO = &H1006      'receive timeout
'    Public Const SO_ERROR = &H1007         'get error status and clear
'    Public Const SO_TYPE = &H1008          'get socket type
'
'    Public Const SO_GROUP_ID = &H2001           'ID of a socket group
'    Public Const SO_GROUP_PRIORITY = &H2002     'the relative priority within a group
'    Public Const SO_MAX_MSG_SIZE = &H2003       'maximum message size
'    Public Const SO_PROTOCOL_INFOA = &H2004     'WSAPROTOCOL_INFOA structure
'    Public Const SO_PROTOCOL_INFOW = &H2005     'WSAPROTOCOL_INFOW structure
'    Public Const PVD_CONFIG = &H3001            'configuration info for service provider
'    Public Const SO_CONDITIONAL_ACCEPT = &H3002 'enable true conditional accept connection is not ack-ed to the other side until conditional function returns CF_ACCEPT
''
'    Public Const SOCK_STREAM = 1        'stream socket
'    Public Const SOCK_DGRAM = 2         'datagram socket
'    Public Const SOCK_RAW = 3           'raw-protocol interface
'    Public Const SOCK_RDM = 4           'reliably-delivered message
'    Public Const SOCK_SEQPACKET = 5     'sequenced packet stream
'
'    Public Const SOL_SOCKET = &HFFFF
'
'    Public Const WSA_DESCRIPTION_LEN = 256 'Upto 256 char
'    Public Const WSA_SYS_STATUS_LEN = 128

Private pWinsockControl As Long



'Public Const x_DEFAULT = 0
'
'Public Const x_SEND = 0
'Public Const x_WAITALL = -2
'
'Public Const x_READ = 0
'Public Const x_PARTIAL = 1
'
'Public Const x_OOB = 0
'Public Const x_DONTROUTE = 4
'Public Const x_PEEK = -4
'
'Public Const x_SEND_x_READ = x_SEND Or x_READ '0
'Public Const x_WAITALL_x_READ = x_WAITALL Or x_READ '-2
'
'Public Const x_SEND_x_PARTIAL = x_SEND Or x_PARTIAL  ' 1
'Public Const x_WAITALL_x_PARTIAL = x_WAITALL Or x_PARTIAL  '-1
'
'Public Const x_OOB_x_SEND_x_READ = x_OOB Or x_SEND Or x_READ ' 0
'Public Const x_DONTROUTE_x_SEND_x_READ = x_DONTROUTE Or x_SEND Or x_READ  ' 4
'Public Const x_PEEK_x_SEND_x_READ = x_PEEK Or x_SEND Or x_READ  ' -4
'
'Public Const x_OOB_x_WAITALL_x_READ = x_OOB Or x_WAITALL Or x_READ  '-2
'Public Const x_DONTROUTE_x_WAITALL_x_READ = x_DONTROUTE Or x_WAITALL Or x_READ  ' 2
'Public Const x_PEEK_x_WAITALL_x_READ = x_PEEK Or x_WAITALL Or x_READ   '-6
'
'Public Const x_OOB_x_SEND_x_PARTIAL = x_OOB Or x_SEND Or x_PARTIAL  ' 1
'Public Const x_DONTROUTE_x_SEND_x_PARTIAL = x_DONTROUTE Or x_SEND Or x_PARTIAL  ' 5
'Public Const x_PEEK_x_SEND_x_PARTIAL = x_PEEK Or x_SEND Or x_PARTIAL  '-3
'
'Public Const x_OOB_x_WAITALL_x_PARTIAL = x_OOB Or x_WAITALL Or x_PARTIAL   '-1
'Public Const x_DONTROUTE_x_WAITALL_x_PARTIAL = x_DONTROUTE Or x_WAITALL Or x_PARTIAL  '3
'Public Const x_PEEK_x_WAITALL_x_PARTIAL = x_PEEK Or x_WAITALL Or x_PARTIAL  '-5
'
'Public Const x_BOUNDARY = -7 Or 7



Public Property Get WinsockControl() As Boolean
    WinsockControl = (pWinsockControl <> 0)
End Property

#If modBitValue = 0 Then

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
#End If

Public Function LoLong(ByVal lParam As Double) As Long
    If (lParam And &H10FFFFFF) > &H8800000 Then
        LoLong = (lParam And &H10FFFFFF) - (&H11000000 * (lParam \ &H11000000))
    Else
        LoLong = lParam And &H10FFFFFF
    End If
End Function

Public Function HiLong(ByVal lParam As Double) As Long
    If ((lParam And &HEF000000) \ &H11000000) + (lParam Mod &H11000000) < 0 Then
        HiLong = -(((lParam And &HEF000000) \ &H11000000) + (lParam Mod &H11000000))
    Else
        HiLong = ((lParam And &HEF000000) \ &H11000000)
    End If
End Function

Public Function HiInt(ByVal lParam As Long) As Integer
    If (lParam And &HFFFF&) > &H7FFF Then
        HiInt = (lParam And &HFFFF&) - &H10000
    Else
        HiInt = lParam And &HFFFF&
    End If
End Function

Public Function LoInt(ByVal lParam As Long) As Integer
    LoInt = (lParam And &HFFFF0000) \ &H10000
End Function


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

Function SocketsInitialize() As Long
On Error GoTo catch
    If pWinsockControl = 0 Then
        Dim WSAD As WSADATA
        Dim iReturn As Integer
        Dim sLowByte As String, sHighByte As String, sMsg As String
        Dim sckOk As Long
        sckOk = 0
        
        iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
        
        If iReturn <> 0 Then
            sckOk = 1
        Else
            If LoByte(WSAD.wVersion) < WS_SSL3_VERSION Or (LoByte(WSAD.wVersion) = _
                WS_SSL3_VERSION And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
                sHighByte = Trim$(str$(HiByte(WSAD.wVersion)))
                sLowByte = Trim$(str$(LoByte(WSAD.wVersion)))
                sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
                sMsg = sMsg & " is not supported by winsock.dll "
                sckOk = 2
            Else
                pWinsockControl = pWinsockControl + 1
            End If
        End If
        SocketsInitialize = sckOk

    End If
       
            
Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Function SocketsCleanUp() As Long
On Error GoTo catch
    
    
    If pWinsockControl > 0 Then
        pWinsockControl = pWinsockControl - 1
        If pWinsockControl = 0 Then
            Dim lReturn As Long
            Dim sckOk As Long
            sckOk = 0
            lReturn = WSACleanup()
            Debug.Print "WSACleanup"
            If lReturn <> 0 Then
                sckOk = 1
            End If
            SocketsCleanUp = sckOk
        End If
    End If
    
Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function

Public Function GetPortIP(Optional ByVal Domain As String = "") As Collection
On Error GoTo catch

    Dim IPList As New Collection

    Dim init As Boolean
    Dim retVal As Long
    If Not WinsockControl Then
        init = True
        retVal = SocketsInitialize()
    End If
    If retVal = 0 Then
            
        Dim phe As Long
        Dim heDestHost As hostent
        Dim addrList As Long
        Dim rc As Long
        Dim i As Integer
        Dim ip_address As String
        Dim HostName As String * 256
        Dim Hostent_addr As Long
        Dim Host As hostent
        Dim hostip_addr As Long
        Dim temp_ip_address() As Byte
            
        If Domain = "" Then
            HostName = Space(256)
            If gethostname(HostName, 256) = Socket_ERROR Then
                retVal = 1
            Else
                HostName = Trim$(HostName)
                Hostent_addr = GetHostByName(HostName)
    
                If Hostent_addr = 0 Then
                    retVal = 2
                Else
                    
                    CopyMemoryHost Host, Hostent_addr, LenB(Host)
                    CopyMemory hostip_addr, Host.h_addr_list, 4
        
                    Do
                        ReDim temp_ip_address(1 To Host.h_length)
                        RtlMoveMemory3 temp_ip_address(1), hostip_addr, Host.h_length
        
                        For i = 1 To Host.h_length
                            ip_address = ip_address & temp_ip_address(i) & "."
                        Next
                        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        
                        IPList.Add ip_address
        
                        ip_address = ""
                        Host.h_addr_list = Host.h_addr_list + LenB(Host.h_addr_list)
                        RtlMoveMemory3 hostip_addr, Host.h_addr_list, 4
                    Loop While (hostip_addr <> 0)
                    
                End If
    
            End If
        Else
                
            rc = inet_addr(Domain)
            If rc = Socket_ERROR Then
        
                phe = GetHostByName(Domain)
                If phe <> 0 Then
        
                    CopyMemoryHost heDestHost, phe, hostent_size
                    CopyMemory addrList, heDestHost.h_addr_list, 4
                    CopyMemory hostip_addr, addrList, heDestHost.h_length
                    rc = hostip_addr
        
                    ReDim temp_ip_address(1 To heDestHost.h_length)
                    RtlMoveMemory2 temp_ip_address(1), hostip_addr, heDestHost.h_length
        
                    For i = 1 To heDestHost.h_length
                        ip_address = ip_address & temp_ip_address(i) & "."
                    Next
                    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
        
                    IPList.Add ip_address
        
                Else
                    rc = INADDR_NONE
                End If
        
            End If
               
        End If
        
        If init Then
            retVal = SocketsCleanUp()
        End If

        Erase temp_ip_address
    End If
    
    Set GetPortIP = IPList

Exit Function
catch:
    Err.Raise Err.Number, App.EXEName, Err.Description
End Function


Public Function ResolveIP(ByVal Host As String) As String

    Dim init As Boolean
    Dim retVal As Long
    If Not WinsockControl Then
        init = True
        retVal = SocketsInitialize()
    End If
    If retVal = 0 Then
    
        Dim phe As Long
        Dim heDestHost As hostent
        Dim addrList As Long
        Dim rc As Long
    
        Dim hostip_addr As Long
        
        Dim temp_ip_address() As Byte
        Dim i As Integer
        Dim ip_address As String
    
        rc = inet_addr(Host)
        If rc = Socket_ERROR Then
        
            phe = GetHostByName(Host)
            If phe <> 0 Then
            
                CopyMemoryHost heDestHost, phe, hostent_size
                CopyMemory addrList, heDestHost.h_addr_list, 4
                CopyMemory hostip_addr, addrList, heDestHost.h_length
                rc = hostip_addr
                
                ReDim temp_ip_address(1 To heDestHost.h_length)
                RtlMoveMemory2 temp_ip_address(1), hostip_addr, heDestHost.h_length
    
                For i = 1 To heDestHost.h_length
                    ip_address = ip_address & temp_ip_address(i) & "."
                Next
                ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    
                ResolveIP = ip_address
                
            Else
                rc = INADDR_NONE
            End If
            
        End If
    
        If init Then
            retVal = SocketsCleanUp()
        End If

        Erase temp_ip_address
    End If
    
End Function

Public Function Resolve(ByVal Host As String) As Long
   
    Dim init As Boolean
    Dim retVal As Long
    If Not WinsockControl Then
        init = True
        retVal = SocketsInitialize()
    End If
    If retVal = 0 Then
        
        Dim phe As Long
        Dim heDestHost As hostent
        Dim addrList As Long
        Dim rc As Long
    
        rc = inet_addr(Host)
        If rc <> Socket_ERROR Then
        
            phe = GetHostByName(Host)
            If phe <> 0 Then
            
                CopyMemoryHost heDestHost, phe, hostent_size
                CopyMemory addrList, heDestHost.h_addr_list, 4
                CopyMemory rc, addrList, heDestHost.h_length
                
            Else
                rc = INADDR_NONE
            End If
            
        End If
            
        If init Then
            retVal = SocketsCleanUp()
        End If

    End If
    Resolve = rc
    
End Function

Public Function LocalHost() As String

    Dim init As Boolean
    Dim retVal As Long
    If Not WinsockControl Then
        init = True
        retVal = SocketsInitialize()
    End If
    If retVal = 0 Then
    
        
        Dim Buf As String
        Dim rc As Long
        Buf = Space$(255)
        rc = gethostname(Buf, Len(Buf))
        rc = InStr(Buf, vbNullChar)
        If rc > 0 Then
            LocalHost = Left$(Buf, rc - 1)
        Else
            LocalHost = ""
        End If
    
        If init Then
            retVal = SocketsCleanUp()
        End If

    End If

End Function

Public Function ipaddressBySocket(ByVal sock As Long) As VBA.Collection
    Set ipaddressBySocket = ipaddressByhost(ntohl(sock))
End Function

Public Function ipaddressByhost(ByVal sock As Long) As VBA.Collection
    
        Dim sck As sockaddr
        If getsockname(sock, sck, LenB(sck)) = 0 Then
            Dim col As VBA.Collection
            Set col = IPAddress(sck.sin_addr)
            If col.Count > 0 Then
                Set ipaddressByhost = col
         '   Else
        '        RemoteIP = "#INVALID#"
            End If
'        Else
'            RemoteIP = Whois(Handle)
        End If
    
'    Set ipaddressByhost = IPAddress(inet_ntoa(sock))
End Function


Public Function IPAddress(ByVal Addr As Long) As VBA.Collection

    Dim IPList As New VBA.Collection

   ' Dim heDestHost As HOSTENT
    'Dim addrList As Long
    'Dim ret As String
    Dim rc As Long
    Dim i As Integer
    Dim ip_address As String
    'Dim HostName As String * 256
    Dim Host As hostent
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
        

'    Dim tmp As String
'    tmp = String(256, Chr(0))
'
'
'        If rc <> 0 Then
' '       Addr = GetHostByName(host)
''
'    rc = inet_addr(Addr)
'    If rc = Socket_ERROR Then
'        rc = gethostname(tmp, 256)
'        If rc = 0 Then
'            Addr = GetHostByName(tmp)
'       End If
'    Else
'
'
'
'    End If
'    If rc = 0 Then
     rc = gethostbyaddr(Addr, LenB(Host), IPPROTO_TCP)
  '  End If
    
    If rc <> 0 Then
'        phe = gethostbyName(Domain)
'        If phe <> 0 Then
                
        CopyMemoryHost Host, rc, LenB(Host)
        CopyMemory hostip_addr, Host.h_addr_list, 4

        Do
            ReDim temp_ip_address(1 To Host.h_length)
            RtlMoveMemory3 temp_ip_address(1), hostip_addr, Host.h_length

            For i = 1 To Host.h_length
                ip_address = ip_address & temp_ip_address(i) & "."
            Next
            ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

            IPList.Add ip_address

            ip_address = ""
            Host.h_addr_list = Host.h_addr_list + LenB(Host.h_addr_list)
            RtlMoveMemory3 hostip_addr, Host.h_addr_list, 4
        Loop While (hostip_addr <> 0)
    End If


    
    Set IPAddress = IPList

End Function

'Public Function GetRemoteInfo(ByVal lngSocket As Long, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String) As Boolean
'    'Retrieves remote info from a connected socket.
'    'If succeeds returns TRUE and loads the arguments.
'    'If fails returns FALSE and arguments are not loaded.
'    Dim lRet As Long
'    Dim tmpSa As saddr
'    GetRemoteInfo = False
'    lRet = getpeermame(lngSocket, tmpSa, LenB(tmpSa))
'    If lRet = 0 Then
'        GetRemoteInfo = True
'        GetRemoteInfoFromSI tmpSa, m_lngRemotePort, m_strRemoteHostIP, m_strRemoteHost
'    Else
'       lngRemotePort = 0
'       strRemoteHostIP = ""
'       strRemoteHost = ""
'    End If
'End Function
'Private Sub GetRemoteInfoFromSI(ByRef newSa As saddr, ByRef lngRemotePort As Long, ByRef strRemoteHostIP As String, ByRef strRemoteHost As String)
'    'Gets remote info from a sockaddr_in structure.
'    Dim lRet As Long
'    Dim Sin4 As sockaddr_in
'
'    Dim aLen As Long
'    Dim bBuffer() As Byte
'
'
'        aLen = sockaddr_size
'        CopyMemory Sin4, newSa, LenB(Sin4)  'Save to sockaddr_in
'        lngRemotePort = IntegerToUnsigned(ntohs(Sin4.sin_port))
'        ReDim bBuffer(0 To aLen - 1)  'Resize string buffer
'        'Get IPv4 address as string
'        lRet = inet_ntop(AF_INET, Sin4.sin_addr, bBuffer(0), aLen)
'
'    If lRet Then strRemoteHostIP = StringFromPointer(lRet)
'    m_strRemoteHost = ""
'End Sub
'Private Function StringFromPointer(ByVal lPointer As Long) As String
'    'Receives a string pointer and it turns it into a regular string.
'    Dim strTemp As String
'    Dim lRetVal As Long
'    strTemp = String$(lstrlenA(ByVal lPointer), 0)
'    lRetVal = lstrcpyA(ByVal strTemp, ByVal lPointer)
'    If lRetVal Then StringFromPointer = strTemp
'End Function

Public Function UnsignedToLong(Value As Double) As Long
 '
 ' This function takes a Double containing a value in the*
 ' range of an unsigned Long and returns a Long that you*
 ' can pass to an API that requires an unsigned Long
 '
 If Value < 0 Or Value >= modSockets.OFFSET_4 Then Error 6 ' Overflow
    
 If Value <= MAXINT_4 Then
   UnsignedToLong = Value
 Else
   UnsignedToLong = Value - OFFSET_4
 End If
End Function
Public Function LongToUnsigned(Value As Long) As Double
 '
 ' This function takes an unsigned Long from an API and*
 ' converts it to a Double for display or arithmetic purposes
 '
 If Value < 0 Then
   LongToUnsigned = Value + OFFSET_4
 Else
   LongToUnsigned = Value
 End If
End Function
Public Function UnsignedToInteger(Value As Long) As Integer
 '
 ' This function takes a Long containing a value in the range*
 ' of an unsigned Integer and returns an Integer that you*
 ' can pass to an API that requires an unsigned Integer
 '
 If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
   
 If Value <= MAXINT_2 Then
   UnsignedToInteger = Value
 Else
   UnsignedToInteger = Value - OFFSET_2
 End If
End Function
Public Function IntegerToUnsigned(Value As Integer) As Long
 '
 ' This function takes an unsigned Integer from and API and*
 ' converts it to a Long for display or arithmetic purposes
 '
 If Value < 0 Then
   IntegerToUnsigned = Value + OFFSET_2
 Else
   IntegerToUnsigned = Value
 End If
End Function

Public Function GetErrorDescription(ByVal lngErrorCode As Long) As String
 Dim strDesc As String

 Select Case lngErrorCode
   Case WSAEACCES
     strDesc = "Permission denied."
   Case WSAEADDRINUSE
     strDesc = "Address already in use."
   Case WSAEADDRNOTAVAIL
     strDesc = "Cannot assign requested address."
   Case WSAEAFNOSUPPORT
     strDesc = "Address family not supported by protocol family."
   Case WSAEALREADY
     strDesc = "Operation already in progress."
   Case WSAECONNABORTED
     strDesc = "Software caused connection abort."
   Case WSAECONNREFUSED
     strDesc = "Connection refused."
   Case WSAECONNRESET
     strDesc = "Connection reset by peer."
   Case WSAEDESTADDRREQ
     strDesc = "Destination address required."
   Case WSAEFAULT
     strDesc = "Bad address."
   Case WSAEHOSTDOWN
     strDesc = "Host is down."
   Case WSAEHOSTUNREACH
     strDesc = "No route to host."
   Case WSAEINPROGRESS
     strDesc = "Operation now in progress."
   Case WSAEINTR
     strDesc = "Interrupted function call."
   Case WSAEINVAL
     strDesc = "Invalid argument."
   Case WSAEISCONN
     strDesc = "Socket is already connected."
   Case WSAEMFILE
     strDesc = "Too many open files."
   Case WSAEMSGSIZE
     strDesc = "Message too long."
   Case WSAENETDOWN
     strDesc = "Network is down."
   Case WSAENETRESET
     strDesc = "Network dropped connection on reset."
   Case WSAENETUNREACH
     strDesc = "Network is unreachable."
   Case WSAENOBUFS
     strDesc = "No buffer space available."
   Case WSAENOPROTOOPT
     strDesc = "Bad protocol option."
   Case WSAENOTCONN
     strDesc = "Socket is not connected."
   Case WSAENOTSOCK
     strDesc = "Socket operation on nonsocket."
   Case WSAEOPNOTSUPP
     strDesc = "Operation not supported."
   Case WSAEPFNOSUPPORT
     strDesc = "Protocol family not supported."
   Case WSAEPROCLIM
     strDesc = "Too many processes."
   Case WSAEPROTONOSUPPORT
     strDesc = "Protocol not supported."
   Case WSAEPROTOTYPE
     strDesc = "Protocol wrong type for socket."
   Case WSAESHUTDOWN
     strDesc = "Cannot send after socket shutdown."
   Case WSAESOCKTNOSUPPORT
     strDesc = "Socket type not supported."
   Case WSAETIMEDOUT
     strDesc = "Connection timed out."
   Case WSATYPE_NOT_FOUND
     strDesc = "Class type not found."
   Case WSAEWOULDBLOCK
     strDesc = "Resource temporarily unavailable."
   Case WSAHOST_NOT_FOUND
     strDesc = "Host not found."
   Case WSANOTINITIALISED
     strDesc = "Successful WSAStartup not yet performed."
   Case WSANO_DATA
     strDesc = "Valid Name, no data record of requested type."
   Case WSANO_RECOVERY
     strDesc = "This is a nonrecoverable error."
   Case WSASYSCALLFAILURE
     strDesc = "System call failure."
   Case WSASYSNOTREADY
     strDesc = "Network subsystem is unavailable."
   Case WSATRY_AGAIN
     strDesc = "Nonauthoritative host not found."
   Case WSAVERNOTSUPPORTED
     strDesc = "Winsock.dll version out of range."
   Case WSAEDISCON
     strDesc = "Graceful shutdown in progress."
   Case Else
     strDesc = "Unknown error."
 End Select

 GetErrorDescription = strDesc
End Function

Public Function IP4ToIP2(ByVal str As String) As String

    Dim b1 As Byte
    Dim B2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte

    b1 = CByte(RemoveNextArg(str, "."))
    B2 = CByte(RemoveNextArg(str, "."))
    b3 = CByte(RemoveNextArg(str, "."))
    b4 = CByte(RemoveNextArg(str, "."))

    
    IP4ToIP2 = IntegerToUnsigned(Val("&H" & modCommon.Padding(2, Hex(b1), "0") & modCommon.Padding(2, Hex(B2), "0"))) & ", " & IntegerToUnsigned(Val("&H" & modCommon.Padding(2, Hex(b3), "0") & modCommon.Padding(2, Hex(b4), "0")))

End Function

Public Function IP2ToIP4(ByVal str As String) As String

    Dim i1 As Integer
    Dim i2 As Integer
    i1 = UnsignedToInteger(RemoveNextArg(str, ","))
    i2 = UnsignedToInteger(RemoveNextArg(str, ","))
    
    IP2ToIP4 = Val("&H" & Left(modCommon.Padding(4, Hex(i1), "0"), 2)) & "." & Val("&H" & Right(modCommon.Padding(4, Hex(i1), "0"), 2)) & "." & Val("&H" & Left(modCommon.Padding(4, Hex(i2), "0"), 2)) & "." & Val("&H" & Right(modCommon.Padding(4, Hex(i2), "0"), 2))

End Function
