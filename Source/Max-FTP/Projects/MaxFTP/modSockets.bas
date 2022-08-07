Attribute VB_Name = "modSockets"
#Const modSockets = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32" () As Long
Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function gethostname Lib "wsock32" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "wsock32" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)

Private Function SocketsInitialize() As Long

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

End Function

Private Function SocketsCleanup() As Long

Dim lReturn As Long
Dim sckOk As Integer
sckOk = 0
lReturn = WSACleanup()

If lReturn <> 0 Then
    sckOk = 1
    End If
SocketsCleanup = sckOk

End Function

Public Sub GetAllIPs(ByRef IPList)

    Dim retVal As Long
    retVal = SocketsInitialize()
    If retVal = 0 Then
        Dim hostname As String * 256
        Dim hostent_addr As Long
        Dim host As HOSTENT
        Dim hostip_addr As Long
        Dim temp_ip_address() As Byte
        Dim i As Integer
        Dim ip_address As String
    
        If gethostname(hostname, 256) = SOCKET_ERROR Then
            retVal = 1
        Else
            hostname = Trim$(hostname)
            hostent_addr = gethostbyname(hostname)
    
            If hostent_addr = 0 Then
                retVal = 2
            Else
                
                RtlMoveMemory host, hostent_addr, LenB(host)
                RtlMoveMemory hostip_addr, host.hAddrList, 4

                Do
                    ReDim temp_ip_address(1 To host.hLength)
                    RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
    
                    For i = 1 To host.hLength
                        ip_address = ip_address & temp_ip_address(i) & "."
                        Next
                    ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
    
                    IPList.AddItem ip_address
    
                    ip_address = ""
                    host.hAddrList = host.hAddrList + LenB(host.hAddrList)
                    RtlMoveMemory hostip_addr, host.hAddrList, 4
                Loop While (hostip_addr <> 0)
                retVal = SocketsCleanup()
            End If
        
        End If
    End If

End Sub



