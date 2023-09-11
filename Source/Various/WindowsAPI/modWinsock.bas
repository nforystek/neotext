Attribute VB_Name = "modWinsock"

Option Explicit



' *************************************************************************************************
' Accept
' *************************************************************************************************
'
' The accept function permits an incoming connection attempt on a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
' s:            [in] Descriptor identifying a socket that has been placed in a listening state
'               with the listen function. The connection is actually made with the socket that is
'               returned by accept.
'
' addr:         [out] Optional pointer to a buffer that receives the address of the connecting
'               entity, as known to the communications layer. The exact format of the addr
'               parameter is determined by the address family that was established when the socket
'               from the sockaddr structure was created.
'
' addrlen:      [out] Optional pointer to an integer that contains the length of addr.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new
' socket. This returned value is a handle for the socket on which the actual connection is made.
'
' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved
' by calling WSAGetLastError.
'
' The integer referred to by addrlen initially contains the amount of space pointed to by addr.
' On return it will contain the actual length in bytes of the address returned.
'
' *************************************************************************************************

Public Declare Function accept Lib "ws2_32.dll" _
                                    (ByVal s As Long, _
                                     ByRef Addr As API_SOCKADDR_IN, _
                                     ByRef AddrLen As Long) As Long


' *************************************************************************************************
' AcceptEx
' *************************************************************************************************
'
' The AcceptEx function accepts a new connection, returns the local and remote address, and
' receives the first block of data sent by the client application. The AcceptEx function is
' not supported on Windows Me/98/95.
'
' Note This function is a Microsoft-specific extension to the Windows Sockets specification.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' sListenSocket:        [in] Descriptor identifying a socket that has already been called with the
'                       listen function. A server application waits for attempts to connect on
'                       this socket.
' sAcceptSocket:        [in] Descriptor identifying a socket on which to accept an incoming
'                       connection. This socket must not be bound or connected.
' lpOutputBuffer:       [in] Pointer to a buffer that receives the first block of data sent on a
'                       new connection, the local address of the server, and the remote address of
'                       the client. The receive data is written to the first part of the buffer
'                       starting at offset zero, while the addresses are written to the latter
'                       part of the buffer. This parameter must be specified.
' dwReceiveDataLength:  [in] Number of bytes in lpOutputBuffer that will be used for actual
'                       receive data at the beginning of the buffer. This size should not include
'                       the size of the local address of the server, nor the remote address of the
'                       client; they are appended to the output buffer. If dwReceiveDataLength is
'                       zero, accepting the connection will not result in a receive operation.
'                       Instead, AcceptEx completes as soon as a connection arrives, without
'                       waiting for any data.
' dwLocalAddressLength: [in] Number of bytes reserved for the local address information. This
'                       value must be at least 16 bytes more than the maximum address length for
'                       the transport protocol in use.
' dwRemoteAddressLength:[in] Number of bytes reserved for the remote address information. This
'                       value must be at least 16 bytes more than the maximum address length for
'                       the transport protocol in use. Cannot be zero.
' lpdwBytesReceived:    [out] Pointer to a DWORD that receives the count of bytes received. This
'                       parameter is set only if the operation completes synchronously. If it
'                       returns ERROR_IO_PENDING and is completed later, then this DWORD is never
'                       set and you must obtain the number of bytes read from the completion
'                       notification mechanism.
' lpOverlapped:         [in] An OVERLAPPED structure that is used to process the request. This
'                       parameter must be specified; it cannot be null.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, the AcceptEx function completed successfully and a value of TRUE is
' returned.
'
' If the function fails, AcceptEx returns FALSE. The WSAGetLastError function can then be called
' to return extended error information. If WSAGetLastError returns ERROR_IO_PENDING, then the
' operation was successfully initiated and is still in progress.
'
' *************************************************************************************************

Public Declare Function AcceptEx Lib "ws2_32.dll" _
                                    (ByVal sListenSocket As Long, _
                                     ByVal sAcceptSocket As Long, _
                                     ByRef lpOutputBuffer As Any, _
                                     ByVal dwReceiveDataLength As Long, _
                                     ByVal dwLocalAddressLength As Long, _
                                     ByVal dwRemoteAddressLength As Long, _
                                     ByRef lpdwBytesReceived As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED) As Long


' *************************************************************************************************
' Bind
' *************************************************************************************************
'
' The bind function associates a local address with a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying an unbound socket.
' name:                 [in] Address to assign to the socket from the sockaddr structure.
' namelen:              [in] Length of the value in the name parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, bind returns zero. Otherwise, it returns SOCKET_ERROR, and a specific error
' code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function bind Lib "ws2_32.dll" _
                                    (ByVal s As Long, _
                                     ByRef Name As API_SOCKADDR_IN, _
                                     ByRef namelen As Long) As Long


' *************************************************************************************************
' CloseSocket
' *************************************************************************************************
'
' The closesocket function closes an existing socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket to close.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, closesocket returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function closesocket Lib "ws2_32.dll" _
                                    (ByVal s As Long) As Long


' *************************************************************************************************
' Connect
' *************************************************************************************************
'
' The connect function establishes a connection to a specified socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying an unconnected socket.
' name:                 [in] Name of the socket in the sockaddr structure to which the connection
'                       should be established.
' namelen:              [in] Length of name, in bytes
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, connect returns zero. Otherwise, it returns SOCKET_ERROR, and a specific
' error code can be retrieved by calling WSAGetLastError.
'
' On a blocking socket, the return value indicates success or failure of the connection attempt.
'
' With a nonblocking socket, the connection attempt cannot be completed immediately. In this case,
' connect will return SOCKET_ERROR, and WSAGetLastError will return WSAEWOULDBLOCK. In this case,
' there are three possible scenarios:
'
' Use the select function to determine the completion of the connection request by checking to see
' if the socket is writeable.
'
' If the application is using WSAAsyncSelect to indicate interest in connection events, then the
' application will receive an FD_CONNECT notification indicating that the connect operation is
' complete (successfully or not).
'
' If the application is using WSAEventSelect to indicate interest in connection events, then the
' associated event object will be signaled indicating that the connect operation is complete
' (successfully or not).
'
' Until the connection attempt completes on a nonblocking socket, all subsequent calls to connect
' on the same socket will fail with the error code WSAEALREADY, and WSAEISCONN when the connection
' completes successfully. Due to ambiguities in version 1.1 of the Windows Sockets specification,
' error codes returned from connect while a connection is already pending may vary among
' implementations. As a result, it is not recommended that applications use multiple calls to
' connect to detect connection completion. If they do, they must be prepared to handle WSAEINVAL
' and WSAEWOULDBLOCK error values the same way that they handle WSAEALREADY, to assure robust
' execution.
'
' If the error code returned indicates the connection attempt failed (that is, WSAECONNREFUSED,
' WSAENETUNREACH, WSAETIMEDOUT) the application can call connect again for the same socket.
'
' *************************************************************************************************

Public Declare Function connect Lib "ws2_32.dll" _
                                    (ByVal s As Long, _
                                     ByRef Name As API_SOCKADDR_IN, _
                                     ByVal namelen As Long) As Long
                                     

' *************************************************************************************************
' EnumProtocols
' *************************************************************************************************
'
' The EnumProtocols function retrieves information about a specified set of network protocols that
' are active on a local host.
'
' Note The EnumProtocols function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' the reference material is included. The WSAEnumProtocols function provides equivalent
' functionality in Windows Sockets 2.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpiProtocols:         Pointer to a null-terminated array of protocol identifiers. The
'                       EnumProtocols function retrieves information about the protocols specified
'                       by this array.
'
'                       If lpiProtocols is NULL, the function retrieves information about all
'                       available protocols.
'lpProtocolBuffer:      Pointer to a buffer that the function fills with an array of
'                       PROTOCOL_INFO data structures.
'lpdwBufferLength:      Pointer to a variable that, on input, specifies the size, in bytes, of the
'                       buffer pointed to by lpProtocolBuffer.
'
'                       On output, the function sets this variable to the minimum buffer size
'                       needed to retrieve all of the requested information. For the function to
'                       succeed, the buffer must be at least this size.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is the number of PROTOCOL_INFO data structures
' written to the buffer pointed to by lpProtocolBuffer.

' If the function fails, the return value is SOCKET_ERROR(–1). To get extended error information,
' call GetLastError, which returns the following extended error code.
'
' *************************************************************************************************

Public Declare Function EnumProtocols Lib "Wsock32.lib" _
                                    (ByVal lpiProtocols As Long, _
                                     ByRef lpProtocolBuffer As Any, _
                                     ByVal lpdeBufferLength As Long) As Long


' *************************************************************************************************
' GetAcceptExSockaddrs
' *************************************************************************************************
'
' The GetAcceptExSockaddrs function parses the data obtained from a call to the AcceptEx function
' and passes the local and remote addresses to a sockaddr structure.
'
' Note This function is a Microsoft-specific extension to the Windows Sockets specification.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpOutputBuffer:       [in] Pointer to a buffer that receives the first block of data sent on a
'                       connection resulting from an AcceptEx call. Must be the same lpOutputBuffer
'                       parameter that was passed to the AcceptEx function.
' dwReceiveDataLength:  [in] Number of bytes in the buffer used for receiving the first data. This
'                       value must be equal to the dwReceiveDataLength parameter that was passed to
'                       the AcceptEx function.
' dwLocalAddressLength: [in] Number of bytes reserved for the local address information. Must be
'                       equal to the dwLocalAddressLength parameter that was passed to the AcceptEx
'                       function.
' dwRemoteAddressLength:[in] Number of bytes reserved for the remote address information. This
'                       value must be equal to the dwRemoteAddressLength parameter that was passed
'                       to the AcceptEx function.
' LocalSockaddr:        [out] Pointer to the sockaddr structure that receives the local address of
'                       the connection (the same information that would be returned by the
'                       getsockname function). This parameter must be specified.
' LocalSockaddrLength:  [out] Size of the local address, in bytes. This parameter must be specified.
' RemoteSockaddr:       [out] Pointer to the sockaddr structure that receives the remote address of
'                       the connection (the same information that would be returned by the
'                       getpeername function). This parameter must be specified.
' RemoteSockaddrLength: [out] Size of the local address, in bytes. This parameter must be
'                       specified.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' This function does not return a value.
'
' *************************************************************************************************

Public Declare Sub GetAcceptExSockaddrs Lib "wsock32.dll" _
                                    (ByRef lpOutputBuffer As Any, _
                                     ByVal dwReceiveDataLength As Long, _
                                     ByVal dwLocalAddressLength As Long, _
                                     ByVal dwRemoteAddressLength As Long, _
                                     ByRef LocalSockaddr As API_SOCKADDR_IN, _
                                     ByRef LocalSockaddrLength As Long, _
                                     ByRef RemoteSockaddr As API_SOCKADDR_IN, _
                                     ByRef RemoteSockaddrLength As Long)


' *************************************************************************************************
' GetAddressByName
' *************************************************************************************************
'
' The GetAddressByName function queries a namespace, or a set of default namespaces, to retrieve
' network address information for a specified network service. This process is known as service
' name resolution. A network service can also use the function to obtain local address information
' that it can use with the bind function.
'
' Note The GetAddressByName function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' the reference material is as follows. The functions detailed in Protocol-Independent Name
' Resolution provide equivalent functionality in Windows Sockets 2.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' dwNameSpace:          Specifies the namespace, or set of default namespaces, that the operating
'                       system queries for network address information.
'
'                       Most calls to GetAddressByName should use the special value NS_DEFAULT.
'                       This lets a client get by with no knowledge of which namespaces are
'                       available on an internetwork. The system administrator determines
'                       namespace access. Namespaces can come and go without the client having to
'                       be aware of the changes.
'
' lpServiceType:        Pointer to a globally unique identifier (GUID) that specifies the type of
'                       the network service. The header file Svcguid.h includes definitions of
'                       several GUID service types, and macros for working with them.
' lpServiceName:        Pointer to a zero-terminated string that uniquely represents the service
'                       name. For example, "MY SNA SERVER".
'
'                       Setting lpServiceName to NULL is the equivalent of setting dwResolution to
'                       RES_SERVICE. The function operates in its second mode, obtaining the local
'                       address to which a service of the specified type should bind. The function
'                       stores the local address within the LocalAddr member of the CSADDR_INFO
'                       structures stored into *lpCsaddrBuffer.
'
'                       If dwResolution is set to RES_SERVICE, the function ignores the
'                       lpServiceName parameter.
'
'                       If dwNameSpace is set to NS_DNS, *lpServiceName is the name of the host.
' lpiProtocols:         Pointer to a zero-terminated array of protocol identifiers. The function
'                       restricts a name resolution attempt to namespace providers that offer
'                       these protocols. This lets the caller limit the scope of the search.
'
'                       If lpiProtocols is null, the function retrieves information on all
'                       available protocols.
' dwResolution:         Set of bit flags that specify aspects of the service name resolution
'                       process. The following bit flags are defined.
' Value Meaning:        RES_SERVICE If set, the function retrieves the address to which a service
'                       of the specified type should bind. This is the equivalent of setting
'                       lpServiceName to NULL.
'
'                       If this flag is clear, normal name resolution occurs.
'
'                       RES_FIND_MULTIPLE   If this flag is set, the operating system performs an
'                                           extensive search of all namespaces for the service. It
'                                           asks every appropriate namespace to resolve the service
'                                           name. If this flag is clear, the operating system stops
'                                           looking for service addresses as soon as one is found.
'                       RES_SOFT_SEARCH     This flag is valid if the namespace supports multiple
'                                           levels of searching.
'
'                       If this flag is valid and set, the operating system performs a simple and
'                       quick search of the namespace. This is useful if an application only needs
'                       to obtain easy-to-find addresses for the service.
'
'                       If this flag is valid and clear, the operating system performs a more
'                       extensive search of the namespace.
' lpServiceAsyncInfo:   Reserved for future use; must be set to NULL.
' lpCsaddrBuffer:       Pointer to a buffer to receive one or more CSADDR_INFO data structures.
'                       The number of structures written to the buffer depends on the amount of
'                       information found in the resolution attempt. You should assume that
'                       multiple structures will be written, although in many cases there will
'                       only be one.
' lpdwBufferLength:     Pointer to a variable that, upon input, specifies the size, in bytes, of
'                       the buffer pointed to by lpCsaddrBuffer.
'
'                       Upon output, this variable contains the total number of bytes required to
'                       store the array of CSADDR_INFO structures. If this value is less than or
'                       equal to the input value of *lpdwBufferLength, and the function is
'                       successful, this is the number of bytes actually stored in the buffer. If
'                       this value is greater than the input value of *lpdwBufferLength, the
'                       buffer was too small, and the output value of *lpdwBufferLength is the
'                       minimal required buffer size.
' lpAliasBuffer:        Pointer to a buffer to receive alias information for the network service.
'
'                       If a namespace supports aliases, the function stores an array of
'                       zero-terminated name strings into the buffer pointed to by lpAliasBuffer.
'                       There is a double zero-terminator at the end of the list. The first name
'                       in the array is the service's primary name. Names that follow are aliases.
'                       An example of a namespace that supports aliases is DNS.
'
'                       If a namespace does not support aliases, it stores a double zero-terminator
'                       into the buffer.
'
'                       This parameter is optional, and can be set to NULL.
' lpdwAliasBufferLength:Pointer to a variable that, upon input, specifies the size, in bytes, of
'                       the buffer pointed to by lpAliasBuffer.
'
'                       Upon output, this variable contains the total number of bytes required to
'                       store the array of name strings. If this value is less than or equal to
'                       the input value of *lpdwAliasBufferLength, and the function is successful,
'                       this is the number of bytes actually stored in the buffer. If this value
'                       is greater than the input value of *lpdwAliasBufferLength, the buffer was
'                       too small, and the output value of *lpdwAliasBufferLength is the minimal
'                       required buffer size.
'
'                       If lpAliasBuffer is NULL, lpdwAliasBufferLength is meaningless and can
'                       also be NULL.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is the number of CSADDR_INFO data structures written
' to the buffer pointed to by lpCsaddrBuffer.
'
' If the function fails, the return value is SOCKET_ERROR( – 1). To get extended error information,
' call GetLastError, which returns the following extended error value.
'
' *************************************************************************************************

Public Declare Function GetAddressByNameA Lib "wsock32.dll" _
                                    (ByVal dwNameSpace As Long, _
                                     ByRef lpServiceType As API_GUID, _
                                     ByVal lpServiceName As String, _
                                     ByRef lpiProtocols As String, _
                                     ByVal dwResolution As Long, _
                                     ByVal Reserved As Long, _
                                     ByRef lpCsaddrBuffer As API_CSADDR_INFO, _
                                     ByRef lpdwBufferLength As Long, _
                                     ByVal lpAliasBuffer As String, _
                                     ByRef lpdwAliasBufferLength As Long) As Long


' *************************************************************************************************
' GetAddrInfo
' *************************************************************************************************
'
' The getaddrinfo function provides protocol-independent translation from host name to address.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' nodename:             [in] Pointer to a null-terminated string containing a host (node) name or
'                       a numeric host address string. The numeric host address string is a
'                       dotted-decimal IPv4 address or an IPv6 hex address.
' servname:             [in] Pointer to a null-terminated string containing either a service name
'                       or port number.
' hints:                [in] Pointer to an addrinfo structure that provides hints about the type
'                       of socket the caller supports. See Remarks.
' res:                  [out] Pointer to a linked list of one or more addrinfo structures
'                       containing response information about the host.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' Success returns zero. Failure returns a nonzero Windows Sockets error code, as found in the
' Windows Sockets Error Codes.
'
' Nonzero error codes returned by the getaddrinfo function also map to the set of errors outlined
' by IETF recommendations. The following table shows these error codes and their WSA* equivalents.
' It is recommended that the WSA* error codes be used, as they offer familiar and comprehensive
' error information for Winsock programmers.
'
' *************************************************************************************************

Public Declare Function getaddrinfo Lib "ws2_32.dll" _
                                    (ByRef nodename As String, _
                                     ByRef servname As String, _
                                     ByRef hints As API_ADDRINFO, _
                                     ByRef res As Long) As Long


' *************************************************************************************************
' GetHostByAddr
' *************************************************************************************************
'
' The gethostbyaddr function retrieves the host information corresponding to a network address.
'
' Note The gethostbyaddr function has been deprecated by the introduction of the getnameinfo
' function. Developers creating Windows Sockets 2 applications are urged to use the getnameinfo
' function instead of the gethostbyaddr function. See Remarks.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' addr:                 [in] Pointer to an address in network byte order.
' len:                  [in] Length of the address, in bytes.
' type:                 [in] Type of the address, such as the AF_INET address family type
'                       (defined as TCP, UDP, and other associated Internet protocols). Address
'                       family types and their corresponding values are defined in the Winsock2.h
'                       header file.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, gethostbyaddr returns a pointer to the hostent structure. Otherwise, it
' returns a null pointer, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function gethostbyaddr Lib "ws2_32.dll" _
                                    (ByRef Address As Long, _
                                     ByVal AddrLen As Long, _
                                     ByVal AddrType As Long) As Long


' *************************************************************************************************
' GetHostByName
' *************************************************************************************************
'
' The gethostbyname function retrieves host information corresponding to a host name from a host
' database.
'
' Note The gethostbyname function has been deprecated by the introduction of the getaddrinfo
' function. Developers creating Windows Sockets 2 applications are urged to use the getaddrinfo
' function instead of gethostbyname.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' name:                 [in] Pointer to the null-terminated name of the host to resolve.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, gethostbyname returns a pointer to the hostent structure described above.
' Otherwise, it returns a null pointer and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function gethostbyname Lib "ws2_32.dll" _
                                    (ByVal Name As String) As Long



' *************************************************************************************************
' GetHostName
' *************************************************************************************************
'
' The gethostname function retrieves the standard host name for the local computer.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' name:                 [out] Pointer to a buffer that receives the local host name.
' namelen:              [in] Length of the buffer, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, gethostname returns zero. Otherwise, it returns SOCKET_ERROR and a specific
' error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function gethostname Lib "ws2_32.dll" _
                                    (ByVal HostName As String, _
                                     ByVal namelen As Long) As Long


' *************************************************************************************************
' GetNameByType
' *************************************************************************************************
'
' The GetNameByType function retrieves the name of a network service for the specified service
' type.
'
' Note The GetNameByType function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' the reference material is as follows.
'
' Note The functions detailed in Protocol-Independent Name Resolution provide equivalent
' functionality in Windows Sockets 2.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpServiceType:        Pointer to a globally unique identifier (GUID) that specifies the type of
'                       the network service. The header file Svcguid.h includes definitions of
'                       several GUID service types, and macros for working with them.
' lpServiceName:        Pointer to a buffer to receive a zero-terminated string that uniquely
'                       represents the name of the network service.
' dwNameLength:         Pointer to a variable that, on input, specifies the size of the buffer
'                       pointed to by lpServiceName. On output, the variable contains the actual
'                       size of the service name string, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is not SOCKET_ERROR (–1).
'
' If the function fails, the return value is SOCKET_ERROR (–1). To get extended error information,
' call GetLastError.
'
' *************************************************************************************************

Public Declare Function GetNameByTypeA Lib "wsock32.dll" _
                                    (ByRef lpServiceType As API_GUID, _
                                     ByVal lpServiceName As String, _
                                     ByRef dwNameLength As Long) As Long


' *************************************************************************************************
' GetNameInfo
' *************************************************************************************************
'
' The getnameinfo function provides name resolution from an address to the host name.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' sa:                   [in] Pointer to a socket address structure containing the address and port
'                       number of the socket. For IPv4, the sa parameter points to a sockaddr_in
'                       structure; for IPv6, the sa parameter points to a sockaddr_in6 structure.
' salen:                [in] Length of the structure pointed to in the sa parameter, in bytes.
' host:                 [out] Pointer to the host name. The host name is returned as a Fully
'                       Qualified Domain Name (FQDN) by default.
' hostlen:              [in] Length of the buffer pointed to by the host parameter, in bytes. The
'                       caller must provide a buffer large enough to hold the host name, including
'                       terminating NULL characters. A value of zero indicates the caller does not
'                       want to receive the string provided in host.
' serv:                 [out] Pointer to the service name associated with the port number.
' servlen:              [in] Length of the buffer pointed to by the serv parameter, in bytes. The
'                       caller must provide a buffer large enough to hold the service name,
'                       including terminating null characters. A value of zero indicates the caller
'                       does not want to receive the string provided in serv.
' Flags:                [in] Flags used to customize processing of the getnameinfo function. See
'                       Remarks.

' Return Values
' -------------------------------------------------------------------------------------------------
'
' Success returns zero. Any nonzero return value indicates failure. Use the WSAGetLastError
' function to retrieve error information.
'
' *************************************************************************************************

Public Declare Function getnameinfo Lib "ws2_32.dll" _
                                    (ByRef sa As API_SOCKADDR_IN, _
                                     ByVal salen As Long, _
                                     ByRef host As String, _
                                     ByVal hostlen As Long, _
                                     ByRef serv As String, _
                                     ByVal servlen As Long, _
                                     ByVal Flags As Long) As Long


' *************************************************************************************************
' GetPeerName
' *************************************************************************************************
'
' The getpeername function retrieves the name of the peer to which a socket is connected.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' name:                 [out] The SOCKADDR structure that receives the name of the peer.
' NameLen:              [in, out] Pointer to the size of the name structure, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' if no error occurs, getpeername returns zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function getpeername Lib "wsock32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef PearName As API_SOCKADDR_IN, _
                                     ByRef namelen As Long) As Long


' *************************************************************************************************
' GetProtoByName
' *************************************************************************************************
'
' The getprotobyname function retrieves the protocol information corresponding to a protocol name.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' name:                 [in] Pointer to a null-terminated protocol name.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getprotobyname returns a pointer to the protoent. Otherwise, it returns a
' null pointer and a specific error number can be retrieved by calling WSAGetLastError.
'
' Error code            Meaning
' -------------------------------------------------------------------------------------------------
'
' *************************************************************************************************

Public Declare Function getprotobyname Lib "ws2_32.dll" _
                                    (ByVal Name As String) As Long


' *************************************************************************************************
' GetProtoByNumber
' *************************************************************************************************
' The getprotobynumber function retrieves protocol information corresponding to a protocol number.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' number:               [in] Protocol number, in host byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getprotobynumber returns a pointer to the protoent structure. Otherwise,
' it returns a null pointer and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' Error code            Meaning
' -------------------------------------------------------------------------------------------------
'
' *************************************************************************************************

Public Declare Function getprotobynumber Lib "ws2_32.dll" _
                                    (ByVal Number As Long) As Long


' *************************************************************************************************
' GetServByName
' *************************************************************************************************
'
' The getservbyname function retrieves service information corresponding to a service name and
' protocol.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' name:                 [in] Pointer to a null-terminated service name.
' proto:                [in] Optional pointer to a null-terminated protocol name. If this pointer
'                       is NULL, getservbyname returns the first service entry where name matches
'                       the s_name member of the servent structure or the s_aliases member of the
'                       servent structure. Otherwise, getservbyname matches both the name and the
'                       proto.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getservbyname returns a pointer to the servent structure. Otherwise, it
' returns a null pointer and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function getservbyname Lib "ws2_32.dll" _
                                    (ByVal Name As String, _
                                     ByVal Proto As Long) As Long


' *************************************************************************************************
' GetServByPort
' *************************************************************************************************
'
' The getservbyport function retrieves service information corresponding to a port and protocol.

' Parameters
' -------------------------------------------------------------------------------------------------
'
' port:                 [in] Port for a service, in network byte order.
' Proto:                [in] Optional pointer to a protocol name. If this is null, getservbyport
'                       returns the first service entry for which the port matches the s_port of
'                       the servent structure. Otherwise, getservbyport matches both the port and
'                       the proto parameters.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getservbyport returns a pointer to the servent structure. Otherwise, it
' returns a null pointer and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function getservbyport Lib "ws2_32.dll" _
                                    (ByVal Port As Long, _
                                     ByVal Proto As Long) As Long


' *************************************************************************************************
' GetService
' *************************************************************************************************
'
' The GetService function retrieves information about a network service in the context of a set of
' default namespaces or a specified namespace. The network service is specified by its type and
' name. The information about the service is obtained as a set of NS_SERVICE_INFO data structures.
'
' Note The GetService function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' this reference material is included.
'
' Note The functions detailed in Protocol-Independent Name Resolution provide equivalent
' functionality in Windows Sockets 2.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' dwNameSpace:          Specifies the namespace, or a set of default namespaces, that the
'                       operating system queries for information about the specified network
'                       service.
'
'                       Most calls to GetService should use the special value NS_DEFAULT. This
'                       lets a client get by without knowing available namespaces on an
'                       internetwork. The system administrator determines namespace access.
'                       Namespaces can come and go without the client having to be aware of the
'                       changes.
' lpGuiD:               Pointer to a globally unique identifier (GUID) that specifies the type of
'                       the network service. The header file Svcguid.h includes GUID service types
'                       from many well-known services within the DNS and SAP namespaces.
' lpServiceName:        Pointer to a zero-terminated string that uniquely represents the service
'                       name. For example, "MY SNA SERVER."
' dwProperties:         Set of bit flags that specify the service information that the function
'                       retrieves. Each of these bit flag constants, other than PROP_ALL,
'                       corresponds to a particular member of the SERVICE_INFO data structure.
'                       If the flag is set, the function puts information into the corresponding
'                       member of the data structures stored in *lpBuffer.
' lpBuffer:             Pointer to a buffer to receive an array of NS_SERVICE_INFO structures
'                       and associated service information. Each NS_SERVICE_INFO structure
'                       contains service information in the context of a particular namespace.
'                       Note that if dwNameSpace is NS_DEFAULT, the function stores more than one
'                       structure into the buffer; otherwise, just one structure is stored.
'
'                       Each NS_SERVICE_INFO structure contains a SERVICE_INFO structure.
'                       The members of these SERVICE_INFO structures will contain valid data based
'                       on the bit flags that are set in the dwProperties parameter. If a member's
'                       corresponding bit flag is not set in dwProperties, the member's value is
'                       undefined.
'
'                       The function stores the NS_SERVICE_INFO structures in a consecutive array,
'                       starting at the beginning of the buffer. The pointers in the contained
'                       SERVICE_INFO structures point to information that is stored in the buffer
'                       between the end of the NS_SERVICE_INFO structures and the end of the buffer.
' lpdwBufferSize:       Pointer to a variable that, on input, contains the size, in bytes, of the
'                       buffer pointed to by lpBuffer. On output, this variable contains the number
'                       of bytes required to store the requested information. If this output value
'                       is greater than the input value, the function has failed due to insufficient
'                       buffer size.
' lpServiceAsyncInfo:   Reserved for future use. Must be set to null.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is the number of NS_SERVICE_INFO structures stored in
' *lpBuffer. Zero indicates that no structures were stored.
'
' If the function fails, the return value is SOCKET_ERROR ( – 1). To get extended error
' information, call GetLastError, which returns one of the following extended error values.
'
' *************************************************************************************************

Public Declare Function GetServiceA Lib "wsock32.dll" _
                                    (ByVal dwNameSpace As Long, _
                                     ByRef lpGuid As API_GUID, _
                                     ByVal lpServiceName As String, _
                                     ByVal dwProperties As Long, _
                                     ByRef lpBuffer As API_NS_SERVICE_INFO, _
                                     ByRef lpdwBufferSize As Long, _
                                     ByVal Reserved As Long) As Long


' *************************************************************************************************
' GetSockName
' *************************************************************************************************
'
' The getsockname function retrieves the local name for a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' name:                 [out] Pointer to a SOCKADDR structure that receives the address (name) of
'                       the socket.
' NameLen:              [in, out] Size of the name buffer, in bytes
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getsockname returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function getsockname Lib "wsock32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Name As API_SOCKADDR_IN, _
                                     ByRef namelen As Long) As Long


' *************************************************************************************************
' GetSockOpt
' *************************************************************************************************
'
' The getsockopt function retrieves a socket option.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' Level:                [in] Level at which the option is defined; the supported levels include
'                       SOL_SOCKET and IPPROTO_TCP. See Winsock Annexes for more information on
'                       protocol-specific levels.
' optname:              [in] Socket option for which the value is to be retrieved.
' optval:               [out] Pointer to the buffer in which the value for the requested option is
'                       to be returned.
' optlen:               [in, out] Pointer to the size of the optval buffer, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, getsockopt returns zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function getsockopt Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal Level As Long, _
                                     ByVal OptName As Long, _
                                     ByRef OptVal As Any, _
                                     ByRef OptLen As Long) As Long


' *************************************************************************************************
' GetTypeByName
' *************************************************************************************************
'
' The GetTypeByName function retrieves a service type GUID for a network service specified by
' name.
'
' Note The GetTypeByName function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' this reference material is included. The functions detailed in Protocol-Independent Name
' Resolution provide equivalent functionality in Windows Sockets 2.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' pServiceName:         Pointer to a zero-terminated string that uniquely represents the name of
'                       the service. For example, "MY SNA SERVER."
' lpServiceType:        Pointer to a variable to receive a globally unique identifier (GUID) that
'                       specifies the type of the network service. The header file Svcguid.h
'                       includes definitions of several GUID service types and macros for working
'                       with them.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is zero.
'
' If the function fails, the return value is SOCKET_ERROR( – 1). To get extended error
' information, call GetLastError, which returns the following extended error value.
'
' Error code                    Meaning
' -------------------------------------------------------------------------------------------------
'
' ERROR_SERVICE_DOES_NOT_EXIST  The specified service type is unknown.
'
' *************************************************************************************************

Public Declare Function GetTypeByNameA Lib "wsock32.dll" _
                                    (ByVal lpServiceName As String, _
                                     ByRef lpServiceType As API_GUID) As Long


' *************************************************************************************************
' htonl
' *************************************************************************************************
'
' The htonl function converts a u_long from host to TCP/IP network byte order
' (which is big endian).
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hostlong:             [in] 32-bit number in host byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The htonl function returns the value in TCP/IP's network byte order.
'
'
' *************************************************************************************************

Public Declare Function htonl Lib "ws2_32.dll" _
                                    (ByVal HostLong As Long) As Long


' *************************************************************************************************
' htons
' *************************************************************************************************
'
' The htons function converts a u_short from host to TCP/IP network byte order
' (which is big-endian).
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hostshort:            [in] 16-bit number in host byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The htons function returns the value in TCP/IP network byte order.
'
'
' *************************************************************************************************

Public Declare Function htons Lib "ws2_32.dll" _
                                    (ByVal HostShort As Integer) As Integer


' *************************************************************************************************
' Inet_Addr
' *************************************************************************************************
'
' The inet_addr function converts a string containing an (Ipv4) Internet Protocol dotted address
' into a proper address for the IN_ADDR structure.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' cp:                   [in] Null-terminated character string representing a number expressed in
'                       the Internet standard "." (dotted) notation.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, inet_addr returns an unsigned long value containing a suitable binary
' representation of the Internet address given. If the string in the cp parameter does not
' contain a legitimate Internet address, for example if a portion of an "a.b.c.d" address exceeds
' 255, then inet_addr returns the value INADDR_NONE.
'
' Internet Addresses
' -------------------------------------------------------------------------------------------------
'
' Values specified using the "." notation take one of the following forms:
'
' a.b.c.d a.b.c a.b a
'
' When four parts are specified, each is interpreted as a byte of data and assigned, from left to
' right, to the 4 bytes of an Internet address. When an Internet address is viewed as a 32-bit
' integer quantity on the Intel architecture, the bytes referred to above appear as "d.c.b.a".
' That is, the bytes on an Intel processor are ordered from right to left.
'
' The parts that make up an address in "." notation can be decimal, octal or hexadecimal as
' specified in the C language. Numbers that start with "0x" or "0X" imply hexadecimal. Numbers
' that start with "0" imply octal. All other numbers are interpreted as decimal.

' Internet address      Meaning
' -------------------------------------------------------------------------------------------------
' "4.3.2.16"            Decimal
' "004.003.002.020"     Octal
' "0x4.0x3.0x2.0x10"    Hexadecimal
' "4.003.002.0x10"      Mix
'
' Note The following notations are only used by Berkeley, and nowhere else on the Internet.
' For compatibility with their software, they are supported as specified.
'
' When a three-part address is specified, the last part is interpreted as a 16-bit quantity and
' placed in the right-most 2 bytes of the network address. This makes the three-part address
' format convenient for specifying Class B network addresses as "128.net.host"
'
' When a two-part address is specified, the last part is interpreted as a 24-bit quantity and
' placed in the right-most 3 bytes of the network address. This makes the two-part address format
' convenient for specifying Class A network addresses as "net.host''.
'
' When only one part is given, the value is stored directly in the network address without any byte
' rearrangement.
'
' *************************************************************************************************

Public Declare Function inet_addr Lib "ws2_32.dll" _
                                    (ByVal IPAddress As String) As Long


' *************************************************************************************************
' Inet_ntoa
' *************************************************************************************************
'
' The inet_ntoa function converts an (Ipv4) Internet network address into a string in Internet
' standard dotted format.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' in:                   [in] Pointer to an in_addr structure that represents an Internet host address.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, inet_ntoa returns a character pointer to a static buffer containing the text
' address in standard "." notation. Otherwise, it returns NULL.
'
'
' *************************************************************************************************

Public Declare Function inet_ntoa Lib "ws2_32.dll" _
                                    (ByVal InAddr As Long) As Long


' *************************************************************************************************
' IOCtlSocket
' *************************************************************************************************
'
' The ioctlsocket function controls the I/O mode of a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' cmd:                  [in] Command to perform on the socket s.
' argp:                 [in, out] Pointer to a parameter for cmd.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' Upon successful completion, the ioctlsocket returns zero. Otherwise, a value of SOCKET_ERROR is
' returned, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function ioctlsocket Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal Cmd As Long, _
                                     ByRef CmdParam As Long) As Long


' *************************************************************************************************
' Listen
' *************************************************************************************************
'
' The listen function places a socket in a state in which it is listening for an incoming
' connection.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a bound, unconnected socket.
' backlog:              [in] Maximum length of the queue of pending connections. If set to
'                       SOMAXCONN, the underlying service provider responsible for socket s will
'                       set the backlog to a maximum reasonable value. There is no standard
'                       provision to obtain the actual backlog value.
'
'                       The backlog parameter is limited (silently) to a reasonable value as
'                       determined by the underlying service provider. Illegal values are replaced
'                       by the nearest legal value. There is no standard provision to find out the
'                       actual backlog value.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, listen returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function listen Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal Backlog As Long) As Long


' *************************************************************************************************
' ntohl
' *************************************************************************************************
'
' The ntohl function converts a u_long from TCP/IP network order to host byte order (which is
' little-endian on Intel processors).
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' netlong:              [in] 32-bit number in TCP/IP network byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The ntohl function always returns a value in host byte order. If the netlong parameter was
' already in host byte order, then no operation is performed.
'
'
' *************************************************************************************************

Public Declare Function ntohl Lib "ws2_32.dll" _
                                    (ByVal NetLong As Long) As Long


' *************************************************************************************************
' ntohs
' *************************************************************************************************
'
' The ntohs function converts a u_short from TCP/IP network byte order to host byte order
' (which is little-endian on Intel processors).
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' netshort:             [in] 16-bit number in TCP/IP network byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The ntohs function returns the value in host byte order. If the netshort parameter was already
' in host byte order, then no operation is performed.
'
' *************************************************************************************************

Public Declare Function ntohs Lib "ws2_32.dll" _
                                    (ByVal NetShort As Integer) As Integer


' *************************************************************************************************
' Recv
' *************************************************************************************************
'
' The recv function receives data from a connected or bound socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' buf:                  [out] Buffer for the incoming data.
' len:                  [in] Length of buf, in bytes
' flags:                [in] Flag specifying the way in which the call is made.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, recv returns the number of bytes received. If the connection has been
' gracefully closed, the return value is zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function recv Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long, _
                                     ByVal Flags As Long) As Long


' *************************************************************************************************
' RecvFrom
' *************************************************************************************************
'
' The recvfrom function receives a datagram and stores the source address.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a bound socket.
' buf:                  [out] Buffer for the incoming data.
' len:                  [in] Length of buf, in bytes.
' Flags:                [in] Indicator specifying the way in which the call is made.
' From:                 [out] Optional pointer to a buffer in a sockaddr structure that will hold
'                       the source address upon return.
' fromlen:              [in, out] Optional pointer to the size, in bytes, of the from buffer.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, recvfrom returns the number of bytes received. If the connection has been
' gracefully closed, the return value is zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function recvfrom Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long, _
                                     ByVal Flags As Long, _
                                     ByRef From As API_SOCKADDR_IN, _
                                     ByRef FromLen As Long) As Long


' *************************************************************************************************
' Select
' *************************************************************************************************
'
' The select function determines the status of one or more sockets, waiting if necessary, to
' perform synchronous I/O.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' nfds:                 [in] Ignored. The nfds parameter is included only for compatibility with
'                       Berkeley sockets.
' readfds;              [in, out] Optional pointer to a set of sockets to be checked for
'                       readability.
' writefds:             [in, out] Optional pointer to a set of sockets to be checked for
'                       writability.
' exceptfds:            [in, out] Optional pointer to a set of sockets to be checked for errors.
' timeout:              [in] Maximum time for select to wait, provided in the form of a TIMEVAL
'                       structure. Set the timeout parameter to null for blocking operations.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The select function returns the total number of socket handles that are ready and contained in
' the fd_set structures, zero if the time limit expired, or SOCKET_ERROR if an error occurred.
' If the return value is SOCKET_ERROR, WSAGetLastError can be used to retrieve a specific error
' code.
'
' *************************************************************************************************

Public Declare Function wsselect Lib "ws2_32.dll" Alias "select" _
                                    (ByVal Reserved As Long, _
                                     ByRef ReadFds As Any, _
                                     ByRef WriteFds As Any, _
                                     ByRef ExceptFds As Any, _
                                     ByVal TimeOut As Long) As Long


' *************************************************************************************************
' Send
' *************************************************************************************************
'
' The send function sends data on a connected socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' buf:                  [in] Buffer containing the data to be transmitted.
' len:                  [in] Length of the data in buf, in bytes
' Flags:                [in] Indicator specifying the way in which the call is made.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, send returns the total number of bytes sent, which can be less than the
' number indicated by len. Otherwise, a value of SOCKET_ERROR is returned, and a specific error
' code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function send Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long, _
                                     ByVal Flags As Long) As Long


' *************************************************************************************************
' SendTo
' *************************************************************************************************
'
' The sendto function sends data to a specific destination.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a (possibly connected) socket.
' buf:                  [in] Buffer containing the data to be transmitted.
' len:                  [in] Length of the data in buf, in bytes.
' Flags:                [in] Indicator specifying the way in which the call is made.
' to:                   [in] Optional pointer to a sockaddr structure that contains the address of
'                       the target socket.
' tolen:                [in] Size of the address in to, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, sendto returns the total number of bytes sent, which can be less than the
' number indicated by len. Otherwise, a value of SOCKET_ERROR is returned, and a specific error
' code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function sendt Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long, _
                                     ByVal Flags As Long, _
                                     ByRef ToSocket As API_SOCKADDR_IN, _
                                     ByVal ToLength As Long) As Long


' *************************************************************************************************
' SetService
' *************************************************************************************************
'
' The SetService function registers or removes from the registry a network service within one or
' more namespaces. The function can also add or remove a network service type within one or more
' namespaces.
'
' Note The SetService function is obsolete. The functions detailed in Protocol-Independent Name
' Resolution provide equivalent functionality in Windows Sockets 2. For the convenience of Windows
' Sockets 1.1 developers, the reference material is as follows.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' dwNameSpace:          Namespace, or a set of default namespaces, within which the function will
'                       operate.
' dwOperation:          Specifies the operation that the function will perform.
' dwFlags:              Set of bit flags that modify the function's operation.
' lpServiceInfo:        Pointer to a SERVICE_INFO structure that contains information about the
'                       network service or service type.
' lpServiceAsyncInfo:   Reserved for future use. Must be set to NULL.
' lpdwStatusFlags:      Set of bit flags that receive function status information.
' Value Meaning:        Set of bit flags that receive function status information.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function fails, the return value is SOCKET_ERROR. To get extended error information, call
' GetLastError. GetLastError can return the following extended error value.
'
' *************************************************************************************************

Public Declare Function SetServiceA Lib "ws2_23.dll" _
                                    (ByVal dwNameSpace As Long, _
                                     ByVal dwOperation As Long, _
                                     ByVal dwFlags As Long, _
                                     ByRef lpServiceInfo As API_SERVICE_INFO, _
                                     ByVal Reserved As Long, _
                                     ByRef lpdwStatusFlags As Long) As Long


' *************************************************************************************************
' SetSockOpt
' *************************************************************************************************
'
' The setsockopt function sets a socket option.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' Level:                [in] Level at which the option is defined; the supported levels include
'                       SOL_SOCKET and IPPROTO_TCP. See Winsock Annexes for more information on
'                       protocol-specific levels.
' OptName:              [in] Socket option for which the value is to be set.
' OptVal:               [in] Pointer to the buffer in which the value for the requested option is
'                       specified.
' OptLen:               [in] Size of the optval buffer, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, setsockopt returns zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function setsockopt Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal Level As Long, _
                                     ByVal OptionName As Long, _
                                     ByRef OptionValue As Any, _
                                     ByVal OptionLength As Long) As Long


' *************************************************************************************************
' Shutdown
' *************************************************************************************************
'
' The shutdown function disables sends or receives on a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' how:                  [in] Flag that describes what types of operation will no longer be allowed.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, shutdown returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function shutdown Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal How As Long) As Long


' *************************************************************************************************
' Socket
' *************************************************************************************************
'
' The socket function creates a socket that is bound to a specific service provider.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
'  af:                  [in] Address family specification.
' type:                 [in] Type specification for the new socket.
'
'                       In Windows Sockets 2, many new socket types will be introduced and no
'                       longer need to be specified, since an application can dynamically discover
'                       the attributes of each available transport protocol through the
'                       WSAEnumProtocols function. Socket type definitions appear in Winsock2.h,
'                       which will be periodically updated as new socket types, address families,
'                       and protocols are defined.
' Protocol:             [in] Protocol to be used with the socket that is specific to the indicated
'                       address family.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, socket returns a descriptor referencing the new socket. Otherwise, a value
' of INVALID_SOCKET is returned, and a specific error code can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function socket Lib "ws2_32.dll" _
                                    (ByVal AddressFamily As Long, _
                                     ByVal SocketType As Long, _
                                     ByVal Protocol As Long) As Long


' *************************************************************************************************
' TransmitFile
' *************************************************************************************************
'
' The TransmitFile function transmits file data over a connected socket handle. This function
' uses the operating system's cache manager to retrieve the file data, and provides
' high-performance file data transfer over sockets.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hSocket:              Handle to a connected socket. The TransmitFile function will transmit the
'                       file data over this socket. The socket specified by hSocket must be a
'                       connection-oriented socket; the TransmitFile function does not support
'                       datagram sockets. Sockets of type SOCK_STREAM, SOCK_SEQPACKET, or
'                       SOCK_RDM are connection-oriented sockets.
' hFile:                Handle to the open file that the TransmitFile function transmits. Since
'                       operating system reads the file data sequentially, you can improve
'                       caching performance by opening the handle with FILE_FLAG_SEQUENTIAL_SCAN.
'                       The hFile parameter is optional; if the hFile parameter is null, only
'                       data in the header and/or the tail buffer is transmitted; any additional
'                       action, such as socket disconnect or reuse, is performed as specified by
'                       the dwFlags parameter.
' nNumberOfBytesToWrite:Number of file bytes to transmit. The TransmitFile function completes
'                       when it has sent the specified number of bytes, or when an error occurs,
'                       whichever occurs first.
'
'                       Set nNumberOfBytesToWrite to zero in order to transmit the entire file.
' nNumberOfBytesPerSend:Size of each block of data sent in each send operation, in bytes. This
'                       specification is used by Windows' sockets layer. To select the default
'                       send size, set nNumberOfBytesPerSend to zero.
'
'                       The nNumberOfBytesPerSend parameter is useful for message protocols that
'                       have limitations on the size of individual send requests.
' lpOverlapped:         Pointer to an OVERLAPPED structure. If the socket handle has been opened
'                       as overlapped, specify this parameter in order to achieve an overlapped
'                       (asynchronous) I/O operation. By default, socket handles are opened as
'                       overlapped.
'
'                       You can use lpOverlapped to specify an offset within the file at which to
'                       start the file data transfer by setting the Offset and OffsetHigh member
'                       of the OVERLAPPED structure. If lpOverlapped is null, the transmission of
'                       data always starts at the current byte offset in the file.
'
'                       When lpOverlapped is not null, the overlapped I/O might not finish before
'                       TransmitFile returns. In that case, the TransmitFile function returns
'                       FALSE, and GetLastError returns ERROR_IO_PENDING. This enables the caller
'                       to continue processing while the file transmission operation completes.
'                       Windows will set the event specified by the hEvent member of the
'                       OVERLAPPED structure, or the socket specified by hSocket, to the signaled
'                       state upon completion of the data transmission request.
' lpTransmitBuffers:    Pointer to a TRANSMIT_FILE_BUFFERS data structure that contains pointers
'                       to data to send before and after the file data is sent. Set the
'                       lpTransmitBuffers parameter to null if you want to transmit only the file
'                       data.
' dwFlags:              The dwFlags parameter.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the TransmitFile function succeeds, the return value is TRUE. Otherwise, the return value is
' FALSE. To get extended error information, call WSAGetLastError. The function returns FALSE if
' an overlapped I/O operation is not complete before TransmitFile returns. In that case,
' WSAGetLastError returns ERROR_IO_PENDING or WSA_IO_PENDING. Applications should handle either
' ERROR_IO_PENDING or WSA_IO_PENDING.
'
' *************************************************************************************************

Public Declare Function TransmitFile Lib "wsock32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal hFile As Long, _
                                     ByVal nNumberOfBytesToWrite As Long, _
                                     ByVal nNumberOfBytesPerSend As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByRef lpTransmitBuffers As API_TRANSMIT_FILE_BUFFERS, _
                                     ByVal dwFlags As Long) As Long


' *************************************************************************************************
' TransmitPackets
' *************************************************************************************************
'
' The TransmitPackets function transmits in-memory data or file data over a connected socket. The
' TransmitPackets function uses the operating system cache manager to retrieve file data, locking
' memory for the minimum time required to transmit and resulting in efficient, high-performance
' transmission.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hSocket:              Handle to the connected socket to be used in the transmission. Although
'                       the socket does not need to be a connection-oriented circuit, the default
'                       destination/peer should have been established using the connect,
'                       WSAConnect, accept, WSAAccept, AcceptEx, or WSAJoinLeaf function.
' lpPacketArray:        Array of type TRANSMIT_PACKETS_ELEMENT, describing the data to be
'                       transmitted.
' nElementCount:        Number of elements in lpPacketArray.
' nSendSize:            Size of the data block used in the send operation, in bytes. Set
'                       nSendSize to zero to let the sockets layer select a default send size.
'
'                       Setting nSendSize to 0xFFFFFFF enables the caller to control the size and
'                       content of each send request, achieved by using the TP_ELEMENT_EOP flag
'                       in the TRANSMIT_PACKETS_ELEMENT array pointed to in the lpPacketArray
'                       parameter. This capability is useful for message protocols that place
'                       limitations on the size of individual send requests.
' lpOverlapped:         Pointer to an OVERLAPPED structure. If the socket handle specified in the
'                       hSocket parameter has been opened as overlapped, use this parameter to
'                       achieve asynchronous (overlapped) I/O operation. Socket handles are
'                       opened as overlapped by default.
' dwFlags:              Flags used to customize processing of the TransmitPackets function.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' Success returns TRUE, failure returns FALSE. Use the WSAGetLastError function to retrieve
' extended error information.
'
' *************************************************************************************************

Public Declare Function TransmitPackets Lib "wsock32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpPacketArray As API_TRANSMIT_PACKETS_ELEMENT, _
                                     ByVal nElementCount As Long, _
                                     ByVal nSendSize As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal dwFlags As Long) As Long

' *************************************************************************************************
' WSAAccept
' *************************************************************************************************
'
' The WSAAccept function conditionally accepts a connection based on the return value of a
' condition function, provides quality of service flow specifications, and allows the transfer of
' connection data.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket that is listening for connections
'                       after a call to the listen function.
' addr:                 [out] Optional pointer to a buffer in a sockaddr structure that receives
'                       the address of the connecting entity, as known to the communications layer.
'                       The exact format of the addr parameter is determined by the address family
'                       established when the socket was created.
' AddrLen:              [in, out] Optional pointer to an integer that contains the length of the
'                       address addr, in bytes.
' lpfnCondition:        [in] Procedure instance address of the optional, application-specified
'                       condition function that will make an accept/reject decision based on the
'                       caller information passed in as parameters, and optionally create or join
'                       a socket group by assigning an appropriate value to the result parameter
'                       g of this function.
' dwCallbackData:       [in] Callback data passed back to the application as the value of the
'                       dwCallbackData parameter of the condition function. This parameter is not
'                       interpreted by Windows Sockets.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAAccept returns a value of type SOCKET that is a descriptor for the
' accepted socket. Otherwise, a value of INVALID_SOCKET is returned, and a specific error code
' can be retrieved by calling WSAGetLastError.
'
' The integer referred to by addrlen initially contains the amount of space pointed to by addr.
' On return it will contain the actual length in bytes of the address returned.
'
' A prototype of the condition function is as follows:
'
' Public Function ConditionFunc (ByRef lpCallerId As WSABUF, _
'                                ByRef lpCallerData As WSABUF, _
'                                ByRef lpSQOS As FLOWSPEC, _
'                                ByVal Reserved As Long, _
'                                ByRef lpCalleeId As WSABUF, _
'                                ByRef lpCalleeData As WSABUF, _
'                                ByRef Group As Long, _
'                                ByVal dwCallbackData As Long) As Long
'
' *************************************************************************************************

Public Declare Function WSAAccept Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef SocketAddress As API_SOCKADDR_IN, _
                                     ByRef AddressLength As Long, _
                                     ByVal lpfnCondition As Long, _
                                     ByVal dwCallbackData As Long) As Long


' *************************************************************************************************
' WSAAddressToString
' *************************************************************************************************
'
' The WSAAddressToString function converts all components of a sockaddr structure into a
' human-readable string representation of the address.

' This is intended to be used mainly for display purposes. If the caller wants the translation to
' be done by a particular provider, it should supply the corresponding WSAPROTOCOL_INFO structure
' in the lpProtocolInfo parameter.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpsaAddress:          [in] Pointer to the sockaddr structure to translate into a string.
' dwAddressLength:      [in] Length of the address in sockaddr, in bytes, which may vary in size
'                       with different protocols.
' lpProtocolInfo:       [in] (Optional) The WSAPROTOCOL_INFO structure for a particular provider.
'                       If this is NULL, the call is routed to the provider of the first protocol
'                       supporting the address family indicated in lpsaAddress.
' lpszAddressString:    [in, out] Buffer that receives the human-readable address string.
' lpdwAddressStringLength: [in, out] On input, the length of the AddressString buffer, in bytes.
'                       On output, returns the length of the string actually copied into the
'                       buffer. If the specified buffer is not large enough, the function fails
'                       with a specific error of WSAEFAULT and this parameter is updated with the
'                       required size in characters.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAAddressToString returns a value of zero. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAAddressToStringA Lib "ws2_32.dll" _
                                    (ByRef lpsaAddress As API_SOCKADDR_IN, _
                                     ByVal dwAddressLength As Long, _
                                     ByRef lpProtocolInfo As API_WSAPROTOCOL_INFO, _
                                     ByVal lpszAddressString As String, _
                                     ByRef lpdwAddressStringLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetHostByAddr
' *************************************************************************************************
'
' The WSAAsyncGetHostByAddr function asynchronously retrieves host information that corresponds
' to an address.
'
' Note The WSAAsyncGetHostByAddr function is not designed to provide parallel resolution of
' several addresses. Therefore, applications that issue several requests should not expect them
' to be executed concurrently. Alternatively, applications can start another thread and use the
' getnameinfo function to resolve addresses in an IP-version agnostic manner. Developers creating
' Windows Sockets 2 applications are urged to use the getnameinfo function to enable smooth
' transition to IPv6 compatibility.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that will receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' addr:                 [in] Pointer to the network address for the host. Host addresses are
'                       stored in network byte order.
' len:                  [in] Length of the address, in bytes.
' type:                 [in] Type of the address.
' buf:                  [out] Pointer to the data area to receive the hostent data. The data area
'                       must be larger than the size of a hostent structure because the data area
'                       is used by Windows Sockets to contain a hostent structure and all of the
'                       data referenced by members of the hostent structure. A buffer of
'                       MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [in] Size of data area for the buf parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetHostByAddr returns a nonzero value of type HANDLE that is the
' asynchronous task handle (not to be confused with a Windows HTASK) for the request. This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetHostByAddr returns a zero
' value, and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As
' described above, they can be extracted from the lParam in the reply message using the
' WSAGETASYNCERROR macro.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetHostByAddr Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal Address As Long, _
                                     ByVal AddressLength As Long, _
                                     ByVal AddressType As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetHostByName
' *************************************************************************************************
'
' The WSAAsyncGetHostByName function asynchronously retrieves host information that corresponds
' to a host name.
'
' Note The WSAAsyncGetHostByName function is not designed to provide parallel resolution of
' several names. Therefore, applications that issue several requests should not expect them to be
' executed concurrently. Alternatively, applications can start another thread and use the
' getaddrinfo function to resolve names in an IP-version agnostic manner. Developers creating
' Windows Sockets 2 applications are urged to use the getaddrinfo function to enable smooth
' transition to IPv6 compatibility.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that will receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' name:                 [in] Pointer to the null-terminated name of the host.
' buf:                  [out] Pointer to the data area to receive the hostent data. The data area
'                       must be larger than the size of a hostent structure because the specified
'                       data area is used by Windows Sockets to contain a hostent structure and
'                       all of the data referenced by members of the hostent structure. A buffer
'                       of MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [in] Size of data area for the buf parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetHostByName returns a nonzero value of type HANDLE that is the
' asynchronous task handle (not to be confused with a Windows HTASK) for the request. This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetHostByName returns a zero
' value, and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message.
' As described above, they can be extracted from the lParam in the reply message.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetHostByName Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal HostName As String, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetProtoByName
' *************************************************************************************************
'
' The WSAAsyncGetProtoByName function asynchronously retrieves protocol information that
' corresponds to a protocol name.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that will receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' name:                 [in] Pointer to the null-terminated protocol name to be resolved.
' buf:                  [out] Pointer to the data area to receive the protoent data. The data
'                       area must be larger than the size of a protoent structure because the
'                       data area is used by Windows Sockets to contain a protoent structure and
'                       all of the data that is referenced by members of the protoent structure.
'                       A buffer of MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [out] Size of data area for the buf parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetProtoByName returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages, by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetProtoByName returns a zero
' value, and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As
' described above, they can be extracted from the lParam in the reply message.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetProtoByName Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal ProtocolName As String, _
                                     ByRef Buffer As Any, _
                                     ByRef BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetProtoByNumber
' *************************************************************************************************
'
' The WSAAsyncGetProtoByNumber function asynchronously retrieves protocol information that
' corresponds to a protocol number.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that will receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' Number:               [in] Protocol number to be resolved, in host byte order.
' buf:                  [out] Pointer to the data area to receive the protoent data. The data
'                       area must be larger than the size of a protoent structure because the
'                       data area is used by Windows Sockets to contain a protoent structure and
'                       all of the data that is referenced by members of the protoent structure.
'                       A buffer of MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [in] Size of data area for the buf parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetProtoByNumber returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages, by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetProtoByNumber returns a zero
' value, and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As
' described above, they can be extracted from the lParam in the reply message.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetProtoByNumber Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal ProtocolNumber As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetServByName
' *************************************************************************************************
'
' The WSAAsyncGetServByName function asynchronously retrieves service information that
' corresponds to a service name and port.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that should receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' name:                 [in] Pointer to a null-terminated service name.
' Proto:                [in] Pointer to a protocol name. This can be NULL, in which case
'                       WSAAsyncGetServByName will search for the first service entry for which
'                       s_name or one of the s_aliases matches the given name. Otherwise,
'                       WSAAsyncGetServByName matches both name and proto.
' buf:                  [out] Pointer to the data area to receive the servent data. The data area
'                       must be larger than the size of a servent structure because the data area
'                       is used by Windows Sockets to contain a servent structure and all of the
'                       data that is referenced by members of the servent structure. A buffer of
'                       MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [in] Size of data area for the buf parameter, in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetServByName returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages, by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncServByName returns a zero value,
' and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message.
' As described above, they can be extracted from the lParam in the reply message.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetServByName Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal ServiceName As String, _
                                     ByVal ProtocolName As String, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncGetServByPort
' *************************************************************************************************
'
' The WSAAsyncGetServByPort function asynchronously retrieves service information that
' corresponds to a port and protocol.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hWnd:                 [in] Handle of the window that should receive a message when the
'                       asynchronous request completes.
' wMsg:                 [in] Message to be received when the asynchronous request completes.
' Port:                 [in] Port for the service, in network byte order.
' Proto:                [in] Pointer to a protocol name. This can be NULL, in which case
'                       WSAAsyncGetServByPort will search for the first service entry for which
'                       s_port match the given port. Otherwise, WSAAsyncGetServByPort matches
'                       both port and proto.
' buf:                  [out] Pointer to the data area to receive the servent data. The data area
'                       must be larger than the size of a servent structure because the data area
'                       is used by Windows Sockets to contain a servent structure and all of the
'                       data that is referenced by members of the servent structure. A buffer of
'                       MAXGETHOSTSTRUCT bytes is recommended.
' buflen:               [in] Size of data area for the buf parameter, in bytes.
'
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value specifies whether or not the asynchronous operation was successfully initiated.
' It does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetServByPort returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value
' can be used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or
' it can be used to match up asynchronous operations and completion messages, by examining the
' wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetServByPort returns a zero
' value, and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As
' described above, they can be extracted from the lParam in the reply message.
'
' *************************************************************************************************

Public Declare Function WSAAsyncGetServByPort Lib "ws2_32.dll" _
                                    (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal Port As Long, _
                                     ByVal ProtocolName As String, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long) As Long


' *************************************************************************************************
' WSAAsyncSelect
' *************************************************************************************************
'
' The WSAAsyncSelect function requests Windows message-based notification of network events for a
' socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket for which event notification is
'                       required.
' hWnd:                 [in] Handle identifying the window that will receive a message when a
'                       network event occurs.
' wMsg:                 [in] Message to be received when a network event occurs.
' lEvent:               [in] Bitmask that specifies a combination of network events in which the
'                       application is interested.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the WSAAsyncSelect function succeeds, the return value is zero provided the application's
' declaration of interest in the network event set was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAAsyncSelect Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal lEvent As Long) As Long


' *************************************************************************************************
' WSACancelAsyncRequest
' *************************************************************************************************
'
' The WSACancelAsyncRequest function cancels an incomplete asynchronous operation.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hAsyncTaskHandle:     [in] Handle that specifies the asynchronous operation to be canceled.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The value returned by WSACancelAsyncRequest is zero if the operation was successfully canceled.
' Otherwise, the value SOCKET_ERROR is returned, and a specific error number can be retrieved by
' calling WSAGetLastError.
'
' Note It is unclear whether the application can usefully distinguish between WSAEINVAL and
' WSAEALREADY, since in both cases the error indicates that there is no asynchronous operation in
' progress with the indicated handle. (Trivial exception: zero is always an invalid asynchronous
' task handle.) The Windows Sockets specification does not prescribe how a conformant
' Windows Sockets provider should distinguish between the two cases. For maximum portability, a
' Windows Sockets application should treat the two errors as equivalent.
'
' *************************************************************************************************

Public Declare Function WSACancelAsyncRequest Lib "ws2_32.dll" _
                                    (ByVal hAsyncTaskHandle As Long) As Long




' *************************************************************************************************
' WSACancelBlockingCall
' *************************************************************************************************

' CancelS a blocking call which is currently in progress.
'
' The WSACancelBlockingCall function has been removed in compliance with the Windows Sockets 2
' specification, revision 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll and Windows Sockets 2 applications
' should not use this function. Windows Sockets 1.1 applications that call this function are
' still supported through the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during
' calls to blocking functions. Instead of using blocking hooks, an applications should use a
' separate thread (separate from the main GUI thread) for network activity.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The value returned by WSACancelBlockingCall() is 0 if the operation was successfully canceled.
' Otherwise the value SOCKET_ERROR is returned, and a specific error number may be retrieved by
' calling WSAGetLastError().
'
' *************************************************************************************************

Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long


' *************************************************************************************************
' WSACleanup
' *************************************************************************************************
'
' The WSACleanup function terminates use of the WS2_32.DLL.

' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' Attempting to call WSACleanup from within a blocking hook and then failing to check the return
' code is a common programming error in Windows Socket 1.1 applications. If an application needs
' to quit while a blocking call is outstanding, the application must first cancel the blocking
' call with WSACancelBlockingCall then issue the WSACleanup call once control has been returned
' to the application.
'
' In a multithreaded environment, WSACleanup terminates Windows Sockets operations for all
' threads.
'
' *************************************************************************************************

Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long


' *************************************************************************************************
' WSACloseEvent
' *************************************************************************************************
'
' The WSACloseEvent function closes an open event object handle.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hEvent:               [in] Object handle identifying the open event.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is TRUE.
'
' If the function fails, the return value is FALSE. To get extended error information, call
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSACloseEvent Lib "ws2_32.dll" _
                                    (ByVal hEvent As Long) As Long


' *************************************************************************************************
' WSAConnect
' *************************************************************************************************
'
' The WSAConnect function establishes a connection to another socket application, exchanges
' connect data, and specifies required quality of service based on the specified FLOWSPEC
' structure.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying an unconnected socket.
' name:                 [in] Name of the socket in a sockaddr structure in the other application
'                       to which to connect.
' NameLen:              [in] Length of name, in bytes.
' lpCallerData:         [in] Pointer to the user data that is to be transferred to the other
'                       socket during connection establishment. See Remarks.
' lpCalleeData:         [out] Pointer to the user data that is to be transferred back from the
'                       other socket during connection establishment. See Remarks.
' lpSQOS:               [in] Pointer to the FLOWSPEC structures for socket s, one for each
'                       direction.
' lpGQOS:               [in] Reserved for future use with socket groups. A pointer to the F
'                       LOWSPEC structures for the socket group (if applicable). Should be NULL.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAConnect returns zero. Otherwise, it returns SOCKET_ERROR, and a specific
' error code can be retrieved by calling WSAGetLastError. On a blocking socket, the return value
' indicates success or failure of the connection attempt.
'
' With a nonblocking socket, the connection attempt cannot be completed immediately.
' In this case, WSAConnect will return SOCKET_ERROR, and WSAGetLastError will return
' WSAEWOULDBLOCK; the application could therefore:
'
' - Use select to determine the completion of the connection request by checking if the socket is
'   writeable.
' - If your application is using WSAAsyncSelect to indicate interest in connection events, then
'   your application will receive an FD_CONNECT notification when the connect operation is
'   complete(successful or not).
' - If your application is using WSAEventSelect to indicate interest in connection events, then
'   the associated event object will be signaled when the connect operation is complete
'   (successful or not).
'
' For a nonblocking socket, until the connection attempt completes all subsequent calls to
' WSAConnect on the same socket will fail with the error code WSAEALREADY.
'
' If the return error code indicates the connection attempt failed (that is, WSAECONNREFUSED,
' WSAENETUNREACH, WSAETIMEDOUT) the application can call WSAConnect again for the same socket.
'
' *************************************************************************************************

Public Declare Function WSAConnect Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal SocketName As API_SOCKADDR_IN, _
                                     ByVal NameLength As Long, _
                                     ByRef lpCallerData As API_WSABUF, _
                                     ByRef lpCalleeData As API_WSABUF, _
                                     ByRef lpSQOS As API_FLOWSPEC, _
                                     ByVal Reserved As Long) As Long


' *************************************************************************************************
' WSACreateEvent
' *************************************************************************************************
'
' The WSACreateEvent function creates a new event object.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSACreateEvent returns the handle of the event object. Otherwise, the
' return value is WSA_INVALID_EVENT. To get extended error information, call WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSACreateEvent Lib "ws2_32.dll" () As Long



' *************************************************************************************************
' WSADuplicateSocket
' *************************************************************************************************
'
' The WSADuplicateSocket function returns a WSAPROTOCOL_INFO structure that can be used to create
' a new socket descriptor for a shared socket. The WSADuplicateSocket function cannot be used on
' a QOS-enabled socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the local socket.
' dwProcessId:          [in] Process identifier of the target process in which the duplicated
'                       socket will be used.
' lpProtocolInfo:       [out] Pointer to a buffer, allocated by the client, that is large enough
'                       to contain a WSAPROTOCOL_INFO structure. The service provider copies the
'                       protocol information structure contents to this buffer.
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSADuplicateSocket returns zero. Otherwise, a value of SOCKET_ERROR is
' returned, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSADuplicateSocketA Lib "ws2)32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal dwProcessId As Long, _
                                     ByRef lpProtocolInfo As API_WSAPROTOCOL_INFO) As Long


' *************************************************************************************************
' WSAEnumNameSpaceProviders
' *************************************************************************************************
'
' The WSAEnumNameSpaceProviders function retrieves information about available namespaces.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpdwBufferLength:     [in, out] On input, the number of bytes contained in the buffer pointed
'                       to by lpnspBuffer. On output (if the function fails, and the error is
'                       WSAEFAULT), the minimum number of bytes to pass for the lpnspBuffer to
'                       retrieve all the requested information. The passed-in buffer must be
'                       sufficient to hold all of the namespace information.
' lpnspBuffer:          [out] Buffer that is filled with WSANAMESPACE_INFO structures. The
'                       returned structures are located consecutively at the head of the buffer.
'                       Variable sized information referenced by pointers in the structures point
'                       to locations within the buffer located between the end of the fixed sized
'                       structures and the end of the buffer. The number of structures filled in
'                       is the return value of WSAEnumNameSpaceProviders.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The WSAEnumNameSpaceProviders function returns the number of WSANAMESPACE_INFO structures copied
' into lpnspBuffer. Otherwise, the value SOCKET_ERROR is returned, and a specific error number can
' be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAEnumNameSpaceProvidersA Lib "ws2_32.dll" _
                                    (ByRef lpdwBufferLength As Long, _
                                     ByRef lpnspBuffer As API_WSANAMESPACE_INFO) As Long


' *************************************************************************************************
' WSAEnumNetworkEvents
' *************************************************************************************************
'
' The WSAEnumNetworkEvents function discovers occurrences of network events for the indicated
' socket, clear internal network event records, and reset event objects (optional).
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket.
' hEventObject:         [in] Optional handle identifying an associated event object to be reset.
' lpNetworkEvents:      [out] Pointer to a WSANETWORKEVENTS structure that is filled with a
'                       record of network events that occurred and any associated error codes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAEnumNetworkEvents Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal hEventObject As Long, _
                                     ByRef lpNetworkEvents As API_WSANETWORKEVENTS) As Long



' *************************************************************************************************
' WSAEnumProtocols
' *************************************************************************************************
'
' The WSAEnumProtocols function retrieves information about available transport protocols.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpiProtocols:         [in] Null-terminated array of iProtocol values. This parameter is
'                       optional; if lpiProtocols is NULL, information on all available protocols
'                       is returned. Otherwise, information is retrieved only for those protocols
'                       listed in the array.
' lpProtocolBuffer:     [out] Buffer that is filled with WSAPROTOCOL_INFO structures.
' lpdwBufferLength:     [in, out] On input, number of bytes in the lpProtocolBuffer buffer passed
'                       to WSAEnumProtocols. On output, the minimum buffer size that can be passed
'                       to WSAEnumProtocols to retrieve all the requested information. This
'                       routine has no ability to enumerate over multiple calls; the passed-in
'                       buffer must be large enough to hold all entries in order for the routine
'                       to succeed. This reduces the complexity of the API and should not pose a
'                       problem because the number of protocols loaded on a computer is typically
'                       small.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAEnumProtocols returns the number of protocols to be reported. Otherwise,
' a value of SOCKET_ERROR is returned and a specific error code can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAEnumProtocolsA Lib "ws2_32.dll" _
                                    (ByVal lpiProtocols As Long, _
                                     ByRef lpProtocolBuffer As Any, _
                                     ByRef lpdwBufferLength As Long) As Long


' *************************************************************************************************
' WSAEventSelect
' *************************************************************************************************
'
' The WSAEventSelect function specifies an event object to be associated with the specified set of
' FD_XXX network events.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket.
' hEventObject:         [in] Handle identifying the event object to be associated with the
'                       specified set of FD_XXX network events.
' lNetworkEvents:       [in] Bitmask that specifies the combination of FD_XXX network events in
'                       which the application has interest.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the application's specification of the network events and the
' associated event object was successful. Otherwise, the value SOCKET_ERROR is returned, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAEventSelect Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal hEventObject As Long, _
                                     ByVal lNetworkEvents As Long) As Long


' *************************************************************************************************
' WSAGetLastError
' *************************************************************************************************
'
' The WSAGetLastError function returns the error status for the last operation that failed.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value indicates the error code for this thread's last Windows Sockets operation that
' failed.
'
' *************************************************************************************************

Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long


' *************************************************************************************************
' WSAGetOverlappedResult
' *************************************************************************************************
'
' The WSAGetOverlappedResult function retrieves the results of an overlapped operation on the
' specified socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket. This is the same socket that was
'                       specified when the overlapped operation was started by a call to WSARecv,
'                       WSARecvFrom, WSASend, WSASendTo, or WSAIoctl.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure that was specified when the
'                       overlapped operation was started.
' lpcbTransfer:         [out] Pointer to a 32-bit variable that receives the number of bytes that
'                       were actually transferred by a send or receive operation, or by WSAIoctl.
' fWait:                [in] Flag that specifies whether the function should wait for the pending
'                       overlapped operation to complete. If TRUE, the function does not return
'                       until the operation has been completed. If FALSE and the operation is
'                       still pending, the function returns FALSE and the WSAGetLastError function
'                       returns WSA_IO_INCOMPLETE. The fWait parameter may be set to TRUE only if
'                       the overlapped operation selected the event-based completion notification.
' lpdwFlags:            [out] Pointer to a 32-bit variable that will receive one or more flags that
'                       supplement the completion status. If the overlapped operation was initiated
'                       through WSARecv or WSARecvFrom, this parameter will contain the results
'                       value for lpFlags parameter.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If WSAGetOverlappedResult succeeds, the return value is TRUE. This means that the overlapped
' operation has completed successfully and that the value pointed to by lpcbTransfer has been
' updated. If WSAGetOverlappedResult returns FALSE, this means that either the overlapped
' operation has not completed, the overlapped operation completed but with errors, or the
' overlapped operation's completion status could not be determined due to errors in one or more
' parameters to WSAGetOverlappedResult. On failure, the value pointed to by lpcbTransfer will not
' be updated. Use WSAGetLastError to determine the cause of the failure (either of
' WSAGetOverlappedResult or of the associated overlapped operation).
'
' *************************************************************************************************

Public Declare Function WSAGetOverlappedResult Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpcbTransfer As Long, _
                                     ByVal fWait As Long, _
                                     ByRef lpdwFlags As Long) As Long


' *************************************************************************************************
' WSAGetQOSByName
' *************************************************************************************************
'
' The WSAGetQOSByName function initializes a QOS structure based on a named template, or it
' supplies a buffer to retrieve an enumeration of the available template names.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' lpQOSName:            [in, out] Pointer to a specific quality of service template.
' lpQOS:                [out] Pointer to the QOS structure to be filled.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If WSAGetQOSByName succeeds, the return value is TRUE. If the function fails, the return value
' is FALSE. To get extended error information, call WSAGetLastError.
'
'
' *************************************************************************************************

Public Declare Function WSAGetQOSByName Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpQOSName As API_WSABUF, _
                                     ByRef lpQOS As API_WSA_QOS) As Long


' *************************************************************************************************
' WSAGetServiceClassInfo
' *************************************************************************************************
'
' The WSAGetServiceClassInfo function retrieves the class information (schema) pertaining to a
' specified service class from a specified namespace provider.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpProviderId:         [in] Pointer to a GUID that identifies a specific namespace provider.
' lpServiceClassId:     [in] Pointer to a GUID identifying the service class.
' lpdwBufferLength:     [in, out] On input, the number of bytes contained in the buffer pointed
'                       to by lpServiceClassInfo. On output, if the function fails and the error
'                       is WSAEFAULT, then it contains the minimum number of bytes to pass for
'                       the lpServiceClassInfo to retrieve the record.
' lpServiceClassInfo:   [out] Pointer to the service class information from the indicated
'                       namespace provider for the specified service class.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the WSAGetServiceClassInfo was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAGetServiceClassInfoA Lib "ws2_32.dll" _
                                    (ByRef lpProviderId As API_GUID, _
                                     ByRef lpServiceClassId As API_GUID, _
                                     ByRef lpdwBufferLength As Long, _
                                     ByRef lpServiceClassInfo As API_WSASERVICECLASSINFO) As Long


' *************************************************************************************************
' WSAGetServiceClassNameByClassId
' *************************************************************************************************
'
' The WSAGetServiceClassNameByClassId function retrieves the name of the service associated with
' the specified type. This name is the generic service name, like FTP or SNA, and not the name of
' a specific instance of that service
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpServiceClassId:     [in] Pointer to the GUID for the service class.
' lpszServiceClassName: [out] Pointer to the service name.
' lpdwBufferLength:     [in, out] On input, the length of the buffer returned by
'                       lpszServiceClassName, in characters. On output, the length of the service
'                       name copied into lpszServiceClassName, in characters.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The WSAGetServiceClassNameByClassId function returns a value of zero if successful. Otherwise,
' the value SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAGetServiceClassNameByClassIdA Lib "ws2_32.dllL" _
                                    (ByRef lpServiceClassId As API_GUID, _
                                     ByVal lpszServiceClassName As String, _
                                     ByRef lpdwBufferLength As Long) As Long


' *************************************************************************************************
' WSAHtonl
' *************************************************************************************************
'
' The WSAHtonl function converts a u_long from host byte order to network byte order.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' HostLong:             [in] 32-bit number in host byte order.
' lpnetlong:            [out] Pointer to a 32-bit number in network byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAHtonl returns zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAHtonl Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal HostLong As Long, _
                                     ByRef lpNetLong As Long) As Long


' *************************************************************************************************
' WSAHtons
' *************************************************************************************************
'
' The WSAHtons function converts a u_short from host byte order to network byte order.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' HostShort:            [in] 16-bit number in host byte order.
' lpnetshort:           [out] Pointer to a 16-bit number in network byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAHtons returns zero. Otherwise, a value of SOCKET_ERROR is returned, and
' a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAHtons Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal HostShort As Integer, _
                                     ByRef lpNetShort As Integer) As Long


' *************************************************************************************************
' WSAInstallServiceClass
' *************************************************************************************************
'
' The WSAInstallServiceClass function registers a service class schema within a namespace. This
' schema includes the class name, class identifier, and any namespace-specific information that
' is common to all instances of the service, such as the SAP identifier or object identifier.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpServiceClassInfo:   [in] Service class to namespace specific–type mapping information.
'                       Multiple mappings can be handled at one time.
'
'                       See the section Service Class Data Structures for a description of
'                       pertinent data structures.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAInstallServiceClassA Lib "ws2_32.dll" _
                                    (ByRef lpServiceClassInfo As API_WSASERVICECLASSINFO) As Long


' *************************************************************************************************
' WSAIoctl
' *************************************************************************************************
'
' The WSAIoctl function controls the mode of a socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' dwIoControlCode:      [in] Control code of operation to perform.
' lpvInBuffer:          [in] Pointer to the input buffer.
' cbInBuffer:           [in] Size of the input buffer, in bytes.
' lpvOutBuffer:         [out] Pointer to the output buffer.
' cbOutBuffer:          [in] Size of the output buffer, in bytes.
' lpcbBytesReturned:    [out] Pointer to actual number of bytes of output.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped
'                       sockets).
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the operation has been
'                       completed (ignored for nonoverlapped sockets).
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' Upon successful completion, the WSAIoctl returns zero. Otherwise, a value of SOCKET_ERROR is
' returned, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAIoctl Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal dwIoControlCode As Long, _
                                     ByRef In_Buffer As Any, _
                                     ByVal In_BufferLen As Long, _
                                     ByRef Out_Buffer As Any, _
                                     ByVal Out_BufferLen As Long, _
                                     ByRef lpcbBytesReturned As Long, _
                                     ByRef lpOverlapped As Any, _
                                     ByVal lpCompletionRoutine As Long) As Long



' *************************************************************************************************
' WSAIsBlocking
' *************************************************************************************************
'
' Determines if a blocking call is in progress.
'
' This function has been removed in compliance with the Windows Sockets 2 specification,
' revision 2.2.0.
'
' The Windows Socket WSAIsBlocking function is not exported directly by the Ws2_32.dll, and
' Windows Sockets 2 applications should not use this function. Windows Sockets 1.1 applications
' that call this function are still supported through the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during
' calls to blocking functions. Instead of using blocking hooks, an applications should use a
' separate thread (separate from the main GUI thread) for network activity.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is TRUE if there is an outstanding blocking function awaiting completion in the
' current thread.  Otherwise, it is FALSE.
'
' *************************************************************************************************

Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long


' *************************************************************************************************
' WSAJoinLeaf
' *************************************************************************************************
'
' The WSAJoinLeaf function joins a leaf node into a multipoint session, exchanges connect data,
' and specifies needed quality of service based on the specified FLOWSPEC structures.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a multipoint socket.
' name:                 [in] Name of the peer to which the socket is to be joined.
' NameLen:              [in] Length of name, in bytes.
' lpCallerData:         [in] Pointer to the user data that is to be transferred to the peer
'                       during multipoint session establishment.
' lpCalleeData:         [out] Pointer to the user data that is to be transferred back from the
'                       peer during multipoint session establishment.
' lpSQOS:               [in] Pointer to the FLOWSPEC structures for socket s, one for each
'                       direction.
' lpGQOS:               [in] Reserved for future use with socket groups. A pointer to the
'                       FLOWSPEC structures for the socket group (if applicable).
' dwFlags:              [in] Flags to indicate that the socket is acting as a sender
'                       (JL_SENDER_ONLY), receiver (JL_RECEIVER_ONLY), or both (JL_BOTH).
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSAJoinLeaf returns a value of type SOCKET that is a descriptor for the
' newly created multipoint socket. Otherwise, a value of INVALID_SOCKET is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
'
' On a blocking socket, the return value indicates success or failure of the join operation.
'
' With a nonblocking socket, successful initiation of a join operation is indicated by a return of
' a valid socket descriptor. Subsequently, an FD_CONNECT indication will be given on the original
' socket s when the join operation completes, either successfully or otherwise. The application
' must use either WSAAsyncSelect or WSAEventSelect with interest registered for the FD_CONNECT
' event in order to determine when the join operation has completed and checks the associated
' error code to determine the success or failure of the operation. The select function cannot be
' used to determine when the join operation completes.
'
' Also, until the multipoint session join attempt completes all subsequent calls to WSAJoinLeaf
' on the same socket will fail with the error code WSAEALREADY. After the WSAJoinLeaf operation
' completes successfully, a subsequent attempt will usually fail with the error code WSAEISCONN.
' An exception to the WSAEISCONN rule occurs for a c_root socket that allows root-initiated joins.
' In such a case, another join may be initiated after a prior WSAJoinLeaf operation completes.
'
' If the return error code indicates the multipoint session join attempt failed (that is,
' WSAECONNREFUSED, WSAENETUNREACH, WSAETIMEDOUT) the application can call WSAJoinLeaf again for
' the same socket.
'
' *************************************************************************************************

Public Declare Function WSAJoinLeaf Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef PeerName As API_SOCKADDR_IN, _
                                     ByVal PeerNameLength As Long, _
                                     ByRef lpCallerData As API_WSABUF, _
                                     ByRef lpCalleeData As API_WSABUF, _
                                     ByRef lpSQOS As API_FLOWSPEC, _
                                     ByVal Reserved As Long, _
                                     ByVal dwFlags As Long) As Long


' *************************************************************************************************
' WSALookupServiceBegin
' *************************************************************************************************
'
' The WSALookupServiceBegin function initiates a client query that is constrained by the
' information contained within a WSAQUERYSET structure. WSALookupServiceBegin only returns a
' handle, which should be used by subsequent calls to WSALookupServiceNext to get the actual
' results.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpqsRestrictions:     [in] Pointer to the search criteria. See the following for details.
' dwControlFlags:       [in] Flag that controls the depth of the search.
' lphLookup:            [out] Handle to be used when calling WSALookupServiceNext in order to
'                       start retrieving the results set.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSALookupServiceBeginA Lib "ws2_32.dll" _
                                    (ByRef lpqsRestrictions As API_WSAQUERYSET, _
                                     ByVal dwControlFlags As Long, _
                                     ByRef lphLookup As Long) As Long


' *************************************************************************************************
' WSALookupServiceEnd
' *************************************************************************************************
'
' The WSALookupServiceEnd function is called to free the handle after previous calls to
' WSALookupServiceBegin and WSALookupServiceNext.
'
' If you call WSALookupServiceEnd from another thread while an existing WSALookupServiceNext is
' blocked, the end call will have the same effect as a cancel and will cause the
' WSALookupServiceNext call to return immediately.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hLookup:              [in] Handle previously obtained by calling WSALookupServiceBegin.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSALookupServiceEnd Lib "ws2_32.dll" _
                                    (ByVal hLookup As Long) As Long


' *************************************************************************************************
' WSALookupServiceNext
' *************************************************************************************************
'
' The WSALookupServiceNext function is called after obtaining a handle from a previous call to
' WSALookupServiceBegin in order to retrieve the requested service information.
'
' The provider will pass back a WSAQUERYSET structure in the lpqsResults buffer. The client
' should continue to call this function until it returns WSA_E_NO_MORE, indicating that all of
' WSAQUERYSET has been returned.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hLookup:              [in] Handle returned from the previous call to WSALookupServiceBegin.
' dwControlFlags:       [in] Flags to control the next operation. Currently only
'                       LUP_FLUSHPREVIOUS is defined as a means to cope with a result set which
'                       is too large. If an application does not (or cannot) supply a large enough
'                       buffer, setting LUP_FLUSHPREVIOUS instructs the provider to discard the
'                       last result set—which was too large—and move on to the next set for this
'                       call.
' lpdwBufferLength:     [in, out] On input, the number of bytes contained in the buffer pointed to
'                       by lpqsResults. On output, if the function fails and the error is
'                       WSAEFAULT, then it contains the minimum number of bytes to pass for the
'                       lpqsResults to retrieve the record.
' lpqsResults:          [out] Pointer to a block of memory, which will contain one result set in a
'                       WSAQUERYSET structure on return.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSALookupServiceNextA Lib "ws2_32.dll" _
                                    (ByVal hLookup As Long, _
                                     ByVal dwControlFlags As Long, _
                                     ByRef lpdwBufferLength As Long, _
                                     ByRef lpqsResults As API_WSAQUERYSET) As Long
                                     

' *************************************************************************************************
' WSANtohl
' *************************************************************************************************
'
' The WSANtohl function converts a u_long from network byte order to host byte order.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' NetLong:              [in] 32-bit number in network byte order.
' lpHostLong:           [out] Pointer to a 32-bit number in host byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSANtohl returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSANtohl Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal NetLong As Long, _
                                     ByRef lpHostLong As Long) As Long


' *************************************************************************************************
' WSANtohs
' *************************************************************************************************
'
' The WSANtohs function converts a u_short from network byte order to host byte order.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' NetShort:             [in] 16-bit number in network byte order.
' lpHostShort:          [out] Pointer to a 16-bit number in host byte order.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSANtohs returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSANtohs Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByVal NetShort As Integer, _
                                     ByRef lpHostShort As Integer) As Long


' *************************************************************************************************
' WSAProviderConfigChange
' *************************************************************************************************
'
' The WSAProviderConfigChange function notifies the application when the provider configuration is
' changed.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpNotificationHandle: [in, out] Pointer to notification handle. If the notification handle is
'                       set to NULL (the handle value not the pointer itself), this function
'                       returns a notification handle in the location pointed to by
'                       lpNotificationHandle.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure.
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the provider change
'                       notification is received.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs the WSAProviderConfigChange returns 0. Otherwise, a value of SOCKET_ERROR is
' returned and a specific error code may be retrieved by calling WSAGetLastError. The error code
' WSA_IO_PENDING indicates that the overlapped operation has been successfully initiated and that
' completion (and thus change event) will be indicated at a later time.
'
' *************************************************************************************************

Public Declare Function WSAProviderConfigChange Lib "ws2_43.dll" _
                                    (ByRef lpNotificationHandle As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpCompletionRoutine As Long) As Long


' *************************************************************************************************
' WSARecv
' *************************************************************************************************
'
' The WSARecv function receives data from a connected socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' lpBuffers:            [in, out] Pointer to an array of WSABUF structures. Each WSABUF structure
'                       contains a pointer to a buffer and the length of the buffer, in bytes.
' dwBufferCount:        [in] Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesRecvd: [out] Pointer to the number of bytes received by this call if the receive
'                       operation completes immediately.
' lpFlags:              [in, out] Pointer to flags.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped
'                       sockets).
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the receive operation
'                       has been completed (ignored for nonoverlapped sockets).
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs and the receive operation has completed immediately, WSARecv returns zero. In
' this case, the completion routine will have already been scheduled to be called once the calling
' thread is in the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates
' that the overlapped operation has been successfully initiated and that completion will be
' indicated at a later time. Any other error code indicates that the overlapped operation was not
' successfully initiated and no completion indication will occur.
'
' *************************************************************************************************

Public Declare Function WSARecv Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpBuffers As API_WSABUF, _
                                     ByVal dwBufferCount As Long, _
                                     ByRef lpNumberOfBytesRecvd As Long, _
                                     ByRef lpFlags As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpCompletionRoutine As Long) As Long


' *************************************************************************************************
' WSARecvDisconnect
' *************************************************************************************************
'
' The WSARecvDisconnect function terminates reception on a socket, and retrieves the disconnect
' data if the socket is connection oriented.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                        [in] Descriptor identifying a socket.
' lpInboundDisconnectData:  [out] Pointer to the incoming disconnect data.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSARecvDisconnect returns zero. Otherwise, a value of SOCKET_ERROR is
' returned, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSARecvDisconnect Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpInboundDisconnectData As API_WSABUF) As Long


' *************************************************************************************************
' WSARecvEx
' *************************************************************************************************
'
' The WSARecvEx function is identical to the recv function, except that the flags parameter is an
' [in, out] parameter. When a partial message is received while using datagram protocol, the
' MSG_PARTIAL bit is set in the flags parameter on return from the function.
'
' Note The WSARecvEx function is a Microsoft-specific extension to the Windows Sockets
' specification.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' buf:                  [out] Buffer for the incoming data.
' len:                  [in] Length of buf, in bytes.
' Flags:                [in, out] Indicator specifying whether the message is fully or partially received for datagram sockets.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSARecvEx returns the number of bytes received. If the connection has been
' closed, it returns zero. Additionally, if a partial message was received, the MSG_PARTIAL bit
' is set in the flags parameter. If a complete message was received, MSG_PARTIAL is not set in
' flags
'
' Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved by
' calling WSAGetLastError.
'
' Important For a stream oriented-transport protocol, MSG_PARTIAL is never set on return from
' WSARecvEx. This function behaves identically to the recv function for stream-transport protocols.
'
' *************************************************************************************************

Public Declare Function WSARecvEx Lib "wsock32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef Buffer As Any, _
                                     ByVal BufferLength As Long, _
                                     ByRef Flags As Long) As Long


' *************************************************************************************************
' WSARecvFrom
' *************************************************************************************************
'
' The WSARecvFrom function receives a datagram and stores the source address.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a socket.
' lpBuffers:            [in, out] Pointer to an array of WSABUF structures. Each WSABUF structure
'                       contains a pointer to a buffer and the length of the buffer.
' dwBufferCount:        [in] Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesRecvd: [out] Pointer to the number of bytes received by this call if the recv
'                       operation completes immediately.
' lpFlags:              [in, out] Pointer to flags.
' lpFrom:               [out] Optional pointer to a buffer that will hold the source address upon
'                       the completion of the overlapped operation.
' lpFromlen:            [in, out] Pointer to the size of the from buffer, in bytes, required only
'                       if lpFrom is specified.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped
'                       sockets).
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the recv operation has
'                       been completed (ignored for nonoverlapped sockets).
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs and the receive operation has completed immediately, WSARecvFrom returns zero.
' In this case, the completion routine will have already been scheduled to be called once the
' calling thread is in the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING
' indicates that the overlapped operation has been successfully initiated and that completion will
' be indicated at a later time. Any other error code indicates that the overlapped operation was
' not successfully initiated and no completion indication will occur.
'
' *************************************************************************************************

Public Declare Function WSARecvFrom Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpBuffers As API_WSABUF, _
                                     ByVal dwBufferCount As Long, _
                                     ByRef lpNumberOfBytesRecvd As Long, _
                                     ByRef lpFlags As Long, _
                                     ByRef lpFrom As API_SOCKADDR_IN, _
                                     ByRef lpFromlen As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpCompletionRoutine As Long) As Long


' *************************************************************************************************
' WSARecvMsg
' *************************************************************************************************
'
' The WSARecvMsg function receives data and optional control information from connected and
' unconnected sockets. The WSARecvMsg function can be used in place of the WSARecv and WSARecvFrom
' functions.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying the socket.
' lpMsg:                [in] A WSAMSG structure based on Posix.1g specification for the msghdr
'                       structure.
' lpdwNumberOfBytesRecvd: [out] Pointer to a variable the receives the number of bytes received.
'                               Available when the WSARecvMsg function call completes immediately.
' lpOverlapped;         [in] Pointer to a WSAOVERLAPPED structure. Ignored for nonoverlapped structures.
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the receive operation completes. Ignored for nonoverlapped structures.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' Success and immediate completion returns zero. When zero is returned, the specified completion
' routine is called once the calling thread is in the alertable state.
'
' A return value of SOCKET_ERROR and a subsequent call to WSAGetLastError that returns
' WSA_IO_PENDING indicates the overlapped operation has been successfully initiated, and
' completion will be indicated using other means (such as through events or completion ports).
'
' Failure returns SOCKET_ERROR and a subsequent call to the WSAGetLastError function returns an
' error other than WSA_IO_PENDING.
'
' *************************************************************************************************

Public Declare Function WSARecvMsg Lib "ws2_32.dll" _
                                    (ByVal s As Long, _
                                     ByVal lpdwNumberOfBytesRecvd As Long, _
                                     ByVal lpOverlapped As Long, _
                                     ByVal lpCompletionRoutine As Long) As Long


' *************************************************************************************************
' WSARemoveServiceClass
' *************************************************************************************************
'
' The WSARemoveServiceClass function permanently removes the service class schema from the registry.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpServiceClassId:     [in] Pointer to the GUID for the service class you want to remove.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSARemoveServiceClass Lib "ws2_32.dll" _
                                    (ByRef lpServiceClassId As API_GUID) As Long


' *************************************************************************************************
' WSAResetEvent
' *************************************************************************************************
'
' The WSAResetEvent function resets the state of the specified event object to nonsignaled.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hEvent:               [in] Handle that identifies an open event object handle.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the WSAResetEvent function succeeds, the return value is TRUE. If the function fails, the
' return value is FALSE. To get extended error information, call WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAResetEvent Lib "ws2_32.dll" _
                                    (ByRef hEvent As Long) As Long


' *************************************************************************************************
' WSASend
' *************************************************************************************************
'
' The WSASend function sends data on a connected socket.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a connected socket.
' lpBuffers:            [in] Pointer to an array of WSABUF structures. Each WSABUF structure
'                       contains a pointer to a buffer and the length of the buffer, in bytes.
'                       This array must remain valid for the duration of the send operation.
' dwBufferCount:        [in] Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesSent:  [out] Pointer to the number of bytes sent by this call if the I/O
'                       operation completes immediately.
' dwFlags:              [in] Flags used to modify the behavior of the WSASend function call. See
'                       Using dwFlags in the Remarks section for more information.
' lpOverlapped:         [in] Pointer to a WSAOVERLAPPED structure. This parameter is ignored for
'                       nonoverlapped sockets.
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the send operation has
'                       been completed. This parameter is ignored for nonoverlapped sockets.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs and the send operation has completed immediately, WSASend returns zero. In
' this case, the completion routine will have already been scheduled to be called once the calling
' thread is in the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates
' that the overlapped operation has been successfully initiated and that completion will be
' indicated at a later time. Any other error code indicates that the overlapped operation was not
' successfully initiated and no completion indication will occur.
'
' *************************************************************************************************

Public Declare Function WSASend Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpBuffers As API_WSABUF, _
                                     ByVal dwBufferCount As Long, _
                                     ByRef lpNumberOfBytesSent As Long, _
                                     ByVal dwFlags As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpCompletionRoutine As Long) As Long


' *************************************************************************************************
' WSASendDisconnect
' *************************************************************************************************
'
' The WSASendDisconnect function initiates termination of the connection for the socket and sends
' disconnect data.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                            [in] Descriptor identifying a socket.
' lpOutboundDisconnectData:     [in] Pointer to the outgoing disconnect data.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSASendDisconnect returns zero. Otherwise, a value of SOCKET_ERROR is
' returned, and a specific error code can be retrieved by calling WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSASendDisconnect Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpOutboundDisconnectData As API_WSABUF) As Long


' *************************************************************************************************
' WSASendTo
' *************************************************************************************************
'
' The WSASendTo function sends data to a specific destination, using overlapped I/O where
' applicable.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' s:                    [in] Descriptor identifying a (possibly connected) socket.
' lpBuffers:            [in] Pointer to an array of WSABUF structures. Each WSABUF structure
'                       contains a pointer to a buffer and the length of the buffer, in bytes.
'                       This array must remain valid for the duration of the send operation.
' dwBufferCount:        [in] Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesSent:  [out] Pointer to the number of bytes sent by this call if the I/O
'                       operation completes immediately.
' dwFlags:              [in] Indicator specifying the way in which the call is made.
' lpTo:                 [in] Optional pointer to the address of the target socket in the SOCKADDR
'                       structure.
' iToLen:               [in] Size of the address in lpTo, in bytes.
' lpOverlapped:         [in] A pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped
'                       sockets).
' lpCompletionRoutine:  [in] Pointer to the completion routine called when the send operation has
'                       been completed (ignored for nonoverlapped sockets).
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs and the send operation has completed immediately, WSASendTo returns zero. In
' this case, the completion routine will have already been scheduled to be called once the calling
' thread is in the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates
' that the overlapped operation has been successfully initiated and that completion will be
' indicated at a later time. Any other error code indicates that the overlapped operation was not
' successfully initiated and no completion indication will occur.
'
' *************************************************************************************************

Public Declare Function WSASendTo Lib "ws2_32.dll" _
                                    (ByVal hSocket As Long, _
                                     ByRef lpBuffers As API_WSABUF, _
                                     ByVal dwBufferCount As Long, _
                                     ByRef lpNumberOfBytesSent As Long, _
                                     ByVal dwFlags As Long, _
                                     ByRef lpTo As API_SOCKADDR_IN, _
                                     ByVal iToLen As Long, _
                                     ByRef lpOverlapped As API_WSAOVERLAPPED, _
                                     ByVal lpCompletionRoutine As Long) As Long

 
' *************************************************************************************************
' WSASetBlockingHook
' *************************************************************************************************
'
' Establish an application-supplied blocking hook function.
'
' This function has been removed in compliance with the Windows Sockets 2 specification, revision
' 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll, and Windows Sockets 2 applications
' should not use this function. Windows Sockets 1.1 applications that call this function are
' still supported through the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during
' calls to blocking functions. Instead of using blocking hooks, an application should use a
' separate thread separate from the main GUI thread) for network activity.
'
' Parameters
' -------------------------------------------------------------------------------------------------

' lpBlockFunc [in] A pointer to the procedure instance address of the blocking function to be installed.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is a pointer to the procedure-instance of the previously installed blocking
' function. The application or library that calls the WSASetBlockingHook() function should save
' this return value so that it can be restored if necessary.  (If "nesting" is not important, the
' application may simply discard the value returned by WSASetBlockingHook() and eventually use
' WSAUnhookBlockingHook() to restore the default mechanism.)  If the operation fails, a NULL
' pointer is returned, and a specific error number may be retrieved by calling WSAGetLastError().
'
' *************************************************************************************************

Public Declare Function WSASetBlockingHook Lib "wsock32.dll" _
                                    (ByVal lpBlockFunc As Long) As Long


' *************************************************************************************************
' WSASetEvent
' *************************************************************************************************
'
' The WSASetEvent function sets the state of the specified event object to signaled.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' hEvent:               [in] Handle that identifies an open event object.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is TRUE.
'
' If the function fails, the return value is FALSE. To get extended error information, call
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSASetEvent Lib "ws2_32.dll" _
                                    (ByVal hEvent As Long) As Long



' *************************************************************************************************
' WSASetLastError
' *************************************************************************************************
'
' The WSASetLastError function sets the error code that can be retrieved through the
' WSAGetLastError function.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' iError:               [in] Integer that specifies the error code to be returned by a subsequent
'                       WSAGetLastError call.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' This function generates no return values.
'
' *************************************************************************************************

Public Declare Sub WSASetLastError Lib "ws2_32.dll" _
                                    (ByVal iError As Long)


' *************************************************************************************************
' WSASetService
' *************************************************************************************************
'
' The WSASetService function registers or removes from the registry a service instance within one
' or more namespaces. This function can be used to affect a specific namespace provider, all
' providers associated with a specific namespace, or all providers across all namespaces.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' lpqsRegInfo:          [in] Pointer to the service information for registration or
'                       deregistration.
' essOperation:         [in] Enumeration whose values include the following: RNRSERVICE_REGISTER,
'                       RNRSERVICE_DEREGISTER, RNRSERVICE_DELETE
' dwControlFlags:       [in] Meaning of dwControlFlags is dependent on the following values:
'                       SERVICE_MULTIPLE
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value for WSASetService is zero if the operation was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSASetServiceA Lib "ws2_32.dll" _
                                    (ByRef lpqsRegInfo As API_WSAQUERYSET, _
                                     ByVal essOperation As Long, _
                                     ByVal dwControlFlags As Long) As Long


' *************************************************************************************************
' WSASocket
' *************************************************************************************************
'
' The WSASocket function creates a socket that is bound to a specific transport-service provider.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' af:                   [in] Address family specification.
' type:                 [in] Type specification for the new socket.
' Protocol:             [in] Protocol to be used with the socket that is specific to the indicated
'                       address family.
' lpProtocolInfo:       [in] Pointer to a WSAPROTOCOL_INFO structure that defines the
'                       characteristics of the socket to be created.
' g:                    [in] Reserved.
' dwFlags:              [in] Flag that specifies the socket attribute.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If no error occurs, WSASocket returns a descriptor referencing the new socket. Otherwise, a
' value of INVALID_SOCKET is returned, and a specific error code can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSASocketA Lib "ws2_32.dll" _
                                    (ByVal AddressFamily As Long, _
                                     ByVal SocketType As Long, _
                                     ByVal Protocol As Long, _
                                     ByRef lpProtocolInfo As API_WSAPROTOCOL_INFO, _
                                     ByVal Reserved As Long, _
                                     ByVal dwFlags As Long) As Long


' *************************************************************************************************
' WSAStartup
' *************************************************************************************************
'
' The WSAStartup function initiates use of WS2_32.DLL by a process.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' wVersionRequested:    [in] Highest version of Windows Sockets support that the caller can use.
'                       The high-order byte specifies the minor version (revision) number; the
'                       low-order byte specifies the major version number.
' lpWSAData:            [out] Pointer to the WSADATA data structure that is to receive details of
'                       the Windows Sockets implementation.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The WSAStartup function returns zero if successful. Otherwise, it returns an error code
'
' An application cannot call WSAGetLastError to determine the error code as is normally done in
' Windows Sockets if WSAStartup fails. The WS2_32.DLL will not have been loaded in the case of a
' failure so the client data area where the last error information is stored could not be
' established.
'
' *************************************************************************************************

Public Declare Function WSAStartup Lib "ws2_32.dll" _
                                    (ByVal wVersionRequested As Integer, _
                                     ByRef lpWSAData As API_WSADATA) As Long


' *************************************************************************************************
' WSAStringToAddress
' *************************************************************************************************
'
' The WSAStringToAddress function converts a numeric string to a sockaddr structure, suitable for
' passing to Windows Sockets routines that take such a structure.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' AddressString:        [in] Pointer to the zero-terminated human-readable numeric string to
'                       convert.
' AddressFamily:        [in] Address family to which the string belongs.
' lpProtocolInfo:       [in] (optional) The WSAPROTOCOL_INFO structure associated with the
'                       provider to be used. If this is NULL, the call is routed to the provider
'                       of the first protocol supporting the indicated AddressFamily.
' lpAddress:            [out] Buffer that is filled with a single sockaddr.
' lpAddressLength:      [in, out] Length of the Address buffer, in bytes. Returns the size of the
'                       resultant sockaddr structure. If the specified buffer is not large enough,
'                       the function fails with a specific error of WSAEFAULT and this parameter
'                       is updated with the required size in bytes.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value for WSAStringToAddress is zero if the operation was successful. Otherwise, the
' value SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
'
' *************************************************************************************************

Public Declare Function WSAStringToAddressA Lib "ws2_32.dll" _
                                    (ByVal AddressString As String, _
                                     ByVal AddressFamily As Long, _
                                     ByRef lpProtocolInfo As API_WSAPROTOCOL_INFO, _
                                     ByRef lpAddress As API_SOCKADDR_IN, _
                                     ByRef lpAddressLength As Long) As Long


' *************************************************************************************************
' WSAUnhookBlockingHook
' *************************************************************************************************
'
' Restores the default blocking hook function.
'
' This function has been removed in compliance with the Windows Sockets 2 specification, revision
' 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll, and Windows Sockets 2 applications
' should not use this function. Windows Sockets 1.1 applications that call this function are
' still supported through the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during
' calls to blocking functions. Instead of using blocking hooks, an application should use a
' separate thread (separate from the main GUI thread) for network activity.

' Return Values
' -------------------------------------------------------------------------------------------------
'
' The return value is 0 if the operation was successful.  Otherwise the value SOCKET_ERROR is
' returned, and a specific error number may be retrieved by calling WSAGetLastError().
'
' *************************************************************************************************

Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long


' *************************************************************************************************
' WSAWaitForMultipleEvents
' *************************************************************************************************
'
' The WSAWaitForMultipleEvents function returns either when one or all of the specified event
' objects are in the signaled state, or when the time-out interval expires.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' cEvents:              [in] Indicator specifying the number of event object handles in the array
'                       pointed to by lphEvents. The maximum number of event object handles is
'                       WSA_MAXIMUM_WAIT_EVENTS. One or more events must be specified.
' lphEvents:            [in] Pointer to an array of event object handles.
' fWaitAll:             [in] Indicator specifying the wait type. If TRUE, the function returns
'                       when the state of all objects in the lphEvents array is signaled. If
'                       FALSE, the function returns when any of the event objects is signaled. In
'                       the latter case, the return value indicates the event object whose state
'                       caused the function to return.
' dwTimeout:            [in] Indicator specifying the time-out interval, in milliseconds. The
'                       function returns if the interval expires, even if conditions specified by
'                       the fWaitAll parameter are not satisfied. If dwTimeout is zero, the
'                       function tests the state of the specified event objects and returns
'                       immediately. If dwTimeout is WSA_INFINITE, the function's time-out
'                       interval never expires.
' fAlertable:           [in] Indicator specifying whether the function returns when the system
'                       queues an I/O completion routine for execution by the calling thread. If
'                       TRUE, the completion routine is executed and the function returns. If
'                       FALSE, the completion routine is not executed when the function returns.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the WSAWaitForMultipleEvents function succeeds, the return value indicates the event object
' that caused the function to return.
'
' If the function fails, the return value is WSA_WAIT_FAILED. To get extended error information,
' call WSAGetLastError.
'
' The return value upon success is one of the following values.
'
' *************************************************************************************************

Public Declare Function WSAWaitForMultipleEvents Lib "ws2_32.dll" _
                                    (ByVal cEvents As Long, _
                                     ByVal lphEvents As Long, _
                                     ByVal fWaitAll As Long, _
                                     ByVal dwTimeout As Long, _
                                     ByVal fAlertable As Long) As Long



' *************************************************************************************************
' GetIfTable
' *************************************************************************************************
'
' The GetIfTable function retrieves the MIB-II interface table.
'
' Parameters
' -------------------------------------------------------------------------------------------------
'
' Parameters
' pIfTable:             [out] Pointer to a buffer that receives the interface table as a MIB_IFTABLE
'                       structure.
' pdwSize:              [in, out] On input, specifies the size of the buffer pointed to by the
'                       pIfTable parameter. On output, if the buffer is not large enough to hold the
'                       returned interface table, the function sets this parameter equal to the
'                       required buffer size.
' bOrder:               [in] Specifies whether the returned interface table should be sorted in
'                       ascending order by interface index. If this parameter is TRUE, the table is
'                       sorted.
'
' Return Values
' -------------------------------------------------------------------------------------------------
'
' If the function succeeds, the return value is NO_ERROR.
'
' If the function fails, the return value is one of the following error codes.
'
' ERROR_INSUFFICIENT_BUFFER     The buffer pointed to by the pIfTable parameter is not large enough.
'                               The required size is returned in the DWORD variable pointed to by the
'                               pdwSize parameter.
' ERROR_INVALID_PARAMETER       The pdwSize parameter is NULL, or GetIfTable is unable to write to
'                               the memory pointed to by the pdwSize parameter.
' ERROR_NOT_SUPPORTED           This function is not supported on the operating system in use on the
'                               local system.
' Other                         Use FormatMessage to obtain the message string for the returned error.
'
' *************************************************************************************************

Public Declare Function api_GetIfTable Lib "iphlpapi" (ByRef pIfRowTable As Any, _
                                                       ByRef pdwSize As Long, _
                                                       ByVal bOrder As Long) As Long


' *************************************************************************************************
' addrinfo
' *************************************************************************************************
'
' The addrinfo structure is used by the getaddrinfo function to hold host address information.
'
' *************************************************************************************************
Public Type API_ADDRINFO
    
    ' Flags that indicate options used in the getaddrinfo function. See AI_PASSIVE, AI_CANONNAME,
    ' and AI_NUMERICHOST.
    ai_flags      As Long
    
    ' Protocol family, such as PF_INET.
    ai_family     As Long
    
    ' Socket type, such as SOCK_RAW, SOCK_STREAM, or SOCK_DGRAM.
    ai_socktype   As Long
    
    ' Protocol, such as IPPROTO_TCP or IPPROTO_UDP. For protocols other than IPv4 and IPv6, set
    ' ai_protocol to zero.
    ai_protocol   As Long
    
    ' Length of the ai_addr member, in bytes.
    ai_addrlen    As Long
    
    ' Canonical name for the host.
    ai_canonname  As String
    
    ' Pointer to a sockaddr structure.
    ai_addr       As Long
    
    ' Pointer to the next structure in a linked list. This parameter is set to NULL in the last
    ' addrinfo structure of a linked list.
    ai_next       As Long
End Type


' *************************************************************************************************
' AFPROTOCOLS
' *************************************************************************************************
'
' The AFPROTOCOLS structure supplies a list of protocols to which application programmers can
' constrain queries. The AFPROTOCOLS structure is used for query purposes only.
'
' *************************************************************************************************
Public Type API_AFPROTOCOLS

    ' Address family to which the query is to be constrained.
    iAddressFamily  As Integer
    
    ' Protocol to which the query is to be constrained.
    iProtocol       As Integer
End Type


' *************************************************************************************************
' BLOB
' *************************************************************************************************
'
' The BLOB structure, derived from Binary Large Object, contains information about a block of data.
'
' *************************************************************************************************
Public Type API_BLOB

    ' Size of the block of data pointed to by pBlobData, in bytes.
    cbSize    As Long
    
    ' Pointer to a block of data.
    pBlobData As Byte
End Type

' *************************************************************************************************
' SOCKET_ADDRESS
' *************************************************************************************************
'
' The SOCKET_ADDRESS structure stores protocol-specific address information.
'
' *************************************************************************************************
Public Type API_SOCKET_ADDRESS

    ' Pointer to a socket address
    lpSockaddr      As Long
    
    ' Length of the socket address, in bytes.
    iSockaddrLength As Long
End Type


' *************************************************************************************************
' SOCKET_ADDRESS_LIST
' *************************************************************************************************
Public Type API_SOCKET_ADDRESS_LIST

    ' number of address structures in the list
    iAddressCount   As Long
    
    ' array of protocol family specific address structures.
    lpAddresses     As Long
End Type


' *************************************************************************************************
' CSADDR_INFO
' *************************************************************************************************
'
' The CSADDR_INFO structure contains Windows Sockets address information for a network service or
' namespace provider. The GetAddressByName function obtains Windows Sockets address information
' using CSADDR_INFO structures.
'
' *************************************************************************************************
Type API_CSADDR_INFO

    ' In a client application, pass this address to the bind function to obtain access to a network
    ' service.
    ' In a network service, pass this address to the bind function so that the service is bound to
    ' the appropriate local address.
    LocalAddr   As API_SOCKET_ADDRESS
    
    ' You can use this remote address to connect to the service through the connect function.
    ' You can also use this remote address with the sendto function when you are communicating
    ' over a connectionless (datagram) protocol.
    RemoteAddr  As API_SOCKET_ADDRESS
    
    ' Type of Windows socket.
    iSocketType As Integer
    
    ' Value to pass as the protocol parameter to the socket function to open a socket for this service.
    iProtocol   As Integer
End Type


' *************************************************************************************************
' fd_set
' *************************************************************************************************
'
' The fd_set structure is used by various Windows Sockets functions and service providers, such
' as the select function, to place sockets into a "set" for various purposes, such as testing a
' given socket for readability using the readfds parameter of the select function.
'
' *************************************************************************************************
Public Type API_FD_SET

    ' Number of sockets in the set.
    fd_count     As Long
    
    ' Array of sockets that are in the set.
    fd_array(63) As Long
End Type

' *************************************************************************************************
' FLOWSPEC
' *************************************************************************************************
'
' Flow Specifications for each direction of data flow.
'
' *************************************************************************************************
Public Type API_FLOWSPEC

    ' In Bytes/sec
    TokenRate          As Long
    
    ' In Bytes
    TokenBucketSize    As Long
    
    ' In Bytes/sec
    PeakBandwidth      As Long
    
    ' In microseconds
    Latency            As Long
    
    ' In microseconds
    DelayVariation     As Long
    
    ' ServiceType
    ServiceType        As Long
    
    ' In Bytes
    MaxSduSize         As Long
    
    ' In Bytes
    MinimumPolicedSize As Long
End Type


' *************************************************************************************************
' Guid
' *************************************************************************************************
'
' The GUID data type is a text string representing a Class identifier(ID). COM must be able to
' convert the string to a valid Class ID. All GUIDs must be authored in uppercase. The valid
' format for a GUID is {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} where X is a hex digit.
'
' *************************************************************************************************
Public Type API_GUID
  Data1    As Long
  Data2    As Integer
  Data3    As Integer
  Data4(0 To 7) As Byte
End Type


' *************************************************************************************************
' hostent
' *************************************************************************************************
'
' The hostent structure is used by functions to store information about a given host, such as
' host name, IP address, and so forth. An application should never attempt to modify this
' structure or to free any of its components. Furthermore, only one copy of the hostent structure
' is allocated per thread, and an application should therefore copy any information that it needs
' before issuing any other Windows Sockets API calls.
'
' *************************************************************************************************
Public Type API_HOSTENT

    ' Official name of the host (PC). If using the DNS or similar resolution system, it is the
    ' Fully Qualified Domain Name (FQDN) that caused the server to return a reply. If using a
    ' local hosts file, it is the first entry after the IP address.
    hName           As Long
    
    ' NULL terminated array of alternate names.
    hAliases        As Long
    
    ' Type of address being returned.
    hAddrType       As Integer
    
    ' Length of each address, in bytes.
    hLength         As Integer
    
    ' NULL terminated list of addresses for the host. Addresses are returned in network byte
    ' order. The macro h_addr is defined to be h_addr_list[0] for compatibility with older
    ' software.
    hAddrList       As Long
End Type


' *************************************************************************************************
' in_addr
' *************************************************************************************************
'
' The in_addr structure represents a host by its Internet address.
'
' NOTE: Whenver a function calls for a "IN_ADDR" structure, use long instead and pass the results
' of inet_addr("xxx.xxx.xxx.xxx")
' *************************************************************************************************
Public Type API_IN_ADDR

    ' Address of the host formatted as a u_long
    S_addr          As Long
End Type


' *************************************************************************************************
' in_pktinfo
' *************************************************************************************************
'
' The in_pktinfo structure is used to store received packet address information, and is used by
' Windows to return information about received packets.
'
' *************************************************************************************************
Public Type API_IN_PKTINFO
    
    ' Destination address from the IP header of the received packet.
    ipi_addr        As API_IN_ADDR

    ' Interface on which the packet was received.
    ipi_ifindex     As Long
End Type

' *************************************************************************************************
' in6_addr
' *************************************************************************************************
'
' The in6_addr structure represents a host by its Internet address.
'
' *************************************************************************************************
Public Type API_IN_ADDR6

    ' IP v6 address
    s6_addr(16)     As Byte
End Type


' *************************************************************************************************
' in6_pktinfo
' *************************************************************************************************
'
' The in6_pktinfo structure is used to store received IPv6 packet address information, and is used
' by Windows to return information about received packets.
'
' *************************************************************************************************
Public Type API_IN6_PKTINFO

    ' Destination IPv6 address from the IP header of the received packet.
    ipi6_addr       As API_IN_ADDR6
    
    ' Interface on which the packet was received.
    ipi6_ifindex    As Long
End Type


' *************************************************************************************************
' ip_mreq
' *************************************************************************************************
'
' The ip_mreq structure provides multicast group information, and is used with the
' IP_ADD_MEMBERSHIP and IP_DROP_MEMBERSHIP socket options.
'
' *************************************************************************************************
Public Type API_IP_MREQ

    'IP address of the multicast group.
    imr_multiaddr   As API_IN_ADDR

    ' Local IP address of the interface on which the multicast group should be joined or dropped.
    imr_interface  As API_IN_ADDR
End Type


' *************************************************************************************************
' ip_mreq_source
' *************************************************************************************************
'
' The ip_mreq_source structure provides multicast group information for IGMPv3.
'
' *************************************************************************************************
Public Type API_IP_MREQ_SOURCE

    ' IP address of the multicast group.
    imr_multiaddr   As API_IN_ADDR
    
    ' IP address of the multicast source.
    imr_sourceaddr  As API_IN_ADDR
    
    ' Local IP address of the interface on which the multicast group should be joined, dropped,
    ' blocked, or unblocked.
    imr_interface   As API_IN_ADDR
End Type


' *************************************************************************************************
' ip_msfilter
' *************************************************************************************************
'
' The ip_msfilter structure provides multicast filtering parameters for IGMPv3.
'
' *************************************************************************************************
Public Type API_IP_MSFILTER

    ' IP address of the multicast group.
    imsf_multiaddr  As API_IN_ADDR
    
    ' Local IP address of the interface.
    imsf_interface  As API_IN_ADDR
    
    ' Filter mode to be used, either INCLUDE to include particular multicast source(s), or
    ' EXCLUDE to exclude traffic from specified source(s). Set to zero for INCLUDE, set to 1 for
    ' EXCLUDE.
    imsf_fmode      As API_IN_ADDR
    
    ' Number of sources in the imsf_slist member.
    imsf_numsrc     As Long
    
    ' Array of in_addr structures specifying the multicast sources to include or exclude.
    ' This should actually be treated as a block of memory of size: imsf_numsrc * sizeof(in_addr).
    imsf_slist      As API_IN_ADDR
End Type


' *************************************************************************************************
' ipv6_mreq
' *************************************************************************************************
'
' The ipv6_mreq structure provides multicast group information for IPv6 addresses.
'
' *************************************************************************************************
Public Type API_IPV6_MREQ

    ' Address of the IPv6 multicast group.
    ipv6mr_multiaddr    As API_IN_ADDR6
        
    ' Local IPv6 address of the interface on which the multicast group should be joined or dropped.
    ipv6mr_interface    As Long
End Type


' *************************************************************************************************
' linger
' *************************************************************************************************
'
' The linger structure maintains information about a specific socket that specifies how that
' socket should behave when data is queued to be sent and the closesocket function is called on
' the socket.
'
' *************************************************************************************************
Public Type API_LINGER
    
    ' Specifies whether a socket should remain open for a specified amount of time after a
    ' closesocket function call to enable queued data to be sent.
    l_onoff     As Integer

    ' Enabling SO_LINGER also disables SO_DONTLINGER, and vice versa. Note that if SO_DONTLINGER
    ' is DISABLED (that is, SO_LINGER is ENABLED) then no time-out value is specified. In this
    ' case, the time-out used is implementation dependent. If a previous time-out has been
    ' established for a socket (by enabling SO_LINGER), this time-out value should be reinstated
    ' by the service provider.
    l_linger    As Integer
End Type


' *************************************************************************************************
' SERVICE_ADDRESS
' *************************************************************************************************
'
' The SERVICE_ADDRESS structure contains address information for a service. The structure can
' accommodate many types of interprocess communications (IPC) mechanisms and their address forms,
' including remote procedure calls (RPC), named pipes, and sockets.
'
' *************************************************************************************************
Public Type API_SERVICE_ADDRESS
    
    ' Address family to which the socket address pointed to by lpAddress belongs.
    dwAddressType       As Long
    
    ' Set of bit flags that specify properties of the address.
    dwAddressFlags      As Long
    
    ' Size of the address, in bytes.
    dwAddressLength     As Long
    
    ' Reserved for future use. Must be zero.
    dwPrincipalLength   As Long
    
    ' Pointer to a socket address of the appropriate type.
    lpAddress           As Long
    
    ' Reserved for future use. Must be NULL.
    lpPrincipal         As Long
End Type


Public Type API_SERVICE_ADDRESSES

    ' Specifies the number of SERVICE_ADDRESS structures in the Addresses array.
    dwAddressCount  As Long
    
    ' An array of SERVICE_ADDRESS data structures. Each SERVICE_ADDRESS structure contains
    ' information about a network service address.
    Addresses()     As API_SERVICE_ADDRESS
End Type


' *************************************************************************************************
' SERVICE_INFO
' *************************************************************************************************
'
' The SERVICE_INFO structure contains information about a network service or a network service
' type.
'
' *************************************************************************************************

Public Type API_SERVICE_INFO

    ' Pointer to a GUID that is the type of the network service.
    lpServiceType       As Long
    
    ' Pointer to a zero-terminated string that is the name of the network service.
    lpServiceName       As Long
    
    ' Pointer to a zero-terminated string that is a comment or description for the network service.
    lpComment           As String
    
    ' Pointer to a zero-terminated string that contains locale information.
    lpLocale            As String
    
    ' Specifies a hint as to how to display the network service in a network browsing user
    ' interface.
    dwDisplayHint       As Long
    
    ' Version information for the network service. The high word of this value specifies a major
    ' version number. The low word of this value specifies a minor version number.
    dwVersion           As Long
    
    ' Reserved for future use. Must be set to zero.
    Reserved            As Long
    
    ' Pointer to a zero-terminated string that is the name of the computer on which the network
    ' service is running.
    lpMachineName       As String
    
    ' Pointer to a SERVICE_ADDRESSES structure that contains an array of SERVICE_ADDRESS
    ' structures. Each SERVICE_ADDRESS structure contains information about a network service
    ' address.
    lpServiceAddress    As Long
    
    ' A BLOB structure that specifies service-defined information. Note that In general, the data
    ' pointed to by the BLOB structure's pBlobData member must not contain any pointers. That is
    ' because only the network service knows the format of the data; copying the data without such
    ' knowledge would lead to pointer invalidation. If the data pointed to by pBlobData contains
    ' variable-sized elements, offsets from pBlobData can be used to indicate the location of
    ' those elements.
    ServiceSpecificInfo As API_BLOB
End Type


' *************************************************************************************************
' NS_SERVICE_INFO
' *************************************************************************************************
'
' The NS_SERVICE_INFO structure contains information about a network service or a network service
' type in the context of a specified namespace, or a set of default namespaces.
'
' *************************************************************************************************
Public Type API_NS_SERVICE_INFO
  
  ' Specifies the name space or a set of default name spaces to which this service information
  ' applies.  Use one of the following constant values to specify a name space: NS_DEFAULT,
  ' NS_DNS, NS_MS, NS_NDS, NS_NETBT, NS_NIS, NS_SAP, NS_STDA, NS_TCPIP_HOSTS, NS_TCPIP_LOCAL,
  ' NS_WINS, NS_X500
    dwNameSpace     As Long
  
  ' A SERVICE_INFO structure that contains information about a network service or network service
  ' type.
    ServiceInfo     As API_SERVICE_INFO

End Type


' *************************************************************************************************
' PROTOCOL_INFO
' *************************************************************************************************
'
' The PROTOCOL_INFO structure contains information about a protocol.
'
' *************************************************************************************************
Public Type API_PROTOCOL_INFO 'Requires Windows Sockets 1.1 or later

    ' A set of bit flags that specifies the services provided by the protocol.  One or more of the
    ' XP_* bit flags may be set.
    dwServiceFlags  As Long
    
    ' Value to pass as the af parameter when the socket function is called to open a socket for
    ' the protocol. This address family value uniquely defines the structure of protocol
    ' addresses, also known as sockaddr structures, used by the protocol.
    iAddressFamily  As Long
    
    ' Maximum length of a socket address supported by the protocol.
    iMaxSockAddr    As Long
    
    ' Minimum length of a socket address supported by the protocol.
    iMinSockAddr    As Long
    
    ' Value to pass as the type parameter when the socket function is called to open a socket for
    ' the protocol. Note that if XP_PSEUDO_STREAM is set in dwServiceFlags, the application can
    ' specify SOCK_STREAM as the type parameter to socket, regardless of the value of iSocketType.
    iSocketType     As Long
    
    ' Value to pass as the protocol parameter when the socket function is called to open a socket
    ' for the protocol.
    iProtocol       As Long
    
    ' Maximum message size supported by the protocol. This is the maximum size of a message that
    ' can be sent from or received by the host. For protocols that do not support message framing,
    ' the actual maximum size of a message that can be sent to a given address may be less than
    ' this value.
    dwMessageSize   As Long

    ' Points to a zero-terminated string that supplies a name for the protocol; for example, "SPX2"
    lpProtocol      As String
End Type


' *************************************************************************************************
' protoent
' *************************************************************************************************
'
' The protoent structure contains the name and protocol numbers that correspond to a given
' protocol name. Applications must never attempt to modify this structure or to free any of its
' components. Furthermore, only one copy of this structure is allocated per thread, and therefore,
' the application should copy any information it needs before issuing any other Windows Sockets
' function calls.
'
' *************************************************************************************************
Public Type API_PROTOENT

    ' Official name of the protocol
    p_name    As Long
    
    ' Null-terminated array of alternate names
    p_aliases As Long
    
    ' Protocol number, in host byte order
    p_proto   As Integer
End Type



' *************************************************************************************************
' WSABUF
' *************************************************************************************************
'
' The WSABUF structure enables the creation or manipulation of a data buffer.
'
' *************************************************************************************************
Public Type API_WSABUF

    ' The length of the buffer
    len     As Long
    
    ' The pointer to the buffer
    buf     As Long
End Type


' *************************************************************************************************
' QOS
' *************************************************************************************************
'
' QualityOfService
'
' *************************************************************************************************
Public Type API_WSA_QOS

    ' The flow spec for data sending
    SendingFlowspec   As API_FLOWSPEC
    
    ' The flow spec for data receiving
    ReceivingFlowspec As API_FLOWSPEC
    
    ' Additional provider specific stuff
    ProviderSpecific  As API_WSABUF
End Type


' *************************************************************************************************
' servent
' *************************************************************************************************
'
' The servent structure is used to store or return the name and service number for a given service
' name.
'
' *************************************************************************************************
Public Type API_SERVENT

    ' Official name of the service.
    s_name    As Long
    
    ' Null-terminated array of alternate names.
    s_aliases As Long
    
    ' Port number at which the service can be contacted. Port numbers are returned in network byte
    ' order.
    s_port    As Integer
    
    ' Name of the protocol to use when contacting the service.
    s_proto   As Long
End Type


' *************************************************************************************************
' sockaddr
' *************************************************************************************************
'
' The sockaddr structure varies depending on the protocol selected. Except for the sin*_family
' parameter, sockaddr contents are expressed in network byte order.
'
' Winsock functions using sockaddr are not strictly interpreted to be pointers to a sockaddr
' structure. The structure is interpreted differently in the context of different address families.
' The only requirements are that the first u_short is the address family and the total size of the
' memory buffer in bytes is namelen.
'
' The structures below are used with IPv4 and IPv6, respectively. Other protocols use similar
' structures.
'
' *************************************************************************************************

Public Type API_SOCKADDR_IN
    sin_family          As Integer
    sin_port            As Integer
    sin_addr            As API_IN_ADDR
    sin_zero(1 To 8)    As Byte
End Type

Public Type API_SOCKADDR_IN6
    sin6_family         As Integer
    sin6_port           As Integer
    sin6_flowinfo       As Long
    sin6_addr           As API_IN_ADDR6
    sin6_scope_id       As Long
End Type

Public Type API_SOCKADDR_GEN    ' Generic
    AddressIn                           As API_SOCKADDR_IN
    filler(0 To 7)                      As Byte
End Type


' *************************************************************************************************
' interface_info
' *************************************************************************************************
'
' Get interface info
'
' *************************************************************************************************

Public Type API_INTERFACE_INFO
    iiFlags             As Long                 'Interface flags
    iiAddress           As API_SOCKADDR_GEN     'Interface address
    iiBroadcastAddress  As API_SOCKADDR_GEN     'Broadcast address
    iiNetmask           As API_SOCKADDR_GEN     'Network mask
End Type


Public Type API_INTERFACEINFO
    iInfo(0 To 7) As API_INTERFACE_INFO
End Type


' Possible flags for the iiFlags - bitmask '

Public Const IFF_UP             As Long = &H1        ' Interface is up
Public Const IFF_BROADCAST      As Long = &H2        ' Broadcast is supported
Public Const IFF_LOOPBACK       As Long = &H4        ' this is loopback interface
Public Const IFF_POINTTOPOINT   As Long = &H8        ' this is point-to-point interface
Public Const IFF_MULTICAST      As Long = &H10       ' multicast is supported


Public Enum InterfaceInfoFlags
    Int_Up = IFF_UP
    Int_Broadcast = IFF_BROADCAST
    Int_Loopback = IFF_LOOPBACK
    Int_PointToPoint = IFF_POINTTOPOINT
    Int_Multicast = IFF_MULTICAST
End Enum


' *************************************************************************************************
' timeval
' *************************************************************************************************
'
' The timeval structure is used to specify time values. It is associated with the Berkeley
' Software Distribution (BSD) file Time.h.
'
' *************************************************************************************************
Public Type API_TIMEVAL

    ' Seconds
    tv_sec  As Long
    
    ' Microseconds
    tv_usec As Long
End Type


' *************************************************************************************************
' TRANSMIT_FILE_BUFFERS
' *************************************************************************************************
'
' The TRANSMIT_FILE_BUFFERS structure specifies data to be transmitted before and after file data
' during a TransmitFile function file transfer operation.
'
' *************************************************************************************************
Public Type API_TRANSMIT_FILE_BUFFERS

    ' Pointer to a buffer that contains data to be transmitted before the file data is transmitted.
    Head        As Long
    
    ' Size of the buffer pointed to by Head, in bytes, to be transmitted.
    HeadLength  As Long
    
    'Pointer to a buffer that contains data to be transmitted after the file data is transmitted.
    Tail        As Long
    
    ' Size of the buffer pointed to Tail, in bytes, to be transmitted.
    TailLength  As Long
End Type


' *************************************************************************************************
' TRANSMIT_PACKETS_ELEMENT
' *************************************************************************************************
'
' The TRANSMIT_PACKETS_ELEMENT structure specifies a single data element to be transmitted by the
' TransmitPackets function.
'
' *************************************************************************************************
Public Type API_TRANSMIT_PACKETS_ELEMENT
    
    ' Flags used to describe the contents of the packet array element, and to customize
    ' TransmitPackets function processing.
    dwElFlags       As Long
    
    ' Number of bytes to transmit. If zero, the entire file is transmitted.
    cLength         As Long

    ' File offset at which to begin the transfer. Valid only if TP_ELEMENT_FILE is specified in
    ' dwEIFlags. When set to –1, transmission begins at the current byte offset.
    nFileOffset     As Currency
    
    ' Handle to an open file to be transmitted. Valid only if TP_ELEMENT_FILE is specified in
    ' dwEIFlags. Windows reads the file sequentially; caching performance is improved by opening
    ' this handle with FILE_FLAG_SEQUENTIAL_SCAN.
    hFile           As Long
End Type


' Use this in the case TP_ELEMENT_MEMORY
Public Type API_TRANSMIT_PACKETS_ELEMENT2
    
    ' Flags used to describe the contents of the packet array element, and to customize
    ' TransmitPackets function processing.
    dwElFlags       As Long
    
    ' Number of bytes to transmit. If zero, the entire file is transmitted.
    cLength         As Long
    
    ' Pointer to the data in memory to be sent. Valid only if TP_ELEMENT_MEMORY is specified in
    ' dwEIFlags.
    pBuffer         As Long
End Type


' *************************************************************************************************
' WSADATA
' *************************************************************************************************
'
' The WSADATA structure contains information about the Windows Sockets implementation.
'
' NOTE : An application should ignore the iMaxsockets, iMaxUdpDg, and lpVendorInfo members in
' WSAData if the value in wVersion after a successful call to WSAStartup is at least 2. This
' is because the architecture of Windows Sockets has been changed in version 2 to support
' multiple providers, and WSAData no longer applies to a single vendor's stack. Two new
' socket options are introduced to supply provider-specific information: SO_MAX_MSG_SIZE
' (replaces the iMaxUdpDg element) and PVD_CONFIG (allows any other provider-specific
' configuration to occur).
'
' *************************************************************************************************
Public Type API_WSADATA

    ' Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
    wVersion       As Integer
    
    ' Highest version of the Windows Sockets specification that this .dll can support
    ' (also encoded as above). Normally this is the same as wVersion.
    wHighVersion   As Integer
    
    ' Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows
    ' Sockets implementation. The text (up to 256 characters in length) can contain any
    ' characters except control and formatting characters: the most likely use that an application
    ' can put this to is to display it (possibly truncated) in a status message.
    szDescription  As String * 256
    
    ' Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or
    ' configuration information. The Ws2_32.dll should use this parameter only if the information
    ' might be useful to the user or support staff: it should not be considered as an extension of
    ' the szDescription parameter.
    szSystemStatus As String * 128
    
    ' Retained for backward compatibility, but should be ignored for Windows Sockets version 2 and
    ' later, as no single value can be appropriate for all underlying service providers.
    iMaxSockets    As Integer
    
    ' Ignored for Windows Sockets version 2 and onward. iMaxUdpDg is retained for compatibility
    ' with Windows Sockets specification 1.1, but should not be used when developing new
    ' applications. For the actual maximum message size specific to a particular Windows Sockets
    ' service provider and socket type, applications should use getsockopt to retrieve the value
    ' of option SO_MAX_MSG_SIZE after a socket has been created.
    iMaxUdpDg      As Integer
    
    ' Ignored for Windows Sockets version 2 and onward. It is retained for compatibility with
    ' Windows Sockets specification 1.1. Applications needing to access vendor-specific
    ' configuration information should use getsockopt to retrieve the value of option PVD_CONFIG.
    ' The definition of this value (if utilized) is beyond the scope of this specification.
    lpVendorInfo   As Long
End Type


' *************************************************************************************************
' WSANAMESPACE_INFO
' *************************************************************************************************
'
' The WSANAMESPACE_INFO structure contains all registration information for a namespace provider.
'
' *************************************************************************************************
Public Type API_WSANAMESPACE_INFO
    
    ' Unique identifier for this name-space provider.
    NSProviderId   As API_GUID
    
    ' Name space supported by this implementation of the provider.
    dwNameSpace    As Long
    
    ' If TRUE, indicates that this provider is active. If FALSE, the provider is inactive and is
    ' not accessible for queries, even if the query specifically references this provider.
    fActive        As Long
    
    ' Name space–version identifier.
    dwVersion      As Long
    
    ' Display string for the provider.
    lpszIdentifier As String
End Type


' *************************************************************************************************
' WSANETWORKEVENTS
' *************************************************************************************************
'
' The WSANETWORKEVENTS structure is used to store a socket's internal information about network
' events.
'
' *************************************************************************************************
Public Type API_WSANETWORKEVENTS

    ' Indicates which of the FD_XXX network events have occurred.
    lNetworkEvents As Long
    
    ' An array that contains any associated error codes, with an array index that corresponds to
    ' the position of event bits in lNetworkEvents. The identifiers FD_READ_BIT, FD_WRITE_BIT and
    ' other can be used to index the iErrorCode array.
    iErrorCode(7)  As Long
End Type


' *************************************************************************************************
' WSANSCLASSINFO
' *************************************************************************************************
'
' The WSANSCLASSINFO structure provides individual parameter information for a specific Windows
' Sockets namespace.
'
' *************************************************************************************************
Public Type API_WSANSCLASSINFO

    ' String value associated with the parameter, such as SAPID, TCPPORT, and so forth.
    lpszName    As String
    
    ' GUID associated with the namespace.
    dwNameSpace As Long
    
    ' Value type for the parameter, such as REG_DWORD or REG_SZ, and so forth.
    dwValueType As Long
    
    ' Size of the parameter provided in lpValue, in bytes.
    dwValueSize As Long
    
    ' Pointer to the value of the parameter.
    lpValue     As Long
End Type


' *************************************************************************************************
' WSAOVERLAPPED
' *************************************************************************************************
'
' The WSAOVERLAPPED structure provides a communication medium between the initiation of an
' overlapped I/O operation and its subsequent completion. The WSAOVERLAPPED structure is designed
' to be compatible with the Windows OVERLAPPED structure:
'
' *************************************************************************************************
Public Type API_WSAOVERLAPPED

    ' Reserved for internal use. The Internal member is used internally by the entity that
    ' implements overlapped I/O. For service providers that create sockets as installable file
    ' system (IFS) handles, this parameter is used by the underlying operating system. Other
    ' service providers (non-IFS providers) are free to use this parameter as necessary.
    Internal     As Long
    
    ' Reserved. Used internally by the entity that implements overlapped I/O. For service
    ' providers that create sockets as IFS handles, this parameter is used by the underlying
    ' operating system. NonIFS providers are free to use this parameter as necessary.
    InternalHigh As Long
    
    ' Reserved for use by service providers.
    Offset       As Long
    
    ' Reserved for use by service providers.
    OffsetHigh   As Long
    
    ' If an overlapped I/O operation is issued without an I/O completion routine
    ' (lpCompletionRoutine is null), then this parameter should either contain a valid handle to
    ' a WSAEVENT object or be null. If lpCompletionRoutine is non-null then applications are free
    ' to use this parameter as necessary.
    hEvent       As Long
End Type


' *************************************************************************************************
' WSAPROTOCOLCHAIN
' *************************************************************************************************
'
' The WSAPROTOCOLCHAIN structure contains a counted list of Catalog Entry identifiers that comprise
' a protocol chain.
'
' *************************************************************************************************
Public Const MAX_PROTOCOL_CHAIN As Long = 7

Public Const BASE_PROTOCOL      As Long = 1
Public Const LAYERED_PROTOCOL   As Long = 0

Public Type API_WSAPROTOCOLCHAIN

    ' Length of the chain. The following settings apply: Setting ChainLen to zero indicates a
    ' layered protocol, Setting ChainLen to one indicates a base protocol, Setting ChainLen to
    ' greater than one indicates a protocol chain
    ChainLen        As Long
    
    ' Array of protocol chain entries.
    ChainEntries(MAX_PROTOCOL_CHAIN - 1) As Long
End Type


'Public Enum ProtocolChainType
'    ' If the length of the chain is 0, this WSAPROTOCOL_INFO entry represents a layered protocol
'    ' which has Windows Sockets 2 SPI as both its top and bottom edges.
'    LayeredProtocol = LAYERED_PROTOCOL
'
'    ' If the length of the chain equals 1, this entry represents a base protocol whose Catalog
'    ' Entry identifier is in the dwCatalogEntryId member of the WSAPROTOCOL_INFO structure.
'    BaseProtocol = BASE_PROTOCOL
'
'    ' If the length of the chain is larger than 1, this entry represents a protocol chain which
'    ' consists of one or more layered protocols on top of a base protocol.
'    LayeredProtocolChain = &H2
'End Enum


' *************************************************************************************************
' PROTOCOL_INFO
' *************************************************************************************************
'
' The WSAPROTOCOL_INFO structure is used to store or retrieve complete information for a given protocol.
'
' *************************************************************************************************
Public Type API_WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As API_GUID
    dwCatalogEntryId As Long
    ProtocolChain As API_WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type


' *************************************************************************************************
' Byte ordering
' *************************************************************************************************

Public Const BO_BIGENDIAN As Long = &H0
Public Const BO_LITTLEENDIAN As Long = &H1


'Public Enum ByteOrder
'    BOrd_BigEndian = BO_BIGENDIAN
'    BOrd_LittleEndian = BO_LITTLEENDIAN
'End Enum


' *************************************************************************************************
' Data type and manifest constants for socket groups
' *************************************************************************************************

Public Type API_GROUP
    grp As Integer     'an unsigned integer
End Type
  
Public Const SG_UNCONSTRAINED_GROUP = &H1
Public Const SG_CONSTRAINED_GROUP = &H2

' *************************************************************************************************
' WSAVERSION
' *************************************************************************************************
'
' The WSAVERSION structure provides version comparison in Windows Sockets.
'
' *************************************************************************************************
Public Type API_WSAVERSION

    ' Version of Windows Sockets.
   dwVersion As Long
   
   ' WSAECOMPARATOR enumeration, used in the comparison.
   ecHow     As Long
End Type


' *************************************************************************************************
' WSAQUERYSET
' *************************************************************************************************
'
' The WSAQUERYSET structure provides relevant information about a given service, including service
' class ID, service name , applicable name-space identifier and protocol information, as well as
' a set of transport addresses at which the service listens.
'
' *************************************************************************************************
Public Type API_WSAQUERYSET

    ' Must be set to sizeof(WSAQUERYSET). This is a versioning mechanism.
    dwSize                  As Long
    
    ' Ignored for queries.
    lpszServiceInstanceName As String
    
    ' (Optional) Referenced string contains service name. The semantics for using wildcards
    ' within the string are not defined, but can be supported by certain name space providers.
    lpServiceClassId        As API_GUID
    
    ' (Required) The GUID corresponding to the service class.
    lpVersion               As API_WSAVERSION
    
    ' (Optional) References desired version number and provides version comparison semantics
    ' (that is, version must match exactly, or version must be not less than the value supplied).
    lpszComment             As String
    
    ' Ignored for queries.
    dwNameSpace             As Long
    
    ' Identifier of a single name space in which to constrain the search, or NS_ALL to include all
    ' name spaces.
    lpNSProviderId          As API_GUID
    
    ' (Optional) References the GUID of a specific name-space provider, and limits the query to
    ' this provider only.
    lpszContext             As String
    
    ' (Optional) Specifies the starting point of the query in a hierarchical name space.
    dwNumberOfProtocols     As String
    
    ' Size of the protocol constraint array, can be zero.
    lpafpProtocols          As API_AFPROTOCOLS
    
    ' (Optional) References an array of AFPROTOCOLS structure. Only services that utilize these
    ' protocols will be returned.
    lpszQueryString()       As String
    
    ' (Optional) Some name spaces (such as Whois++) support enriched SQL-like queries that are
    ' contained in a simple text string. This parameter is used to specify that string.
    dwNumberOfCsAddrs       As Long
    
    ' Ignored for queries.
    lpcsaBuffer             As Long         'CSADDR_INFO
    
    ' Ignored for queries.
    dwOutputFlags           As Long
    
    ' (Optional) This is a pointer to a provider-specific entity.
    lpBlob                  As Long        'BLOB
End Type


' *************************************************************************************************
' WSASERVICECLASSINFO
' *************************************************************************************************
'
' The WSASERVICECLASSINFO structure contains information about a specified service class. For each service class in Windows Sockets 2, there is a single WSASERVICECLASSINFO structure.
'
' *************************************************************************************************
Public Type API_WSASERVICECLASSINFO

    ' Unique Identifier (GUID) for the service class.
    lpServiceClassId     As Long           'GUID
    
    ' Well known associated with the service class.
    lpszServiceClassName As String
    
    ' Number of entries in lpClassInfos.
    dwCount              As Long
    
    ' Array of WSANSCLASSINFOW structures that contains information about the service class.
    lpClassInfos()       As API_WSANSCLASSINFO
End Type


' *************************************************************************************************
' TCP_KEEPALIVE
' *************************************************************************************************

Public Type API_TCP_KEEPALIVE
    onoff               As Long
    KeepAliveTime       As Long
    KeepAliveInterval   As Long
End Type



' *************************************************************************************************
' All Windows Sockets error constants are biased by WSABASEERR from the "normal"
' *************************************************************************************************

Public Const WSABASEERR              As Long = 10000

' *************************************************************************************************
' Windows Sockets definitions of regular Microsoft C error constants
' *************************************************************************************************

Public Const WSAEINTR                As Long = (WSABASEERR + 4)
Public Const WSAEBADF                As Long = (WSABASEERR + 9)
Public Const WSAEACCES               As Long = (WSABASEERR + 13)
Public Const WSAEFAULT               As Long = (WSABASEERR + 14)
Public Const WSAEINVAL               As Long = (WSABASEERR + 22)
Public Const WSAEMFILE               As Long = (WSABASEERR + 24)

' *************************************************************************************************
' Windows Sockets definitions of regular Berkeley error constants
' *************************************************************************************************

Public Const WSAEWOULDBLOCK          As Long = (WSABASEERR + 35)
Public Const WSAEINPROGRESS          As Long = (WSABASEERR + 36)
Public Const WSAEALREADY             As Long = (WSABASEERR + 37)
Public Const WSAENOTSOCK             As Long = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ         As Long = (WSABASEERR + 39)
Public Const WSAEMSGSIZE             As Long = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE           As Long = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT          As Long = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT      As Long = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT      As Long = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP           As Long = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT         As Long = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT         As Long = (WSABASEERR + 47)
Public Const WSAEADDRINUSE           As Long = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL        As Long = (WSABASEERR + 49)
Public Const WSAENETDOWN             As Long = (WSABASEERR + 50)
Public Const WSAENETUNREACH          As Long = (WSABASEERR + 51)
Public Const WSAENETRESET            As Long = (WSABASEERR + 52)
Public Const WSAECONNABORTED         As Long = (WSABASEERR + 53)
Public Const WSAECONNRESET           As Long = (WSABASEERR + 54)
Public Const WSAENOBUFS              As Long = (WSABASEERR + 55)
Public Const WSAEISCONN              As Long = (WSABASEERR + 56)
Public Const WSAENOTCONN             As Long = (WSABASEERR + 57)
Public Const WSAESHUTDOWN            As Long = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS         As Long = (WSABASEERR + 59)
Public Const WSAETIMEDOUT            As Long = (WSABASEERR + 60)
Public Const WSAECONNREFUSED         As Long = (WSABASEERR + 61)
Public Const WSAELOOP                As Long = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG         As Long = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN            As Long = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH         As Long = (WSABASEERR + 65)
Public Const WSAENOTEMPTY            As Long = (WSABASEERR + 66)
Public Const WSAEPROCLIM             As Long = (WSABASEERR + 67)
Public Const WSAEUSERS               As Long = (WSABASEERR + 68)
Public Const WSAEDQUOT               As Long = (WSABASEERR + 69)
Public Const WSAESTALE               As Long = (WSABASEERR + 70)
Public Const WSAEREMOTE              As Long = (WSABASEERR + 71)

' *************************************************************************************************
' Extended Windows Sockets error constant definitions
' *************************************************************************************************

Public Const WSASYSNOTREADY          As Long = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED      As Long = (WSABASEERR + 92)
Public Const WSANOTINITIALISED       As Long = (WSABASEERR + 93)
Public Const WSAEDISCON              As Long = (WSABASEERR + 101)
Public Const WSAENOMORE              As Long = (WSABASEERR + 102)
Public Const WSAECANCELLED           As Long = (WSABASEERR + 103)
Public Const WSAEINVALIDPROCTABLE    As Long = (WSABASEERR + 104)
Public Const WSAEINVALIDPROVIDER     As Long = (WSABASEERR + 105)
Public Const WSAEPROVIDERFAILEDINIT  As Long = (WSABASEERR + 106)
Public Const WSASYSCALLFAILURE       As Long = (WSABASEERR + 107)
Public Const WSASERVICE_NOT_FOUND    As Long = (WSABASEERR + 108)
Public Const WSATYPE_NOT_FOUND       As Long = (WSABASEERR + 109)
Public Const WSA_E_NO_MORE           As Long = (WSABASEERR + 110)
Public Const WSA_E_CANCELLED         As Long = (WSABASEERR + 111)
Public Const WSAEREFUSED             As Long = (WSABASEERR + 112)


' *************************************************************************************************
' Define QOS related error return codes
' *************************************************************************************************

' at least one Reserve has arrived
Public Const WSA_QOS_RECEIVERS                As Long = (WSABASEERR + 1005)

' at least one Path has arrived
Public Const WSA_QOS_SENDERS                  As Long = (WSABASEERR + 1006)

' there are no senders
Public Const WSA_QOS_NO_SENDERS               As Long = (WSABASEERR + 1007)

' there are no receivers
Public Const WSA_QOS_NO_RECEIVERS             As Long = (WSABASEERR + 1008)

' Reserve has been confirmed
Public Const WSA_QOS_REQUEST_CONFIRMED        As Long = (WSABASEERR + 1009)

' error due to lack of resources
Public Const WSA_QOS_ADMISSION_FAILURE        As Long = (WSABASEERR + 1010)

' rejected for administrative reasons - bad credentials
Public Const WSA_QOS_POLICY_FAILURE           As Long = (WSABASEERR + 1011)

' unknown or conflicting style
Public Const WSA_QOS_BAD_STYLE                As Long = (WSABASEERR + 1012)

' problem with some part of the filterspec or providerspecific buffer in general
Public Const WSA_QOS_BAD_OBJECT               As Long = (WSABASEERR + 1013)

' problem with some part of the flowspec
Public Const WSA_QOS_TRAFFIC_CTRL_ERROR       As Long = (WSABASEERR + 1014)

' general error
Public Const WSA_QOS_GENERIC_ERROR            As Long = (WSABASEERR + 1015)

' An invalid or unrecognized service type was found in the flowspec
Public Const WSA_QOS_ESERVICETYPE             As Long = (WSABASEERR + 1016)

' An invalid or inconsistent flowspec was found in the QOS structure
Public Const WSA_QOS_EFLOWSPEC                As Long = (WSABASEERR + 1017)

' Invalid QOS provider-specific buffer
Public Const WSA_QOS_EPROVSPECBUF             As Long = (WSABASEERR + 1018)

' An invalid QOS filter style was used
Public Const WSA_QOS_EFILTERSTYLE             As Long = (WSABASEERR + 1019)

' An invalid QOS filter type was used
Public Const WSA_QOS_EFILTERTYPE              As Long = (WSABASEERR + 1020)

' An incorrect number of QOS FILTERSPECs were specified in the FLOWDESCRIPTOR
Public Const WSA_QOS_EFILTERCOUNT             As Long = (WSABASEERR + 1021)

' An object with an invalid ObjectLength field was specified in the QOS provider-specific buffer
Public Const WSA_QOS_EOBJLENGTH               As Long = (WSABASEERR + 1022)

' An incorrect number of flow descriptors was specified in the QOS structure
Public Const WSA_QOS_EFLOWCOUNT               As Long = (WSABASEERR + 1023)

' An unrecognized object was found in the QOS provider-specific buffer
Public Const WSA_QOS_EUNKOWNPSOBJ             As Long = (WSABASEERR + 1024)

' An invalid policy object was found in the QOS provider-specific buffer
Public Const WSA_QOS_EPOLICYOBJ               As Long = (WSABASEERR + 1025)

' An invalid QOS flow descriptor was found in the flow descriptor list
Public Const WSA_QOS_EFLOWDESC                As Long = (WSABASEERR + 1026)

' An invalid or inconsistent flowspec was found in the QOS provider-specific buffer
Public Const WSA_QOS_EPSFLOWSPEC              As Long = (WSABASEERR + 1027)

' An invalid FILTERSPEC was found in the QOS provider-specific buffer
Public Const WSA_QOS_EPSFILTERSPEC            As Long = (WSABASEERR + 1028)

' An invalid shape discard mode object was found in the QOS provider-specific buffer
Public Const WSA_QOS_ESDMODEOBJ               As Long = (WSABASEERR + 1029)

' An invalid shaping rate object was found in the QOS provider-specific buffer
Public Const WSA_QOS_ESHAPERATEOBJ            As Long = (WSABASEERR + 1030)

' A reserved policy element was found in the QOS provider-specific buffer
Public Const WSA_QOS_RESERVED_PETYPE          As Long = (WSABASEERR + 1031)


' *************************************************************************************************
' Error return codes from gethostbyname() and gethostbyaddr()
' (when using the resolver). Note that these errors are
' retrieved via WSAGetLastError() and must therefore follow
' the rules for avoiding clashes with error numbers from
' specific implementations or language run-time systems.
' For this reason the codes are based at WSABASEERR+1001.
' Note also that [WSA]NO_ADDRESS is defined only for
' compatibility purposes.
' *************************************************************************************************

' Authoritative Answer: Host not found
Public Const WSAHOST_NOT_FOUND       As Long = (WSABASEERR + 1001)
'Public Const HOST_NOT_FOUND          As Long = WSAHOST_NOT_FOUND

' Non-Authoritative: Host not found, or SERVERFAI
Public Const WSATRY_AGAIN            As Long = (WSABASEERR + 1002)
'Public Const TRY_AGAIN               As Long = WSATRY_AGAIN

' Non-recoverable errors, FORMERR, REFUSED, NOTIMP
Public Const WSANO_RECOVERY          As Long = (WSABASEERR + 1003)
'Public Const NO_RECOVERY             As Long = WSANO_RECOVERY

' Valid name, no data record of requested type
Public Const WSANO_DATA              As Long = (WSABASEERR + 1004)
'Public Const NO_DATA                 As Long = WSANO_DATA

' no address, look for MX record
Public Const WSANO_ADDRESS           As Long = WSANO_DATA
'Public Const NO_ADDRESS              As Long = WSANO_ADDRESS


Public Enum ResolverErrorCodeType
    ResolveErr_HostNotFound = WSAHOST_NOT_FOUND
    ResolveErr_TryAgain = WSATRY_AGAIN
    ResolveErr_NonRecovebrable = WSANO_RECOVERY
    ResolveErr_NoData = WSANO_DATA
    ResolveErr_NoAddress = WSANO_ADDRESS
End Enum
    

' *************************************************************************************************
' General error and return codes
' *************************************************************************************************

Public Const WSA_IO_PENDING          As Long = 997&
Public Const WSA_IO_INCOMPLETE       As Long = 996&
Public Const WSA_INVALID_HANDLE      As Long = 6&
Public Const WSA_INVALID_PARAMETER   As Long = 87&
Public Const WSA_NOT_ENOUGH_MEMORY   As Long = 8&
Public Const WSA_OPERATION_ABORTED   As Long = 995&

Public Const WSA_MAXIMUM_WAIT_EVENTS As Long = 64
Public Const WSA_WAIT_EVENT_0        As Long = &H0
Public Const WSA_WAIT_IO_COMPLETION  As Long = &HC0&
Public Const WSA_WAIT_TIMEOUT        As Long = &H102&
Public Const WSA_INFINITE            As Long = &HFFFFFFFF

' General socket errors
Public Const SOCKET_ERROR            As Long = -1
Public Const INVALID_SOCKET          As Long = SOCKET_ERROR


'' *************************************************************************************************
'' A list of error codes which may be returned by various winsock functions.
'' *************************************************************************************************
'Public Enum ErrorCodeType
'
'    ' The operation completed successfuly
'    NoError = 0
'
'    ' A blocking operation was interrupted by a call to WSACancelBlockingCall.
'    InterruptedFunctionCall = WSAEINTR
'
'    ' An attempt was made to access a socket in a way forbidden by its access permissions.
'    ' An example is using a broadcast address for sendto without broadcast permission being
'    ' set using setsockopt(SO_BROADCAST).
'    '
'    ' Another possible reason for the WSAEACCES error is that when the bind function is called
'    ' (on Windows NT 4 SP4 or later), another application, service, or kernel mode driver is
'    ' bound to the same address with exclusive access. Such exclusive access is a new feature
'    ' of Windows NT 4 SP4 and later, and is implemented by using the SO_EXCLUSIVEADDRUSE option.
'    PermissionDenied = WSAEACCES
'
'    ' The system detected an invalid pointer address in attempting to use a pointer argument of
'    ' a call. This error occurs if an application passes an invalid pointer value, or if the
'    ' length of the buffer is too small. For instance, if the length of an argument, which is a
'    ' sockaddr structure, is smaller than the sizeof(sockaddr).
'    BadAddress = WSAEFAULT
'
'    ' Some invalid argument was supplied (for example, specifying an invalid level to the
'    ' setsockopt function). In some instances, it also refers to the current state of the
'    ' socket—for instance, calling accept on a socket that is not listening.
'    InvalidArgument = WSAEINVAL
'
'    ' Each implementation may have a maximum number of socket handles available, either globally,
'    ' per process, or per thread.
'    TooManyOpenFiles = WSAEMFILE
'
'    ' This error is returned from operations on nonblocking sockets that cannot be completed
'    ' immediately, for example recv when no data is queued to be read from the socket. It is a
'    ' nonfatal error, and the operation should be retried later. It is normal for WSAEWOULDBLOCK
'    ' to be reported as the result from calling connect on a nonblocking SOCK_STREAM socket,
'    ' since some time must elapse for the connection to be established.
'    ResourceTemporarilyUnavailable = 10035
'
'    ' A blocking operation is currently executing. Windows Sockets only allows a single blocking
'    ' operation—per- task or thread—to be outstanding, and if any other function call is made
'    ' (whether or not it references that or any other socket) the function fails with the
'    ' WSAEINPROGRESS error.
'    OperationNowInProgress = WSAEINPROGRESS
'
'    ' An operation was attempted on a nonblocking socket with an operation already in
'    ' progress—that is, calling connect a second time on a nonblocking socket that is already
'    ' connecting, or canceling an asynchronous request (WSAAsyncGetXbyY) that has already been
'    ' canceled or completed.
'    OperationAlreadyInProgress = WSAEALREADY
'
'    ' An operation was attempted on something that is not a socket. Either the socket handle
'    ' parameter did not reference a valid socket, or for select, a member of an fd_set was not
'    ' valid.
'    SocketOperationOnNonSocket = WSAENOTSOCK
'
'    ' A required address was omitted from an operation on a socket. For example, this error is
'    ' returned if sendto is called with the remote address of ADDR_ANY.
'    DestinationAddressRequired = WSAEDESTADDRREQ
'
'    ' A message sent on a datagram socket was larger than the internal message buffer or some
'    ' other network limit, or the buffer used to receive a datagram was smaller than the
'    ' datagram itself.
'    MessageTooLong = WSAEMSGSIZE
'
'    ' A protocol was specified in the socket function call that does not support the semantics
'    ' of the socket type requested. For example, the ARPA Internet UDP protocol cannot be
'    ' specified with a socket type of SOCK_STREAM.
'    ProtocolWrongTypeForSocket = 10041
'
'    ' An unknown, invalid or unsupported option or level was specified in a getsockopt or
'    ' setsockopt call.
'    BadProtocolOption = WSAENOPROTOOPT
'
'    ' The requested protocol has not been configured into the system, or no implementation for
'    ' it exists. For example, a socket call requests a SOCK_DGRAM socket, but specifies a stream
'    ' protocol.
'    ProtocolNotSupported = WSAEPROTONOSUPPORT
'
'    ' The support for the specified socket type does not exist in this address family.
'    ' For example, the optional type SOCK_RAW might be selected in a socket call, and the
'    ' implementation does not support SOCK_RAW sockets at all.
'    SocketTypeNotSupported = WSAESOCKTNOSUPPORT
'
'    ' The attempted operation is not supported for the type of object referenced. Usually this
'    ' occurs when a socket descriptor to a socket that cannot support this operation is trying
'    ' to accept a connection on a datagram socket.
'    OperationNotSupported = WSAEOPNOTSUPP
'
'    ' The protocol family has not been configured into the system or no implementation for it
'    ' exists. This message has a slightly different meaning from WSAEAFNOSUPPORT. However, it
'    ' is interchangeable in most cases, and all Windows Sockets functions that return one of
'    ' these messages also specify WSAEAFNOSUPPORT.
'    ProtocolFamilyNotSupported = WSAEPFNOSUPPORT
'
'    ' An address incompatible with the requested protocol was used. All sockets are created with
'    ' an associated address family (that is, AF_INET for Internet Protocols) and a generic
'    ' protocol type (that is, SOCK_STREAM). This error is returned if an incorrect protocol is
'    ' explicitly requested in the socket call, or if an address of the wrong family is used for
'    ' a socket, for example, in sendto.
'    AddressFamilyNotSupportedByProtocolFamily = WSAEAFNOSUPPORT
'
'    ' Typically, only one usage of each socket address (protocol/IP address/port) is permitted.
'    ' This error occurs if an application attempts to bind a socket to an IP address/port that
'    ' has already been used for an existing socket, or a socket that was not closed properly, or
'    ' one that is still in the process of closing. For server applications that need to bind
'    ' multiple sockets to the same port number, consider using setsockopt (SO_REUSEADDR).
'    ' Client applications usually need not call bind at all—connect chooses an unused port
'    ' automatically. When bind is called with a wildcard address (involving ADDR_ANY), a
'    ' WSAEADDRINUSE error could be delayed until the specific address is committed. This could
'    ' happen with a call to another function later, including connect, listen, WSAConnect, or
'    ' WSAJoinLeaf.
'    AddressAlreadyInUse = WSAEADDRINUSE
'
'    ' The requested address is not valid in its context. This normally results from an attempt to
'    ' bind to an address that is not valid for the local computer. This can also result from
'    ' connect, sendto, WSAConnect, WSAJoinLeaf, or WSASendTo when the remote address or port is
'    ' not valid for a remote computer (for example, address or port 0).
'    CannotAssignRequestedAddress = WSAEADDRNOTAVAIL
'
'    ' A socket operation encountered a dead network. This could indicate a serious failure of the
'    ' network system (that is, the protocol stack that the Windows Sockets DLL runs over), the
'    ' network interface, or the local network itself.
'    NetworkIsDown = WSAENETDOWN
'
'    ' A socket operation was attempted to an unreachable network. This usually means the local
'    ' software knows no route to reach the remote host.
'    NetworkIsUnreachable = WSAENETUNREACH
'
'    ' The connection has been broken due to keep-alive activity detecting a failure while the
'    ' operation was in progress. It can also be returned by setsockopt if an attempt is made to
'    ' set SO_KEEPALIVE on a connection that has already failed.
'    NetworkDroppedConnectionOnReset = WSAENETRESET
'
'    ' An established connection was aborted by the software in your host computer, possibly due
'    ' to a data transmission time-out or protocol error.
'    SoftwareCausedConnectionAbort = WSAECONNABORTED
'
'    ' An existing connection was forcibly closed by the remote host. This normally results if the
'    ' peer application on the remote host is suddenly stopped, the host is rebooted, the host or
'    ' remote network interface is disabled, or the remote host uses a hard close (see setsockopt
'    ' for more information on the SO_LINGER option on the remote socket). This error may also
'    ' result if a connection was broken due to keep-alive activity detecting a failure while one
'    ' or more operations are in progress. Operations that were in progress fail with WSAENETRESET.
'    ' Subsequent operations fail with WSAECONNRESET.
'    ConnectionResetByPeer = WSAECONNRESET
'
'    ' An operation on a socket could not be performed because the system lacked sufficient buffer
'    ' space or because a queue was full.
'    NoBufferSpaceAvailable = WSAENOBUFS
'
'    ' A connect request was made on an already-connected socket. Some implementations also
'    ' return this error if sendto is called on a connected SOCK_DGRAM socket (for SOCK_STREAM
'    ' sockets, the to parameter in sendto is ignored) although other implementations treat this
'    ' as a legal occurrence.
'    SocketIsAlreadyConnected = WSAEISCONN
'
'    ' A request to send or receive data was disallowed because the socket is not connected and
'    ' (when sending on a datagram socket using sendto) no address was supplied. Any other type
'    ' of operation might also return this error—for example, setsockopt setting SO_KEEPALIVE if
'    ' the connection has been reset.
'    SocketIsNotConnected = WSAENOTCONN
'
'    ' A request to send or receive data was disallowed because the socket had already been shut
'    ' down in that direction with a previous shutdown call. By calling shutdown a partial close
'    ' of a socket is requested, which is a signal that sending or receiving, or both have been
'    ' discontinued.
'    CannotSendAfterSocketShutdown = WSAESHUTDOWN
'
'    ' A connection attempt failed because the connected party did not properly respond after a
'    ' period of time, or the established connection failed because the connected host has failed
'    ' to respond.
'    ConnectionTimedOut = WSAETIMEDOUT
'
'    ' No connection could be made because the target computer actively refused it. This usually
'    ' results from trying to connect to a service that is inactive on the foreign host—that is,
'    ' one with no server application running.
'    ConnectionRefused = WSAECONNREFUSED
'
'    ' A socket operation failed because the destination host is down. A socket operation
'    ' encountered a dead host. Networking activity on the local host has not been initiated.
'    ' These conditions are more likely to be indicated by
'    HostIsDown = WSAEHOSTDOWN
'
'    ' A socket operation was attempted to an unreachable host. See WSAENETUNREACH.
'    NoRouteToHost = WSAEHOSTUNREACH
'
'    ' A Windows Sockets implementation may have a limit on the number of applications that can
'    ' use it simultaneously. WSAStartup may fail with this error if the limit has been reached.
'    TooManyProcesses = WSAEPROCLIM
'
'    ' This error is returned by WSAStartup if the Windows Sockets implementation cannot function
'    ' at this time because the underlying system it uses to provide network services is currently
'    ' unavailable. Users should check:
'    '
'    ' - That the appropriate Windows Sockets DLL file is in the current path.
'    '
'    ' - That they are not trying to use more than one Windows Sockets implementation
'    '   simultaneously. If there is more than one Winsock DLL on your system, be sure the first
'    '   one in the path is appropriate for the network subsystem currently loaded.
'    '
'    ' - The Windows Sockets implementation documentation to be sure all necessary components are
'    '   currently installed and configured correctly.
'    '
'    NetworkSubsystemIsUnavailable = WSASYSNOTREADY
'
'    ' The current Windows Sockets implementation does not support the Windows Sockets
'    ' specification version requested by the application. Check that no old Windows Sockets DLL
'    ' files are being accessed.
'    WinsockDllVersionOutOfRange = WSAVERNOTSUPPORTED
'
'    ' Either the application has not called WSAStartup or WSAStartup failed. The application may
'    ' be accessing a socket that the current active task does not own (that is, trying to share a
'    ' socket between tasks), or WSACleanup has been called too many times.
'    SuccessfulWSAStartupNotYetPerformed = WSANOTINITIALISED
'
'    ' Returned by WSARecv and WSARecvFrom to indicate that the remote party has initiated a
'    ' graceful shutdown sequence.
'    GracefulShutdownInProgress = WSAEDISCON
'
'    ' The specified class was not found.
'    ClassTypeNotFound = WSATYPE_NOT_FOUND
'
'    ' No such host is known. The name is not an official host name or alias, or it cannot be
'    ' found in the database(s) being queried. This error may also be returned for protocol and
'    ' service queries, and means that the specified name could not be found in the relevant
'    ' database.
'    HostNotFound = WSAHOST_NOT_FOUND
'
'    ' This is usually a temporary error during host name resolution and means that the local
'    ' server did not receive a response from an authoritative server. A retry at some time later
'    ' may be successful.
'    NonAuthoritativeHostNotFound = WSATRY_AGAIN
'
'    ' This indicates some sort of nonrecoverable error occurred during a database lookup. This
'    ' may be because the database files (for example, BSD-compatible HOSTS, SERVICES, or
'    ' PROTOCOLS files) could not be found, or a DNS request was returned by the server with a
'    ' severe error.
'    NonRecoverableError = WSANO_RECOVERY
'
'    ' The requested name is valid and was found in the database, but it does not have the
'    ' correct associated data being resolved for. The usual example for this is a host
'    ' name-to-address translation attempt (using gethostbyname or WSAAsyncGetHostByName)
'    ' which uses the DNS (Domain Name Server). An MX record is returned but no A
'    ' record—indicating the host itself exists, but is not directly reachable.
'    ValidNameButNoDataRecordOfRequestedType = WSANO_DATA
'
'    ' Specified event object handle is invalid.
'    ' An application attempts to use an event object, but the specified handle is not valid.
'    InvalidHandle = WSA_INVALID_HANDLE ' OS Dependant
'
'    ' An application used a Windows Sockets function which directly maps to a Windows function.
'    ' The Windows function is indicating a problem with one or more parameters.
'    InvalidParameter = WSA_INVALID_PARAMETER ' OS Dependant
'
'    ' Overlapped I/O event object not in signaled state.
'    ' The application has tried to determine the status of an overlapped operation which is not
'    ' yet completed. Applications that use WSAGetOverlappedResult (with the fWait flag set to
'    ' FALSE) in a polling mode to determine when an overlapped operation has completed.
'    IOEventIncomplete = WSA_IO_INCOMPLETE  ' OS Dependant
'
'    ' Overlapped operations will complete later.
'    ' The application has initiated an overlapped operation that cannot be completed immediately.
'    ' A completion indication will be given later when the operation has been completed.
'    IOOperationPending = WSA_IO_PENDING ' OS Dependant
'
'    ' Insufficient memory available.
'    ' An application used a Windows Sockets function that directly maps to a Windows function.
'    ' The Windows function is indicating a lack of required memory resources.
'    InsufficientMemory = WSA_NOT_ENOUGH_MEMORY ' OS Dependant
'
'    ' Overlapped operation aborted.
'    ' An overlapped operation was canceled due to the closure of the socket, or the execution of
'    ' the SIO_FLUSH command in WSAIoctl.
'    OverlappedOperationAborted = WSA_OPERATION_ABORTED ' OS Dependant
'
'    ' Invalid procedure table from service provider.
'    ' A service provider returned a bogus procedure table to Ws2_32.dll. (Usually caused by one
'    ' or more of the function pointers being null.)
'    InvalidProcedureTable = WSAEINVALIDPROCTABLE ' OS Dependant
'
'    ' Invalid service provider version number.
'    ' A service provider returned a version number other than 2.0.
'    InvalidServiceProviderVersion = WSAEINVALIDPROVIDER ' OS Dependant
'
'    ' Unable to initialize a service provider.
'    ' Either a service provider's DLL could not be loaded (LoadLibrary failed) or the provider's
'    ' WSPStartup/NSPStartup function failed.
'    UnableToInitializeServiceProvider = WSAEPROVIDERFAILEDINIT ' OS Dependant
'
'    ' System call failure.
'    ' Generic error code, returned under various conditions.
'    '
'    ' - Returned when a system call that should never fail does fail. For example, if a call to
'    '   WaitForMultipleEvents fails or one of the registry functions fails trying to manipulate
'    '   the protocol/namespace catalogs.
'    '
'    ' - Returned when a provider does not return SUCCESS and does not provide an extended error
'    '   code. Can indicate a service provider implementation error.
'    SystemCallFailure = WSASYSCALLFAILURE ' OS Dependant
'End Enum


' *************************************************************************************************
' Specifies the address families that can be used when creating a socket.
' *************************************************************************************************

Public Const AF_UNSPEC      As Long = 0      ' unspecified
Public Const AF_UNIX        As Long = 1      ' local to host (pipes, portals)
Public Const AF_INET        As Long = 2      ' internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK     As Long = 3      ' arpanet imp addresses
Public Const AF_PUP         As Long = 4      ' pup protocols: e.g. BSP
Public Const AF_CHAOS       As Long = 5      ' mit CHAOS protocols
Public Const AF_NS          As Long = 6      ' XEROX NS protocols
Public Const AF_IPX         As Long = AF_NS  ' IPX protocols: IPX, SPX, etc.
Public Const AF_ISO         As Long = 7      ' ISO protocols
Public Const AF_OSI         As Long = AF_ISO ' OSI is ISO
Public Const AF_ECMA        As Long = 8      ' european computer manufacturers
Public Const AF_DATAKIT     As Long = 9      ' datakit protocols
Public Const AF_CCITT       As Long = 10     ' CCITT protocols, X.25 etc
Public Const AF_SNA         As Long = 11     ' IBM SNA
Public Const AF_DECnet      As Long = 12     ' DECnet
Public Const AF_DLI         As Long = 13     ' Direct data link interface
Public Const AF_LAT         As Long = 14     ' LAT
Public Const AF_HYLINK      As Long = 15     ' NSC Hyperchannel
Public Const AF_APPLETALK   As Long = 16     ' AppleTalk
Public Const AF_NETBIOS     As Long = 17     ' NetBios-style addresses
Public Const AF_VOICEVIEW   As Long = 18     ' VoiceView
Public Const AF_FIREFOX     As Long = 19     ' Protocols from Firefox
Public Const AF_BAN         As Long = 21     ' Banyan
Public Const AF_ATM         As Long = 22     ' Native ATM Services
Public Const AF_INET6       As Long = 23     ' Internetwork Version 6
Public Const AF_CLUSTER     As Long = 24     ' Microsoft Wolfpack
Public Const AF_12844       As Long = 25     ' IEEE 1284.4 WG AF
Public Const AF_IRDA        As Long = 26     ' IrDA
Public Const AF_NETDES      As Long = 28     ' Network Designers OSI & gateway enabled protocols
Public Const AF_MAX         As Long = 29


'Public Enum AddressFamilyType
'    AddFam_Unknown = -1                      ' Unknown
'    AddFam_Unspecified = AF_UNSPEC           ' unspecified
'    AddFam_Unix = AF_UNIX                    ' local to host (pipes' portals)
'    AddFam_InterNetwork = AF_INET            ' internetwork: UDP' TCP' etc.
'    AddFam_ImpLink = AF_IMPLINK              ' arpanet imp addresses
'    AddFam_Pup = AF_PUP                      ' pup protocols: e.g. BSP
'    AddFam_Chaos = AF_CHAOS                  ' mit CHAOS protocols
'    AddFam_NS = AF_NS                        ' XEROX NS protocols
'    AddFam_Ipx = AF_NS                       ' IPX and SPX
'    AddFam_Iso = AF_ISO                      ' ISO protocols
'    AddFam_Osi = AF_ISO                      ' OSI is ISO
'    AddFam_Ecma = AF_ECMA                    ' european computer manufacturers
'    AddFam_DataKit = AF_DATAKIT              ' datakit protocols
'    AddFam_Ccitt = AF_CCITT                  ' CCITT protocols' X.25 etc
'    AddFam_Sna = AF_SNA                      ' IBM SNA
'    AddFam_DecNet = AF_DECnet                ' DECnet
'    AddFam_DataLink = AF_DLI                 ' Direct data link interface
'    AddFam_Lat = AF_LAT                      ' LAT
'    AddFam_HyperChannel = AF_HYLINK          ' NSC Hyperchannel
'    AddFam_AppleTalk = AF_APPLETALK          ' AppleTalk
'    AddFam_NetBios = AF_NETBIOS              ' NetBios-style addresses
'    AddFam_VoiceView = AF_VOICEVIEW          ' VoiceView
'    AddFam_FireFox = AF_FIREFOX              ' FireFox
'    AddFam_Banyan = AF_BAN                   ' Banyan
'    AddFam_ATM = AF_ATM                      ' Native ATM Services
'    AddFam_InterNetwork6 = AF_INET6          ' Internetwork Version 6
'    AddFam_Cluster = AF_CLUSTER              ' Microsoft Wolfpack
'    AddFam_Ieee12844 = AF_12844              ' IEEE 1284.4 WG AF
'    AddFam_Irda = AF_IRDA                    ' IrDA
'    AddFam_NetworkDesigners = AF_NETDES      ' Network Designers OSI & gateway enabled protocols
'    AddFam_Max = AF_MAX                      ' Max
'End Enum


' *************************************************************************************************
' Specifies the protocol families (same as the address families for now)
' *************************************************************************************************

Public Const PF_UNSPEC      As Long = AF_UNSPEC
Public Const PF_UNIX        As Long = AF_UNIX
Public Const PF_INET        As Long = AF_INET
Public Const PF_IMPLINK     As Long = AF_IMPLINK
Public Const PF_PUP         As Long = AF_PUP
Public Const PF_CHAOS       As Long = AF_CHAOS
Public Const PF_NS          As Long = AF_NS
Public Const PF_IPX         As Long = AF_IPX
Public Const PF_ISO         As Long = AF_ISO
Public Const PF_OSI         As Long = AF_OSI
Public Const PF_ECMA        As Long = AF_ECMA
Public Const PF_DATAKIT     As Long = AF_DATAKIT
Public Const PF_CCITT       As Long = AF_CCITT
Public Const PF_SNA         As Long = AF_SNA
Public Const PF_DECnet      As Long = AF_DECnet
Public Const PF_DLI         As Long = AF_DLI
Public Const PF_LAT         As Long = AF_LAT
Public Const PF_HYLINK      As Long = AF_HYLINK
Public Const PF_APPLETALK   As Long = AF_APPLETALK
Public Const PF_VOICEVIEW   As Long = AF_VOICEVIEW
Public Const PF_FIREFOX     As Long = AF_FIREFOX
Public Const PF_BAN         As Long = AF_BAN
Public Const PF_ATM         As Long = AF_ATM
Public Const PF_INET6       As Long = AF_INET6
Public Const PF_MAX         As Long = AF_MAX


Public Enum IPProtocolFamilyType
    ProtFam_Unknown = -1
    ProtFam_Unspecified = PF_UNSPEC
    ProtFam_Unix = PF_UNIX
    ProtFam_InterNetwork = PF_INET
    ProtFam_ImpLink = PF_IMPLINK
    ProtFam_Pup = PF_PUP
    ProtFam_Chaos = PF_CHAOS
    ProtFam_NS = PF_NS
    ProtFam_Ipx = PF_NS
    ProtFam_Iso = PF_ISO
    ProtFam_Osi = PF_ISO
    ProtFam_Ecma = PF_ECMA
    ProtFam_DataKit = PF_DATAKIT
    ProtFam_Ccitt = PF_CCITT
    ProtFam_Sna = PF_SNA
    ProtFam_DecNet = PF_DECnet
    ProtFam_DataLink = PF_DLI
    ProtFam_Lat = PF_LAT
    ProtFam_HyperChannel = PF_HYLINK
    ProtFam_AppleTalk = PF_APPLETALK
    ProtFam_VoiceView = PF_VOICEVIEW
    ProtFam_FireFox = PF_FIREFOX
    ProtFam_Banyan = PF_BAN
    ProtFam_ATM = PF_ATM
    ProtFam_Max = PF_MAX
End Enum


' *************************************************************************************************
' Specifies the type of a socket.
' *************************************************************************************************

Public Const SOCK_STREAM        As Long = 1  ' stream socket
Public Const SOCK_DGRAM         As Long = 2  ' datagram socket
Public Const SOCK_RAW           As Long = 3  ' raw-protocol interface
Public Const SOCK_RDM           As Long = 4  ' reliably-delivered message
Public Const SOCK_SEQPACKET     As Long = 5  ' sequenced packet stream


'Public Enum SocketType
'    SockType_stream = SOCK_STREAM            ' stream socket
'    SockType_Dgram = SOCK_DGRAM              ' datagram socket
'    SockType_Raw = SOCK_RAW                  ' raw-protocolinterface
'    SockType_Rdm = SOCK_RDM                  ' reliably-delivered message
'    SockType_Seqpacket = SOCK_SEQPACKET      ' sequenced packet stream
'    SockType_Unknown = -1                    ' unknown socket type
'End Enum



' *************************************************************************************************
' Specifies the protocol that a socket can use.
' *************************************************************************************************

Public Const IPPROTO_IP     As Long = 0      ' dummy for IP
Public Const IPPROTO_ICMP   As Long = 1      ' control message protocol
Public Const IPPROTO_IGMP   As Long = 2      ' internet group management protocol
Public Const IPPROTO_GGP    As Long = 3      ' gateway^2 (deprecated)
Public Const IPPROTO_TCP    As Long = 6      ' tcp
Public Const IPPROTO_PUP    As Long = 12     ' pup
Public Const IPPROTO_UDP    As Long = 17     ' user datagram protocol
Public Const IPPROTO_IDP    As Long = 22     ' xns idp
Public Const IPPROTO_ND     As Long = 77     ' UNOFFICIAL net disk proto
Public Const IPPROTO_IPX    As Long = 1000   ' xns idp
Public Const IPPROTO_SPX    As Long = 1256   ' UNOFFICIAL net disk proto
Public Const IPPROTO_SPXII  As Long = 1257   ' raw IP packet
Public Const IPPROTO_RAW    As Long = 255    ' raw IP packet
Public Const IPPROTO_MAX    As Long = 256


'Public Enum IPProtocolType
'    Proto_IP = IPPROTO_IP                    ' dummy for IP
'    Proto_Icmp = IPPROTO_ICMP                ' control message protocol
'    Proto_Igmp = IPPROTO_IGMP                ' group management protocol
'    Proto_Ggp = IPPROTO_GGP                  ' gateway^2 (deprecated)
'    Proto_Tcp = IPPROTO_TCP                  ' tcp
'    Proto_Pup = IPPROTO_PUP                  ' pup
'    Proto_Udp = IPPROTO_UDP                  ' user datagram protocol
'    Proto_Idp = IPPROTO_IDP                  ' xns idp
'    Proto_ND = IPPROTO_ND                    ' UNOFFICIAL net disk proto
'    Proto_Raw = IPPROTO_RAW                  ' raw IP packet
'    Proto_Unspecified = 0                    ' unspecified
'    Proto_Ipx = IPPROTO_IPX                  ' IPX protocol
'    Proto_Spx = IPPROTO_SPX                  ' SPX protocol
'    Proto_SpxII = IPPROTO_SPXII              ' SPXII protocol
'    Proto_Unknown = -1                       ' unknown protocol type
'    Proto_Max = IPPROTO_MAX
'End Enum


' *************************************************************************************************
' Specifies port/socket numbers: network standard functions
' *************************************************************************************************

Public Const IPPORT_ECHO            As Long = 7
Public Const IPPORT_DISCARD         As Long = 9
Public Const IPPORT_SYSTAT          As Long = 11
Public Const IPPORT_DAYTIME         As Long = 13
Public Const IPPORT_NETSTAT         As Long = 15
Public Const IPPORT_FTP             As Long = 21
Public Const IPPORT_TELNET          As Long = 23
Public Const IPPORT_SMTP            As Long = 25
Public Const IPPORT_TIMESERVER      As Long = 37
Public Const IPPORT_NAMESERVER      As Long = 42
Public Const IPPORT_WHOIS           As Long = 43
Public Const IPPORT_MTP             As Long = 57

' *************************************************************************************************
' Specifies port/socket numbers: host specific functions
' *************************************************************************************************

Public Const IPPORT_TFTP            As Long = 69
Public Const IPPORT_RJE             As Long = 77
Public Const IPPORT_FINGER          As Long = 79
Public Const IPPORT_TTYLINK         As Long = 87
Public Const IPPORT_SUPDUP          As Long = 95

' *************************************************************************************************
' Specifies UNIX TCP sockets
' *************************************************************************************************

Public Const IPPORT_EXECSERVER      As Long = 512
Public Const IPPORT_LOGINSERVER     As Long = 513
Public Const IPPORT_CMDSERVER       As Long = 514
Public Const IPPORT_EFSSERVER       As Long = 520

' *************************************************************************************************
' Specifies UNIX UDP sockets
' *************************************************************************************************

Public Const IPPORT_BIFFUDP         As Long = 512
Public Const IPPORT_WHOSERVER       As Long = 513
Public Const IPPORT_ROUTESERVER     As Long = 520

' *************************************************************************************************
' Ports < IPPORT_RESERVED are reserved for privileged processes (e.g. root).
' *************************************************************************************************

Public Const IPPORT_RESERVED        As Long = 1024


' *************************************************************************************************
' Common ports/sockets
' *************************************************************************************************

Public Enum IPPortType

    Port_Echo = IPPORT_ECHO
    Port_Discard = IPPORT_DISCARD
    Port_Systat = IPPORT_SYSTAT
    Port_DayTime = IPPORT_DAYTIME
    Port_NetState = IPPORT_NETSTAT
    Port_FileTransferProtocol = IPPORT_FTP
    Port_Telnet = IPPORT_TELNET
    Port_SimpleMailTransferProtocol = IPPORT_SMTP
    Port_TimerServer = IPPORT_TIMESERVER
    Port_NameServer = IPPORT_NAMESERVER
    Port_Whois = IPPORT_WHOIS
    Port_MTP = IPPORT_MTP
    Port_TFTP = IPPORT_TFTP
    Port_RJE = IPPORT_RJE
    Port_Finger = IPPORT_FINGER
    Port_TTYLink = IPPORT_TTYLINK
    Port_SUPDUP = IPPORT_SUPDUP
    Port_ExecServer = IPPORT_EXECSERVER
    Port_LoginServer = IPPORT_LOGINSERVER
    Port_CommandServer = IPPORT_CMDSERVER
    Port_EFSServer = IPPORT_EFSSERVER
    Port_BIFFUDP = IPPORT_BIFFUDP
    Port_WhoServer = IPPORT_WHOSERVER
    Port_EouteServer = IPPORT_ROUTESERVER
    Port_Reserved = IPPORT_RESERVED
End Enum



' *************************************************************************************************
' Protocol info service flags
' *************************************************************************************************
'
' Set of bit flags that specifies the services provided by the protocol. One or more of the
' following bit flags may be set.
'
' *************************************************************************************************

Public Const XP_GUARANTEED_DELIVERY     As Long = &H2
Public Const XP_CONNECTIONLESS          As Long = &H1
Public Const XP_GUARANTEED_ORDER        As Long = &H4
Public Const XP_MESSAGE_ORIENTED        As Long = &H8
Public Const XP_PSEUDO_STREAM           As Long = &H10
Public Const XP_GRACEFUL_CLOSE          As Long = &H20
Public Const XP_EXPEDITED_DATA          As Long = &H40
Public Const XP_CONNECT_DATA            As Long = &H80
Public Const XP_DISCONNECT_DATA         As Long = &H100
Public Const XP_SUPPORTS_BROADCAST      As Long = &H200
Public Const XP_SUPPORTS_MULTICAST      As Long = &H400
Public Const XP_BANDWIDTH_ALLOCATION    As Long = &H800
Public Const XP_FRAGMENTATION           As Long = &H1000
Public Const XP_ENCRYPTS                As Long = &H2000


Public Enum ProtocolServiceFlags

    ' If this flag is set, the protocol guarantees that all data sent will reach the intended
    ' destination. If this flag is clear, there is no such guarantee.
    ServFlag_GuaranteedDelivery = XP_GUARANTEED_DELIVERY
    
    ' If this flag is set, the protocol provides connectionless (datagram) service. If this
    ' flag is clear, the protocol provides connection-oriented data transfer.
    ServFlag_Connectionless = XP_CONNECTIONLESS
    
    ' If this flag is set, the protocol guarantees that data will arrive in the order in which it
    ' was sent. Note that this characteristic does not guarantee delivery of the data, only its
    ' order. If this flag is clear, the order of data sent is not guaranteed.
    ServFlag_GuaranteedOrder = XP_GUARANTEED_ORDER
    
    ' If this flag is set, the protocol is message-oriented. A message-oriented protocol honors
    ' message boundaries. If this flag is clear, the protocol is stream oriented, and the concept
    ' of message boundaries is irrelevant.
    ServFlag_MessageOriented = XP_MESSAGE_ORIENTED
    
    ' If this flag is set, the protocol is a message-oriented protocol that ignores message
    ' boundaries for all receive operations. This optional capability is useful when you do not
    ' want the protocol to frame messages. An application that requires stream-oriented
    ' characteristics can open a socket with type SOCK_STREAM for transport protocols that
    ' support this functionality, regardless of the value of iSocketType.
    ServFlag_PseudoStream = XP_PSEUDO_STREAM
    
    ' If this flag is set, the protocol supports two-phase close operations, also known as
    ' graceful close operations. If this flag is clear, the protocol supports only abortive
    ' close operations.
    ServFlag_GracefulClose = XP_GRACEFUL_CLOSE
    
    ' If this flag is set, the protocol supports expedited data, also known as urgent data.
    ServFlag_ExpeditedData = XP_EXPEDITED_DATA
    
    ' If this flag is set, the protocol supports connect data.
    ServFlag_ConnectData = XP_CONNECT_DATA
    
    ' If this flag is set, the protocol supports disconnect data.
    ServFlag_DisconnectData = XP_DISCONNECT_DATA
    
    ' If this flag is set, the protocol supports a broadcast mechanism.
    ServFlag_SupportsBroadcast = XP_SUPPORTS_BROADCAST
    
    ' If this flag is set, the protocol supports a multicast mechanism.
    ServFlag_SupportsMulticast = XP_SUPPORTS_MULTICAST
    
    ' If this flag is set, the protocol supports a mechanism for allocating a guaranteed bandwidth to an application.
    ServFlag_BandwidthAllocation = XP_BANDWIDTH_ALLOCATION
    
    ' If this flag is set, the protocol supports message fragmentation; physical network MTU is hidden from applications.
    ServFlag_Fragmentation = XP_FRAGMENTATION
    
    ' If this flag is set, the protocol supports data encryption.
    ServFlag_Encrypts = XP_ENCRYPTS
End Enum



' *************************************************************************************************
' Protocol info provider flags 1
' *************************************************************************************************


Public Const XP1_CONNECTIONLESS         As Long = &H1
Public Const XP1_GUARANTEED_DELIVERY    As Long = &H2
Public Const XP1_GUARANTEED_ORDER       As Long = &H4
Public Const XP1_MESSAGE_ORIENTED       As Long = &H8
Public Const XP1_PSEUDO_STREAM          As Long = &H10
Public Const XP1_GRACEFUL_CLOSE         As Long = &H20
Public Const XP1_EXPEDITED_DATA         As Long = &H40
Public Const XP1_CONNECT_DATA           As Long = &H80
Public Const XP1_DISCONNECT_DATA        As Long = &H100
Public Const XP1_INTERRUPT              As Long = &H4000
Public Const XP1_SUPPORT_BROADCAST      As Long = &H200
Public Const XP1_SUPPORT_MULTIPOINT     As Long = &H400
Public Const XP1_MULTIPOINT_CONTROL_PLANE As Long = &H800
Public Const XP1_MULTIPOINT_DATA_PLANE  As Long = &H1000
Public Const XP1_QOS_SUPPORTED          As Long = &H2000
Public Const XP1_UNI_SEND               As Long = &H8000
Public Const XP1_UNI_RECV               As Long = &H10000
Public Const XP1_IFS_HANDLES            As Long = &H20000
Public Const XP1_PARTIAL_MESSAGE        As Long = &H40000


Public Enum ProtocolServiceFlags1

    ' Provides connectionless (datagram) service. If not set, the protocol supports
    ' connection-oriented data transfer.
    ServFlag1_Connectionless = XP1_CONNECTIONLESS
    
    ' Guarantees that all data sent will reach the intended destination.
    ServFlag1_GuaranteedDelivery = XP1_GUARANTEED_DELIVERY
    
    ' Guarantees that data only arrives in the order in which it was sent and that it is not
    ' duplicated. This characteristic does not necessarily mean that the data is always
    ' delivered, but that any data that is delivered is delivered in the order in which it was
    ' sent.
    ServFlag1_GuaranteedOrder = XP1_GUARANTEED_ORDER
    
    ' Honors message boundaries as opposed to a stream-oriented protocol where there is no
    ' concept of message boundaries.
    ServFlag1_MessageOriented = XP1_MESSAGE_ORIENTED
    
    ' A message-oriented protocol, but message boundaries are ignored for all receipts. This is
    ' convenient when an application does not desire message framing to be done by the protocol.
    ServFlag1_PseudoStream = XP1_PSEUDO_STREAM
    
    ' Supports two-phase (graceful) close. If not set, only abortive closes are performed.
    ServFlag1_GracefulClose = XP1_GRACEFUL_CLOSE
    
    ' Supports expedited (urgent) data.
    ServFlag1_ExpeditedData = XP1_EXPEDITED_DATA
    
    ' Supports connect data.
    ServFlag1_ConnectData = XP1_CONNECT_DATA
    
    ' Supports disconnect data.
    ServFlag1_DisconnectData = XP1_DISCONNECT_DATA
    
    ' Bit is reserved.
    ServFlag1_Interrupt = XP1_INTERRUPT
    
    ' Supports a broadcast mechanism.
    ServFlag1_SupportBroadcast = XP1_SUPPORT_BROADCAST
    
    ' Supports a multipoint or multicast mechanism. Control and data
    ' plane attributes are indicated below.
    ServFlag1_SupportMultipoint = XP1_SUPPORT_MULTIPOINT
    
    ' Indicates whether the control plane is rooted (value = 1) or nonrooted (value = 0).
    ServFlag1_MultipointControlPlane = XP1_MULTIPOINT_CONTROL_PLANE
    
    ' Indicates whether the data plane is rooted (value = 1) or nonrooted (value = 0).
    ServFlag1_MultipointDataPlane = XP1_MULTIPOINT_DATA_PLANE
    
    ' Supports quality of service requests.
    ServFlag1_QOSSupported = XP1_QOS_SUPPORTED
    
    ' Protocol is unidirectional in the send direction.
    ServFlag1_UniSend = XP1_UNI_SEND
    
    ' Protocol is unidirectional in the recv direction.
    ServFlag1_UniRecv = XP1_UNI_RECV
    
    ' Socket descriptors returned by the provider are operating system Installable File System (IFS) handles.
    ServFlag1_IFSHandles = XP1_IFS_HANDLES
    
    ' The MSG_PARTIAL flag is supported in WSASend and WSASendTo.
    ServFlag1_PartialMessage = XP1_PARTIAL_MESSAGE
End Enum


' *************************************************************************************************
' Protocol info provider flags
' *************************************************************************************************

Public Const PFL_MULTIPLE_PROTO_ENTRIES     As Long = &H1
Public Const PFL_RECOMMENDED_PROTO_ENTRY    As Long = &H2
Public Const PFL_HIDDEN                     As Long = &H4
Public Const PFL_MATCHES_PROTOCOL_ZERO      As Long = &H8


Public Enum ProtocolInfoProviderFlags

    ' Indicates that this is one of two or more entries for a single protocol (from a given
    ' provider) which is capable of implementing multiple behaviors. An example of this is
    ' SPX which, on the receiving side, can behave either as a message-oriented or a
    ' stream-oriented protocol.
    PFlags_MultipleProtocolEntries
    
    ' Indicates that this is the recommended or most frequently used entry for a protocol that is
    ' capable of implementing multiple behaviors.
    PFlags_RecommendedProtocolEntry
    
    ' Set by a provider to indicate to the Ws2_32.dll that this protocol should not be returned
    ' in the result buffer generated by WSAEnumProtocols. Obviously, a Windows Sockets 2
    ' application should never see an entry with this bit set.
    PFlags_Hidden
    
    ' Indicates that a value of zero in the protocol parameter of socket or WSASocket matches
    ' this protocol entry.
    PFlags_MatchesProtocolZero
End Enum


' *************************************************************************************************
' Name spaces
' *************************************************************************************************

Public Const NS_DEFAULT     As Long = 0
Public Const NS_DNS         As Long = 12
Public Const NS_NDS         As Long = 2
Public Const NS_NETBT       As Long = 13
Public Const NS_SAP         As Long = 1
Public Const NS_TCPIP_HOSTS As Long = 11
Public Const NS_TCPIP_LOCAL As Long = 10
Public Const NS_PEER_BROWSE As Long = 3
Public Const NS_WINS        As Long = 14
Public Const NS_NBP         As Long = 20
Public Const NS_MS          As Long = 30
Public Const NS_STDA        As Long = 31
Public Const NS_NTDS        As Long = 32
Public Const NS_X500        As Long = 40
Public Const NS_NIS         As Long = 41
Public Const NS_VNS         As Long = 50


Public Enum NamespaceType

    ' A set of default name spaces. The function queries each name space within this set. The set
    ' of default name spaces typically includes all the name spaces installed on the system.
    ' System administrators, however, can exclude particular name spaces from the set. This is
    ' the value that most applications should use for dwNameSpace.
    Namespace_Default = NS_DEFAULT

    ' The Domain Name System used in the Internet for host name resolution.
    Namespace_DNS = NS_DNS

    ' The NetWare 4 provider.
    Namespace_NDS = NS_NDS
    
    ' The NetBIOS over TCP/IP layer. All Windows NT/Windows 2000 systems register their computer
    ' names with NetBIOS. This name space is used to convert a computer name to an IP address
    ' that uses this registration. Note that NS_NETBT can access a WINS server to perform the
    ' resolution.
    Namespace_NetBios = NS_NETBT
    
    ' The Netware Service Advertising Protocol. This can access the Netware bindery if
    ' appropriate. NS_SAP is a dynamic name space that allows registration of services.
    Namespace_SAP = NS_SAP

    ' Lookup value in the <systemroot>\system32\drivers\etc\hosts file.
    Namespace_TcpIpHosts = NS_TCPIP_HOSTS

    ' Local TCP/IP name resolution mechanisms, including comparisons against the local host name
    ' and looks up host names and IP addresses in cache of host to IP address mappings.
    Namespace_TcpIpLocal = NS_TCPIP_LOCAL

    Namespace_PeerBworse = NS_PEER_BROWSE
    Namespace_Wins = NS_WINS
    Namespace_NBP = NS_NBP
    Namespace_MS = NS_MS
    Namespace_STA = NS_STDA
    Namespace_NTDS = NS_NTDS
    Namespace_X500 = NS_X500
    Namespace_NIS = NS_NIS
    Namespace_VNS = NS_VNS
End Enum




' *************************************************************************************************
' Domain name resolution
' *************************************************************************************************

Public Const RES_SOFT_SEARCH    As Long = &H1
Public Const RES_FIND_MULTIPLE  As Long = &H2
Public Const RES_SERVICE        As Long = &H4


Public Enum ResolutionType

    ' This flag is valid if the name space supports multiple levels of searching.  If this flag
    ' is valid and set, the operating system performs a simple and quick search of the name space.
    ' This is useful if an application only needs to obtain easy-to-find addresses for the service.
    ' If this flag is valid and clear, the operating system performs a more extensive search of
    ' the name space.
    Res_SoftSearch = RES_SOFT_SEARCH
    
    ' If this flag is set, the operating system performs an extensive search of all name spaces
    ' for the service. It asks every appropriate name space to resolve the service name. If this
    ' flag is clear, the operating system stops looking for service addresses as soon as one is
    ' found.
    Res_FindMultiple = RES_FIND_MULTIPLE
    
    ' If set, the function obtains the address to which a service of the specified type should
    ' bind. This is the equivalent of setting lpServiceName to NULL.  If this flag is clear,
    ' normal name resolution occurs.
    Res_ServiceType = RES_SERVICE
End Enum



' *************************************************************************************************
' Service properties
' *************************************************************************************************
'
' Set of bit flags that specify the service information that the function retrieves. Each of
' these bit flag constants, other than PROP_ALL, corresponds to a particular member of the
' SERVICE_INFO data structure. If the flag is set, the function puts information into the
' corresponding member of the data structures stored in *lpBuffer. The following bit flags are
' defined.
'
' *************************************************************************************************

Public Const PROP_COMMENT       As Long = &H1
Public Const PROP_LOCALE        As Long = &H2
Public Const PROP_DISPLAY_HINT  As Long = &H4
Public Const PROP_VERSION       As Long = &H8
Public Const PROP_START_TIME    As Long = &H10
Public Const PROP_MACHINE       As Long = &H20
Public Const PROP_ADDRESSES     As Long = &H100
Public Const PROP_SD            As Long = &H200
Public Const PROP_ALL           As Long = &H80000000


Public Enum ServicePropertyFlagsType
    
    ' If this flag is set, the function stores data in the lpComment member of the data
    ' structures stored in *lpBuffer.
    ServProp_Comment = PROP_COMMENT
    
    ' If this flag is set, the function stores data in the lpLocale member of the data structures
    ' stored in *lpBuffer.
    ServProp_Locale = PROP_LOCALE
    
    ' If this flag is set, the function stores data in the dwDisplayHint member of the data
    ' structures stored in *lpBuffer.
    ServProp_DisplayHint = PROP_DISPLAY_HINT
    
    ' If this flag is set, the function stores data in the dwVersion member of the data
    ' structures stored in *lpBuffer.
    ServProp_Version = PROP_VERSION
    
    ' If this flag is set, the function stores data in the dwTime member of the data structures
    ' stored in *lpBuffer.
    ServProp_StartTime = PROP_START_TIME
    
    ' If this flag is set, the function stores data in the lpMachineName member of the data
    ' structures stored in *lpBuffer.
    ServProp_Machine = PROP_MACHINE
    
    ' If this flag is set, the function stores data in the lpServiceAddress member of the data
    ' structures stored in *lpBuffer.
    ServProp_Addresses = PROP_ADDRESSES
    
    ' If this flag is set, the function stores data in the ServiceSpecificInfo member of the data
    ' structures stored in *lpBuffer.
    ServProp_SD = PROP_SD
    
    ' If this flag is set, the function stores data in all of the members of the data structures
    ' stored in *lpBuffer.
    ServProp_All = PROP_ALL
End Enum



' *************************************************************************************************
' Service address flags
' *************************************************************************************************


Public Const SERVICE_ADDRESS_FLAG_RPC_CN As Long = &H1
Public Const SERVICE_ADDRESS_FLAG_RPC_DG As Long = &H2
Public Const SERVICE_ADDRESS_FLAG_RPC_NB As Long = &H4


Public Enum ServiceAddressFlagsType

    ' If this bit flag is set, the service supports connection-oriented RPC over this transport
    ' protocol.
    ServAddr_RpcConnectionOriented = SERVICE_ADDRESS_FLAG_RPC_CN
    
    ' If this bit flag is set, the service supports datagram-oriented RPC over this transport
    ' protocol.
    ServAddr_RpcDatagramOriented = SERVICE_ADDRESS_FLAG_RPC_DG
    
    ' If this bit flag is set, the service supports NetBIOS RPC over this transport protocol.
    ServAddr_RpcNetviosOriented = SERVICE_ADDRESS_FLAG_RPC_NB
End Enum


' *************************************************************************************************
' Socket options and option levels
' *************************************************************************************************

' Socket option levels
Public Const SOL_SOCKET         As Long = &HFFFF&
Public Const SOL_IP             As Long = IPPROTO_IP
Public Const SOL_TCP            As Long = IPPROTO_TCP
Public Const SOL_UDP            As Long = IPPROTO_UDP
Public Const SOL_IPX            As Long = IPPROTO_IPX


' Socket level socket options
Public Const SO_ACCEPTCONN      As Long = &H2
Public Const SO_BROADCAST       As Long = &H20
Public Const SO_DEBUG           As Long = &H1
Public Const SO_DONTROUTE       As Long = &H10
Public Const SO_ERROR           As Long = &H1007
Public Const SO_GROUP_ID        As Long = &H2001
Public Const SO_GROUP_PRIORITY  As Long = &H2002
Public Const SO_KEEPALIVE       As Long = &H8
Public Const SO_LINGER          As Long = &H80
Public Const SO_DONTLINGER      As Long = Not SO_LINGER
Public Const SO_MAX_MSG_SIZE    As Long = &H2003
Public Const SO_OOBINLINE       As Long = &H100
Public Const SO_PROTOCOL_INFO   As Long = &H2004
Public Const SO_RCVBUF          As Long = &H1002
Public Const SO_RCVLOWAT        As Long = &H1004
Public Const SO_RCVTIMEO        As Long = &H1006
Public Const SO_REUSEADDR       As Long = &H4
Public Const SO_EXCLUSIVEADDRUSE As Long = Not SO_REUSEADDR
Public Const SO_SNDBUF          As Long = &H1001
Public Const SO_SNDLOWAT        As Long = &H1003
Public Const SO_SNDTIMEO        As Long = &H1005
Public Const SO_TYPE            As Long = &H1008

' Extended socket level socket options
Public Const SO_CONNDATA        As Long = &H7000
Public Const SO_CONNOPT         As Long = &H7001
Public Const SO_DISCDATA        As Long = &H7002
Public Const SO_DISCOPT         As Long = &H7003
Public Const SO_CONNDATALEN     As Long = &H7004
Public Const SO_CONNOPTLEN      As Long = &H7005
Public Const SO_DISCDATALEN     As Long = &H7006
Public Const SO_DISCOPTLEN      As Long = &H7007

' IP level socket options
Public Const IP_OPTIONS         As Long = 1
Public Const IP_HDRINCL         As Long = 2
Public Const IP_TOS             As Long = 3
Public Const IP_TTL             As Long = 4
Public Const IP_MULTICAST_IF    As Long = 9
Public Const IP_MULTICAST_TTL   As Long = 10
Public Const IP_MULTICAST_LOOP  As Long = 11
Public Const IP_ADD_MEMBERSHIP  As Long = 12
Public Const IP_DROP_MEMBERSHIP As Long = 13
Public Const IP_DONTFRAGMENT    As Long = 14

' TCP level socket options
Public Const TCP_NODELAY        As Long = &H1
Public Const TCP_BSDURGENT      As Long = &H7000

' UDP level socket options
Public Const UDP_NOCHECKSUM     As Long = 1

' IPX level socket options
Public Const IPX_PTYPE              As Long = &H4000
Public Const IPX_FILTERPTYPE        As Long = &H4001
Public Const IPX_DSTYPE             As Long = &H4002
Public Const IPX_EXTENDED_ADDRESS   As Long = &H4004
Public Const IPX_RECVHDR            As Long = &H4005
Public Const IPX_MAXSIZE            As Long = &H4006
Public Const IPX_ADDRESS            As Long = &H4007
Public Const IPX_GETNETINFO         As Long = &H4008
Public Const IPX_GETNETINFO_NORIP   As Long = &H4009
Public Const IPX_SPXGETCONNECTIONSTATUS As Long = &H400B
Public Const IPX_ADDRESS_NOTIFY     As Long = &H400C
Public Const IPX_MAX_ADAPTER_NUM    As Long = &H400D
Public Const IPX_RERIPNETNUMBER     As Long = &H400E
Public Const IPX_IMMEDIATESPXACK    As Long = &H4010
Public Const IPX_STOPFILTERPTYPE    As Long = &H4003
Public Const IPX_RECEIVE_BROADCAST  As Long = &H400F


Public Enum SocketOptionLevel
    Optlev_Socket = SOL_SOCKET         ' Indicates socket options apply to the socket itself.
    Optlev_IP = IPPROTO_IP             ' Indicates socket options apply to IP sockets.
    Optlev_Tcp = IPPROTO_TCP           ' Indicates socket options apply to Tcp sockets.
    Optlev_Udp = IPPROTO_UDP           ' Indicates socket options apply to Udp sockets.
    Optlev_Ipx                         ' Indicates socket options apply to Ipx sockets.
End Enum


Public Enum SocketOptionName
    
    Sock_Debugging = SO_DEBUG                 ' turn on debugging info recording
    Sock_AcceptConnection = SO_ACCEPTCONN     ' socket has had listen()
    Sock_ReuseAddress = SO_REUSEADDR          ' allow local address reuse
    Sock_KeepAlive = SO_KEEPALIVE             ' keep connections alive
    Sock_DontRoute = SO_DONTROUTE             ' just use interface addresses
    Sock_Broadcast = SO_BROADCAST             ' permit sending of broadcast msgs
    Sock_UseLoopback = &H40                   ' bypass hardware when possible
    Sock_Linger = SO_LINGER                   ' linger on close if data present
    Sock_OutOfBandInline = SO_OOBINLINE       ' leave received OOB data in line
    Sock_DontLinger = SO_DONTLINGER           ' dont linger
    Sock_ExclusiveAddressUse = SO_EXCLUSIVEADDRUSE  ' disallow local address reuse
    Sock_SendBuffer = SO_SNDBUF               ' send buffer size
    Sock_ReceiveBuffer = SO_RCVBUF            ' receive buffer size
    Sock_SendLowWater = SO_SNDLOWAT           ' send low-water mark
    Sock_ReceiveLowWater = SO_RCVLOWAT        ' receive low-water mark
    Sock_SendTimeout = SO_SNDTIMEO            ' send timeout
    Sock_ReceiveTimeout = SO_RCVTIMEO         ' receive timeout
    Sock_Error = SO_ERROR                     ' get error status and clear
    Sock_TypeOfSocket = SO_TYPE               ' get socket type
    Sock_MaxConnections = &H7FFFFFFF          ' Maximum q length specifiable by listen.
    Sock_ConnectionData = SO_CONNDATA
    Sock_ConnectionOptions = SO_CONNOPT
    Sock_DisconnectData = SO_DISCDATA
    Sock_DisconnectOptions = SO_DISCOPT
    Sock_ConnectionDataLength = SO_CONNDATALEN
    Sock_ConnectionOptionLength = SO_CONNOPTLEN
    Sock_DisconnectDataLength = SO_DISCDATALEN
    Sock_DisconnectOptionLength = SO_DISCOPTLEN

    IP_IPOptions = IP_OPTIONS
    IP_HeaderIncluded = IP_HDRINCL            ' Header is included with data.
    IP_TypeOfService = IP_TOS                 ' IP type of service and preced.
    IP_IpTimeToLive = IP_TTL                  ' IP time to live.
    IP_MulticastInterface = IP_MULTICAST_IF   ' IP multicast interface.
    IP_MulticastTimeToLive = IP_MULTICAST_TTL ' IP multicast time to live.
    IP_MulticastLoopback = IP_MULTICAST_LOOP  ' IP Multicast loopback.
    IP_AddMembership = IP_ADD_MEMBERSHIP      ' Add an IP group membership.
    IP_DropMembership = IP_DROP_MEMBERSHIP    ' Drop an IP group membership.
    IP_DontFragDatagrams = IP_DONTFRAGMENT    ' Don't fragment IP datagrams.
    IP_AddSourceMembership = 15               ' Join IP group/source.
    IP_DropSourceMembership = 16              ' Leave IP group/source.
    IP_BlockSource = 17                       ' Block IP group/source.
    IP_UnblockSource = 18                     ' Unblock IP group/source.
    IP_PacketInformation = 19                 ' RIP_DONTFRAGMENTeceive packet information for ipv4.

    Tcp_WithoutDelay = TCP_NODELAY            ' Disables the Nagle algorithm for send coalescing.
    Tcp_Urgent = TCP_BSDURGENT
    Tcp_Expedited = TCP_BSDURGENT

    Udp_DontCalculateChecksum = UDP_NOCHECKSUM
    
    Ipx_PacketType = IPX_PTYPE
    Ipx_FilterPacketType = IPX_FILTERPTYPE
    Ipx_DataSreeamType = IPX_DSTYPE
    Ipx_ExtendedAddress = IPX_EXTENDED_ADDRESS
    Ipx_RecieveHeader = IPX_RECVHDR
    Ipx_MaximumSize = IPX_MAXSIZE
    Ipx_AddressInfo = IPX_ADDRESS
    Ipx_ObtainNetInfo = IPX_GETNETINFO
    Ipx_ObtainNetInfo_NoRIP = IPX_GETNETINFO_NORIP
    Ipx_SPXObtainConnectionStatus = IPX_SPXGETCONNECTIONSTATUS
    Ipx_ObtainAddressNotification = IPX_ADDRESS_NOTIFY
    Ipx_MaxAdapterNumber = IPX_MAX_ADAPTER_NUM
    Ipx_ObtainNetInfo_RIP = IPX_RERIPNETNUMBER
    Ipx_SendAckImmediately = IPX_IMMEDIATESPXACK
    Ipx_StopFiltering = IPX_STOPFILTERPTYPE
    Ipx_RecieveBroadcast = IPX_RECEIVE_BROADCAST
End Enum

' *************************************************************************************************
' Addresses
' *************************************************************************************************

' Any IP Address (0.0.0.0)
Public Const INADDR_ANY         As Long = &H0

' The loopback address, 127.0.0.1
Public Const INADDR_LOOPBACK    As Long = 16777343

' The broadcast address 255.255.255.255
Public Const INADDR_BROADCAST   As Long = &HFFFFFFFF

' No address is defined as 255.255.255.255
Public Const INADDR_NONE        As Long = &HFFFFFFFF


Public Enum CommonAddressTypes
    AddrType_Any = INADDR_ANY
    AddrType_Loopback = INADDR_LOOPBACK
    AddrType_Broadcast = INADDR_BROADCAST
    AddrType_None = INADDR_NONE
End Enum


' *************************************************************************************************
' Message communication flags
' *************************************************************************************************

Public Const MSG_OOB        As Long = &H1       ' Process out-of-band data
Public Const MSG_PEEK       As Long = &H2       ' Peek at incoming message
Public Const MSG_DONTROUTE  As Long = &H4       ' Send without using routing tables
Public Const MSG_PARTIAL    As Long = &H8000    ' Partial send or recv for message xport


'Public Enum MessageCommunicationFlags
'    Msg_None = 0
'    Msg_OutOfBand = MSG_OOB
'    Msg_PeekData = MSG_PEEK
'    Msg_NoRoutingTables = MSG_DONTROUTE
'    Msg_PartialSendRecv = MSG_PARTIAL
'End Enum


' *************************************************************************************************
' Service Operation
' *************************************************************************************************

Public Const SERVICE_REGISTER       As Long = &H1
Public Const SERVICE_DEREGISTER     As Long = &H2
Public Const SERVICE_FLUSH          As Long = &H3
Public Const SERVICE_ADD_TYPE       As Long = &H4
Public Const SERVICE_DELETE_TYPE    As Long = &H5


Public Enum ServiceOperationType

    ' Register the network service with the name space. This operation can be used with the
    ' SERVICE_FLAG_DEFER and SERVICE_FLAG_HARD bit flags.
    ServOp_RegisterService = SERVICE_REGISTER
    
    ' Remove from the registry the network service from the name space. This operation can be used
    ' with the SERVICE_FLAG_DEFER and SERVICE_FLAG_HARD bit flags.
    ServOp_UnregisterService = SERVICE_DEREGISTER
    
    ' Perform any operation that was called with the SERVICE_FLAG_DEFER bit flag set to one.
    ServOp_Flush = SERVICE_FLUSH
    
    ' Add a service type to the name space.  For this operation, use the ServiceSpecificInfo
    ' member of the SERVICE_INFO structure pointed to by lpServiceInfo to pass a
    ' SERVICE_TYPE_INFO_ABS structure. You must also set the ServiceType member of the
    ' SERVICE_INFO structure. Other SERVICE_INFO members are ignored.
    ServOp_AddType = SERVICE_ADD_TYPE
    
    ' Remove a service type, added by a previous call specifying the SERVICE_ADD_TYPE operation,
    ' from the name space.
    ServOp_DeleteType = SERVICE_DELETE_TYPE
End Enum


' *************************************************************************************************
' Service Flags
' *************************************************************************************************

Public Const SERVICE_FLAG_DEFER As Long = &H1
Public Const SERVICE_FLAG_HARD  As Long = &H2


Public Enum ServiceFlagsType
    ' This bit flag is valid only if the operation is SERVICE_REGISTER or SERVICE_DEREGISTER.
    ' If this bit flag is one, and it is valid, the name-space provider should defer the
    ' registration or deregistration operation until a SERVICE_FLUSH operation is requested.
    ServFlg_Defer = SERVICE_FLAG_DEFER
    
    ' This bit flag is valid only if the operation is SERVICE_REGISTER or SERVICE_DEREGISTER.  If
    ' this bit flag is one, and it is valid, the name-space provider updates any relevant
    ' persistent store information when the operation is performed.  For example: If the
    ' operation involves deregistration in a name space that uses a persistent store, the
    ' name-space provider would remove the relevant persistent store information.
    ServFlg_Hard = SERVICE_FLAG_HARD
End Enum

' *************************************************************************************************
' Service Status Flags
' *************************************************************************************************

Public Const SET_SERVICE_PARTIAL_SUCCESS As Long = &H1


Public Enum ServiceStatusFlags

    ' One or more name-space providers were unable to successfully perform the requested operation.
    ServStatus_PartialSuccess = SET_SERVICE_PARTIAL_SUCCESS
End Enum



' *************************************************************************************************
' Socket Shutdown types
' *************************************************************************************************

Public Const SD_RECEIVE As Long = &H0
Public Const SD_SEND    As Long = &H1
Public Const SD_BOTH    As Long = &H2


'Public Enum SocketShutdownType
'
'    ' If the how parameter is SD_RECEIVE, subsequent calls to the recv function on the socket
'    ' will be disallowed. This has no effect on the lower protocol layers. For TCP sockets,
'    ' if there is still data queued on the socket waiting to be received, or data arrives
'    ' subsequently, the connection is reset, since the data cannot be delivered to the user. For
'    ' UDP sockets, incoming datagrams are accepted and queued. In no case will an ICMP error
'    ' packet be generated.
'    SDType_Recieve = SD_RECEIVE
'
'    ' If the how parameter is SD_SEND, subsequent calls to the send function are disallowed. For
'    ' TCP sockets, a FIN will be sent after all data is sent and acknowledged by the receiver.
'    SDType_Send = SD_SEND
'
'    ' Setting how to SD_BOTH disables both sends and receives as described above.
'    SDType_Both = SD_BOTH
'End Enum


' *************************************************************************************************
' Transmit File Flags
' *************************************************************************************************

Public Const TF_DISCONNECT          As Long = &H1
Public Const TF_REUSE_SOCKET        As Long = &H2
Public Const TF_WRITE_BEHIND        As Long = &H4
Public Const TF_USE_DEFAULT_WORKER  As Long = &H0
Public Const TF_USE_SYSTEM_THREAD   As Long = &H10
Public Const TF_USE_KERNEL_APC      As Long = &H20


Public Enum TransmifFileFlagsType

    ' Start a transport-level disconnect after all the file data has been queued for transmission.
    tff_Disconnect = TF_DISCONNECT
    
    ' Prepare the socket handle to be reused. When the TransmitFile request completes, the socket
    ' handle can be passed to the AcceptEx function. It is only valid if TF_DISCONNECT is also
    ' specified.
    tff_ResuseSocket = TF_REUSE_SOCKET
    
    ' Complete the TransmitFile request immediately, without pending. If this flag is specified
    ' and TransmitFile succeeds, then the data has been accepted by the system but not necessarily
    ' acknowledged by the remote end. Do not use this setting with the TF_DISCONNECT and
    ' TF_REUSE_SOCKET flags.
    tff_WriteBehind = TF_WRITE_BEHIND
    
    ' Directs the Windows Sockets service provider to use the system's default thread to process
    ' long TransmitFile requests. The system default thread can be adjusted using the following
    ' registry parameter as a REG_DWORD:CurrentControlSet\Services\afd\Parameters\TransmitWorker
    tff_UseDefaultWorker = TF_USE_DEFAULT_WORKER
    
    ' Directs the Windows Sockets service provider to use system threads to process long
    ' TransmitFile requests.
    tff_UseSystemThread = TF_USE_SYSTEM_THREAD
    
    ' Directs the driver to use kernel Asynchronous Procedure Calls (APCs) instead of worker
    ' threads to process long TransmitFile requests. Long TransmitFile requests are defined as
    ' requests that require more than a single read from the file or a cache; the request therefore depends on the size of the file and the specified length of the send packet.  Use of TF_USE_KERNEL_APC can deliver significant performance benefits. It is possible (though unlikely), however, that the thread in which context TransmitFile is initiated is being used for heavy computations; this situation may prevent APCs from launching. Note that the Windows Sockets kernel mode driver uses normal kernel APCs, which launch whenever a thread is in a wait state, which differs from user-mode APCs, which launch whenever a thread is in an alertable wait state initiated in user mode).
    tff_UseKernelAPC = TF_USE_KERNEL_APC
End Enum



' *************************************************************************************************
' Async Select flags
' *************************************************************************************************

Public Const FD_READ        As Long = &H1
Public Const FD_WRITE       As Long = &H2
Public Const FD_OOB         As Long = &H4
Public Const FD_ACCEPT      As Long = &H8
Public Const FD_CONNECT     As Long = &H10
Public Const FD_CLOSE       As Long = &H20
Public Const FD_QOS         As Long = &H40
Public Const FD_GROUP_QOS   As Long = &H80
Public Const FD_ROUTING_INTERFACE_CHANGE As Long = &H100
Public Const FD_ADDRESS_LIST_CHANGE As Long = &H200
Public Const FD_MAX_EVENTS  As Long = 8

Public Enum AsyncSelectFlagsType

    ' INPUT  = Wants to receive notification of readiness for reading.
    ' OUTPUT = Socket ready for reading.
    ASFlag_Read = FD_READ
    
    ' INPUT  = Wants to receive notification of readiness for writing.
    ' OUTPUT = Socket ready for writing.
    ASFlag_Write = FD_WRITE
    
    ' INPUT  = Wants to receive notification of the arrival of OOB data.
    ' OUTPUT = OOB data ready for reading on socket.
    ASFlag_OutOfBand = FD_OOB
    
    ' INPUT  = Wants to receive notification of incoming connections.
    ' OUTPUT = Socket ready for accepting a new incoming connection.
    ASFlag_Accept = FD_ACCEPT
    
    ' INPUT  = Wants to receive notification of completed connection or multipoint join operation.
    ' OUTPUT = Connection or multipoint join operation initiated on socket completed.
    ASFlag_Connect = FD_CONNECT
    
    ' INPUT  = Wants to receive notification of socket closure.
    ' OUTPUT = Connection identified by socket has been closed.
    ASFlag_Close = FD_CLOSE
    
    ' INPUT  = Wants to receive notification of socket Quality of Service (QOS) changes.
    ' OUTPUT = Quality of Service associated with socket has changed.
    ASFlag_QualityOfService = FD_QOS
    
    ' INPUT  = Reserved
    ' OUTPUT = Reserved
    'ASFlag_GroupQualityOfService = FD_GROUP_QOS
    
    ' INPUT  = Wants to receive notification of routing interface changes for the specified
    ' destination(s).
    ' OUTPUT = Local interface that should be used to send to the specified destination has
    ' changed.
    ASFlag_RoutingInterfaceChange = FD_ROUTING_INTERFACE_CHANGE
    
    ' INPUT  = Wants to receive notification of local address list changes for the socket's
    ' protocol family.
    ' OUTPUT = The list of addresses of the socket's protocol family to which the application
    ' client can bind has changed.
    ASFlag_AddressListChange = FD_ADDRESS_LIST_CHANGE
    
End Enum



' *************************************************************************************************
' Lookup service control flags
' *************************************************************************************************

Public Const LUP_DEEP           As Long = &H1
Public Const LUP_CONTAINERS     As Long = &H2
Public Const LUP_NOCONTAINERS   As Long = &H4
Public Const LUP_FLUSHCACHE     As Long = &H1000
Public Const LUP_FLUSHPREVIOUS  As Long = &H2000
Public Const LUP_NEAREST        As Long = &H8
Public Const LUP_RES_SERVICE    As Long = &H8000
Public Const LUP_RETURN_ALIASES As Long = &H400
Public Const LUP_RETURN_NAME    As Long = &H10
Public Const LUP_RETURN_TYPE    As Long = &H20
Public Const LUP_RETURN_VERSION As Long = &H40
Public Const LUP_RETURN_COMMENT As Long = &H80
Public Const LUP_RETURN_ADDR    As Long = &H100
Public Const LUP_RETURN_BLOB    As Long = &H200
Public Const LUP_RETURN_ALL     As Long = &HFF0


Public Enum LookupServiceControlFlagsType

        ' Queries deep as opposed to just the first level.
        Loopuk_Deep
        
        ' Returns containers only.
        Lookup_Containers
        
        ' Does not return any containers.
        Lookup_NoContainers
        
        ' If the provider has been caching information, ignores the cache, and queries the name
        ' space itself.
        Lookup_FlushCache
        
        ' Used as a value for the dwControlFlags parameter in WSALookupServiceNext. Setting this
        ' flag instructs the provider to discard the last result set, which was too large for the
        ' supplied buffer, and move on to the next result set.
        Lookup_FlushPrevious
        
        ' If possible, returns results in the order of distance. The measure of distance is
        ' provider specific.
        Lookup_Nearest
        
        ' This indicates whether prime response is in the remote or local part of CSADDR_INFO
        ' structure. The other part needs to be usable in either case.
        Lookup_ResService
        
        ' Any available alias information is to be returned in successive calls to
        ' WSALookupServiceNext, and each alias returned will have the RESULT_IS_ALIAS flag set.
        Lookup_ReturnAliases
        
        ' Retrieves the name as lpszServiceInstanceName.
        Lookup_ReturnName
        
        ' Retrieves the type as lpServiceClassId.
        Lookup_ReturnType
        
        ' Retrieves the version as lpVersion.
        Lookup_ReturnVersion
        
        ' Retrieves the comment as lpszComment.
        Lookup_ReturnComment
        
        ' Retrieves the addresses as lpcsaBuffer.
        Lookup_ReturnAddress
        
        ' Retrieves the private data as lpBlob.
        Lookup_ReturnBlob
        
        ' Retrieves all of the information.
        Lookup_ReturnAll
End Enum


' *************************************************************************************************
' WSA Set Service operations
' *************************************************************************************************

Public Const RNRSERVICE_REGISTER    As Long = 0
Public Const RNRSERVICE_DEREGISTER  As Long = 1
Public Const RNRSERVICE_DELETE      As Long = 2


Public Enum WSAServiceOperationType

    ' Register the service. For SAP, this means sending out a periodic broadcast. This is an NOP
    ' for the DNS name space. For persistent data stores, this means updating the address
    ' information.
    WSAServOp_Register = RNRSERVICE_REGISTER
    
    ' Remove the service from the registry. For SAP, this means stop sending out the periodic
    ' broadcast. This is an NOP for the DNS name space. For persistent data stores this means
    ' deleting address information.
    WSAServOp_Unregister = RNRSERVICE_DEREGISTER
    
    ' Delete the service from dynamic name and persistent spaces. For services represented by
    ' multiple CSADDR_INFO structures (using the SERVICE_MULTIPLE flag), only the supplied
    ' address will be deleted, and this must match exactly the corresponding CSADDR_INFO
    ' structure that was supplied when the service was registered.
    WSAServOp_Delete = RNRSERVICE_DELETE
End Enum
    

' *************************************************************************************************
' WSA Set Service control flags
' *************************************************************************************************


Public Const SERVICE_MULTIPLE   As Long = &H1 ' Controls scope of operation. When clear, service addresses are managed as a group. A register or removal from the registry invalidates all existing addresses before adding the given address set. When set, the action is only performed on the given address set. A register does not invalidate existing addresses and a removal from the registry only invalidates the given set of addresses.


' *************************************************************************************************
' WSASocket flags
' *************************************************************************************************

Public Const WSA_FLAG_OVERLAPPED        As Long = &H1
Public Const WSA_FLAG_MULTIPOINT_C_ROOT As Long = &H2
Public Const WSA_FLAG_MULTIPOINT_C_LEAF As Long = &H4
Public Const WSA_FLAG_MULTIPOINT_D_ROOT As Long = &H8
Public Const WSA_FLAG_MULTIPOINT_D_LEAF As Long = &H10


Public Enum WSASocketFlagsType

    ' This flag causes an overlapped socket to be created. Overlapped sockets can utilize WSASend,
    ' WSASendTo, WSARecv, WSARecvFrom, and WSAIoctl for overlapped I/O operations, which allow
    ' multiple operations to be initiated and in progress simultaneously. All functions that
    ' allow overlapped operation (WSASend, WSARecv, WSASendTo, WSARecvFrom, WSAIoctl) also
    ' support nonoverlapped usage on an overlapped socket if the values for parameters related to
    ' overlapped operations are NULL.
    WSASock_Overlapped = WSA_FLAG_OVERLAPPED
    
    ' Indicates that the socket created will be a c_root in a multipoint session. Only allowed if
    ' a rooted control plane is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to
    ' Multipoint and Multicast Semantics for additional information.
    WSASock_MultipointCRoot = WSA_FLAG_MULTIPOINT_C_ROOT
    
    ' Indicates that the socket created will be a c_leaf in a multicast session. Only allowed if
    ' XP1_SUPPORT_MULTIPOINT is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to
    ' Multipoint and Multicast Semantics for additional information.
    WSASock_MultipointCLeaf = WSA_FLAG_MULTIPOINT_C_LEAF
    
    ' Indicates that the socket created will be a d_root in a multipoint session. Only allowed if
    ' a rooted data plane is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to
    ' Multipoint and Multicast Semantics for additional information.
    WSASock_MultipointDRoot = WSA_FLAG_MULTIPOINT_D_ROOT
    
    'Indicates that the socket created will be a d_leaf in a multipoint session. Only allowed if
    ' XP1_SUPPORT_MULTIPOINT is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to
    ' Multipoint and Multicast Semantics for additional information.
    WSASock_MultipointDLeaf = WSA_FLAG_MULTIPOINT_D_LEAF
End Enum


' *************************************************************************************************
' Socket input output control
' *************************************************************************************************

' Masks
Public Const IOC_VOID       As Long = &H20000000      ' no parameters
Public Const IOC_OUT        As Long = &H40000000      ' copy out parameters
Public Const IOC_IN         As Long = &H80000000      ' copy in parameters
Public Const IOC_INOUT      As Long = (IOC_IN Or IOC_OUT)

Public Const IOC_UNIX       As Long = &H0&
Public Const IOC_WS2        As Long = &H8000000
Public Const IOC_PROTOCOL   As Long = &H10000000
Public Const IOC_VENDOR     As Long = &H18000000


' Enables or disables nonblocking mode on socket s. lpvInBuffer points at an unsigned long, which
' is nonzero if nonblocking mode is to be enabled and zero if it is to be disabled. When a socket
' is created, it operates in blocking mode (that is, nonblocking mode is disabled). This is
' consistent with Berkeley Software Distribution (BSD) sockets.
Public Const FIONBIO        As Long = &H8004667E

' Determines the amount of data that can be read atomically from socket s. lpvOutBuffer points at
' an unsigned long in which WSPIoctl stores the result. If s is stream oriented (for example,
' type SOCK_STREAM), FIONREAD returns the total amount of data that can be read in a single
' receive operation; this is normally the same as the total amount of data queued on the socket.
' If s is message oriented (for example, type SOCK_DGRAM), FIONREAD returns the size of the first
' datagram (message) queued on the socket.
Public Const FIONREAD       As Long = &H4004667F

' Determines whether or not all OOB data has been read. This applies only to a socket of stream
' style (for example, type SOCK_STREAM) that has been configured for inline reception of any OOB
' data (SO_OOBINLINE). If no OOB data is waiting to be read, the operation returns TRUE.
' Otherwise, it returns FALSE, and the next receive operation performed on the socket will
' retrieve some or all of the data preceding the mark; the Windows Sockets SPI client should use
' the SIOCATMARK operation to determine whether any remains. If there is any normal data
' preceding the urgent (OOB) data, it will be received in order. (Note that receive operations
' will never mix OOB and normal data in the same call.) lpvOutBuffer points at a bool in which
' WSPIoctl stores the result.
Public Const SIOCATMARK     As Long = &H40047307

' Enables a socket to receive all IP packets on the network. The socket handle passed to the
' WSAIoctl function must be of AF_INET address family, SOCK_RAW socket type, and IPPROTO_IP
' protocol. The socket also must be bound to an explicit local interface, which means that you
' cannot bind to INADDR_ANY.
'
' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return IP
' datagrams passing through the given interface. Note that you must supply a sufficiently large
' buffer. Setting this ioctl requires Administrator privilege on the local computer. SIO_RCVALL
' is available in Windows 2000 and later versions of Windows.
Public Const SIO_RCVALL                    As Long = (IOC_IN Or IOC_VENDOR Or 1)

' Enables a socket to receive all multicast IP traffic on the network (that is, all IP packets
' destined for IP addresses in the range of 224.0.0.0 to 239.255.255.255). The socket handle
' passed to the WSAIoctl function must be of AF_INET address family, SOCK_RAW socket type, and
' IPPROTO_UDP protocol. The socket also must bind to an explicit local interface, which means
' that you cannot bind to INADDR_ANY. The socket should bind to port zero.

' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return
' multicast IP datagrams passing through the given interface. Note that you must supply a
' sufficiently large buffer. Setting this ioctl requires Administrator privilege on the local
' computer. SIO_RCVALL_MCAST is available only in Windows 2000 and later versions of Windows.
Public Const SIO_RCVALL_MCAST              As Long = (IOC_IN Or IOC_VENDOR Or 2)

' Enables a socket to receive all IGMP multicast IP traffic on the network, without receiving
' other multicast IP traffic. The socket handle passed to the WSAIoctl function must be of
' AF_INET address family, SOCK_RAW socket type, and IPPROTO_IGMP protocol. The socket also must
' be bound to an explicit local interface, which means that you cannot bind to INADDR_ANY.
'
' Once the socket is bound and the ioctl set, calls to the WSARecv or recv functions return
' multicast IP datagrams passing through the given interface. Note that you must supply a
' sufficiently large buffer. Setting this ioctl requires Administrator privilege on the local
' computer. SIO_RCVALL_IGMPMCAST is available only in Windows 2000 and later versions of Windows.
Public Const SIO_RCVALL_IGMPMCAST          As Long = (IOC_IN Or IOC_VENDOR Or 3)

' Enables the per-connection setting of keep-alive option, keepalive time, and keepalive interval.
Public Const SIO_KEEPALIVE_VALS            As Long = (IOC_IN Or IOC_VENDOR Or 4)

' Associates this socket with the specified handle of a companion interface. The input buffer
' contains the integer value corresponding to the manifest constant for the companion interface
' (for example, TH_NETDEV and TH_TAPI), followed by a value that is a handle of the specified
' companion interface, along with any other required information. The handle associated by this
' IOCTL can be retrieved using SIO_TRANSLATE_HANDLE.
Public Const SIO_ASSOCIATE_HANDLE          As Long = (IOC_IN Or IOC_WS2 Or 1)

' Indicates to a message-oriented service provider that a newly arrived message should never be
' dropped because of a buffer queue overflow. Instead, the oldest message in the queue should be
' eliminated in order to accommodate the newly arrived message. No input and output buffers are
' required. Note that this IOCTL is only valid for sockets associated with unreliable,
' message-oriented protocols.
Public Const SIO_ENABLE_CIRCULAR_QUEUEING  As Long = (IOC_VOID Or IOC_WS2 Or 2)

' This IOCTL fills the output buffer with a sockaddr structure containing a suitable broadcast
' address for use with WSPSendTo.
Public Const SIO_GET_BROADCAST_ADDRESS     As Long = (IOC_OUT Or IOC_WS2 Or 5)

' Retrieves a pointer to the specified extension function supported by the associated service
' provider. The input buffer contains a GUID whose value identifies the extension function
' in question. The pointer to the desired function is returned in the output buffer. Extension
' function identifiers are established by service provider vendors and should be included in
' vendor documentation that describes extension function capabilities and semantics.
Public Const SIO_GET_EXTENSION_FUNCTION_POINTER As Long = (IOC_INOUT Or IOC_WS2 Or 6)

' Controls whether data sent in a multipoint session will also be received by the same socket on
' the local host. A value of TRUE causes loopback reception to occur while a value of FALSE
' prohibits this.
Public Const SIO_MULTIPOINT_LOOPBACK       As Long = (IOC_IN Or IOC_WS2 Or 9)

' Specifies the scope over which multicast transmissions will occur. Scope is defined as the
' number of routed network segments to be covered. A scope of zero would indicate that the
' multicast transmission would not be placed on the wire, but could be disseminated across
' sockets within the local host. A scope value of 1 (the default) indicates that the transmission
' will be placed on the wire, but will not cross any routers. Higher scope values determine the
' number of routers that can be crossed. Note that this corresponds to the time-to-live (TTL)
' parameter in IP multicasting.
Public Const SIO_MULTICAST_SCOPE           As Long = (IOC_IN Or IOC_WS2 Or 10)

' To obtain a corresponding handle for socket s that is valid in the context of a companion
' interface (for example, TH_NETDEV and TH_TAPI). A manifest constant identifying the companion
' interface along with any other needed parameters are specified in the input buffer.
' The corresponding handle will be available in the output buffer upon completion of this
' function.
Public Const SIO_TRANSLATE_HANDLE          As Long = (IOC_INOUT Or IOC_WS2 Or 13)

' To obtain the address of the local interface (represented as sockaddr structure) that should be
' used to send to the remote address specified in the input buffer (as sockaddr). Remote multicast
' addresses may be submitted in the input buffer to get the address of the preferred interface for
' multicast transmission. In any case, the interface address returned may be used by the
' application in a subsequent bind request.
Public Const SIO_ROUTING_INTERFACE_QUERY   As Long = (IOC_WS2 Or IOC_INOUT Or 20)

' To obtain a list of local transport addresses of the socket's protocol family to which the
' Windows Sockets SPI client can bind.
Public Const SIO_ADDRESSLIST_QUERY         As Long = (IOC_WS2 Or IOC_INOUT Or 22)



' When issued, this IOCTL requests that the route to the remote address specified as a sockaddr
' in the input buffer be discovered. If the address already exists in the local cache, its entry
' is invalidated. In the case of Novell's IPX, this call initiates an IPX GetLocalTarget (GLT),
' that queries the network for the given remote address.
Public Const SIO_FIND_ROUTE                As Long = (IOC_OUT Or IOC_WS2 Or 3)

' Discards current contents of the sending queue associated with this socket. No input and output
' buffers are required. The WSAENOPROTOOPT error code is indicated for service providers that do
' not support this IOCTL.
Public Const SIO_FLUSH                     As Long = (IOC_VOID Or IOC_WS2 Or 4)

' Retrieves the QOS structure associated with the socket. The input buffer is optional. Some
' protocols (for example, RSVP) allow the input buffer to be used to qualify a QOS request. The
' QOS structure will be copied into the output buffer. The output buffer must be sized large
' enough to be able to contain the full QOS structure. The WSAENOPROTOOPT error code is indicated
' for service providers that do not support quality of service.
Public Const SIO_GET_QOS                   As Long = (IOC_INOUT Or IOC_WS2 Or 7)

' Reserved.
Public Const SIO_GET_GROUP_QOS             As Long = (IOC_INOUT Or IOC_WS2 Or 8)

' Associate the supplied QOS structure with the socket. No output buffer is required, the QOS
' structure will be obtained from the input buffer. The WSAENOPROTOOPT error code is indicated
' for service providers that do not support quality of service.
Public Const SIO_SET_QOS                   As Long = (IOC_IN Or IOC_WS2 Or 11)

' Reserved.
Public Const SIO_SET_GROUP_QOS             As Long = (IOC_IN Or IOC_WS2 Or 12)

' To receive notification of the interface change that should be used to reach the remote address
' in the input buffer (specified as a sockaddr structure). No output information will be provided
' upon completion of this IOCTL; the completion merely indicates that the routing interface for a
' given destination has changed and should be queried again through SIO_ROUTING_INTERFACE_QUERY.
Public Const SIO_ROUTING_INTERFACE_CHANGE  As Long = (IOC_WS2 Or IOC_INOUT Or 21)

' To receive notification of changes in the list of local transport addresses of the socket's
' protocol family to which the Windows Sockets SPI client can bind. No output information will be
' provided upon completion of this IOCTL; the completion merely indicates that the list of
' available local addresses has changed and should be queried again through
' SIO_ADDRESS_LIST_QUERY.
Public Const SIO_ADDRESS_LIST_CHANGE       As Long = (IOC_WS2 Or IOC_INOUT Or 23)

' Used to retrieve interface info
Public Const SIO_GET_INTERFACE_LIST         As Long = &H4004747F


Public Enum SocketInputOutputOptionType
    sockio_FIONBIO = FIONBIO
    sockio_fionread = FIONREAD
    sockio_catmark = SIOCATMARK
    sockio_RecieveAll = SIO_RCVALL
    sockio_RecieveAllMulticast = SIO_RCVALL_MCAST
    sockio_RecieveAllIGMPMulticast = SIO_RCVALL_IGMPMCAST
    sockio_KeepAliveVals = SIO_KEEPALIVE_VALS
    sockio_AssociateHandle = SIO_ASSOCIATE_HANDLE
    sockio_EnableCircularQueueing = SIO_ENABLE_CIRCULAR_QUEUEING
    sockio_GetBroadcastAddress = SIO_GET_BROADCAST_ADDRESS
    sockio_GetExtenssionFunctionPointer = SIO_GET_EXTENSION_FUNCTION_POINTER
    sockio_MultipointLoopback = SIO_MULTIPOINT_LOOPBACK
    sockio_MulticastScope = SIO_MULTICAST_SCOPE
    sockio_TranslateHandle = SIO_TRANSLATE_HANDLE
    sockio_RoutingInterfaceQuery = SIO_ROUTING_INTERFACE_QUERY
    sockio_AddresslistQuery = SIO_ADDRESSLIST_QUERY
    sockio_FindRoute = SIO_FIND_ROUTE
    sockio_Flush = SIO_FLUSH
    sockio_GetQualityOfService = SIO_GET_QOS
    'sockio_GetGroupQualityOfService = SIO_GET_GROUP_QOS
    sockio_SetQualityOfService = SIO_SET_QOS
    'sockio_SetGroupQualityOfService = SIO_SET_GROUP_QOS
    sockio_RoutingInterfaceChange = SIO_ROUTING_INTERFACE_CHANGE
    sockio_AddresslistChange = SIO_ADDRESS_LIST_CHANGE
    sockio_GetInterfaceList = SIO_GET_INTERFACE_LIST
End Enum


' Dunno where this belongs. Stick it at the bottom
Public Const MAXGETHOSTSTRUCT As Long = 1024

' Dunno where this belongs either
Public Const SOMAXCONN        As Long = &H7FFFFFFF


