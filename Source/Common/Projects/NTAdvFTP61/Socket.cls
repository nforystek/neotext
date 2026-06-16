VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Socket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Implements ISocket

Private mSSL As Boolean

Private mCtx As Long
Private mServerName As String
Private mInTls As Boolean
Private mHandshakeStarted As Boolean
Private mEncryptingData As Boolean

Private WithEvents wClient As ISocket
Attribute wClient.VB_VarHelpID = -1

Public Event Connected()
Public Event Disconnected()
Public Event DataArriving()
Public Event SendComplete()
Public Event Connection(ByRef handle As Long)
Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)


Public Property Get SSL() As Boolean
    SSL = mSSL
End Property
Public Property Let SSL(ByVal RHS As Boolean)
    If Connected And (Not SSL) And RHS Then
        mSSL = RHS
        mServerName = wClient.Host
        StartTlsHandshake
        
'    ElseIf (Not RHS) Then
'        If Not (SSL And Connected) Then
'            mSSL = RHS
'        End If
    Else
        mSSL = RHS
    End If
End Property

Private Sub Class_Initialize()
    Set wClient = New ISocket
End Sub

Private Sub Class_Terminate()

    Set wClient = Nothing
End Sub

Private Function ISocket_Accept(Optional handle As Long = 0&) As Boolean
    ISocket_Accept = Me.Accept(handle)
End Function

Private Property Let ISocket_AutoRate(ByVal RHS As Boolean)
    Me.AutoRate = RHS
End Property

Private Property Get ISocket_AutoRate() As Boolean
    ISocket_AutoRate = Me.AutoRate
End Property

Private Function ISocket_Connect(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
    ISocket_Connect = Me.Connect(nHost, nPort)
End Function

Private Property Get ISocket_Connected() As Boolean
    ISocket_Connected = Me.Connected
End Property

Private Function ISocket_Decline(Optional handle As Long = 0&) As Boolean
    ISocket_Decline = Me.Decline(handle)
End Function

Private Sub ISocket_Disconnect()
    Me.Disconnect
End Sub

Private Property Get ISocket_Handle() As Long
    ISocket_Handle = Me.handle
End Property

Private Property Let ISocket_host(ByVal RHS As String)
    Me.Host = RHS
End Property

Private Property Get ISocket_host() As String
    ISocket_host = Me.Host
End Property

'Private Function ISocket_Incoming(ByVal Accept As Boolean) As Boolean
'    ISocket_Incoming = Me.Incoming(Accept)
'End Function

Private Function ISocket_Listen(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
    ISocket_Listen = Me.Listen(nHost, nPort)
End Function

Private Property Get ISocket_Listening() As Boolean
    ISocket_Listening = Me.Listening
End Property

Private Function ISocket_LocalHost() As String
    ISocket_LocalHost = Me.LocalHost
End Function

Private Function ISocket_LocalIP() As String
    ISocket_LocalIP = Me.LocalIP
End Function

Private Function ISocket_LocalIPs() As Collection
    Set ISocket_LocalIPs = Me.LocalIPs
End Function


Private Property Let ISocket_Port(ByVal RHS As Long)
    Me.Port = RHS
End Property

Private Property Get ISocket_Port() As Long
    ISocket_Port = Me.Port
End Property

Private Function ISocket_Read() As String
    ISocket_Read = Me.Read
End Function

Private Function ISocket_ReadBytes(data() As Byte) As Boolean
    ISocket_ReadBytes = Me.ReadBytes(data)
End Function


Private Function ISocket_ReadString(data As String) As Boolean
    ISocket_ReadString = ISocket.ReadString(data)
End Function

Private Function ISocket_Ready() As Boolean
    ISocket_Ready = Me.Ready
End Function

Private Property Let ISocket_RecvRate(ByVal RHS As Long)
    Me.RecvRate = RHS
End Property

Private Property Get ISocket_RecvRate() As Long
    ISocket_RecvRate = Me.RecvRate
End Property

Private Function ISocket_RemoteIP() As String
    ISocket_RemoteIP = Me.RemoteIP
End Function

Private Function ISocket_ResolveIP(ByVal Host As String) As String
    ISocket_ResolveIP = Me.ResolveIP(Host)
End Function

Private Sub ISocket_Send(ByVal data As String)
    Me.Send data
End Sub

Private Function ISocket_SendBytes(data() As Byte) As Boolean
    ISocket_SendBytes = Me.SendBytes(data)
End Function

Private Property Let ISocket_SendRate(ByVal RHS As Long)
    Me.SendRate = RHS
End Property

Private Property Get ISocket_SendRate() As Long
    ISocket_SendRate = Me.SendRate
End Property

Private Function ISocket_SendString(data As String) As Boolean
    ISocket_SendString = Me.SendString(data)
End Function

Private Property Get ISocket_State(bBit As Byte) As Boolean
    ISocket_State = Me.State(bBit)
End Property

Private Property Let ISocket_Timeout(ByVal RHS As Long)
    Me.timeout = RHS
End Property

Private Property Get ISocket_Timeout() As Long
    ISocket_Timeout = Me.timeout
End Property

Private Property Let ISocket_Transmission(ByVal RHS As Direction)
    Me.Transmission = RHS
End Property

Private Property Get ISocket_Transmission() As Direction
    ISocket_Transmission = Me.Transmission
End Property

Private Function ISocket_WhoAmI(ByVal sock As Long) As String
    ISocket_WhoAmI = Me.WhoAmI(sock)
End Function

'Private Function ISocket_Whois(ByVal Handle As Long) As String
'    ISocket_Whois = Me.WhoIs(Handle)
'End Function

Private Function ISocket_WhosIs(ByVal sock As Long) As String
    ISocket_WhosIs = Me.WhosIs(sock)
End Function



Private Sub wClient_Connected()

    If Not SSL Then
        RaiseEvent Connected

    Else

        StartTlsHandshake
        mInTls = True
        mHandshakeStarted = True

    End If

End Sub

Private Sub wClient_Connection(handle As Long)
    RaiseEvent Connection(handle)
End Sub


Private Sub wClient_DataArriving()


    If Not SSL Then

        RaiseEvent DataArriving

    Else
        
        If mHandshakeStarted Then
            Dim data() As Byte
            wClient.ReadBytes data
            HandleTlsHandshakeData data
            Exit Sub   ' ? DO NOT propagate to FTP layer
        End If
        
        RaiseEvent DataArriving
           

    End If

End Sub

Private Sub wClient_Disconnected()
    RaiseEvent Disconnected
    If mCtx <> 0 Then
        TlsClose mCtx
        mCtx = 0
    End If
    mInTls = False
    mHandshakeStarted = False
End Sub

Private Sub wClient_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    RaiseEvent Error(Number, Source, Description)
End Sub

Private Sub wClient_SendComplete()
    wClient.AutoRatedEvent AutoRated.SendEvent
    If SSL Then
    
        If mInTls And Not mHandshakeStarted Then
            RaiseEvent SendComplete
        End If
     
    Else
        RaiseEvent SendComplete
    End If
End Sub

Public Function Accept(Optional handle As Long = 0&) As Boolean
    Accept = wClient.Accept(handle)
End Function

Public Property Let AutoRate(ByVal RHS As Boolean)
    wClient.AutoRate = RHS
End Property

Public Property Get AutoRate() As Boolean
    AutoRate = wClient.AutoRate
End Property

Public Function Connect(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
    mInTls = False
    mHandshakeStarted = False
    mServerName = nHost

    Connect = wClient.Connect(nHost, nPort)
End Function

Public Property Get Connected() As Boolean
    Connected = wClient.Connected
End Property

Public Function Decline(Optional handle As Long = 0&) As Boolean
    Decline = wClient.Decline(handle)
End Function

Public Sub Disconnect()
    wClient.Disconnect
End Sub

Public Property Get handle() As Long
    handle = wClient.handle
End Property

Public Property Let Host(ByVal RHS As String)
    wClient.Host = RHS
End Property

Public Property Get Host() As String
    Host = wClient.Host
End Property

Public Function Listen(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
    Listen = wClient.Listen(nHost, nPort)
End Function

Public Property Get Listening() As Boolean
    Listening = wClient.Listening
End Property

Public Function LocalHost() As String
    LocalHost = wClient.LocalHost
End Function

Public Function LocalIP() As String
    LocalIP = wClient.LocalIP
End Function

Public Function LocalIPs() As Collection
    Set LocalIPs = wClient.LocalIPs
End Function

Public Property Let Port(ByVal RHS As Long)
    wClient.Port = RHS
End Property

Public Property Get Port() As Long
    Port = wClient.Port
End Property

Private Sub StartTlsHandshake()
    Dim errCode As Long
    Dim outBuf() As Byte
    Dim outLen As Long

    Debug.Print "Starting TLS handshake..."

    mCtx = TlsInit(mServerName, errCode)
    If mCtx = 0 Or errCode <> 0 Then
        Debug.Print "TlsInit failed, err=" & errCode
        Exit Sub
    End If

    mInTls = True
    mHandshakeStarted = True

    ReDim outBuf(0 To 8191)

    outLen = TlsHandshake(mCtx, ByVal 0&, 0, _
                          VarPtr(outBuf(LBound(outBuf))), UBound(outBuf) + IIf(LBound(outBuf) = 0, 1, 0), errCode)

    If errCode <> 0 Then
        Debug.Print "TlsHandshake initial failed, err=" & errCode
        Exit Sub
    End If

    If outLen > 0 Then
        ReDim Preserve outBuf(0 To outLen - 1)
        wClient.SendBytes outBuf
       ' Debug.Print "TLS OUT: " & outLen & " bytes (ClientHello)"
    End If
End Sub

Private Sub HandleTlsHandshakeData(enc() As Byte)
    Dim outBuf() As Byte
    Dim outLen As Long
    Dim err As Long

    ReDim outBuf(0 To 16383)

    outLen = TlsHandshake(mCtx, ByVal VarPtr(enc(LBound(enc))), UBound(enc) + IIf(LBound(enc) = 0, 1, 0), _
                          ByVal VarPtr(outBuf(LBound(outBuf))), UBound(outBuf) + IIf(LBound(outBuf) = 0, 1, 0), err)

    If err <> 0 Then
        Debug.Print "TlsHandshake failed, err=" & err
        Exit Sub
    End If

    If outLen > 0 Then
        ReDim Preserve outBuf(0 To outLen - 1)
        wClient.SendBytes outBuf
        'Debug.Print "TLS OUT: " & outLen & " bytes"
    End If

    If TlsIsHandshakeComplete(mCtx) <> 0 Then
        Debug.Print "TLS handshake complete."
        mHandshakeStarted = False
        mInTls = True
    End If
End Sub


Private Sub HandleTlsAppData(enc() As Byte)
    Dim plain() As Byte
    Dim plainLen As Long
    Dim err As Long

    ReDim plain(0 To 16383)

    ' Decrypt TLS record
    plainLen = TlsRecv(mCtx, ByVal VarPtr(enc(LBound(enc))), UBound(enc) + IIf(LBound(enc) = 0, 1, 0), _
                       ByVal VarPtr(plain(LBound(plain))), UBound(plain) + IIf(LBound(plain) = 0, 1, 0), err)

    If err <> 0 Then
        Debug.Print "TlsRecv failed, err=" & err
        Exit Sub
    End If

    If plainLen > 0 Then enc = plain

End Sub




Public Function Read() As String
    If Not SSL Then
        wClient.ReadString Read
    Else
        If mHandshakeStarted Then
            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function
        End If
        
        Dim data() As Byte
        
        If mInTls Then
            If wClient.ReadBytes(data) Then
                HandleTlsAppDat data
                Read = StringFromBytes(data)
            End If
        Else
            wClient.ReadString Read
'            wClient.AutoRatedEvent AutoRated.ReadCalls
'            Read = SecureSession.OutBuffer
'            wClient.AutoRatedEvent AutoRated.ReadWidth
    
        End If

    End If

End Function

Public Function ReadBytes(data() As Byte) As Boolean
    If Not SSL Then
        ReadBytes = wClient.ReadBytes(data)
    Else
        If mHandshakeStarted Then
            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function
        End If
        
        If mInTls Then
            ReadBytes = wClient.ReadBytes(data)
            If ReadBytes Then HandleTlsAppData data
        Else
            ReadBytes = wClient.ReadBytes(data)
'            wClient.AutoRatedEvent AutoRated.ReadCalls
'            Erase data
'            data = SecureSession.sOutBuffer.Partial
'            ReadBytes = (SecureSession.sOutBuffer.Length > 0)
'            wClient.AutoRatedEvent AutoRated.ReadWidth
    
        End If


    End If
    
End Function


Public Function ReadString(data As String) As Boolean
    If Not SSL Then
        ReadString = wClient.ReadString(data)
    Else
    
        If mHandshakeStarted Then
            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function
        End If
        
        Dim b() As Byte

        If mInTls Then
            ReadString = wClient.ReadBytes(b)
            If ReadBytes Then
                HandleTlsAppData b
                data = StringFromBytes(b)
            End If
        Else
             ReadString = wClient.ReadString(data)
            'wClient.AutoRatedEvent AutoRated.ReadCalls
            'data = wClient.ReadString(data)
            'ReadString = (SecureSession.sOutBuffer.Length > 0)
            'wClient.AutoRatedEvent AutoRated.ReadWidth
        End If

    End If

End Function

Public Function Ready() As Boolean
    Ready = wClient.Ready
End Function

Public Property Let RecvRate(ByVal RHS As Long)
    wClient.RecvRate = RHS
End Property

Public Property Get RecvRate() As Long
    RecvRate = wClient.RecvRate
End Property

Public Function RemoteIP() As String
    RemoteIP = wClient.RemoteIP
End Function



Public Function ResolveIP(ByVal Host As String) As String
    ResolveIP = wClient.ResolveIP(Host)
End Function

Public Sub Send(ByVal data As String)

    If Not SSL Then
        wClient.SendBytes BytesFromString(data)
    Else

'        If mInTls Then
'            SendTlsCommand BytesFromString(data)
'        Else
'            wClient.SendBytes BytesFromString(data)
'        End If

        If mHandshakeStarted Then
            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function
        End If

        
        If mInTls And Not mHandshakeStarted Then
            SendTlsCommand BytesFromString(data)
        Else
            wClient.SendBytes BytesFromString(data)
        End If

'        wClient.AutoRatedEvent AutoRated.SendCalls
'
'        SecureSession.InBuffer = data
'        Call SecureSession.TLSEncrypt
'        data = AddRecHeader(SecureSession.OutBuffer)
'        wClient.SendBytes modMemory.Convert(data)
'        SecureSession.IncSendSeqNum
'
'        wClient.AutoRatedEvent AutoRated.SendWidth
    End If

End Sub

Public Function SendBytes(data() As Byte) As Boolean
    If Not SSL Then
        SendBytes = wClient.SendBytes(data)
    Else
'        If mInTls Then
'            SendBytes = SendTlsCommand(data)
'        Else
'            SendBytes = wClient.SendBytes(data)
        If mHandshakeStarted Then
            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function
        End If
 
        
        
        If mInTls Then
            SendBytes = SendTlsCommand(data)
        Else
            SendBytes = wClient.SendBytes(data)
        End If


'        wClient.AutoRatedEvent AutoRated.SendCalls
'
'        SecureSession.InBuffer = modMemory.Convert(data)
'        Call SecureSession.TLSEncrypt
'        data = modMemory.Convert(AddRecHeader(SecureSession.OutBuffer))
'        SendBytes = wClient.SendBytes(data)
'        SecureSession.IncSendSeqNum
'
'        wClient.AutoRatedEvent AutoRated.SendWidth
    End If

End Function


Public Property Let SendRate(ByVal RHS As Long)
    wClient.SendRate = RHS
End Property

Public Property Get SendRate() As Long
    SendRate = wClient.SendRate
End Property

Public Function SendString(data As String) As Boolean
    If Not SSL Then
        SendString = wClient.SendBytes(BytesFromString(data))
    Else
'        If mInTls Then
'            SendString = SendTlsCommand(BytesFromString(data))
'        Else
'            SendString = wClient.SendBytes(BytesFromString(data))
'        End If

        If mHandshakeStarted Then

            ' BLOCK ALL APPLICATION DATA
            'Debug.Print "Blocked send during TLS handshake"
            Do Until mHandshakeStarted = False And mInTls = True
                DoTasks
            Loop
            'Exit Function

        End If
        
        If mInTls And Not mHandshakeStarted Then
            SendString = SendTlsCommand(BytesFromString(data))
        Else
            SendString = wClient.SendBytes(BytesFromString(data))
        End If
        
'        wClient.AutoRatedEvent AutoRated.SendCalls
'
'
'        SecureSession.InBuffer = data
'        Call SecureSession.TLSEncrypt
'        data = AddRecHeader(SecureSession.OutBuffer)
'        SendString = wClient.SendBytes(modMemory.Convert(data))
'        SecureSession.IncSendSeqNum
'
'        wClient.AutoRatedEvent AutoRated.SendWidth
       
    End If

End Function

Private Function SendTlsCommand(ByRef plain() As Byte) As Boolean
   ' Dim plain() As Byte
    Dim enc() As Byte
    Dim encLen As Long
    Dim err As Long

    ' Convert command to bytes
   ' plain = BytesFromString(cmd & vbCrLf)

    ' Allocate encrypted buffer
    ReDim enc(0 To 16383)

    ' Encrypt
    encLen = TlsSend(mCtx, ByVal VarPtr(plain(LBound(plain))), UBound(plain) + IIf(LBound(plain) = 0, 1, 0), _
                     ByVal VarPtr(enc(LBound(enc))), UBound(enc) + IIf(LBound(enc) = 0, 1, 0), err)

    If err <> 0 Then
        Debug.Print "TlsSend failed, err=" & err
        Exit Function
    End If

    ' Send encrypted TLS record
    If encLen > 0 Then
        ReDim Preserve enc(0 To encLen - 1)
        SendTlsCommand = wClient.SendBytes(enc)
        'Debug.Print "TLS OUT (encrypted): " & encLen & " bytes"
    End If
End Function


Public Property Get State(bBit As Byte) As Boolean
    State = wClient.State(bBit)
End Property

Public Property Let timeout(ByVal RHS As Long)
    wClient.timeout = RHS
End Property

Public Property Get timeout() As Long
    timeout = wClient.timeout
End Property

Public Property Let Transmission(ByVal RHS As Direction)
    wClient.Transmission = RHS
End Property

Public Property Get Transmission() As Direction
    Transmission = wClient.Transmission
End Property

Public Function WhoAmI(ByVal sock As Long) As String
    WhoAmI = wClient.WhoAmI(sock)
End Function

Public Function WhosIs(ByVal sock As Long) As String
    WhosIs = wClient.WhosIs(sock)
End Function





