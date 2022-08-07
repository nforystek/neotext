VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Option Compare Binary


Public Enum Direction
    NormalStream = 1
    TextMessages = 2
    SocketPaused = 6
    StandByPause = 8
End Enum

Public Enum States
    Closed = 1
    Binded = 2
    Request = 4
    Opening = 8
    Opened = 16
    Sending = 32
    Reading = 64
    Closing = 128
End Enum

'Public Enum AutoRated
'    HaltRates = -3
'    OpenRates = -1
'    SendCalls = 1
'    SendWidth = 2
'    ReadCalls = 3
'    ReadWidth = 4
'
'    SendEvent = 5
'    ReadEvent = 6
'End Enum

Private mBanding As sockaddr
Private mAddress As sockaddr

Private mHost As String
Private mPort As Long

Private sendFlag As Long
Private readFlag As Long

Private sendLatency As Single
Private sendElapsed As Single
Private sendBearing As Single
Private sendJumping As Single

Private readLatency As Single
Private readElapsed As Single
Private readBearing As Single
Private readJumping As Single

Private mSocket As Long
Private mState As Byte
Private mTransmission As Direction

Private mAutoRate As Boolean
Private mRecvRate As Long
Private mSendRate As Long
Private mTimeOut As Long

Private mSendData As New Stream
Private mReadData As New Stream

Private mNetAdapter As Integer

Public Event Connected()
Public Event Disconnected()
Public Event DataArriving()
Public Event SendComplete()
Public Event Connection(ByRef Handle As Long)
Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)


Public Property Get Host() As String ' _
Gets or sets the objects remote Host used for Listen or Connect when not supplied as parameters, which if supplied in a call also sets this property.
    Host = mHost
End Property
Public Property Let Host(ByVal newVal As String)
    mHost = newVal
End Property

Public Property Get Port() As Long ' _
Gets or sets the objects remote Port used for Listen or Connect when not supplied as parameters, which if supplied in a call also sets this property.
    Port = mPort
End Property
Public Property Let Port(ByVal newVal As Long)
    mPort = newVal
End Property

Public Property Get Connected() As Boolean ' _
Returns whether the socket is in a active state by a previous call to Connect, any state of Opening, Opned or Closing while not Closed, validates as Connected.
    Connected = ((State(States.Opened) Or State(States.Opening)) And State(States.Closing)) Or (State(States.Opened) And (Not State(States.Closed)))
End Property
Public Property Get Listening() As Boolean ' _
Returns whether the socket is in Listening mode from a previous call to Listen, and would determine the Connection event then is also active if so.
    Listening = State(States.Binded)
End Property

Public Property Get timeout() As Long ' _
Gets or sets the maximum time out in seconds the socket will wait while not receving any activity.  Zero TimeOut is equivelent to KeepAlive.
    timeout = mTimeOut
End Property
Public Property Let timeout(ByVal newVal As Long)
    mTimeOut = newVal
    SetReceive Handle
    SetSending Handle
End Property

Public Property Get SendRate() As Long ' _
Gets or set the number of bytes at which SendBytes will proxy to SendNextPart.  Transmission TextMessages forces this value to 4069.
    SendRate = mSendRate
End Property
Public Property Let SendRate(ByVal newVal As Long)
    mSendRate = IIf(CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages), 4096, newVal)
    SetSending Handle
End Property

Public Property Get RecvRate() As Long ' _
Gets or set the number of bytes at which ReadBytes will proxy to ReadNextPart.  Transmission TextMessages forces this value to 4069.
    RecvRate = mRecvRate
End Property
Public Property Let RecvRate(ByVal newVal As Long)
    mRecvRate = IIf(CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages), 4096, newVal)
    SetReceive Handle
End Property

Public Property Get Transmission() As Direction ' _
Gets or sets the type of transmissions, either TextMessage or NormalStream are basis one or the other, while the flag StandByPause, sets whether the socket will flag SocketPaused upon standby, which will pause it.
    Transmission = mTransmission
End Property
Public Property Let Transmission(ByVal newVal As Direction)
    mTransmission = newVal
    SetReceive Handle
    SetSending Handle
End Property


Public Property Get Handle() As Long ' _
Returns the current active socket handle that this object is associated with, for use with WhosIs, WhoAmI and Decline or Accept functions.
    If mSocket < -1 Then
        Handle = (-mSocket) - 1
    Else
        Handle = mSocket
    End If
End Property


Public Property Get AutoRate() As Boolean ' _
Gets or sets a option that allos the changing of receive and sending buffer sizing to automate, if expereincing net bottle necking, try disabling this and set custom rates.
Attribute AutoRate.VB_Description = "Gets or sets a option that allos the changing of receive and sending buffer sizing to automate, if expereincing net bottle necking, try disabling this and set custom rates."
    AutoRate = mAutoRate
End Property
Public Property Let AutoRate(ByVal RHS As Boolean)
    mAutoRate = RHS
End Property


Friend Property Get Address() As Long
    Address = ObjPtr(Me)
End Property

Friend Property Let State(ByRef bBit As Byte, ByRef nValue As Boolean)
    If (mState And bBit) And (Not nValue) Then
        mState = mState - bBit
    ElseIf (Not (mState And bBit)) And nValue Then
        mState = mState Or bBit
    End If
End Property
Public Property Get State(ByRef bBit As Byte) As Boolean ' _
Gets the current state of the state bit, States, supplied as the parameter retruning true or false determining whether the state is or is not active basis of the State bit.
    State = (mState And bBit)
End Property


Friend Property Get ReadBuff() As Stream
    Set ReadBuff = mReadData
End Property
Friend Property Get SendBuff() As Stream
    Set SendBuff = mSendData
End Property


Friend Function EventRaised(ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo sockerror
    If (Not CBool((mTransmission And Direction.SocketPaused) = Direction.SocketPaused)) Or _
        CBool(((mTransmission And Direction.SocketPaused) = Direction.SocketPaused) And _
        (Not CBool((mTransmission And Direction.StandByPause) = Direction.StandByPause))) Then
        If (WSAGetLastError <> 0) Then
            RaiseEvent Error(8, "Socket", GetErrorDescription(WSAGetLastError))
        Else
            DataValidating wParam


            Select Case lParam 'WSAGetSelectEvent(lParam)
                Case FD_CONNECT
                    DebugEvent "FD_CONNECT"
                    State(States.Opened) = True
                    State(States.Opening) = False
                    RaiseEvent Connected
                    EventRaised = -CInt(State(States.Closing) = False)
                Case FD_ACCEPT
                    DebugEvent "FD_ACCEPT"
                    State(States.Opening) = False
                    EventRaised = -CInt(Incoming)
                Case FD_CLOSE
                    DebugEvent "FD_CLOSE"

                    If (State(States.Opening) = False) And (State(States.Closed) = False) Then
                        EventRaised = -CInt(Ready And ((State(States.Opened) = False) And ((State(States.Closing) - True) Or (State(States.Closed) - True))))
                        If (State(States.Closing) = False) And (State(States.Closed) = False) Then
                            DebugEvent "Disconnected"
                            RaiseEvent Disconnected
                            UnregisterSocket Me
                            State(States.Closed) = True
                            'EventRaised = 1
                        End If
                    ElseIf State(States.Closing) Then
                        State(States.Closed) = Ready
                    End If


                Case FD_WRITE
                    AutoRatedEvent AutoRated.SendEvent
                    DebugEvent "FD_WRITE"
                    If Not State(States.Sending) Then
                        RaiseEvent SendComplete
                    Else
                        State(States.Sending) = Not (mSendData.Length = 0)
                    End If
                    EventRaised = -CInt(Not CloseConnection)

                Case FD_READ, FD_OOB
                    AutoRatedEvent AutoRated.ReadEvent
                    DebugEvent "FD_READ"
                    If (mSendData.Length = 0) Then
                        State(States.Sending) = False
                    End If


                    State(States.Reading) = True
                    EventRaised = -CInt(ReadNextPart = 1)
                    If (mReadData.Length > 0) Then
                        RaiseEvent DataArriving
                    ElseIf EventRaised = 1 Then
                        State(States.Reading) = False
                        EventRaised = -CInt(Not CloseConnection)
                    End If




'                        Case FD_QOS
'                            DebugEvent "FD_QOS"
'                        Case FD_GROUP_QOS
'                            DebugEvent "FD_GROUP_QOS"
'                        Case FD_ROUTING_INTERFACE_CHANGE
'                            DebugEvent "FD_ROUTING_INTERFACE_CHANGE"
'                        Case FD_ADDRESS_LIST_CHANGE
'                            DebugEvent "FD_ADDRESS_LIST_CHANGE"
'                        Case FD_MAX_EVENTS
'                            DebugEvent "FD_MAX_EVENTS"

            End Select
'                Else
'                    RaiseEvent Error(8, "Socket", "Either remote or local socket handle did not match ip address for the connection.")
'                End If
'            Else
'                RaiseEvent Error(8, "Socket", "The remote address did not match the ip address for the connection.")
'            End If
        End If
    ElseIf CBool(((mTransmission And Direction.SocketPaused) = Direction.SocketPaused) And _
        CBool((mTransmission And Direction.StandByPause) = Direction.StandByPause)) Then
        EventRaised = 0
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to incoming request."
    Resume
End Function



Public Function WhosIs(ByVal sock As Long) As String ' _
Accepts a socket handle as parameter and returns the remote party's IP address to the handle, considerably, whom "I" connected to of two parties.
Attribute WhosIs.VB_Description = "Accepts a socket handle as parameter and returns the remote party's IP address to the handle, considerably, whom "

    Dim sck As sockaddr
    If getpeername(sock, sck, LenB(sck)) = 0 Then
        Dim col As VBA.Collection
        Set col = IPAddress(sck.sin_addr)
        If col.count > 0 Then
            WhosIs = col.Item(1)
        Else
            WhosIs = "#INVALID#"
        End If
    Else
        WhosIs = "#INVALID#"
    End If
End Function
Public Function WhoAmI(ByVal sock As Long) As String ' _
Acccepts a socket handle as parameter and returns the local party's IP address to the handle, considerably, "I", or self, in a two party connection.
Attribute WhoAmI.VB_Description = "Acccepts a socket handle as parameter and returns the local party's IP address to the handle, considerably, "
    Dim sck As sockaddr
    If getsockname(sock, sck, LenB(sck)) = 0 Then
        Dim col As VBA.Collection
        Set col = IPAddress(sck.sin_addr)
        If col.count > 0 Then
            WhoAmI = col.Item(1)
        Else
            WhoAmI = "#INVALID#"
        End If
    Else
        WhoAmI = "#INVALID#"
    End If
End Function

Public Function LocalHost() As String ' _
Returns the local hsot name or computer name associated with the local hosts IP address in domain name system format, i.e. wording.
Attribute LocalHost.VB_Description = "Returns the local hsot name or computer name associated with the local hosts IP address in domain name system format, i.e. wording."
On Error GoTo sockerror
    Dim Buf As String
    Dim rc As Long
    Buf = Space$(255)
    rc = modSockets.gethostname(Buf, Len(Buf))
    rc = InStr(Buf, vbNullChar)
    If rc > 0 Then
        LocalHost = Left$(Buf, rc - 1)
    Else
        LocalHost = ""
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to resolve localhost."
End Function

Public Function LocalIP() As String ' _
Returns the local IP address of which this socket is currently utilizing, in multiple adapter use, this resolves to the local IP used in contact.
Attribute LocalIP.VB_Description = "Returns the local IP address of which this socket is currently utilizing, in multiple adapter use, this resolves to the local IP used in contact."
    LocalIP = WhoAmI(Handle)
    If LocalIP = "#INVALID#" Then
        If IsNumeric(Replace(LocalHost, ".", "")) And InStr(LocalHost, ".") > 0 Then
            LocalIP = LocalHost
        Else
            LocalIP = ResolveIP(LocalHost)
        End If
    End If
End Function

Public Function RemoteIP() As String ' _
Returns the IP address of the remote socket's connection, usually this is same as the property supplied for Connect, or Host resolves to the be the same.
Attribute RemoteIP.VB_Description = "Returns the IP address of the remote socket's connection, usually this is same as the property supplied for Connect, or Host resolves to the be the same."
    RemoteIP = WhosIs(Handle)
End Function

Public Function ResolveIP(ByVal Host As String) As String ' _
Accepts the input of a Internic Netowrk Solutions domain name system reference and returns the IP associated with it, computer name may also be supplied for the IP related.
Attribute ResolveIP.VB_Description = "Accepts the input of a Internic Netowrk Solutions domain name system reference and returns the IP associated with it, computer name may also be supplied for the IP related."
On Error GoTo sockerror
    ResolveIP = modSockets.ResolveIP(Host)
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "socket", "Error attempting to resolve host to ip."
End Function

Public Function Ready() As Boolean ' _
Checks the state of sending or reading, and includes the biffer lengths to determine if the socket is ready, meaning it is a at idle state, or still active.
Attribute Ready.VB_Description = "Checks the state of sending or reading, and includes the biffer lengths to determine if the socket is ready, meaning it is a at idle state, or still active."
    Ready = ((State(States.Reading) = False) And (State(States.Sending) = False))
    If (Not mReadData Is Nothing) Then Ready = Ready And (mReadData.Length = 0)
    If (Not mSendData Is Nothing) Then Ready = Ready And (mSendData.Length = 0)
End Function

Public Sub Disconnect() ' _
Resets the socket from most states to unopened and neither listening, also occurs when the remote party intiaites close but will finishing reading or sending first.
Attribute Disconnect.VB_Description = "Resets the socket from most states to unopened and neither listening, also occurs when the remote party intiaites close but will finishing reading or sending first."
On Error GoTo sockerror

    State(States.Opening) = False
    If Not (Connected Or Listening) Then
        State(States.Closing) = False
    Else
        State(States.Closing) = True
    End If

    If (Connected Or Listening) Then
        CloseConnection
        State(States.Closed) = True
        State(States.Closing) = False
        AutoRatedEvent AutoRated.HaltRates
    End If

Exit Sub
sockerror:
    Err.Clear
End Sub


Public Function Accept(Optional ByRef nHandle As Long = 0) As Boolean ' _
Accepts a incoming connection, or mismatched addressed data arrival for the supplied handle provided by the Connection event raised, maybe the lsitening socket or any closed other socket whom accepts.  Data arrival secuirty will raise the Error event if not accepted or declined.
Attribute Accept.VB_Description = "Accepts a incoming connection, or mismatched addressed data arrival for the supplied handle provided by the Connection event raised, maybe the lsitening socket or any closed other socket whom accepts.  Data arrival secuirty will raise the Error event if not accepted or declined."
    If (State(States.Request) Or State(States.Closed) Or State(States.Binded)) Then
        If State(States.Binded) Then
            State(States.Opening) = True
            State(States.Closed) = False
            socketclose Handle
            UnregisterSocket Me
            State(States.Binded) = False
        End If
        If (SetEvents(True, FD_EVENTS, nHandle) <> Socket_ERROR) Then
            AutoRatedEvent AutoRated.OpenRates
            State(States.Opening) = True
            State(States.Closed) = False
            mSocket = nHandle
            RegisterSocket Me
            mSendData.Reset
            mReadData.Reset
            State(States.Opened) = True
            RaiseEvent Connected
            State(States.Opening) = False
            Accept = True
        End If
    ElseIf (State(States.Opened) Or State(States.Reading) Or State(States.Sending)) And (mSocket < -1) Then
        Accept = UnsetValidatingMode(nHandle)
    Else
        Err.Raise 8, "Connect", "Invalid at current state."
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Accept", "Error attempting to accept."
End Function

Public Function Decline(Optional ByRef nHandle As Long = 0) As Boolean ' _
Declines a active connection request via the handle provided by the Connection event, strictly, enforcing closed activity from and to the peer.
Attribute Decline.VB_Description = "Declines a active connection request via the handle provided by the Connection event, strictly, enforcing closed activity from and to the peer."
On Error GoTo sockerror
    If ((State(States.Request) Or State(States.Closed) Or State(States.Binded))) Or _
         ((State(States.Opened) Or State(States.Reading) Or State(States.Sending)) And (mSocket < -1)) Then
        If nHandle = Handle Then
            State(States.Closing) = True
            SetEvents False, 0, nHandle
        End If
        socketclose nHandle
        UnregisterSocket Me
        If nHandle = Handle Then
            AutoRatedEvent AutoRated.OpenRates
            State(States.Binded) = False
            State(States.Closing) = False
            State(States.Closed) = True
        End If
        nHandle = INVALID_SOCKET
        Decline = True

    Else
        Err.Raise 8, "Decline", "Invalid at current state."
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to decline."
End Function


Public Function Listen(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0) As Boolean ' _
Places the socket into listening mode for the host and post supplied, if none provided, the objects properties are used, returns true if listening is sucess.
Attribute Listen.VB_Description = "Places the socket into listening mode for the host and post supplied, if none provided, the objects properties are used, returns true if listening is sucess."
On Error GoTo sockerror
    If State(States.Closed) Then
        State(States.Request) = False
        State(States.Opening) = False
        Dim rc As Long
        Dim sk As Long
        SetEvents True, ByVal 0&
        If Not nHost = "" Then mHost = nHost
        If Not nPort = 0 Then mPort = nPort
        mSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If (Not (rc = Socket_ERROR)) Then
            With mBanding
                .sin_family = AF_INET
                .sin_port = htons(Port)
                If IsNumeric(Replace(Host, ".", "")) And InStr(Host, ".") > 0 Then
                    .sin_addr = modSockets.Resolve(Host)
                Else
                    .sin_addr = modSockets.Resolve(ResolveIP(Host))
                End If
            End With

            rc = modSockets.Bind(Handle, mBanding, sockaddr_size)
            If rc = 0 Then

                rc = SocketListen(Handle, 0)
                If rc <> Socket_ERROR Then
                    State(States.Binded) = True
                    State(States.Closed) = False

                    rc = SetEvents(True, FD_LISTEN, Handle)
                    If rc <> Socket_ERROR Then
                        RegisterSocket Me
                        AutoRatedEvent AutoRated.HaltRates

                        Listen = True
                    End If
                End If
            End If
        End If
    Else
        Err.Raise 8, "Connect", "Invalid at current state."
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to bind or listen."
End Function

Public Function Connect(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0) As Boolean ' _
Connects to the specified host and port, if none supplied the objects host and post property are evaluated for use, returns whether the connect succeeds and is awaiting connect event.
Attribute Connect.VB_Description = "Connects to the specified host and port, if none supplied the objects host and post property are evaluated for use, returns whether the connect succeeds and is awaiting connect event."
On Error GoTo sockerror
    If Not State(States.Closed) Then
        Disconnect
    End If
    If State(States.Closed) Then
        Dim rc As Long
        State(States.Binded) = False
        State(States.Opening) = True
        State(States.Request) = False
        If Not nHost = "" Then mHost = nHost
        If Not nPort = 0 Then mPort = nPort
        mSocket = Socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
        If (Not (Handle = Socket_ERROR)) Then

            With mAddress
                .sin_family = AF_INET
                .sin_port = htons(Port)
                If IsNumeric(Replace(Host, ".", "")) And InStr(Host, ".") > 0 Then
                    .sin_addr = modSockets.Resolve(Host)
                Else
                    .sin_addr = modSockets.Resolve(ResolveIP(Host))
                End If
            End With

            rc = SetEvents(True, FD_EVENTS)
            If rc <> Socket_ERROR Then
                AutoRatedEvent AutoRated.OpenRates
                State(States.Closed) = False
                RegisterSocket Me
                mSendData.Reset
                mReadData.Reset
                SocketConnect Handle, mAddress, sockaddr_size
                Connect = (WSAGetLastError <> WSAEWOULDBLOCK)
            End If
        End If
    Else
        Err.Raise 8, "Connect", "Invalid at current state."
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Connect", "Error attempting to connect to host."
End Function


Public Sub Send(ByVal Data As String) ' _
Preforms the same as SendString, with out a return value, extra converting occurs for the data string parameter to a byte array.
Attribute Send.VB_Description = "Preforms the same as SendString, with out a return value, extra converting occurs for the data string parameter to a byte array."
On Error GoTo sockerror
    SendBytes modMemory.Convert(Data)
Exit Sub
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to send data."
End Sub

Public Function SendString(ByRef Data As String) As Boolean ' _
Preforms the same funciton as SendBytes, except converts the string parameter supplied to a byte array to supply to sending out.
Attribute SendString.VB_Description = "Preforms the same funciton as SendBytes, except converts the string parameter supplied to a byte array to supply to sending out."
On Error GoTo sockerror
    SendString = SendBytes(modMemory.Convert(Data))
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to send data."
End Function

Public Function ReadString(ByRef Data As String) As Boolean ' _
Preforms the same funciton as ReadBytes, except converts the byte array to the string parameter supplied for output after reading.
Attribute ReadString.VB_Description = "Preforms the same funciton as ReadBytes, except converts the byte array to the string parameter supplied for output after reading."
On Error GoTo sockerror
    Data = ""
    Dim bData() As Byte
    ReadString = ReadBytes(bData)
    If ReadString Then
        Data = modMemory.Convert(bData)
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to read data."
End Function

Public Function Read() As String ' _
Preforms the same as ReadString, returning however, the data string read from the buffer instead of as a parameter or boolean state.
Attribute Read.VB_Description = "Preforms the same as ReadString, returning however, the data string read from the buffer instead of as a parameter or boolean state."
On Error GoTo sockerror
    Dim bData() As Byte
    If ReadBytes(bData) Then
        Read = modMemory.Convert(bData)
    Else
        Read = Empty
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to read data."
End Function

Public Function SendBytes(ByRef Data() As Byte) As Boolean ' _
Push the entire parameter byte arra data current send buffer and returns true if there is current send activity to winsock still occuring.
Attribute SendBytes.VB_Description = "Push the entire parameter byte arra data current send buffer and returns true if there is current send activity to winsock still occuring."
On Error GoTo sockerror
    If (Not (mSendData Is Nothing)) And (Not State(States.Sending)) Then
        AutoRatedEvent AutoRated.SendCalls
        Dim lSize As Long
        Dim nSize As Long
        nSize = modMemory.ArraySize(Data)
        lSize = mSendData.Length
        State(States.Sending) = True

        SendBytes = (mSendData.Length = 0)
        If (nSize > 0) Then
            mSendData.Length = lSize + nSize
            mSendData.Placeat Data, lSize, nSize
        End If

        If (SendNextPart = 1) Then
            SendBytes = True
        Else
            SendBytes = (mSendData.Length = 0)
        End If
        State(States.Sending) = False

        AutoRatedEvent AutoRated.SendWidth
    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to read data."
End Function

Public Function ReadBytes(ByRef Data() As Byte) As Boolean ' _
Pops the entire current read buffer to the data parameter byte array and returns true if winsock continued buffer reading activity is occuring.
Attribute ReadBytes.VB_Description = "Pops the entire current read buffer to the data parameter byte array and returns true if winsock continued buffer reading activity is occuring."
On Error GoTo sockerror
    If (Not (mReadData Is Nothing)) And State(States.Reading) Then

        AutoRatedEvent AutoRated.ReadCalls
        Dim lSize As Long
        Erase Data
        lSize = mReadData.Length
        If lSize > mRecvRate Then lSize = mRecvRate
        If (lSize > 0) Then
            Data = mReadData.Partial(0, lSize)
            mReadData.Push lSize
            ReadBytes = (mReadData.Length = 0)
        End If
       ' ReadBytes = (mReadData.Length = 0)
        State(States.Reading) = (Not (mReadData.Length = 0))
        AutoRatedEvent AutoRated.ReadWidth

    End If
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to read data."
End Function



Public Function LocalIPs() As Collection ' _
A collection that contains all the IP addresses found on the local system network adapters.
    Set LocalIPs = GetPortIP
End Function


Friend Sub Class_Initialize()
    State(States.Closed) = True
    mSocket = INVALID_SOCKET

    mNetAdapter = -1
    mTransmission = NormalStream
    mTimeOut = 30
    mRecvRate = ftpBufferSize
    mSendRate = ftpPacketSize
    mAutoRate = True
    AutoRatedEvent AutoRated.HaltRates

    modWinProc.SocketsInitialize

End Sub

Friend Sub Class_Terminate()
    Disconnect

    Set mSendData = Nothing
    Set mReadData = Nothing

    modWinProc.SocketsCleanUp

End Sub


Friend Function Resolve(ByVal Host As String) As Long
On Error GoTo sockerror
    Resolve = modSockets.Resolve(Host)
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to resolve address."
End Function

Friend Sub AutoRatedEvent(ByVal Func As AutoRated)

    If mAutoRate Then
        Select Case Func
            Case 1 'calculate latency
                sendLatency = (Timer - sendLatency)
            Case 2 'reset latency
                sendLatency = Timer
            Case 3 'calulate latency and elapsed same itteration and do adjust
                readLatency = (Timer - readLatency)
                readElapsed = (Timer - readElapsed)
                Adjust readBearing, readJumping, readLatency, readElapsed, mRecvRate
                SetReceive
            Case 4 'reset latency
                readLatency = Timer
            Case 5 'calculate elapsed and do adjust
                sendElapsed = (Timer - sendElapsed)
                Adjust sendBearing, sendJumping, sendLatency, sendElapsed, mSendRate
                SetSending
            Case 6 'reset elapsed
                readElapsed = Timer
            Case -1 'inital data either way
                readBearing = 0
                readJumping = 0
                sendBearing = 0
                sendJumping = 0
            Case -2 'open prep
                readLatency = 0
                readElapsed = 0
                sendLatency = 0
                sendElapsed = 0
        End Select
    End If

    If Func = -3 Then

        readLatency = 0
        readElapsed = 0
        readBearing = 0
        readJumping = 0

        sendLatency = 0
        sendElapsed = 0
        sendBearing = 0
        sendJumping = 0

    End If
End Sub



Friend Function PauseSocket()
    If Not CBool((mTransmission And Direction.SocketPaused) = Direction.SocketPaused) Then
        mTransmission = mTransmission + Direction.SocketPaused
        SetReceive Handle
        SetSending Handle
    End If
End Function
Friend Function ResumeSocket()
    If CBool((mTransmission And Direction.SocketPaused) = Direction.SocketPaused) Then
        mTransmission = mTransmission - Direction.SocketPaused
        SetReceive Handle
        SetSending Handle
    End If
End Function

Friend Function SetReceive(Optional ByVal sck As Long = -1) As Long
On Error GoTo sockerror
    If sck = -1 Then sck = Handle
    Dim timo As Long
    timo = IIf(CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages), 0, mTimeOut)
    SetReceive = (setsockopt(sck, SOL_SOCKET, SO_RCVTIMEO, ByVal VarPtr(timo), 4&) = 0) _
            And (setsockopt(sck, SOL_SOCKET, SO_RCVBUF, ByVal VarPtr(mRecvRate), 4&) = 0)
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to set socket options."
End Function

Friend Function SetSending(Optional ByVal sck As Long = -1) As Long
On Error GoTo sockerror
    If sck = -1 Then sck = Handle
    Dim timo As Long
    timo = IIf(CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages), 0, mTimeOut)
    SetSending = (setsockopt(sck, SOL_SOCKET, SO_SNDTIMEO, ByVal VarPtr(timo), 4&) = 0) _
                And (setsockopt(sck, SOL_SOCKET, SO_SNDBUF, ByVal VarPtr(mSendRate), 4&) = 0)
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to set socket options."
End Function



Friend Function SendNextPart() As Long ' _
Preforms the activity of the winsock sending from the send buffer and returns 1 if all current buffer was send, otherwise 0.
On Error GoTo sockerror
    If mSendData Is Nothing Then Exit Function
    AutoRatedEvent AutoRated.SendCalls

    If (mSendData.Length > 0) And (WSAGetLastError <> WSAEWOULDBLOCK) And State(States.Sending) Then
        Dim sendActual As Long
        Dim sendExpect As Long
        Static sendFlag As Long
        Dim sendStream As New Stream

        If ((sendStream.Length + mSendData.Length) > mSendRate) Then
            sendExpect = mSendRate
        Else
            sendExpect = mSendData.Length
        End If

        If (mSendData.Length > 0) Then
            sendStream.Clone mSendData, 0, sendExpect
        End If

        If (sendStream.Length > 0) Then
            State(States.Sending) = True
            sendActual = modSockets.SocketSendLngPtr(Handle, ByVal sendStream.Address, sendExpect, 0)

            If (sendActual > 0) And (sendExpect > 0) Then
                mSendData.Push sendActual
                sendStream.Push sendActual
                SendNextPart = -CInt((sendExpect - sendActual) = 0)

            End If
        End If


        Set sendStream = Nothing
    Else
        SendNextPart = 1
    End If
    AutoRatedEvent AutoRated.SendWidth
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to send data."
End Function

Friend Function ReadNextPart() As Long ' _
Preforms the activity of the winsock reading to the read buffer and returns 1 if all incoming data was read, otherwise 0.
On Error GoTo sockerror
    If mReadData Is Nothing Then Exit Function
    AutoRatedEvent AutoRated.ReadCalls
    If (WSAGetLastError <> WSAEWOULDBLOCK) And State(States.Reading) Then
        Dim readActual As Long
        Dim readExpect As Long
        Static readFlag As Long
        Static readStream As New Stream
        readExpect = RecvRate
        If readStream.Length = 0 Then

            readStream.Length = readExpect
        Else
            readExpect = mReadData.Length
            mReadData.Length = readStream.Length + readExpect
            mReadData.Placeat readStream.Partial(0, readStream.Length), readExpect, readStream.Length

            readExpect = (RecvRate - readStream.Length)
            readStream.Length = readStream.Length + readExpect
        End If


        readActual = modSockets.SocketRecvLngPtr(Handle, ByVal mReadData.Address, readExpect, MSG_PEEK)
        If readExpect >= 0 Or (WSAGetLastError <> WSAEWOULDBLOCK) Then

            If readActual > 0 Then readExpect = readActual

            readActual = modSockets.SocketRecvLngPtr(Handle, ByVal readStream.Address + (readStream.Length - readExpect), readExpect, 0)
            If (readExpect > 0) And (readActual > 0) Then
                readExpect = mReadData.Length
                mReadData.Length = readActual + readExpect
                mReadData.Placeat readStream.Partial(0, readActual), readExpect, readActual

                readStream.Length = readActual
                readStream.Push readActual

            End If

        End If
        Set readStream = Nothing

    End If

    ReadNextPart = -CInt(WSAGetLastError <> WSAEWOULDBLOCK)

    AutoRatedEvent AutoRated.ReadWidth
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to read data."
End Function


Friend Function Incoming() As Boolean
    State(States.Request) = True
    Dim rc As Long
    Dim sk As Long
    If (Not getsockopt(Handle, SOL_SOCKET, SO_ACCEPTCONN, ByVal VarPtr(sk), 4)) Then
        If (sk = 1) And State(States.Binded) Then
            rc = SocketAccept(Handle, mAddress, sockaddr_size)
            If (Not (rc = Socket_ERROR)) Then
                Dim nSock As Long
                nSock = rc
                RaiseEvent Connection(nSock)
                If (nSock <> Socket_ERROR) Then
                    Incoming = True
                End If
            End If
        End If
    End If
    State(States.Closing) = False
    State(States.Request) = False
End Function

Friend Function SetEvents(Optional ByVal SetOpt As Boolean = True, Optional ByVal FD_CONST As Long = FD_EVENTS, Optional ByVal sck As Long = -1) As Long
On Error GoTo sockerror

        If sck = -1 Then sck = Handle
        If SetOpt Then SetEvents = SetOptions((FD_CONST = FD_LISTEN), sck)
        SetEvents = SetEvents And WSAAsyncSelect(sck, modWinProc.hwnd, WM_WINSOCK, FD_CONST)

Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to capture events."
End Function



Private Sub Adjust(ByRef adBearing As Single, ByRef adJumping As Single, ByRef adLatency As Single, ByRef adElapsed As Single, ByRef rate As Long)
    If Not ((mTransmission And TextMessages) = TextMessages) Then
        If adLatency < adElapsed And adBearing < 0 Then
            UpRate rate, adJumping

            If adBearing > adElapsed - adLatency Then
                adJumping = adJumping \ 2
                adBearing = -adBearing
            ElseIf adBearing < adElapsed - adLatency Then
                adBearing = Abs(adElapsed - adLatency)
                adJumping = adJumping * 2
            End If

            If adJumping <= 32 Then
                adJumping = 32
            End If
            If adJumping >= rate Then
                adJumping = rate
            End If

        ElseIf adLatency > adElapsed And adBearing > 0 Then

            If adBearing > adLatency - adElapsed Then
                adJumping = adJumping \ 2
                adBearing = -adBearing
            ElseIf adBearing < adLatency - adElapsed Then
                adBearing = Abs(adElapsed - adLatency)
                adJumping = adJumping * 2
            End If

            If adJumping <= 32 Then
                adJumping = 32 * 2
            End If
            If adJumping >= rate Then
                adJumping = rate / 2
            End If

        ElseIf ((Not (adBearing = adElapsed)) And (Not (adBearing = adLatency))) Or adBearing = 0 Then
            adBearing = -Abs(adLatency - adElapsed)
            If adBearing = 0 Then adBearing = adLatency - adElapsed
        End If

        If adBearing > 0 Then
            UpRate rate, adJumping
            adBearing = -adBearing
        ElseIf adBearing < 0 Then
            UpRate rate, adJumping
            adBearing = -adBearing
        End If
    End If
End Sub

Private Sub SetValidatingMode()
    If mSocket > 0 Then
        mSocket = (-mSocket) - 1
    End If
End Sub

Private Function UnsetValidatingMode(ByVal OnHandle As Long, Optional ByVal ErrMsg As String = "") As Boolean
    If (mSocket < -1) And (((-mSocket) - 1) = OnHandle) Then
        'not handled (else accepted, or decline closed)
        mSocket = (-mSocket) - 1
        If ErrMsg <> "" Then
            RaiseEvent Error(8, "Sockets", ErrMsg)
        Else
            UnsetValidatingMode = True
        End If
    End If
End Function

Private Function Whois(ByVal Handle As Long) As String
    Dim col As VBA.Collection
    Set col = IPAddress(Handle)
    If col.count > 0 Then
        Whois = col.Item(1)
    Else
        Whois = "#INVALID#"
    End If
End Function

Private Function SetOptions(Optional ByVal DoAccept As Boolean = False, Optional ByVal sck As Long = -1) As Long
On Error GoTo sockerror
    If sck = -1 Then sck = Handle
    Dim li As LingerType
    Dim rc As Boolean
    rc = True

    rc = rc And (ioctlsocket(Handle, FIONBIO, 0) = 0)
    rc = rc And (setsockopt(sck, SOL_SOCKET, SO_REUSEADDR, ByVal VarPtr(SO_ON), 4&) = 0)
    rc = rc And (setsockopt(sck, SOL_SOCKET, SO_OOBINLINE, ByVal VarPtr(SO_ON), 4&) = 0)
    If ((CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages) Or (mTimeOut = 0))) And _
      ((Not (CBool((mTransmission And Direction.TextMessages) = Direction.TextMessages)) Or (mTimeOut > 0))) Then
        rc = rc And (setsockopt(sck, SOL_SOCKET, SO_KEEPALIVE, ByVal VarPtr(SO_ON), 4&) = 0)
    Else
        rc = rc And (setsockopt(sck, SOL_SOCKET, SO_KEEPALIVE, ByVal VarPtr(SO_OFF), 4&) = 0)
    End If
    rc = rc And (setsockopt(sck, IPPROTO_TCP, TCP_NODELAY, ByVal VarPtr(SO_OFF), 4&) = 0)
    rc = rc And (setsockopt(sck, SOL_SOCKET, SO_LINGER, ByVal VarPtr(li.l_onoff), 8&) = 0)
    If DoAccept Then
        rc = rc And (setsockopt(sck, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, ByVal VarPtr(SO_ON), 4&) = 0)
    Else
        rc = rc And (setsockopt(sck, SOL_SOCKET, SO_CONDITIONAL_ACCEPT, ByVal VarPtr(SO_OFF), 4&) = 0)
    End If

    SetOptions = rc And SetReceive(sck) And SetSending(sck)
Exit Function
sockerror:
    Err.Clear
    On Error GoTo 0
    Err.Raise 8, "Socket", "Error attempting to set socket options."
End Function

Private Function addFlag(ByRef Flags As Long, ByVal newFlag As Long) As Long
    If Not ((Flags And newFlag) = newFlag) Then
        Flags = Flags + newFlag
    End If
    addFlag = Flags
End Function

Private Function delFlag(ByRef Flags As Long, ByVal theFlag As Long) As Long
    If ((Flags And theFlag) = theFlag) Then
        Flags = Flags - theFlag
    End If
    delFlag = Flags
End Function

Private Function hasFlag(ByVal Flags As Long, ByVal theFlag As Long) As Boolean
    hasFlag = ((Flags And theFlag) = theFlag)
End Function

Private Function GetReadFlag(Optional ByVal newFlags As Long = 0) As Long
    If (newFlags <> 0) And (newFlags <> MSG_PEEK) Then
        If Not hasFlag(readFlag, newFlags) Then
            addFlag readFlag, newFlags
        Else
            delFlag readFlag, newFlags
        End If
    End If
    If newFlags <> 0 Then
        GetReadFlag = newFlags
    Else
        GetReadFlag = readFlag
    End If
    delFlag readFlag, MSG_PEEK
End Function

Private Function GetSendFlag(Optional ByVal newFlags As Long = 0) As Long
    If (newFlags <> 0) And (newFlags <> MSG_DONTROUTE) Then
        If Not hasFlag(sendFlag, newFlags) Then
            addFlag sendFlag, newFlags
        Else
            delFlag sendFlag, newFlags
        End If
    End If
    If newFlags <> 0 Then
        GetSendFlag = newFlags
    Else
        GetSendFlag = sendFlag
    End If
    delFlag sendFlag, MSG_DONTROUTE
End Function
Private Sub UpRate(ByRef rate As Long, ByRef jumping As Single)
    If rate + jumping <= modBitValue.IntBound Then
        rate = rate + jumping
    End If
End Sub
Private Sub DownRate(ByRef rate As Long, ByRef jumping As Single)
    If rate - jumping >= 1024 Then
        rate = rate - jumping
    End If
End Sub
Private Sub DebugEvent(ByVal txt As String)
    If ((mTransmission And NormalStream) = NormalStream) Then
        Debug.Print "DATA " & txt
    Else
        Debug.Print "TEXT " & txt
    End If
End Sub
Private Sub DataValidating(ByVal wParam As Long)
        SetValidatingMode
        RaiseEvent Connection(Handle)
        UnsetValidatingMode Handle, "Not validating remote connections."
        
    
'    If Not (WhosIs(wParam) = RemoteIP) Then 'Whois(inet_ntoa(lParam))
'        SetValidatingMode
'        RaiseEvent Connection(Handle)
'        UnsetValidatingMode Handle, "he remote address did not match the ip address for the connection."
'    Else
'        If Not (RemoteIP = WhosIs(Handle) And WhoAmI(Handle) = LocalIP) Then ' ResolveIP(LocalHost) Then
'            SetValidatingMode
'            RaiseEvent Connection(Handle)
'            UnsetValidatingMode Handle, "Either remote or local socket handle did not match ip address for the connection."
'
'        End If
'    End If
End Sub



Private Function CloseConnection() As Boolean

    If Ready And (State(States.Opening) = False) Then
        If (State(States.Closing) = True) And (State(States.Closed) = False) Then

            State(States.Closing) = True
            State(States.Binded) = False
            State(States.Opened) = False

            socketclose Handle

            DebugEvent "Disconnected"
            RaiseEvent Disconnected
            UnregisterSocket Me
            CloseConnection = True
        End If
    End If

End Function


Private Sub DebugState()
    Debug.Print Handle & " ";
    If State(Binded) Then Debug.Print "Binded ";
    If State(Request) Then Debug.Print "Request ";
    If State(Opening) Then Debug.Print "Opening ";
    If State(Opened) Then Debug.Print "Opened ";
    If State(Sending) Then Debug.Print "Sending ";
    If State(Reading) Then Debug.Print "Reading ";
    If State(Closing) Then Debug.Print "Closing ";
    If State(Closed) Then Debug.Print "Closed ";
    Debug.Print
End Sub

