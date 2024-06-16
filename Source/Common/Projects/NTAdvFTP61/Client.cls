VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Public Enum RateTypes
    HardDrive = 0
    Download = 1
    Upload = 2
End Enum

Public Enum ListSystems
    none = 0
    Wins = -1
    Unix = -2
End Enum

Public Enum ErrorReturns
    Wait = -1
    none = 0
    Last = 1
End Enum

Public Enum TransferModes
    none = -1
    Binary = 0
    ASCII = 1
End Enum

Public Enum ConnectionModes
    Passive = 0
    active = 1
End Enum

Public Enum MessageTypes
    Outgoing = 1
    Incoming = 2
    Incorrect = 3
End Enum

Public Enum ProgressTypes
    FileListing = 0
    AllocatingFile = 1
    PositioningFile = 2
    TransferingFile = 3
End Enum

Public Enum AllocateSides
    none = 0
    Client = 1
    Remote = 2
End Enum

Public Enum ConnectedStates
    NotState = 0
    AllStates = -1
    LocalHardDisk = 1
    SocketNetwork = 2
    LoggingInto = 4
    FullyRemoting = 8
    CommandRequest = 16
    ListingToFile = 32
    LocalToLocal = 64
    LocalToRemote = 128
    RemoteToLocal = 256
    DataStreaming = 512
End Enum

Private Enum GatherTypes
    none = 0
    Banner = 1
    MOTD = 2
    Help = 3
    Stat = 4
End Enum

Private Enum ErrEventType
    ReturnError = 0
    SetError = 1
    EventError = 3
    RaiseError = 5
End Enum

Private ftpType As URLTypes
Private ftpLogData As String
Private ftpLastError As String
Private ftpLogBytes As Long
Private ftpTimeOut As Long
Private ftpConnectionMode As String
Private ftpURL As String
Private ftpServer As String
Private ftpPort As Long
Private ftpAccount As String
Private ftpUsername As String
Private ftpPassword As String
Private ftpFolder As String
Private ftpSystemInfo As String
Private ftpBannerMessage As String
Private ftpMessageOfTheDay As String
Private ftpHelpInformation As String
Private ftpStatInfo As String
Private ftpDataPortRange As String
Private fttpModePole As Boolean
Private ftpLargeFileMode As Boolean
Private ftpTransferRates(2) As Long
Private ftpAllocation As AllocateSides
Private ftpGatherType As GatherTypes
Private ftpGettingInfo As Boolean
Private ftpNetAdapter As Integer
Private ftpTransferMode As TransferModes
Private ftpServerUsesCRLF As Boolean
Private ftpAssumeLineFeed As Boolean
Private ftpCancelThread As Boolean
Private ftpFileSize As Double
Private ftpSendIncrement As Double
Private ftpSendCountKnot As Double
Private ftpSendBytesLeft As Double
Private ftpSendPacketLen As Double
Private ftpResumeByte As Double

Private ftpWriteFileNum As Integer
Private ftpWriteTextStream As Scripting.TextStream

Private ftpReadFileNum As Integer
Private ftpReadTextStream As Scripting.TextStream

Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
Public Event LogMessage(ByVal MessageType As MessageTypes, ByVal AddedText As String)
Public Event ItemListing(ByVal ItemName As String, ByVal ItemSize As String, ByVal ItemDate As String, ByVal ItemAccess As String)
Public Event DataProgress(ByVal ProgressType As ProgressTypes, ByVal ReceivedBytes As Double) ' _
Event that is called during a data socket transfering data and has progressed.
Attribute DataProgress.VB_Description = "Event that is called during a data socket transfering data and has progressed."
Public Event DataComplete(ByVal ProgressType As ProgressTypes) ' _
Event that is called when the data socket has finished transfering and has closed. ' _
Attribute DataComplete.VB_Description = "Event that is called when the data socket has finished transfering and has closed. ' _"

Private myObj As Client
Private myServerObj As Client
Private myFso As Scripting.FileSystemObject

Private WithEvents mTCP As Socket
Attribute mTCP.VB_VarHelpID = -1
Private WithEvents mData As Socket
Attribute mData.VB_VarHelpID = -1


Private ftpEventError As Long
Private ftpEventReason As String

Private ftpConnected As ConnectedStates
Private ftpNextCommand As String
Private ftpBuffCommand As String

Private ftpRaiseFTTPComplete As Client

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private ftpElapse As Single

Public Property Get GetLastError() As String
    GetLastError = SetLastError(ReturnError)
'    Static lstErr As String
'    If (ftpEventError = ErrorReturns.Last) Then
'        lstErr = ftpEventReason
'    ElseIf Err.Description <> "" Then
'        lstErr = Err.Description
'    Else
'        lstErr = ""
'    End If
'
'    GetLastError = lstErr
End Property

Private Function SetLastError(ByVal ErrAction As ErrEventType, Optional ByVal dwErrCode As Long = 0, Optional ByVal strMsg As String) As String
    Static lstErr As String
    
    If ErrAction And SetError = SetError Then
        Bit(dwErrCode, Bit29) = 1
        lstErr = strMsg
        modMain.SetLastError dwErrCode
    End If
    
    If ErrAction And EventError = EventError Then
        RaiseEvent Error(dwErrCode, App.EXEName, strMsg)
    ElseIf ErrAction And RaiseError = RaiseError Then
        Err.Raise dwErrCode, App.EXEName, strMsg
    End If
    SetLastError = lstErr

End Function


'Private Function SetError(Optional ByVal errMsg As String = "", Optional ByVal errRaise As Long = 0) As String
'    Static lastErr As String
'    If errMsg <> "" Then
'        lstErr = errMsg
'        ftpEventError = ErrorReturns.None
'        If errRaise <> 0 Then
'            Err.Raise errRaise, "NTAdvFTP61", errMsg
'        End If
'    Else
'        If (ftpEventError = ErrorReturns.Last) Then
'            lstErr = ftpEventReason
'            ftpEventError = ErrorReturns.None
'        End If
'    End If
'    SetError = lstErr
'End Function


Public Property Get ImplicitSSL() As Boolean ' _
Gets or sets whether Secure Sockets Layer is implied on all connections to establish before any communications occurs.
    ImplicitSSL = mTCP.SSL
End Property
Public Property Let ImplicitSSL(ByVal RHS As Boolean)
    If (Not mTCP.Connected) And (Not mData.Connected) Then
        mTCP.SSL = RHS
        mData.SSL = RHS
    Else
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'Err.Raise 8, "SsL", "Can not set this property while connected."
    End If
End Property

Public Property Get AutoRate() As Boolean ' _
Gets or sets whether or not the transfer rate determined by buffer sizes will be automatically adapting.
Attribute AutoRate.VB_Description = "Gets or sets whether or not the transfer rate determined by buffer sizes will be automatically adapting."
    AutoRate = mData.AutoRate
End Property
Public Property Let AutoRate(ByVal RHS As Boolean)
    mData.AutoRate = RHS
End Property
Public Property Get Allocation() As AllocateSides ' _
Gets or sets the level of allocation consideration and FileProgress eventing, exceptions apply with the TransferFile function.
Attribute Allocation.VB_Description = "Gets or sets the level of allocation consideration and FileProgress eventing, exceptions apply with the TransferFile function."
    Allocation = ftpAllocation
End Property
Public Property Let Allocation(ByVal newVal As AllocateSides)
    If ((newVal And AllocateSides.Client) = AllocateSides.Client) And _
        (Not ((ftpAllocation And AllocateSides.Client) = AllocateSides.Client)) Then
        ftpAllocation = ftpAllocation + AllocateSides.Client
    ElseIf (Not ((newVal And AllocateSides.Client) = AllocateSides.Client)) And _
        ((ftpAllocation And AllocateSides.Client) = AllocateSides.Client) Then
        ftpAllocation = ftpAllocation - AllocateSides.Client
    End If
    If ((newVal And AllocateSides.Remote) = AllocateSides.Remote) And _
        (Not ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote)) Then
        ftpAllocation = ftpAllocation + AllocateSides.Remote
    ElseIf (Not ((newVal And AllocateSides.Remote) = AllocateSides.Remote)) And _
        ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote) Then
        ftpAllocation = ftpAllocation - AllocateSides.Remote
    End If
    
End Property

Public Property Get PauseTransfers() As Boolean
    PauseTransfers = ((mTCP.Transmission And SocketPaused) = SocketPaused)
End Property
Public Property Let PauseTransfers(ByVal newVal As Boolean)
    If newVal Then
        If Not ((mTCP.Transmission And SocketPaused) = SocketPaused) Then
            mTCP.Transmission = mTCP.Transmission + SocketPaused
        End If
        If Not ((mData.Transmission And SocketPaused) = SocketPaused) Then
            mData.Transmission = mData.Transmission + SocketPaused
        End If
    Else
        If ((mTCP.Transmission And SocketPaused) = SocketPaused) Then
            mTCP.Transmission = mTCP.Transmission - SocketPaused
        End If
        If ((mData.Transmission And SocketPaused) = SocketPaused) Then
            mData.Transmission = mData.Transmission - SocketPaused
        End If
    End If
End Property

Public Property Get PauseOnStandBy() As Boolean
    PauseOnStandBy = ((mTCP.Transmission And StandByPause) = StandByPause)
End Property
Public Property Let PauseOnStandBy(ByVal newVal As Boolean)
    If newVal Then
        If Not ((mTCP.Transmission And StandByPause) = StandByPause) Then
            mTCP.Transmission = mTCP.Transmission + StandByPause
        End If
        If Not ((mData.Transmission And StandByPause) = StandByPause) Then
            mData.Transmission = mData.Transmission + StandByPause
        End If
    Else
        If ((mTCP.Transmission And StandByPause) = StandByPause) Then
            mTCP.Transmission = mTCP.Transmission - StandByPause
        End If
        If ((mData.Transmission And StandByPause) = StandByPause) Then
            mData.Transmission = mData.Transmission - StandByPause
        End If
    End If
End Property

Friend Property Get ServerUsesCRLF() As Boolean
    ServerUsesCRLF = ftpServerUsesCRLF
End Property

Public Property Get AssumeLineFeed() As Boolean ' _
Gets or sets how the text file line feeds will be handled, when true, all line feeds are changed to match the local system with carriage return line feed.  Transfering to a remote system will be determined by the system type.  When false, no change to li
Attribute AssumeLineFeed.VB_Description = "Gets or sets how the text file line feeds will be handled, when true, all line feeds are changed to match the local system with carriage return line feed.  Transfering to a remote system will be determined by the system type.  When false, no change to li"
    AssumeLineFeed = ftpAssumeLineFeed
End Property

Public Property Let AssumeLineFeed(ByVal newVal As Boolean)
    ftpAssumeLineFeed = newVal
End Property

Friend Property Let SetServerUsesCRLF(ByVal ftpListing As String)
    If ((InStr(UCase(ftpSystemInfo), "UNIX") > 0) Or (InStr(UCase(ftpSystemInfo), "LINUX") > 0) Or (InStr(UCase(ftpSystemInfo), "MAC") > 0)) Or _
            ((InStr(UCase(ftpListing), "TOTAL") > 0) And (InStr(UCase(ftpListing), "<DIR>") <= 0)) Then
        ftpServerUsesCRLF = False
    ElseIf (InStr(UCase(ftpSystemInfo), "WIN") > 0) And _
            ((InStr(UCase(ftpListing), "TOTAL") <= 0) And (InStr(UCase(ftpListing), "<DIR>") <= 0)) Then
        ftpServerUsesCRLF = True
    End If
End Property

Public Property Get TransferRates(ByVal RateType As RateTypes) As Long
    TransferRates = ftpTransferRates(RateType)
End Property

Public Property Let TransferRates(ByVal RateType As RateTypes, ByVal newVal As Long)
    ftpTransferRates(RateType) = newVal
End Property

Public Property Get LargeFileMode() As Boolean
    LargeFileMode = ftpLargeFileMode
End Property
Public Property Let LargeFileMode(ByVal newVal As Boolean)
    ftpLargeFileMode = newVal
End Property

Public Property Get AllAdapters() As Collection ' _
A collection that contains all the IP addresses found on the local system network adapters.
Attribute AllAdapters.VB_Description = "A collection that contains all the IP addresses found on the local system network adapters."
    Set AllAdapters = GetPortIP
End Property
Public Property Get AdapterCount() As Long ' _
Returns the number of network IP addresses (locally based hardware) in the AllAdapters collection.
Attribute AdapterCount.VB_Description = "Returns the number of network IP addresses (locally based hardware) in the AllAdapters collection."
    AdapterCount = GetPortIP.count
End Property

Public Property Get NetAdapter() As Variant ' _
Gets or sets the IP address to be used in the PORT command, this property must be a IP with in the AllAdapters collection, or a numeric index to an element of the AllAdapters collection.
Attribute NetAdapter.VB_Description = "Gets or sets the IP address to be used in the PORT command, this property must be a IP with in the AllAdapters collection, or a numeric index to an element of the AllAdapters collection."
    NetAdapter = ftpNetAdapter
End Property
Public Property Let NetAdapter(ByVal newVal As Variant)
    If IsNumeric(newVal) And (CountWord(newVal, ".") = 0) Then
        If newVal > 0 And newVal <= GetPortIP.count Then
            ftpNetAdapter = newVal
        Else
            SetLastError RaiseError, 10050, "Network adapter does not exist."
            'SetLastError RaiseError,8, "Network adapter does not exist."
        End If
    Else
        Dim tmp As Collection
        Dim cnt As Long
        Dim use As Long
        Set tmp = GetPortIP
        use = 1
        For cnt = 1 To tmp.count
            If LCase(Trim(tmp.Item(cnt))) = LCase(Trim(newVal)) Then
                use = cnt
                Exit For
            End If
        Next
        Set tmp = Nothing
        If use > 0 Then
            SetLastError RaiseError, 10050, "Network adapter does not exist."
            'SetLastError RaiseError,8, "Network adapter does not exist."
        Else
            ftpNetAdapter = use
        End If
    End If
End Property


Public Property Get ConnectedState(Optional ByVal check As ConnectedStates = AllStates) As Boolean ' _
Read-only property that returns the state of the ConnectedStates that is passed to the property, returning either true or false to express the state requested.
Attribute ConnectedState.VB_Description = "Read-only property that returns the state of the ConnectedStates that is passed to the property, returning either true or false to express the state requested."
    If check = AllStates Then
        ConnectedState = (ftpConnected <> NotState)
    ElseIf check = NotState Then
        If (ftpConnected And check) = ConnectedStates.NotState Then
            ConnectedState = True
        Else
            ConnectedState = False
        End If
    Else
        ConnectedState = ((ftpConnected And (check)) = check)
    End If
End Property
Friend Property Let ConnectedState(Optional ByVal check As ConnectedStates = AllStates, ByVal RHS As Boolean)
    If RHS Then
        Select Case check
            Case ListingToFile, LocalToLocal, LocalToRemote, RemoteToLocal
                If (ftpConnected And ListingToFile) = ListingToFile Then
                    ftpConnected = ftpConnected - ListingToFile
                End If
                If (ftpConnected And LocalToLocal) = LocalToLocal Then
                    ftpConnected = ftpConnected - LocalToLocal
                End If
                If (ftpConnected And LocalToRemote) = LocalToRemote Then
                    ftpConnected = ftpConnected - LocalToRemote
                End If
                If (ftpConnected And RemoteToLocal) = RemoteToLocal Then
                    ftpConnected = ftpConnected - RemoteToLocal
                End If
        End Select
    End If

    If (ftpConnected And (check)) And (Not RHS) Then
        ftpConnected = ftpConnected - (check)
    ElseIf (Not (ftpConnected And (check))) And RHS Then
        ftpConnected = ftpConnected Or (check)
    End If

End Property

Public Property Get URLType() As Integer
    URLType = ftpType
End Property

Public Property Get URL() As String
    URL = ftpURL
End Property

Public Property Let URL(ByVal newVal As String)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    
    Dim nURL As New URL
    
    ImplicitSSL = nURL.GetSSL(newVal)
    
    ftpType = nURL.GetType(newVal)
    ftpServer = nURL.GetServer(newVal)
    ftpFolder = nURL.GetFolder(newVal)
    If nURL.GetPort(newVal) <> 0 Then ftpPort = nURL.GetPort(newVal)
    If nURL.GetUserName(newVal) <> "" Then ftpUsername = nURL.GetUserName(newVal)
    If nURL.GetPassword(newVal) <> "" Then ftpPassword = nURL.GetPassword(newVal)
    
    Set nURL = Nothing
    
    ftpURL = newVal
End Property

Public Property Get RemoteSystem() As String
    RemoteSystem = ftpSystemInfo
End Property

Public Property Get MessageOfTheDay() As String
    MessageOfTheDay = ftpMessageOfTheDay
End Property

Public Property Get BannerMessage() As String ' _
Returns the FTP remote systems Banner message, this message is displayed in the protocols messages before login occurs and just after connection.
Attribute BannerMessage.VB_Description = "Returns the FTP remote systems Banner message, this message is displayed in the protocols messages before login occurs and just after connection."
    BannerMessage = ftpBannerMessage
End Property

Public Property Get Transfering() As Boolean
    Dim ret As Boolean
    If Not myObj Is Nothing Then
        ret = ret Or CBool(myObj.Self("Transfering"))
    End If
    If Not myServerObj Is Nothing Then
        ret = ret Or CBool(myServerObj.Self("Transfering"))
    End If
    Transfering = ret Or CBool(Self("Transfering"))
End Property

Public Property Get Server() As String
    Server = ftpServer
End Property

Public Property Let Server(ByVal newVal As String)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    ftpServer = newVal
End Property

Public Property Get Port() As Long ' _
Gets or sets the remote machine's port for connection to, this by default is 21 and is valid as any port listening for incoming connections to provide the FTP protocol service.
Attribute Port.VB_Description = "Gets or sets the remote machine's port for connection to, this by default is 21 and is valid as any port listening for incoming connections to provide the FTP protocol service."
    Port = ftpPort
End Property

Public Property Let Port(ByVal newVal As Long)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    ftpPort = newVal
End Property

Public Property Get Account() As String ' _
Gets or sets the account information for logging on to the FTP remote system.  This is used in conjunction with the Username property and not similar.  Account is sent before the Username in which case Passowrds are ignored.
Attribute Account.VB_Description = "Gets or sets the account information for logging on to the FTP remote system.  This is used in conjunction with the Username property and not similar.  Account is sent before the Username in which case Passowrds are ignored."
    Account = ftpAccount
End Property

Public Property Let Account(ByVal newVal As String)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    ftpAccount = newVal
    ftpUsername = ""
    ftpPassword = ""
End Property

Public Sub Reinitialize()

    If Not ConnectedState() Then
        
        SetLastError RaiseError, 8, "REIN: Not logged in to a remote system."
    End If
    SendFTPCommand "REIN"

End Sub

Public Property Get Username() As String
    Username = ftpUsername
End Property

Public Property Let Username(ByVal newVal As String)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    ftpUsername = newVal
    ftpAccount = ""
End Property

Public Property Get Password() As String
    Password = ftpPassword
End Property

Public Property Let Password(ByVal newVal As String)
    If ConnectedState() Then
        SetLastError RaiseError, 40006, "Can not set this property while connected."
        'SetLastError RaiseError,8, "Can not set this value while connected."
    End If
    ftpPassword = newVal
    ftpAccount = ""
End Property

Public Property Get Folder() As String
    Folder = ftpFolder
End Property

Public Property Get ConnectionMode() As Variant ' _
Gets or sets the connection mode to be used in future ListContents or TransferFile commands.  Valid values are of ConnectionModes.
Attribute ConnectionMode.VB_Description = "Gets or sets the connection mode to be used in future ListContents or TransferFile commands.  Valid values are of ConnectionModes."
    ConnectionMode = UCase(Trim(ftpConnectionMode))
End Property

Public Property Let ConnectionMode(ByVal newVal As Variant)
    If TypeName(ftpConnectionMode) = "String" Then
        ftpConnectionMode = UCase(Trim(newVal))
    End If
    Select Case CStr(ftpConnectionMode)
        Case CStr(Passive)
            ftpConnectionMode = "PASV"
        Case CStr(active)
            ftpConnectionMode = "PORT"
    End Select
End Property

Public Property Get timeout() As Long
    timeout = ftpTimeOut
End Property

Public Property Let timeout(ByVal newVal As Long)
    ftpTimeOut = newVal
    mTCP.timeout = newVal
    mData.timeout = newVal
End Property

Public Property Get LogBytes() As Long
    LogBytes = ftpLogBytes
End Property

Public Property Let LogBytes(ByVal newVal As Long)
    If newVal < 0 Then
        SetLastError RaiseError, 8, "Invalid log byte size (must be a positive number, or zero for no limit)."
    Else
        ftpLogBytes = newVal
    End If
End Property

Public Property Get LogData() As String
    LogData = ftpLogData
End Property

Public Property Let DataPortRange(ByVal newVal As String) ' _
Gets or sets the port ranges used in data connections when the ConnectionMode is Active.  Valid value formatting uses of port numbers are indicated by consisting of the low port number, dash, then the high port number and every port inbetween.
Attribute DataPortRange.VB_Description = "Gets or sets the port ranges used in data connections when the ConnectionMode is Active.  Valid value formatting uses of port numbers are indicated by consisting of the low port number, dash, then the high port number and every port inbetween."
    Dim sHi As String
    Dim sLow As String
    
    newVal = Replace(Replace(Replace(Replace(newVal, " ", ""), vbTab, ""), vbCr, ""), vbLf, "")
    sHi = NextArg(newVal, "-")
    sLow = RemoveArg(newVal, "-")
    If ((Not (InStr(newVal, "-") > 0)) Or (Not IsNumeric(sHi)) Or (Not IsNumeric(sLow)) Or (InStr(sLow, ".") > 0) Or (InStr(sHi, ".") > 0) Or (InStr(sLow, "-") > 0) Or (InStr(sHi, "-") > 0)) Then
        SetLastError RaiseError, 8, "Invalid data port range; (Must be two whole numbers seperated by a single dash greater then zero and less then 65535)"
    End If
    If Not (((CLng(sHi) > 0) And (CLng(sLow) > 0)) And ((CLng(sLow) <= 65535) And (CLng(sHi) <= 65535))) Then
        SetLastError RaiseError, 8, "Invalid data port range; (Must be two whole numbers seperated by a single dash greater then zero and less then 65535)"
    End If
    If (CLng(sHi) < CLng(sLow)) Then
        Dim swp As String
        swp = sHi
        sHi = sLow
        sLow = swp
    End If
    ftpDataPortRange = sLow & "-" & sHi
End Property

Public Property Get DataPortRange() As String
    DataPortRange = ftpDataPortRange
End Property

Private Property Get DataPortRangeLow() As Integer
    DataPortRangeLow = CInt(NextArg(ftpDataPortRange, "-"))
End Property

Private Property Get DataPortRangeHi() As Integer
    DataPortRangeHi = CInt(RemoveArg(ftpDataPortRange, "-"))
End Property

Private Sub ResetDefaultValues()

    ftpCancelThread = False
    ftpGettingInfo = False
    ftpLogData = ""
    ftpConnected = NotState
    ftpEventError = ErrorReturns.none
    ftpEventReason = GetLastError
    ftpMessageOfTheDay = ""
    ftpBannerMessage = ""
    ftpNextCommand = ""
    Set ftpRaiseFTTPComplete = Nothing

End Sub

Private Sub Class_Initialize()
            
    Set mTCP = New NTAdvFTP61.Socket
    Set mData = New NTAdvFTP61.Socket

    mTCP.SSL = False
    mData.SSL = False
    
    
    mTCP.Transmission = Direction.TextMessages
    
    LogBytes = 100000
    timeout = 30
    ConnectionMode = "PASV"
    DataPortRange = "3000-6000"
    fttpModePole = True
    ftpNetAdapter = 1
    ftpServer = "127.0.0.1"
    ResetDefaultValues
    
    ftpLargeFileMode = True
    
    ftpTransferRates(RateTypes.HardDrive) = ftpLocalSize
    ftpTransferRates(RateTypes.Download) = ftpBufferSize
    ftpTransferRates(RateTypes.Upload) = ftpPacketSize
    
End Sub

Private Sub Class_Terminate()
    
    Disconnect
    
    Set mTCP = Nothing
    Set mData = Nothing
    
End Sub

Private Sub SendDataFTPCommand(ByVal ConnectMode As String, ByVal ftpCommand As String, Optional ByVal ftpRestCommand As String = "")
    
    Select Case ConnectionMode
        Case "PORT"
            Static LinkPort As Integer
            
            If LinkPort = 0 Then
                Randomize
                LinkPort = Int(DataPortRangeLow + ((DataPortRangeHi - DataPortRangeLow) * Rnd))
            Else
                LinkPort = LinkPort + 1
                If LinkPort > DataPortRangeHi Then LinkPort = DataPortRangeLow
            End If
            Do While (Not mTCP.Ready) And (Not mData.Ready)
                'DoEvents
                DoLoop
            Loop
           
            If mData.Connected Or mData.Listening Then mData.Disconnect
            Do Until mData.Listen(GetPortIP.Item(ftpNetAdapter), LinkPort)
                
                If LinkPort > 0 Then
                    LinkPort = LinkPort + 1
                ElseIf LinkPort < 0 Then
                    LinkPort = LinkPort - 1
                End If
                
                If Abs(LinkPort) > DataPortRangeHi And LinkPort > 0 Then
                    LinkPort = -DataPortRangeLow
                ElseIf Abs(LinkPort) > DataPortRangeHi And LinkPort < 0 Then
                    SetLastError RaiseError, 8, "Unable to find a valid listening port."
                End If
                mData.Disconnect

            Loop

            'Debug.Print "LINKPORT " & LinkPort
            
'            Dim tmpPort As String
'            tmpPort = Padding(4, Hex(IntegerToUnsigned(LinkPort)), "0")
'
'            If ftpRestCommand <> "" Then
'                SendFTPCommand "PORT " & Replace(GetPortIP.Item(ftpNetAdapter), ".", ",") & "," & CByte(Val("&H" & Left(tmpPort, 2))) & "," & CByte(Val("&H" & Right(tmpPort, 2))) & vbCrLf & ftpRestCommand & vbCrLf & ftpCommand & vbCrLf
'            Else
'                SendFTPCommand "PORT " & Replace(GetPortIP.Item(ftpNetAdapter), ".", ",") & "," & CByte(Val("&H" & Left(tmpPort, 2))) & "," & CByte(Val("&H" & Right(tmpPort, 2))) & vbCrLf & ftpCommand & vbCrLf
'            End If
        
            'Debug.Print Val("&H" & Padding(2, Hex(CByte(Val("&H" & Left(tmpPort, 2)))), "0") & Padding(2, Hex(CByte(Val("&H" & Right(tmpPort, 2)))), "0"))

            Debug.Print LinkPort
            
            If ftpRestCommand <> "" Then
                SendFTPCommand "PORT " & Replace(GetPortIP.Item(ftpNetAdapter), ".", ",") & "," & CLng(LinkPort / 256) & "," & CLng(LinkPort Mod 256) & vbCrLf & ftpRestCommand & vbCrLf & ftpCommand & vbCrLf
            Else
                SendFTPCommand "PORT " & Replace(GetPortIP.Item(ftpNetAdapter), ".", ",") & "," & CLng(LinkPort / 256) & "," & CLng(LinkPort Mod 256) & vbCrLf & ftpCommand & vbCrLf
            End If
            
        Case "PASV"

            If ftpRestCommand <> "" Then
                SendFTPCommand "PASV" & vbCrLf & ftpRestCommand & vbCrLf & ftpCommand & vbCrLf
            Else
                SendFTPCommand "PASV" & vbCrLf & ftpCommand & vbCrLf
            End If
                    
    End Select
   
End Sub

Friend Function SendDataFTTPCommand(ByVal FTTPConnectMode As String) As String
    
    Select Case Left(FTTPConnectMode, 4)
        Case "PORT"
             
             SendFTPCommand FTTPConnectMode
     
        Case "PASV"
             
             Dim dataPort As String

             SendFTPCommand FTTPConnectMode
     
             dataPort = StrReverse(ftpEventReason)
             dataPort = Mid(dataPort, InStr(dataPort, ")") + 1)
             dataPort = StrReverse(Left(dataPort, InStr(dataPort, "(") - 1))
             
    End Select

    SendDataFTTPCommand = dataPort
   
End Function

Friend Function SendFTTPCommand(ByVal ftpCommand As String) As String
    
    SendFTPCommand ftpCommand
    
    SendFTTPCommand = ftpCommand

End Function


Public Sub Connect(Optional ByVal eURL As String = "") ' _
Connects to a remote system and sends the initial commands provided needing for a connection.  Server, Port, Account, Username, Password and Folder property will be handled address upon successful Connect.
Attribute Connect.VB_Description = "Connects to a remote system and sends the initial commands provided needing for a connection.  Server, Port, Account, Username, Password and Folder property will be handled address upon successful Connect."
    If mTCP.Connected Or (Not (ftpConnected = NotState)) Then
        SetLastError RaiseError, 8, "Already connected, you must use disconnect first."
    End If
    
    Dim nURL As New URL
    If eURL <> "" Then
        ftpType = nURL.GetType(eURL)
        ftpServer = nURL.GetServer(eURL)
        ftpFolder = nURL.GetFolder(eURL)
        If nURL.GetPort(eURL) <> 0 Then ftpPort = nURL.GetPort(eURL)
        If nURL.GetUserName(eURL) <> "" Then ftpUsername = nURL.GetUserName(eURL)
        If nURL.GetPassword(eURL) <> "" Then ftpPassword = nURL.GetPassword(eURL)
    
    ElseIf ftpURL <> "" Then
    
    Else
        SetLastError RaiseError, 8, "No URL specifyed."
    End If
    Set nURL = Nothing
    ResetDefaultValues
    
    If ftpType = URLTypes.ftp Then
        
        ftpEventError = ErrorReturns.Wait
        
        If Not mTCP.Connected Then
            ConnectedState(CommandRequest) = True
            ftpGatherType = GatherTypes.Banner

        End If
            
        ConnectedState(LoggingInto) = True

        If Trim(ftpAccount) <> "" Then
            ftpNextCommand = ftpNextCommand & "ACCT " & ftpAccount & vbCrLf
        Else
            ftpNextCommand = ftpNextCommand & "USER " & ftpUsername & vbCrLf
        End If
        
        mTCP.Connect ftpServer, ftpPort
    
        If ftpEventError = ErrorReturns.Last Then
            SetLastError RaiseError, 8, ftpEventReason
            'SetLastError RaiseError,8, ftpEventReason
        End If
    
    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
        
        If Not PathExists(ftpFolder) Then
            ResetDefaultValues
            SetLastError RaiseError, 53, "Path not found"
        End If
        
        ftpServer = "127.0.0.1"
        
        ConnectedState(LocalHardDisk) = True
    
    Else
        ResetDefaultValues
        SetLastError RaiseError, 8, "URL specifyed is of unkown type."
    End If
    
End Sub

Public Sub Disconnect() ' _
Disconnects the Client object from the connected remote system closing all open connections it is active of.  This must be used as well for non internet based locations when pairing, even if a local location or file path is used for the URL.
Attribute Disconnect.VB_Description = "Disconnects the Client object from the connected remote system closing all open connections it is active of.  This must be used as well for non internet based locations when pairing, even if a local location or file path is used for the URL."

    CancelTransfer

    mData.Disconnect
    
    If Not (ftpConnected = NotState) Then
    
        If mTCP.Connected Then
            SendFTPCommand "QUIT"
        Else
            mTCP.Disconnect
        End If
    Else
        mTCP.Disconnect
    End If
    
    ftpConnected = NotState


End Sub

Public Sub CancelTransfer(Optional ByVal AbortSignal As Boolean = False) ' _
Cancels any ListContents or TransferFile action that is still in progress.  This does not send the Abort signal by default, so that any file upload remains partial on the remote system.
Attribute CancelTransfer.VB_Description = "Cancels any ListContents or TransferFile action that is still in progress.  This does not send the Abort signal by default, so that any file upload remains partial on the remote system."
    On Error Resume Next

    
    ftpCancelThread = True
    
    Static stack As Boolean
    If Not stack Then

        stack = True
        If (Not (myObj Is Nothing)) Then
            myObj.CancelTransfer
        End If
        
        If (Not (myServerObj Is Nothing)) Then
            myServerObj.CancelTransfer
        End If
    
        stack = False
    End If


    
    If Transfering Then
        If mData.Connected Or mData.Listening Then mData.Disconnect
    
        If AbortSignal Then
            
            SendFTPCommand "ABOR"
    
        End If
        
    End If
    
    
    
    If Err Then Err.Clear
    
End Sub

Private Sub GetSystemInfo()
    
    SendFTPCommand "SYST"
    
End Sub

Public Sub NoOperation() ' _
Sends the NOOP command to the connected remote system.  NOOP requires a connection and preforms no operation.
Attribute NoOperation.VB_Description = "Sends the NOOP command to the connected remote system.  NOOP requires a connection and preforms no operation."

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "NOOP: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
        SendFTPCommand "NOOP"
    End If
    
End Sub

Public Sub ChangeFolderRelative(ByVal NewFolder As String) ' _
Changes the remote systems folder location with reletivity to the current Folder property.  Special folders "".."", ""."", and root slash may be used.  Upon success the combined new location will reflect in the Folder property.
Attribute ChangeFolderRelative.VB_Description = "Changes the remote systems folder location with reletivity to the current Folder property.  Special folders "".."", ""."", and root slash may be used.  Upon success the combined new location will reflect in the Folder property."

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "CWD: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
        SendFTPCommand "CWD " & NewFolder
    End If
    
    ftpFolder = MapFolder(ftpFolder, NewFolder, ftpType)
    
End Sub

Public Property Get SystemHelp() As String

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "HELP: Not connected to a file system."
    End If

    If ftpType = URLTypes.ftp Then

        ftpHelpInformation = ""
        ftpGatherType = GatherTypes.Help
        SendFTPCommand "HELP"
        SystemHelp = ftpHelpInformation
    End If
    
End Property

Public Property Get StatInformation() As String

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "STAT: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        ftpStatInfo = ""
        ftpGatherType = GatherTypes.Stat
        SendFTPCommand "STAT"
        StatInformation = ftpStatInfo
    End If
    
End Property

Public Sub ChangeFolderAbsolute(ByVal NewFolder As String) ' _
Changes the remote systems folder location absolute to the passed value, and upon success it will reflect in the Folder property entirely to what is passed.
Attribute ChangeFolderAbsolute.VB_Description = "Changes the remote systems folder location absolute to the passed value, and upon success it will reflect in the Folder property entirely to what is passed."

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "CWD: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
        SendFTPCommand "CWD " & NewFolder
    End If
    
    ftpFolder = NewFolder

End Sub

Public Sub Rename(ByVal FromFileName As String, ByVal ToFileName As String)

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "RNFR: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        SendFTPCommand "RNFR " & FromFileName & vbCrLf & "RNTO " & ToFileName & vbCrLf
    
    ElseIf ftpType = URLTypes.File Or ftpType = URLTypes.Remote Then

        Dim errNum As Long
        Dim sFile As String
        Dim dFile As String
            
        sFile = MapFolder(Folder, FromFileName, ftpType)
        dFile = MapFolder(Folder, ToFileName, ftpType)
        
        Name sFile As dFile
   
    End If

End Sub

Public Sub MakeFolder(ByVal NewFolder As String)

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "MKD: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        SendFTPCommand "MKD " & NewFolder
        
    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
    
        If Not PathExists(MapFolder(Folder, NewFolder, ftpType), False) Then
            MkDir MapFolder(Folder, NewFolder, ftpType)
        End If
    
    End If

End Sub

Public Sub RemoveFile(ByVal FileName As String)

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "DELE: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        SendFTPCommand "DELE " & FileName
    
    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
    
        Dim errNum As Long
        Dim dFile As String
        dFile = MapFolder(Folder, FileName, ftpType)

        Kill dFile

    End If

End Sub

Public Sub RemoveFolder(ByVal FolderName As String)

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "RMD: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        SendFTPCommand "RMD " & FolderName
    
    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
    
        Dim errNum As Long
        Dim dFile As String
        dFile = MapFolder(Folder, FolderName, ftpType)

        RmDir dFile

    End If

End Sub

Public Property Let TransferType(ByVal TransType As TransferModes)

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "TYPE: Not connected to a file system."
    End If
    
    If ftpType = URLTypes.ftp Then
    
        If ftpTransferMode <> TransType Then
            ftpTransferMode = TransType
            
            If TransType = TransferModes.Binary Then
    
                SendFTPCommand "TYPE I"
    
            ElseIf TransType = TransferModes.ASCII Then
    
                SendFTPCommand "TYPE A"
    
            End If
        
        End If
        
    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
    
    End If

End Property

Public Property Get TransferType() As TransferModes
    TransferType = ftpTransferMode
End Property


Private Function TransferFileServerToLocal(ByVal SourceFileName As String, ByVal DestFileName As String, Optional ByVal ResumeOffset As Double = 0, Optional ByVal TotalFileSize As Double = 0) As String

    If (ResumeOffset > 0) And (ftpTransferMode = TransferModes.ASCII) Then
        ResumeOffset = 0
    End If

    If (ResumeOffset > 0) Then
        If Not PathExists(DestFileName) Then
            SetLastError RaiseError, 8, "Resume byte position specified does not match destination files size for appending, or is greater then the source file size."
        ElseIf (Not (ResumeOffset = FileSize(DestFileName))) Then
             
            SetLastError RaiseError, 8, "Resume byte position specified does not match destination files size for appending, or is greater then the source file size."
        End If
    End If

    Dim cnt As Double

    If ftpLargeFileMode Or (TotalFileSize > modBitValue.LongBound) Or (ResumeOffset > modBitValue.LongBound) Then
        InitializeThread

        Set myObj = Me
        Set myServerObj = Me

        ConnectedState(RemoteToLocal) = True
        mData.RecvRate = ftpTransferRates(RateTypes.Download)
        mData.SendRate = ftpTransferRates(RateTypes.Upload)
        ftpResumeByte = ResumeOffset
        ftpFileSize = ResumeOffset

         If (ResumeOffset > 0) Then
            Set ftpWriteTextStream = myFso.OpenTextFile(DestFileName, IOMode.ForAppending, False)
            If Not ftpCancelThread Then SendDataFTPCommand ftpConnectionMode, "RETR " & SourceFileName, "REST " & ResumeOffset

        Else

            Set ftpWriteTextStream = myFso.OpenTextFile(DestFileName, IOMode.ForWriting, True)
            If Not ftpCancelThread Then SendDataFTPCommand ftpConnectionMode, "RETR " & SourceFileName
        End If

    Else
        InitializeThread

        Set myObj = Me
        Set myServerObj = Me

        ConnectedState(RemoteToLocal) = True
        mData.RecvRate = ftpTransferRates(RateTypes.Download)
        mData.SendRate = ftpTransferRates(RateTypes.Upload)
        ftpResumeByte = ResumeOffset
        ftpFileSize = ResumeOffset

        If (ResumeOffset > 0) Then

            ftpWriteFileNum = FreeFile
            Open DestFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1

            Seek #ftpWriteFileNum, ResumeOffset + 1

            If (Not ftpCancelThread) Then
                SendDataFTPCommand ftpConnectionMode, "RETR " & SourceFileName, "REST " & ResumeOffset
            End If

        Else
            ftpWriteFileNum = FreeFile
            Open DestFileName For Output Shared As #ftpWriteFileNum
            Close #ftpWriteFileNum

            ftpWriteFileNum = FreeFile
            Open DestFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1

            Seek #ftpWriteFileNum, 1

            If (Not ftpCancelThread) Then
                SendDataFTPCommand ftpConnectionMode, "RETR " & SourceFileName
            End If
        End If

    End If

End Function

Public Sub TransferFile(ByVal SourceFileName As String, ByVal DestinationObject As Client, Optional ByVal ResumeOffset As Double = 0, Optional ByVal TotalFileSize As Double = 0)

    On Error GoTo FileError

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "Source object not connected to a file system."
    End If

    If Not DestinationObject.ConnectedState() Then
        SetLastError RaiseError, 8, "Destination object not connected to a file system."
    End If

    Set ftpRaiseFTTPComplete = Nothing
    
    Dim DestFileName As String
    
    If ftpType = URLTypes.ftp Then

        If DestinationObject.URLType = URLTypes.File Or DestinationObject.URLType = URLTypes.Remote Then
            DestFileName = MapFolder(DestinationObject.Folder, MapFileName(SourceFileName), ftpType)

            TransferFileServerToLocal SourceFileName, DestFileName, ResumeOffset, TotalFileSize

        ElseIf DestinationObject.URLType = URLTypes.ftp Then
            DestFileName = MapFileName(SourceFileName)
            
            TransferFileServerToServer SourceFileName, DestinationObject, ResumeOffset, TotalFileSize

        End If

    ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
        SourceFileName = MapFolder(Folder, MapFileName(SourceFileName), ftpType)

        If DestinationObject.URLType = URLTypes.ftp Then
            DestFileName = MapFileName(SourceFileName)

            DestinationObject.TransferFileLocalToServer Me, SourceFileName, DestFileName, ResumeOffset, TotalFileSize

        ElseIf DestinationObject.URLType = URLTypes.File Or DestinationObject.URLType = URLTypes.Remote Then
            DestFileName = MapFolder(DestinationObject.Folder, MapFileName(SourceFileName), ftpType)

            TransferFileLocalToLocal SourceFileName, DestFileName, ResumeOffset, TotalFileSize

        End If

    End If
    
    Exit Sub
FileError:
    If Err.Description <> "" Then ftpEventReason = Err.Description
    If ftpEventError = ErrorReturns.Last Then
        RaiseEvent Error(8, App.Title, ftpEventReason)
    Else
        RaiseEvent Error(8, App.Title, Err.Description)
    End If

End Sub

Public Function ListContents(Optional ByVal LocalFileName As String = "(Temp)", Optional ByVal TotalFileSize As Double = 0) As String

    On Error GoTo FileError
    
    If LocalFileName = "(Temp)" Then
        LocalFileName = GetTemporaryFile
        ListContents = LocalFileName
    End If

    If Not ConnectedState() Then
        SetLastError RaiseError, 8, "LIST: Not connected to a file system."
    End If
    
    If ftpLargeFileMode Or (TotalFileSize > modBitValue.LongBound) Then
        If ftpType = URLTypes.ftp Then

            InitializeThread
            
            Set myObj = Me
            Set myServerObj = Nothing
            
            ConnectedState(ListingToFile) = True

            mData.RecvRate = ftpTransferRates(RateTypes.Download)
            mData.SendRate = ftpTransferRates(RateTypes.Upload)
            ftpFileSize = 0
            
            Set ftpWriteTextStream = myFso.OpenTextFile(LocalFileName, IOMode.ForWriting, True)
            
            SendDataFTPCommand ftpConnectionMode, "LIST"
    
        ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
            InitializeThread
            
            Set myObj = Me
            Set myServerObj = Nothing
                        
            ConnectedState(ListingToFile) = True

            mData.RecvRate = ftpTransferRates(RateTypes.Download)
            mData.SendRate = ftpTransferRates(RateTypes.Upload)
            ftpFileSize = 0
            
            Set ftpWriteTextStream = myFso.OpenTextFile(LocalFileName, IOMode.ForWriting, True)
            
            mData_SendComplete
            
        End If
    Else
    
        If ftpType = URLTypes.ftp Then
            InitializeThread
            
            Set myObj = Me
            Set myServerObj = Nothing
            
            ConnectedState(ListingToFile) = True

            mData.RecvRate = ftpTransferRates(RateTypes.Download)
            mData.SendRate = ftpTransferRates(RateTypes.Upload)
            
            ftpWriteFileNum = FreeFile
            Open LocalFileName For Output Shared As #ftpWriteFileNum
            Close #ftpWriteFileNum
            
            Open LocalFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1
            
            SendDataFTPCommand ftpConnectionMode, "LIST"
            
        ElseIf ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
            InitializeThread
            
            Set myObj = Me
            Set myServerObj = Nothing

            ConnectedState(ListingToFile) = True

            mData.RecvRate = ftpTransferRates(RateTypes.Download)
            mData.SendRate = ftpTransferRates(RateTypes.Upload)
            
            ftpWriteFileNum = FreeFile
            Open LocalFileName For Output Shared As #ftpWriteFileNum
            Close #ftpWriteFileNum
            
            Open LocalFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1

            mData_SendComplete
            
        End If

    End If


    Exit Function
FileError:
    If Err.Description <> "" Then ftpEventReason = Err.Description
    If ftpEventError = ErrorReturns.Last Then
        RaiseEvent Error(8, App.Title, ftpEventReason)
    Else
        RaiseEvent Error(8, App.Title, Err.Description)
    End If
End Function

Friend Function TransferFileLocalToServer(ByVal SourceObject As Client, ByVal SourceFileName As String, ByVal DestFileName As String, Optional ByVal ResumeOffset As Double = 0, Optional ByVal TotalFileSize As Double = 0) As String
    
    If FileSize(SourceFileName) > modBitValue.HighBound() Then
        SetLastError RaiseError, 8, "Transfering for file sizes over " & CDec(CDbl(modBitValue.HighBound())) & " bytes not implementated."
    End If
    
    If (ResumeOffset > 0) Then
        If (ftpTransferMode = TransferModes.ASCII) Then
            ResumeOffset = 0
            'SetLastError RaiseError, 8,  "Unable to resume in ASCII mode (any single line of WINS text, or line of code, may not exist with out line feed and carriage return)."
        End If
        
        If (ResumeOffset >= FileSize(SourceFileName)) Then
            SetLastError RaiseError, 8, "Resume byte position specified does not match destination files size for appending, or is greater then the source file size."
        End If
    End If
    
    Dim cnt As Double

    If ftpLargeFileMode Or (TotalFileSize > modBitValue.LongBound) Or (ResumeOffset > modBitValue.LongBound) Then
        InitializeThread
        
        Set myObj = SourceObject
        Set myServerObj = Me

        ftpFileSize = FileSize(SourceFileName)
        ConnectedState(LocalToRemote) = True

        mData.RecvRate = ftpTransferRates(RateTypes.Download)
        mData.SendRate = ftpTransferRates(RateTypes.Upload)
        ftpSendPacketLen = ftpTransferRates(RateTypes.Upload)
        ftpSendIncrement = 0
        
        Set ftpReadTextStream = myFso.OpenTextFile(SourceFileName, 1, False)

        If (CDbl(ftpFileSize - ftpResumeByte) > ftpSendPacketLen) Then
            ftpSendCountKnot = ReturnDiv(CDbl(ftpFileSize - ftpResumeByte), ftpSendPacketLen)
            ftpSendBytesLeft = ReturnMod(CDbl(ftpFileSize - ftpResumeByte), ftpSendPacketLen)
        Else
            ftpSendCountKnot = 0
            ftpSendBytesLeft = CDbl(ftpFileSize - ftpResumeByte)
        End If
        
        If Not (ftpFileSize = 0) Then
           
            If (ResumeOffset > 0) Then
                
                If ResumeOffset < modBitValue.IntBound Then
                    ftpReadTextStream.Skip ResumeOffset
                    ftpResumeByte = ResumeOffset
                    SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                Else
                    For cnt = 1 To ReturnDiv(ResumeOffset, modBitValue.IntBound)
                        If ftpCancelThread Then Exit For
                        ftpReadTextStream.Skip modBitValue.IntBound
                        ftpResumeByte = ftpResumeByte + modBitValue.IntBound
                        SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                        DoTasks
                    Next
                    If Not ftpCancelThread Then
                        cnt = ReturnMod(ResumeOffset, modBitValue.IntBound)
                        If cnt > 0 Then
                            ftpReadTextStream.Skip cnt
                            ftpResumeByte = ftpResumeByte + cnt
                            SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                        End If
                    End If
                End If
            Else
                ftpResumeByte = 0
            End If
            
            If Not ftpCancelThread Then
                If (ftpResumeByte > 0) Then
                    SendDataFTPCommand ftpConnectionMode, "APPE " & DestFileName, _
                        IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")
                Else
                    SendDataFTPCommand ftpConnectionMode, "STOR " & DestFileName, _
                        IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")
                End If

            End If
        Else
            ftpResumeByte = 0
            SendDataFTPCommand ftpConnectionMode, "STOR " & DestFileName, _
                IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")

            mData.Disconnect
        End If
        
    Else
        InitializeThread
        
        Set myObj = SourceObject
        Set myServerObj = Me

        ftpSendPacketLen = ftpTransferRates(RateTypes.Upload)
        ConnectedState(LocalToRemote) = True
        mData.RecvRate = ftpTransferRates(RateTypes.Download)
        mData.SendRate = ftpTransferRates(RateTypes.Upload)
        ftpSendIncrement = 0

        ftpReadFileNum = FreeFile
        Open SourceFileName For Binary Access Read Lock Read As #ftpReadFileNum Len = 1

        ftpFileSize = FileSize(SourceFileName)
        
        If ftpFileSize > 0 Then

            If (CDbl(ftpFileSize - ftpResumeByte) > ftpSendPacketLen) Then
                ftpSendCountKnot = ReturnDiv(CDbl(ftpFileSize - ftpResumeByte), ftpSendPacketLen)
                ftpSendBytesLeft = ReturnMod(CDbl(ftpFileSize - ftpResumeByte), ftpSendPacketLen)
            Else
                ftpSendCountKnot = 0
                ftpSendBytesLeft = CDbl(ftpFileSize - ftpResumeByte)
            End If
        
            If (ResumeOffset > 0) Then

                If ResumeOffset < modBitValue.IntBound Then
                    Seek #ftpReadFileNum, ResumeOffset + 1
                    ftpResumeByte = ResumeOffset
                    SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                Else
                    For cnt = 1 To ReturnDiv(ResumeOffset, modBitValue.IntBound)
                        If ftpCancelThread Then Exit For
                        Seek #ftpReadFileNum, ResumeOffset + 1
                        ftpResumeByte = ftpResumeByte + modBitValue.IntBound
                        SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                        DoTasks
                    Next
                    If Not ftpCancelThread Then
                        cnt = ReturnMod(ResumeOffset, modBitValue.IntBound)
                        If cnt > 0 Then
                            Seek #ftpReadFileNum, ResumeOffset + 1
                            ftpResumeByte = ftpResumeByte + modBitValue.IntBound
                            SourceObject.Self "RaiseDataProgress", ProgressTypes.PositioningFile, ftpResumeByte
                        End If
                    End If
                End If
                
            Else
                ftpResumeByte = 0
                Seek #ftpReadFileNum, 1

            End If

            If (ResumeOffset > 0) Then
                SendDataFTPCommand ftpConnectionMode, "APPE " & DestFileName, _
                    IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")
            Else
                SendDataFTPCommand ftpConnectionMode, "STOR " & DestFileName, _
                    IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")
            End If
        Else
            ftpResumeByte = 0
            
            SendDataFTPCommand ftpConnectionMode, "STOR " & DestFileName, _
                IIf(((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote), "ALLO " & (TotalFileSize - ResumeOffset), "")

            mData.Disconnect
        End If
    End If
    
End Function

Private Function TransferFileLocalToLocal(ByVal SourceFileName As String, ByVal DestFileName As String, Optional ByVal ResumeOffset As Double = 0, Optional ByVal TotalFileSize As Double = 0) As String

    If CDec(FileSize(SourceFileName)) > CDec(CDbl(modBitValue.HighBound())) Then
        SetLastError RaiseError, 8, "Transfering for file sizes over " & CDec(CDbl(modBitValue.HighBound())) & " bytes not implementated."
    End If
    
    If (ResumeOffset > 0) Then
        If Not PathExists(DestFileName) Then
            SetLastError RaiseError, 8, "Resume byte position specified does not match destination files size for appending, or is greater then the source file size."
        ElseIf Not (ResumeOffset = FileSize(DestFileName)) Or (ResumeOffset >= FileSize(SourceFileName)) Then
            SetLastError RaiseError, 8, "Resume byte position specified does not match destination files size for appending, or is greater then the source file size."
        End If
    End If

    Dim cnt As Double

    If ftpLargeFileMode Or (TotalFileSize > modBitValue.LongBound) Or (ResumeOffset > modBitValue.LongBound) Then
        InitializeThread
        
        Set myObj = Me
        Set myServerObj = Nothing
        
        ftpFileSize = FileSize(SourceFileName)
        ConnectedState(LocalToLocal) = True
        ftpSendPacketLen = ftpTransferRates(RateTypes.HardDrive)
        ftpSendIncrement = 1

        If (CDbl(ftpFileSize - ResumeOffset) > ftpSendPacketLen) Then
            ftpSendCountKnot = ReturnDiv(CDbl(ftpFileSize - ResumeOffset), ftpSendPacketLen)
            ftpSendBytesLeft = ReturnMod(CDbl(ftpFileSize - ResumeOffset), ftpSendPacketLen)
        Else
            ftpSendCountKnot = 0
            ftpSendBytesLeft = CDbl(ftpFileSize - ResumeOffset)
        End If
            
        Set ftpReadTextStream = myFso.OpenTextFile(SourceFileName, 1, False)
            
        If (ResumeOffset > 0) Then

            If (ResumeOffset > 0) Then

                If ResumeOffset < (modBitValue.LongBound \ 4) Then
                    ftpReadTextStream.Skip ResumeOffset
                    ftpResumeByte = ResumeOffset
                Else
                    For cnt = 1 To ReturnDiv(ResumeOffset, (modBitValue.LongBound \ 4))
                        If ftpCancelThread Then Exit For
                        ftpReadTextStream.Skip (modBitValue.LongBound \ 4)
                        ftpResumeByte = ftpResumeByte + (modBitValue.LongBound \ 4)
                        RaiseEvent DataProgress(ProgressTypes.PositioningFile, ftpResumeByte)
                        'DoTasks
                    Next
                    If Not ftpCancelThread Then
                        cnt = ReturnMod(ResumeOffset, (modBitValue.LongBound \ 4))
                        If cnt > 0 Then
                            ftpReadTextStream.Skip cnt
                            ftpResumeByte = ftpResumeByte + cnt
                            RaiseEvent DataProgress(ProgressTypes.PositioningFile, ftpResumeByte)
                        End If
                    End If
                End If
                
            End If
            
            If (Not ftpCancelThread) Then
                Set ftpWriteTextStream = myFso.OpenTextFile(DestFileName, IOMode.ForAppending, False)
            End If
                            
        ElseIf (Not ftpCancelThread) Then
            ftpResumeByte = 0
            Set ftpWriteTextStream = myFso.OpenTextFile(DestFileName, IOMode.ForWriting, True)
        End If
        
        If (Not ftpCancelThread) Then
            If (Not (ftpFileSize = 0)) And (Not ftpCancelThread) Then
                If (Not ftpCancelThread) Then
    
                    If ftpSendCountKnot > 0 Then
                        ftpWriteTextStream.Write ftpReadTextStream.Read(ftpSendPacketLen)

                    ElseIf ftpSendBytesLeft > 0 Then
                        ftpWriteTextStream.Write ftpReadTextStream.Read(ftpSendBytesLeft)

                    End If
                End If
            End If
        
            If (Not ftpCancelThread) Then
                If (ftpFileSize = 0) Then
    
                    If Not (ftpReadTextStream Is Nothing) Then
                        ftpReadTextStream.Close
                        Set ftpReadTextStream = Nothing
                    End If
    
                    If Not (ftpWriteTextStream Is Nothing) Then
                        ftpWriteTextStream.Close
                        Set ftpWriteTextStream = Nothing
                    End If
    
                    TerminateTransfer
                Else
                    mData_SendComplete
            
                End If
            End If
        
        End If

    Else

        Dim inFileData() As Byte
        InitializeThread
        
        Set myObj = Me
        Set myServerObj = Nothing
        
        ConnectedState(LocalToLocal) = True
        ftpFileSize = FileSize(SourceFileName)
        ftpSendPacketLen = ftpTransferRates(RateTypes.HardDrive)
        ftpResumeByte = ResumeOffset
        ftpSendIncrement = 1
        
        If (CDbl(ftpFileSize - ResumeOffset) > ftpSendPacketLen) Then
            ftpSendCountKnot = ReturnDiv(CDbl(ftpFileSize - ResumeOffset), ftpSendPacketLen)
            ftpSendBytesLeft = ReturnMod(CDbl(ftpFileSize - ResumeOffset), ftpSendPacketLen)
        Else
            ftpSendCountKnot = 0
            ftpSendBytesLeft = CDbl(ftpFileSize - ResumeOffset)
        End If
                    
        ftpReadFileNum = FreeFile
        Open SourceFileName For Binary Access Read Lock Read As #ftpReadFileNum Len = 1
        
        If (ResumeOffset > 0) Then

            ftpWriteFileNum = FreeFile
            Open DestFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1

            Seek #ftpWriteFileNum, ResumeOffset + 1
        Else

            ftpWriteFileNum = FreeFile
            Open DestFileName For Output Shared As #ftpWriteFileNum
            Close #ftpWriteFileNum
        
            ftpWriteFileNum = FreeFile
            Open DestFileName For Binary Access Write Lock Write As #ftpWriteFileNum Len = 1

            Seek #ftpWriteFileNum, 1
        End If
        
        If ftpFileSize > 0 Then
            
            ReDim inFileData(1 To ftpSendPacketLen) As Byte

            If Not ftpCancelThread Then
    
                If (ResumeOffset > 0) Then
                    Seek #ftpReadFileNum, ResumeOffset + 1
                Else
                    Seek #ftpReadFileNum, 1
                End If
                        
                If Not ftpCancelThread Then
                    If ftpSendCountKnot > 0 Then
                        Get #ftpReadFileNum, , inFileData()
                        Put #ftpWriteFileNum, , inFileData()
                        
                    ElseIf ftpSendBytesLeft > 0 Then
                        ReDim inFileData(1 To ftpSendBytesLeft) As Byte
                        Get #ftpReadFileNum, , inFileData()
                        Put #ftpWriteFileNum, , inFileData()

                    End If
                End If
            End If
        End If
        
        
            
        If (Not ftpCancelThread) Then

            If (ftpFileSize = 0) Then

                If ftpReadFileNum > 0 Then
                    Close #ftpReadFileNum
                    ftpReadFileNum = 0
                End If

                If ftpWriteFileNum > 0 Then
                    Close #ftpWriteFileNum
                    ftpWriteFileNum = 0
                End If

                TerminateTransfer
            Else
                mData_SendComplete
            
            End If
        End If
    End If
    
End Function

Private Function TransferFileServerToServer(ByVal SourceFileName As String, ByRef DestinationObject As Client, Optional ByVal ResumeOffset As Long = 0, Optional ByVal TotalFileSize As Double = 0) As String
    SetLastError RaiseError, 8, "Feature disabled.  Refer to RFC 959 of the The Internet Engineering Task Force," & vbCrLf & _
                               "IETF, and IP v4 security in a way-back structured mount to bipass transfering."
On Error GoTo catch:
    Dim dataPort As String
    ftpCancelThread = False
    Set ftpRaiseFTTPComplete = Nothing
    'DestinationObject.Self "ftpRaiseFTTPComplete", Me
If fttpModePole Then
    On Error GoTo method2:
    GoTo skipit:
Else
    GoTo method2:
End If
method1:
Err.Clear
On Error GoTo catch:
skipit:
    dataPort = SendDataFTTPCommand("PASV")
    DestinationObject.SendDataFTTPCommand "PORT " & dataPort
    If ResumeOffset > 0 Then
        If ((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote) Then
            DestinationObject.SendFTPCommand "ALLO " & (TotalFileSize - ResumeOffset)
        End If
        DestinationObject.SendFTTPCommand "APPE " & SourceFileName
    Else
        If ((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote) Then
            DestinationObject.SendFTPCommand "ALLO " & (TotalFileSize - ResumeOffset)
        End If
        DestinationObject.SendFTTPCommand "STOR " & SourceFileName
    End If
    If ResumeOffset > 0 Then
        SendFTPCommand "REST " & ResumeOffset
    End If
    SendFTPCommand "RETR " & SourceFileName
fttpModePole = True
GoTo methodSuccess
method2:
Err.Clear
If fttpModePole Then
    On Error GoTo catch:
Else
    On Error GoTo method1:
End If
    dataPort = DestinationObject.SendDataFTTPCommand("PASV")
    SendDataFTTPCommand "PORT " & dataPort
    If ResumeOffset > 0 Then
        If ((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote) Then
            DestinationObject.SendFTPCommand "ALLO " & (TotalFileSize - ResumeOffset)
        End If
        dataPort = DestinationObject.SendFTTPCommand("APPE " & SourceFileName)
    Else
        If ((TotalFileSize - ResumeOffset) > 0) And ((ftpAllocation And AllocateSides.Remote) = AllocateSides.Remote) Then
            DestinationObject.SendFTPCommand "ALLO " & (TotalFileSize - ResumeOffset)
        End If
        dataPort = DestinationObject.SendFTTPCommand("STOR " & SourceFileName)
    End If
    If ResumeOffset > 0 Then
        SendFTPCommand "REST " & ResumeOffset
    End If
    SendFTPCommand "RETR " & SourceFileName
fttpModePole = Not fttpModePole
methodSuccess:
On Error GoTo catch:
      '  WaitForDataServerClosed
       ' DestinationObject.Self "WaitForDataServerClosed"
Exit Function
catch:
    Set ftpRaiseFTTPComplete = Nothing
    Err.Clear
    On Error GoTo 0
    SetLastError RaiseError, 8, "Unable to make both servers negotiate a connection between each other.  Both servers may each be behind a firewall not allowing the either of the other to make a required exception, or either of both servers may not allow special server to server transfers with out further administration."
End Function

Friend Function SendFTPCommand(ByVal ftpCommand As String) As String
    If (Not ConnectedState(CommandRequest)) And (Not (ftpEventError = ErrorReturns.Wait)) And (ftpNextCommand = "") Then
        ftpEventError = ErrorReturns.none
        ftpNextCommand = ftpNextCommand & ftpCommand & vbCrLf
        DoOutCommands
    Else
        ftpNextCommand = ftpNextCommand & ftpCommand & vbCrLf
    End If
End Function


Private Function GetDelimiter(ByVal strText As String) As String
    If InStr(strText, vbCrLf) > 0 Then
        GetDelimiter = vbCrLf
    ElseIf InStr(strText, vbCr) = 0 Then
        GetDelimiter = vbLf
    End If
End Function

Public Function ParseListing(ByVal ListData As String, Optional ByRef ListItems) As Integer ' _
Attempts to determine listing systems format based on ListData, returning the count of items it has parsed, with optionally raising the ListInformed event per item parsed.  ListItems supplied as an array or collection, will also be populated.
Attribute ParseListing.VB_Description = "Attempts to determine listing systems format based on ListData, returning the count of items it has parsed, with optionally raising the ListInformed event per item parsed.  ListItems supplied as an array or collection, will also be populated."
    Dim has As Boolean
    Dim ary As Boolean
    Dim cnt As Long
    Dim oneLine As String
    Dim delim As String
    Dim lSys As Integer
    lSys = ListSystems.none
    delim = vbLf
    ListData = Replace(Replace(ListData, vbCrLf, vbLf), vbCr, vbNullString)
    If IsMissing(ListItems) Then
        has = False
    Else
        If IsArray(ListItems) Then
            ReDim ListItems(0) As String
            ary = True
        ElseIf TypeName(ListItems) = "Nothing" Then
'            Set ListItems = New NTNodes10.Collection
'            ary = False
'        ElseIf TypeName(ListItems) = "VBA.Collection" Then
            Set ListItems = New VBA.Collection
            ary = False
        Else
            ary = False
        End If
        has = True
    End If
    If Left(LCase(ListData), 5) = "total" Then
        If (InStr(ListData, delim) > 0) Then
            ListData = Mid(ListData, InStr(ListData, delim) + Len(delim))
        End If
    End If
    If Len(ListData) > 0 Then
        If (IsNumeric(Left(ListData, 2)) And Mid(ListData, 3, 1) = "-") Or (IsNumeric(Left(ListData, 1)) And Mid(ListData, 2, 1) = "-") Then
            lSys = ListSystems.Wins
        Else
            If Left(ListData, 1) = "-" Or Left(ListData, 1) = "d" Then
                lSys = ListSystems.Unix
            Else
            End If
        End If
        cnt = 0
        Do Until Trim(ListData) = ""
            If InStr(ListData, delim) > 0 Then
                oneLine = Left(ListData, InStr(ListData, delim) - 1)
                ListData = Mid(ListData, InStr(ListData, delim) + Len(delim))
            Else
                oneLine = ListData
                ListData = ""
            End If
            oneLine = Replace(Replace(Replace(Replace(oneLine, vbCr, ""), vbTab, " "), vbLf, ""), Chr(0), "")
            If Replace(Replace(Replace(oneLine, vbCr, ""), vbLf, ""), " ", "") <> "" Then
                If has And ary Then
                    ReDim Preserve ListItems(cnt) As String
                End If
                Select Case lSys
                    Case ListSystems.Wins
                        oneLine = GetMSDosListParams(oneLine)
                        If Trim(Replace(oneLine, " ", "")) <> "" Then
                            If has Then
                                If ary Then
                                    ListItems(cnt) = oneLine
                                Else
                                    ListItems.Add oneLine
                                End If
                            End If
                            cnt = cnt + 1
                        End If
                    Case ListSystems.Unix
                        oneLine = GetUnixListParams(oneLine)
                        If Trim(Replace(oneLine, " ", "")) <> "" Then
                            If has Then
                                If ary Then
                                    ListItems(cnt) = oneLine
                                Else
                                    ListItems.Add oneLine
                                End If
                            End If
                            cnt = cnt + 1
                        End If
                    Case ListSystems.none
                        oneLine = GetMSDosListParams(oneLine)
                        If Trim(Replace(oneLine, " ", "")) <> "" Then
                            If has Then
                                If ary Then
                                    ListItems(cnt) = oneLine
                                Else
                                    ListItems.Add oneLine
                                End If
                            End If
                            cnt = cnt + 1
                        End If
                End Select
            End If
        Loop
    End If
    If Not has And cnt > 0 Then
        ParseListing = cnt
    Else
        ParseListing = lSys
    End If
End Function

Private Function GetMSDosListParams(ByVal FullStr As String) As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemAccess As String
    On Error GoTo finished
    ItemDate = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    ItemSize = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    If Not IsNumeric(ItemSize) Then
        ItemDate = ItemDate + " " + ItemSize
        FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
        ItemSize = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    End If
    ItemName = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    If ItemName = "./" Or ItemName = "../" Or ItemName = "" Or ItemName = ".." Or ItemName = "." Then
        ItemName = ""
    Else
        If Right(ItemName, 1) = "/" Then
            ItemName = "/" + Left(ItemName, Len(ItemName) - 1)
        Else
            If ItemSize = "<DIR>" And Left(ItemName, 1) <> "/" Then ItemName = "/" + ItemName
            End If
        End If
    If ItemName = "" Then
        GetMSDosListParams = ""
    Else
        GetMSDosListParams = ItemName & " | " & ItemSize & " | " & ItemDate & " | " & ItemAccess
        RaiseEvent ItemListing(ItemName, ItemSize, ItemDate, ItemAccess)
    End If
    Exit Function
finished:
    Err.Clear
    GetMSDosListParams = ""
End Function

Private Function GetUnixListParams(ByVal FullStr As String) As String
    Dim ItemName As String
    Dim ItemSize As String
    Dim ItemDate As String
    Dim ItemAccess As String
    On Error GoTo finished
    ItemAccess = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    Do
        ItemSize = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
        FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    Loop Until Not IsNumeric(Left(FullStr, 1))
    ItemDate = Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    ItemDate = ItemDate & " " & Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    FullStr = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    ItemDate = ItemDate & " " & Trim(Left(FullStr, InStr(FullStr, " ") - 1))
    ItemName = Trim(Mid(FullStr, InStr(FullStr, " ") + 1))
    If ItemName = "./" Or ItemName = "../" Or ItemName = "" Or ItemName = ".." Or ItemName = "." Then
        ItemName = ""
    Else
        If Right(ItemName, 1) = "/" Then
            ItemName = "/" + Left(ItemName, Len(ItemName) - 1)
        Else
            If Left(ItemAccess, 1) = "d" Then
                If Left(ItemName, 1) <> "/" Then ItemName = "/" + ItemName
            End If
        End If
    End If
    If ItemName = "" Then
        GetUnixListParams = ""
    Else
        GetUnixListParams = ItemName & " | " & ItemSize & " | " & ItemDate & " | " & ItemAccess
        RaiseEvent ItemListing(ItemName, ItemSize, ItemDate, ItemAccess)
    End If
    Exit Function
finished:
    Err.Clear
    GetUnixListParams = ""
End Function

Private Sub AddToLog(ByVal msgType As NTAdvFTP61.MessageTypes, ByVal LogData As String)
    If ftpLogBytes > 0 Then
        If Len(ftpLogData) + Len(LogData) > ftpLogBytes Then
            ftpLogData = Mid(ftpLogData, Len(LogData))
        End If
    ElseIf (ftpLogBytes = 0) And (Not (ftpLogData = "")) Then
        ftpLogData = ""
    End If
    ftpLogData = ftpLogData & LogData
    RaiseEvent LogMessage(msgType, LogData)
End Sub

Friend Function Self(ByVal xAction As String, Optional xParam1 As Variant, Optional xParam2 As Variant) As String
    Select Case LCase(Trim(xAction))
        Case "transfering"
            If ftpType = ftp Then
                Self = ConnectedState(DataStreaming) Or mData.Connected Or mData.Listening
            Else
                Self = ConnectedState(DataStreaming)
            End If
        Case "raisedataprogress"
            RaiseEvent DataProgress(xParam1, xParam2)
        Case "raisedatacomplete"
            RaiseEvent DataComplete(xParam1)
    End Select
End Function
Private Sub CloseFiles()
    If ftpLargeFileMode Then
        If Not (ftpWriteTextStream Is Nothing) Then
            ftpWriteTextStream.Close
            Set ftpWriteTextStream = Nothing
        End If
        If Not (ftpReadTextStream Is Nothing) Then
            ftpReadTextStream.Close
            Set ftpReadTextStream = Nothing
        End If
    Else
        If ftpWriteFileNum > 0 Then
            Close #ftpWriteFileNum
            ftpWriteFileNum = 0
        End If
        If ftpReadFileNum > 0 Then
            Close #ftpReadFileNum
            ftpReadFileNum = 0
        End If
    End If
   If ftpLargeFileMode Then
        ftpFileSize = 0
        ftpResumeByte = 0
        Set ftpWriteTextStream = Nothing
        Set ftpReadTextStream = Nothing
    Else
        ftpWriteFileNum = 0
        ftpReadFileNum = 0
    End If
    ftpSendPacketLen = 0
    ftpSendIncrement = 0
    ftpSendCountKnot = 0
    ftpSendBytesLeft = 0
    Set myFso = Nothing
End Sub
Private Sub TerminateTransfer()
    CloseFiles

    If ConnectedState(ListingToFile) Then
       ConnectedState(ListingToFile) = False
        If Not (myObj Is Nothing) Then
            myObj.Self "RaiseDataComplete", ProgressTypes.FileListing
        End If
    ElseIf ConnectedState(RemoteToLocal) Then
        ConnectedState(RemoteToLocal) = False
        If Not (myObj Is Nothing) Then
            myObj.Self "RaiseDataComplete", ProgressTypes.TransferingFile
        End If
    ElseIf ConnectedState(LocalToRemote) Then
       ConnectedState(LocalToRemote) = False
        If Not (myObj Is Nothing) Then
            myObj.Self "RaiseDataComplete", ProgressTypes.TransferingFile
        End If
    ElseIf ConnectedState(LocalToLocal) Then
       ConnectedState(LocalToLocal) = False
        If Not (myObj Is Nothing) Then
            myObj.Self "RaiseDataComplete", ProgressTypes.TransferingFile
        End If
    End If
    ConnectedState(DataStreaming) = False
End Sub


Private Function GetWinDate(ByVal fDate As String) As String
    fDate = Trim(fDate)
    If InStr(fDate, " ") > 0 Then
        GetWinDate = Replace(Left(fDate, InStr(fDate, " ") - 1), "/", "-") & " " & Replace(Mid(fDate, InStr(fDate, " ") + 1), " ", "")
    Else
        GetWinDate = Replace(fDate, "/", "-") & " 00:00xM"
    End If
End Function

Private Sub InitializeThread()
    CloseFiles
    ftpCancelThread = False
    Set myFso = New Scripting.FileSystemObject
    ConnectedState(LocalToLocal) = False
    ConnectedState(LocalToRemote) = False
    ConnectedState(RemoteToLocal) = False
    ConnectedState(ListingToFile) = False
End Sub


Private Sub TimerProgresses()
    If (Not (myObj Is Nothing)) And (Not ftpCancelThread) Then
        If ConnectedState(RemoteToLocal) Then
            myObj.Self "RaiseDataProgress", ProgressTypes.TransferingFile, ftpFileSize
        ElseIf ConnectedState(LocalToLocal) Or ConnectedState(LocalToRemote) Then
            Dim tmpSize As Double
            tmpSize = CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))
            myObj.Self "RaiseDataProgress", ProgressTypes.TransferingFile, IIf(tmpSize <= ftpFileSize, tmpSize, ftpFileSize)
        ElseIf ConnectedState(ListingToFile) Then
            myObj.Self "RaiseDataProgress", ProgressTypes.FileListing, ftpFileSize
        End If
    End If
End Sub

Private Sub mData_Connected()

    ConnectedState(CommandRequest) = True

End Sub

Private Sub mData_Connection(ByRef Handle As Long)
    If mData.Listening Then
    
        If mData.WhosIs(mTCP.Handle) <> mTCP.RemoteIP() Then
            mData.Decline Handle
        Else
            mData.Accept Handle
        End If

    'Else    'must be a data arrival mismatch in ip security
   '     mData.Decline Handle
    End If
End Sub

Private Sub mData_DataArriving()

    If ConnectedState(RemoteToLocal) Or ConnectedState(ListingToFile) Then

        Dim Data() As Byte
        Dim inData As String

        On Error GoTo failerror

        If mData.ReadBytes(Data) Then

            ftpFileSize = ftpFileSize + modMemory.ArraySize(Data)
            inData = StrConv(Data, vbUnicode)

            SetServerUsesCRLF = inData

            If (myObj.ServerUsesCRLF And (myObj.TransferType = ASCII)) And (Not myObj.AssumeLineFeed) Then
                Data = modMemory.Convert(Replace(Replace(modMemory.Convert(Data), vbLf, vbCrLf), vbCr & vbCr, vbCr))
            ElseIf ((Not myObj.ServerUsesCRLF) And (myObj.TransferType = ASCII)) And myObj.AssumeLineFeed Then
                Data = modMemory.Convert(Replace(modMemory.Convert(Data), vbCr, ""))
            End If
            
            If ftpLargeFileMode Then
                ftpWriteTextStream.Write inData
            Else
                Put #ftpWriteFileNum, , Data
            End If
            
            TimerProgresses
        End If
    
    End If
        
    
    Exit Sub
failerror:
    Dim errDesc As String
    errDesc = GetLastError
    If Not ftpCancelThread And errDesc <> "" Then
        SetLastError EventError, 8, errDesc
        
        'RaiseEvent Error(8, App.Title, errDesc)
    End If
    Err.Clear
End Sub

Private Sub mData_Disconnected()
   ' If ConnectedState(RemoteToLocal) Or ConnectedState(ListingToFile) Then
        TerminateTransfer
   ' End If
End Sub

Private Sub mData_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    SetLastError RaiseError, Number, Description
End Sub

Private Sub mData_SendComplete()
On Error GoTo failerror
    Dim lSize As Double
    Dim inFileData() As Byte
    Dim inFileStr As String
    Dim ftpLapse As Single
    
    If ConnectedState(LocalToRemote) Then
        lSize = -1
        Do
                
            If CDbl(ftpSendIncrement * ftpSendPacketLen) > CDbl(ftpFileSize - ftpResumeByte) Then
    
                mData.Disconnect
                TerminateTransfer
            Else
            
                If ftpLargeFileMode Then
                    
                    If ((Not ftpCancelThread) And (Not PauseTransfers)) Then
               
                        If CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))) < ftpSendPacketLen Then
                            lSize = CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen)))
                        Else
                            lSize = ftpSendPacketLen
                        End If
                            
                        If lSize > 0 Then
                            ReDim inFileData(1 To lSize) As Byte
                            inFileData = modMemory.Convert(ftpReadTextStream.Read(lSize))

                            If ((Not myObj.ServerUsesCRLF) And (myObj.TransferType = ASCII)) And myObj.AssumeLineFeed Then
                                inFileStr = modMemory.Convert(Replace(modMemory.Convert(inFileStr), vbCr, ""))
                            ElseIf (myObj.ServerUsesCRLF And (myObj.TransferType = ASCII)) And (Not myObj.AssumeLineFeed) Then
                                inFileStr = modMemory.Convert(Replace(Replace(Replace(modMemory.Convert(inFileStr), vbLf, vbCrLf), vbCr & vbCr, vbCrLf), vbLf & vbLf, vbLf))
                            End If


                            ftpSendIncrement = ftpSendIncrement + 1
                            If Not mData.SendBytes(inFileData) Then Exit Do

                        End If
                    End If
                   TimerProgresses
        
                Else
        
                    If ((Not ftpCancelThread) And (Not PauseTransfers)) Then
    
                        If CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))) < ftpSendPacketLen Then
                            lSize = CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen)))
                        Else
                            lSize = ftpSendPacketLen
                        End If
                            
                        If lSize > 0 Then
    
                            ReDim inFileData(1 To lSize) As Byte
                            
                            Get #ftpReadFileNum, CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen)) + 1, inFileData
                            If ((Not myObj.ServerUsesCRLF) And (myObj.TransferType = ASCII)) And myObj.AssumeLineFeed Then
                                inFileStr = modMemory.Convert(Replace(modMemory.Convert(inFileStr), vbCr, ""))
                            ElseIf (myObj.ServerUsesCRLF And (myObj.TransferType = ASCII)) And (Not myObj.AssumeLineFeed) Then
                                inFileStr = modMemory.Convert(Replace(Replace(Replace(modMemory.Convert(inFileStr), vbLf, vbCrLf), vbCr & vbCr, vbCrLf), vbLf & vbLf, vbLf))
                            End If

                            ftpSendIncrement = ftpSendIncrement + 1
                            If Not mData.SendBytes(inFileData) Then Exit Do
    
                        End If
                    End If
                    TimerProgresses
                
                End If

            End If
                    

        Loop While mData.Connected And (Not ftpCancelThread)

    ElseIf ConnectedState(LocalToLocal) Then
        ConnectedState(DataStreaming) = True
        ftpLapse = Timer
        lSize = -1
        Do
            If ftpLargeFileMode Then
                If CDbl(ftpSendIncrement * ftpSendPacketLen) > CDbl(ftpFileSize - ftpResumeByte) Then
                    Exit Do
                Else
                    If ((Not ftpCancelThread) And (Not PauseTransfers)) Then
                        If CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))) < ftpSendPacketLen Then
                            inFileStr = ftpReadTextStream.Read(CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))))
                        Else
                            inFileStr = ftpReadTextStream.Read(ftpSendPacketLen)
                        End If
                        ftpWriteTextStream.Write inFileStr
                        ftpSendIncrement = ftpSendIncrement + 1
                    End If
                End If
            Else
                If CDbl(ftpSendIncrement * ftpSendPacketLen) > CDbl(ftpFileSize - ftpResumeByte) Then
                    Exit Do
                Else
                    If ((Not ftpCancelThread) And (Not PauseTransfers)) Then
                        If CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen))) < ftpSendPacketLen Then
                            ReDim inFileData(1 To CDbl(ftpFileSize - CDbl(ftpResumeByte + CDbl(ftpSendIncrement * ftpSendPacketLen)))) As Byte
                        Else
                            ReDim inFileData(1 To ftpSendPacketLen) As Byte
                        End If
                        Get #ftpReadFileNum, , inFileData()
                        Put #ftpWriteFileNum, , inFileData()
                        ftpSendIncrement = ftpSendIncrement + 1
                    End If
                End If
            End If
            If CSng(Timer - ftpLapse) > 0.02 Then
                ftpLapse = Timer
                TimerProgresses
            End If
            
        Loop While Not ftpCancelThread
        
        TerminateTransfer

    ElseIf ConnectedState(ListingToFile) And (Not PauseTransfers) Then
        If ftpType = URLTypes.Remote Or ftpType = URLTypes.File Then
            ConnectedState(DataStreaming) = True
            Dim myFolder As Object
            Set myFolder = myFso.GetFolder(Folder)
            Dim listItem As String
            Dim xItem As Object
            Dim kItem As Object
            ftpLapse = Timer
            For Each xItem In myFolder.SubFolders
                If ftpCancelThread Then Exit For
                listItem = GetWinDate(FileDateTime(xItem)) & " <DIR> " & xItem.name
                If ftpLargeFileMode Then
                    ftpWriteTextStream.Write vbCrLf & listItem
                Else
                    Put #ftpWriteFileNum, , vbCrLf & listItem
                End If
                If CSng(Timer - ftpLapse) > 0.02 Then
                    ftpLapse = Timer
                    Self "RaiseDataProgress", ProgressTypes.FileListing, ftpFileSize
                End If
            Next
            For Each kItem In myFolder.Files
                If ftpCancelThread Then Exit For
                listItem = GetWinDate(kItem.DateLastModified) & " " & kItem.Size & " " & kItem.name
                If ftpLargeFileMode Then
                    ftpWriteTextStream.Write vbCrLf & listItem
                Else
                    Put #ftpWriteFileNum, , vbCrLf & listItem
                End If
                If CSng(Timer - ftpLapse) > 0.02 Then
                    ftpLapse = Timer
                    Self "RaiseDataProgress", ProgressTypes.FileListing, ftpFileSize
                End If
            Next
            Set myFolder = Nothing
            TerminateTransfer
        End If
    End If
    Exit Sub
failerror:
    Dim errDesc As String
    If Err.Number = 5 Then
        Err.Clear
        Resume Next
    Else
        errDesc = GetLastError
        If Not ftpCancelThread And errDesc <> "" Then
            SetLastError EventError, 8, errDesc
            'RaiseEvent Error(8, App.Title, errDesc)
        End If
        Err.Clear
    End If
End Sub

Private Sub mTCP_Connected()

    ConnectedState(SocketNetwork) = True
    ConnectedState(LoggingInto) = False
    ConnectedState(FullyRemoting) = False

End Sub

Private Sub mTCP_Connection(Handle As Long)
    'not listening so must be a data arrival security mismatch, decline all
   ' mTCP.Decline Handle
End Sub

Private Sub mTCP_DataArriving()

    Dim delim As String
    Dim oneMsg As String
    Dim inData As String
'
    Dim Data() As Byte

    On Error GoTo catch

    If mTCP.ReadBytes(Data) Then

        inData = inData & modMemory.Convert(Data)
        
        delim = GetDelimiter(inData)
        Do While (Len(inData) > 0)
            delim = GetDelimiter(inData)
            inData = Replace(inData, Chr(0), "")
            If InStr(inData, delim) > 0 Then
                oneMsg = Trim(Left(inData, InStr(inData, delim) - 1))
                inData = Mid(inData, InStr(inData, delim) + Len(delim))
            Else
                oneMsg = Trim(inData)
                inData = ""
            End If
            Select Case Left(oneMsg, 3)
                Case "110", "202", "332", "421", "425", "426", "450", "451", "452", "500", "501", "502", "503", "504", "530", "532", "550", "551", "552", "553"
                    ftpEventError = ErrorReturns.Last
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False
                Case "227", "125", "150"
                    If Left(oneMsg, 3) = "227" Then
                        Dim pServer As String
                        Dim pPHigh As Long
                        Dim pPLow As Long
                        Dim pPort As Integer
                        pServer = StrReverse(oneMsg)
                        pServer = Mid(pServer, InStr(pServer, ")") + 1)
                        pServer = Left(pServer, InStr(pServer, "(") - 1)
                        pPLow = CLng(StrReverse(Left(pServer, InStr(pServer, ",") - 1)))
                        pServer = Mid(pServer, InStr(pServer, ",") + 1)
                        pPHigh = CLng(StrReverse(Left(pServer, InStr(pServer, ",") - 1)))
                        pServer = Replace(StrReverse(Mid(pServer, InStr(pServer, ",") + 1)), ",", ".")
                        'pPort = (Val(pPHigh) * 256) + Val(pPLow)
                       ' modBitValue.HiByte(pPort) = CByte(pPHigh)
                       ' modBitValue.LoByte(pPort) = CByte(pPLow)
                        
                        pPort = IntegerToUnsigned(Val("&H" & Padding(2, Hex(CByte(pPHigh)), "0") & Padding(2, Hex(CByte(pPLow)), "0")))
                        
                        mData.Host = pServer
                        mData.Port = pPort

                    End If
                    
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(DataStreaming) = True
                    
                Case "221"
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False
                    mTCP.Disconnect
                    
                Case "225", "226"
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False
                    
                Case "211"
                    ftpEventReason = oneMsg
                    ftpStatInfo = ftpStatInfo & oneMsg & "\n"
                    If Mid(oneMsg, 4, 1) <> "-" Then
                        ftpGatherType = GatherTypes.none
                        ftpEventError = ErrorReturns.none
                        ConnectedState(CommandRequest) = False
                    Else
                        ftpGatherType = GatherTypes.Stat
                    End If
                Case "214"
                    ftpEventReason = oneMsg
                    ftpHelpInformation = ftpHelpInformation & oneMsg & "\n"
                    If Mid(oneMsg, 4, 1) <> "-" Then
                        ftpGatherType = GatherTypes.none
                        ftpEventError = ErrorReturns.none
                        ConnectedState(CommandRequest) = False
                    Else
                        ftpGatherType = GatherTypes.Help
                    End If
                Case "215"
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ftpSystemInfo = Trim(Mid(oneMsg, 4))
                    ConnectedState(CommandRequest) = False
                Case "220"
                    ftpEventReason = oneMsg
                    ftpBannerMessage = ftpBannerMessage & oneMsg & "\n"
                    If Mid(oneMsg, 4, 1) <> "-" Then
                        ftpGatherType = GatherTypes.none
                        ftpEventError = ErrorReturns.none
                        ConnectedState(CommandRequest) = False
                    Else
                        ftpGatherType = GatherTypes.Banner
                    End If

                Case "230"
                    ftpEventReason = oneMsg
                    ftpMessageOfTheDay = ftpMessageOfTheDay & oneMsg & "\n"
                    If Mid(oneMsg, 4, 1) <> "-" Then
                        ftpGatherType = GatherTypes.none
                        ftpEventError = ErrorReturns.none
                        ConnectedState(CommandRequest) = False
                    Else
                        ftpGatherType = GatherTypes.MOTD
                    End If
                Case "257"
                    ftpFolder = oneMsg
                    ftpFolder = RemoveQuotedArg(ftpFolder, """", """")
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False
                    
                Case "250"
                    
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    If Mid(oneMsg, 4, 1) <> "-" Then
                        ConnectedState(CommandRequest) = False
                    End If
                Case "212", "213", "331", "350"
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False
                Case "200"
                    ftpEventError = ErrorReturns.none
                    ftpEventReason = oneMsg
                    ConnectedState(CommandRequest) = False

                Case Else
                    If oneMsg = "" Then Exit Do
                    Select Case ftpGatherType
                        Case GatherTypes.Banner
                            ftpBannerMessage = ftpBannerMessage & oneMsg & "\n"
                        Case GatherTypes.MOTD
                            ftpMessageOfTheDay = ftpMessageOfTheDay & oneMsg & "\n"
                        Case GatherTypes.Help
                            ftpHelpInformation = ftpHelpInformation & oneMsg & "\n"
                        Case GatherTypes.Stat
                            ftpStatInfo = ftpStatInfo & oneMsg & "\n"
                    End Select
            End Select
            If ftpEventError = ErrorReturns.Last Then
                AddToLog MessageTypes.Incorrect, oneMsg
                SetLastError EventError, 8, oneMsg
                ftpEventError = ErrorReturns.none
                'RaiseEvent Error(8, "NTAdvFTP", oneMsg)
                
            Else
                AddToLog MessageTypes.Incoming, oneMsg
                
            End If
            If Not (ftpEventError = ErrorReturns.Wait) Then DoOutCommands
        Loop
                
    End If
    
    Exit Sub
catch:
    If ftpEventError = ErrorReturns.Last Then
        SetLastError EventError, 8, ftpEventReason
        'RaiseEvent Error(8, App.Title, ftpEventReason)
    Else
        SetLastError EventError, 8, Err.Description
        'RaiseEvent Error(8, App.Title, Err.Description)
    End If
End Sub

Private Sub mTCP_Disconnected()
    On Error Resume Next
    If ftpType = ftp Then
        ftpEventReason = "Disconnected..."
        ftpEventError = False
        If ConnectedState(SocketNetwork) And Not ftpCancelThread Then
            SetLastError EventError, 8, ftpEventReason
           ' RaiseEvent Error(8, "NTAdvFTP", ftpEventReason)
            Err.Clear
        End If
        ConnectedState(SocketNetwork) = False
        ConnectedState(LoggingInto) = False
    Else
        ConnectedState(LocalHardDisk) = False
    End If
    ConnectedState(FullyRemoting) = False

    On Error GoTo 0
End Sub

Private Sub DoOutCommands()
    Static priorSend As String
    
    If mTCP.Connected Then
        Dim ftpCommand As String
        Do
        
            If ftpNextCommand <> "" Then
        
                ftpCommand = RemoveNextArg(ftpNextCommand, vbCrLf)
                Select Case NextArg(ftpCommand, " ")
                    Case "ACCT"
                        If (Not ConnectedState(FullyRemoting)) And (ftpEventError = ErrorReturns.none) Then
                            ConnectedState(LoggingInto) = True
                            If ((ftpUsername <> "") And (ftpPassword = "")) Then
                                ftpNextCommand = "USER " & ftpUsername & vbCrLf & ftpNextCommand
                            End If
                        End If
                    Case "USER"
                        ConnectedState(FullyRemoting) = False
                        ConnectedState(LoggingInto) = True
                        If ((ftpAccount = "") And (ftpPassword <> "")) Then
                            ftpNextCommand = "PASS " & ftpPassword & vbCrLf & ftpNextCommand
                        End If
                    Case "REIN"
                        ConnectedState(FullyRemoting) = False
                        If ((ftpUsername <> "") And (ftpPassword <> "")) Then
                            ftpNextCommand = "USER " & ftpUsername & vbCrLf & ftpNextCommand
                        ElseIf ((ftpAccount <> "") And (ftpPassword = "")) Then
                            ftpNextCommand = "ACCT " & ftpAccount & vbCrLf & ftpNextCommand
                        End If
                    Case "PASS"
                        If (Not ConnectedState(FullyRemoting)) And (ftpEventError = ErrorReturns.none) Then
                            ConnectedState(LoggingInto) = False
                            Select Case ftpTransferMode
                                Case TransferModes.ASCII
                                    ftpNextCommand = "TYPE A" & vbCrLf & ftpNextCommand
                                Case TransferModes.Binary
                                    ftpNextCommand = "TYPE I" & vbCrLf & ftpNextCommand
                            End Select
                            If (ftpSystemInfo = "") Then
                                ftpNextCommand = "SYST" & vbCrLf & ftpNextCommand
                            Else
                                If (ftpFolder <> "") Then
                                    ftpNextCommand = "CWD " & ftpFolder & vbCrLf & ftpNextCommand
                                Else
                                    ftpNextCommand = "PWD" & vbCrLf & ftpNextCommand
                                End If
                            End If
                        End If
                    Case "TYPE"
                        If (Not ConnectedState(FullyRemoting)) And (ftpEventError = ErrorReturns.none) Then
                            If (ftpNextCommand = "") Then
                                If (ftpSystemInfo = "") Then
                                    ftpNextCommand = "SYST" & vbCrLf & ftpNextCommand
                                Else
                                    If (ftpFolder <> "") Then
                                        ftpNextCommand = "CWD " & ftpFolder & vbCrLf & ftpNextCommand
                                    Else
                                        ftpNextCommand = "PWD" & vbCrLf & ftpNextCommand
                                    End If
                                End If
                            End If
                        End If
                    Case "SYST"
                        If (Not ConnectedState(FullyRemoting)) And (ftpEventError = ErrorReturns.none) Then
                            If (ftpFolder <> "") Then
                                ftpNextCommand = "CWD " & ftpFolder & vbCrLf & ftpNextCommand
                            Else
                                ftpNextCommand = "PWD" & vbCrLf & ftpNextCommand
                            End If
                        End If
                    Case "CWD", "PWD"
                        If (Not ConnectedState(FullyRemoting)) And (ftpEventError = ErrorReturns.none) Then
                            ConnectedState(FullyRemoting) = True
                        End If
                    Case "PASV"
                    Case "PORT"
                        'Stop

    
                    Case "STOR", "RETR", "LIST", "REST", "APPE"
    

                        If (ftpConnectionMode = "PASV") And (ftpEventError = ErrorReturns.none) Then

                            If mData.Connected Then mData.Disconnect

                            If Not mData.Connected Then
                                mData.Connect mData.Host, mData.Port

                            End If
                        Else
'                            Do Until mData.Connected
'                                DoLoop
'                            Loop
                        
                        End If

                        
                End Select
                If (ftpCommand <> "") And (ftpEventError = ErrorReturns.none) Then
                    
                    Do While (ftpEventError = ErrorReturns.Wait) And mTCP.Connected And (Not ftpCancelThread) And ConnectedState And Transfering
                        'DoEvents
                        DoLoop
                    Loop
                    ConnectedState(CommandRequest) = True
                    ftpEventError = ErrorReturns.Wait
                    ftpEventReason = GetLastError
                    AddToLog MessageTypes.Outgoing, ftpCommand
                    priorSend = ftpCommand & vbCrLf
                    If Not mTCP.SendBytes(modMemory.Convert(ftpCommand & vbCrLf)) Then
                        Exit Do
                    End If
    
                Else

                    Exit Do
                End If
                
                

            End If
        Loop While (ftpCommand <> "") And (ftpEventError = ErrorReturns.none)

    End If

End Sub

Private Sub mTCP_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    SetLastError RaiseError, Number, Description
End Sub
