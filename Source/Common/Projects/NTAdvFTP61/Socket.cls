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

Private WithEvents SecureSession As Crypto
Attribute SecureSession.VB_VarHelpID = -1
Private WithEvents wClient As ISocket
Attribute wClient.VB_VarHelpID = -1

Public Event Connected()
Public Event Disconnected()
Public Event DataArriving()
Public Event SendComplete()
Public Event Connection(ByRef Handle As Long)
Public Event Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

Private CLIENT_RANDOM As String
Private SERVER_RANDOM As String
Private SESSION_ID As String
Private SESSION_KEY As String
Private RSA_KEY_256V3 As String
Private RSA_KEY_128V3 As String
Private RSA_KEY_LEN As Long
Private PUBLIC_KEY As String
Private PRE_MASTER_KEY As String
Private MASTER_KEY As String

Private Const VERSION_MAJOR As Byte = &H3
Private Const VERSION_MINOR_0 As Byte = &H0
Private Const VERSION_MINOR_1 As Byte = &H1
Private Const TLS1_RT_CHANGE_CIPHER_SPEC As Byte = &H14 '20
Private Const TLS1_RT_ALERT As Byte = &H15              '21
Private Const TLS1_RT_HANDSHAKE As Byte = &H16          '22
Private Const TLS1_RT_APPLICATION_DATA As Byte = &H17   '23
Private Const TLS1_MT_HELLO_REQUEST As Byte = &H0       '0
Private Const TLS1_MT_CLIENT_HELLO As Byte = &H1        '1
Private Const TLS1_MT_SERVER_HELLO As Byte = &H2        '2
Private Const TLS1_MT_CERTIFICATE As Byte = &HB         '11
Private Const TLS1_MT_SERVER_KEY_EXCHANGE As Byte = &HC '12
Private Const TLS1_MT_CERTIFICATE_REQUEST As Byte = &HD '13
Private Const TLS1_MT_SERVER_DONE As Byte = &HE         '14
Private Const TLS1_MT_CERTIFICATE_VERIFY As Byte = &HF  '15
Private Const TLS1_MT_CLIENT_KEY_EXCHANGE As Byte = &H10 '16
Private Const TLS1_MT_FINISHED As Byte = &H14           '20
Private hSESSION_MD5 As Long  'Handle to MD5 Session Hash
Private hSESSION_SHA1 As Long  'Handle to SHA1 Session Hash
Private hCOPY_MD5 As Long  'Handle to Copy of MD5 Session Hash
Private hCOPY_SHA1 As Long  'Handle to Copy of SHA1 Session Hash


Private Layer As Long

Private cert As Certificate

Private Function AddRecHeader(strRecord As String) As String
    'Prepend record with record header
    Dim TLSHeader As String
    TLSHeader = SecureSession.msgType & Chr$(Int(Len(strRecord) / 256)) & Chr$(Len(strRecord) Mod 256)
    AddRecHeader = TLSHeader & strRecord
End Function

Private Function GetClientExchange() As String
    Dim sKey As String
    Dim sCipherSpec As String
    Dim sClientFinish As String
    Dim sClientFinishEnc As String
    'Create Master Key
    SecureSession.ClientRandom = CLIENT_RANDOM
    SecureSession.ServerRandom = SERVER_RANDOM
    'Get random Pre-Master Key
    Call SecureSession.GenRandom(46&)
    PRE_MASTER_KEY = Chr$(&H3) & Chr$(&H1) & SecureSession.RndBuffer
    SecureSession.InBuffer = PRE_MASTER_KEY
    'Set the server Public Key in Crypto
    SecureSession.PubKey = Left(PUBLIC_KEY, RSA_KEY_LEN)
    Call SecureSession.SetPubKey
    Call SecureSession.ExportMasterKey
    'Encrypted Pre-Master Key is now resident in OutBuffer
    '& decrypted Pre-Master Key is in InBuffer
'    PRE_MASTER_KEY = SecureSession.OutBuffer
    sKey = Chr$(TLS1_MT_CLIENT_KEY_EXCHANGE) & Chr$(&H0) & Chr$(&H1) & Chr$(&H2) _
        & Chr$(&H1) & Chr$(&H0) & StrReverse(SecureSession.OutBuffer)
    sCipherSpec = Chr$(TLS1_RT_CHANGE_CIPHER_SPEC) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & Chr$(&H0) & Chr$(&H1) & Chr$(&H1)

    
 
    'Hash data (Change Cipher Spec not included)
    'Call DebugPrintHex("Client Key Exchange (Encrypted)", sKey)
    hSESSION_MD5 = SecureSession.MD5_AddToHash(sKey, hSESSION_MD5)
    hSESSION_SHA1 = SecureSession.SHA1_AddToHash(sKey, hSESSION_SHA1)
    'Make a copy of the hashes
    hCOPY_MD5 = SecureSession.CopyHash(hSESSION_MD5)
    hCOPY_SHA1 = SecureSession.CopyHash(hSESSION_SHA1)
    'Concantenate copy of Hashes
    SecureSession.InBuffer = SecureSession.MD5_GetHash(hCOPY_MD5) _
        & SecureSession.SHA1_GetHash(hCOPY_SHA1)
   ' Call DebugPrintHex("Client MD5 + SHA1 Hashes", SecureSession.InBuffer)
    'Hash Session Hashes
    Call SecureSession.Hash_SHAMD5(False)
    sClientFinish = SecureSession.InBuffer
    'Call DebugPrintHex("Client Session Hash Before Encryption", sClientFinish)
    sClientFinishEnc = SecureSession.OutBuffer
   ' Call DebugPrintHex("Client Session Hash Encrypted", sClientFinishEnc)
    hSESSION_MD5 = SecureSession.MD5_AddToHash(sClientFinish, hSESSION_MD5)
    hSESSION_SHA1 = SecureSession.SHA1_AddToHash(sClientFinish, hSESSION_SHA1)
    GetClientExchange = AddRecHeader(sKey) & sCipherSpec & AddRecHeader(sClientFinishEnc)

    ' Call DebugPrintHex("Signature", GetClientExchange)
 '   Debug.Print "Change Cipher Spec Sent!"
End Function

Private Function GetClientHello() As String
    Call SecureSession.GenRandom(28&)
    CLIENT_RANDOM = GetUnixTime & SecureSession.RndBuffer
    '36 ciphers
    'GetClientHello = Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & CLIENT_RANDOM _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H48) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&HA) & Chr$(&HC0) & Chr$(&H14) & Chr$(&H0) & Chr$(&H88) & Chr$(&H0) & Chr$(&H87) & Chr$(&H0) & Chr$(&H38) _
        & Chr$(&HC0) & Chr$(&HF) & Chr$(&HC0) & Chr$(&H5) & Chr$(&H0) & Chr$(&H84) & Chr$(&H0) & Chr$(&H35) & Chr$(&H0) & Chr$(&H39) & Chr$(&HC0) & Chr$(&H7) & Chr$(&HC0) & Chr$(&H9) & Chr$(&HC0) & Chr$(&H11) _
        & Chr$(&HC0) & Chr$(&H13) & Chr$(&H0) & Chr$(&H45) & Chr$(&H0) & Chr$(&H44) & Chr$(&H0) & Chr$(&H33) & Chr$(&H0) & Chr$(&H32) & Chr$(&HC0) & Chr$(&HC) & Chr$(&HC0) & Chr$(&HE) & Chr$(&HC0) & Chr$(&H2) _
        & Chr$(&HC0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H96) & Chr$(&H0) & Chr$(&H41) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H5) & Chr$(&H0) & Chr$(&H2F) & Chr$(&HC0) & Chr$(&H8) & Chr$(&HC0) & Chr$(&H12) _
        & Chr$(&H0) & Chr$(&H16) & Chr$(&H0) & Chr$(&H13) & Chr$(&HC0) & Chr$(&HD) & Chr$(&HC0) & Chr$(&H3) & Chr$(&HFE) & Chr$(&HFF) & Chr$(&H0) & Chr$(&HA) & Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H2D) _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H13) & Chr$(&H0) & Chr$(&H11) & Chr$(&H0) & Chr$(&H0) & Chr$(&HE) & Chr$(&H77) & Chr$(&H77) & Chr$(&H77) & Chr$(&H2E) & Chr$(&H70) & Chr$(&H61) & Chr$(&H79) _
        & Chr$(&H70) & Chr$(&H61) & Chr$(&H6C) & Chr$(&H2E) & Chr$(&H63) & Chr$(&H6F) & Chr$(&H6D) & Chr$(&H0) & Chr$(&HA) & Chr$(&H0) & Chr$(&H8) & Chr$(&H0) & Chr$(&H6) & Chr$(&H0) & Chr$(&H17) & Chr$(&H0) _
        & Chr$(&H18) & Chr$(&H0) & Chr$(&H19) & Chr$(&H0) & Chr$(&HB) & Chr$(&H0) & Chr$(&H2) & Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H23) & Chr$(&H0) & Chr$(&H0)
    'Both TLS_RSA_WITH_RC4_128_MD5 & TLS_RSA_WITH_RC4_128_SHA
    GetClientHello = Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & CLIENT_RANDOM _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H5) _
        & Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0)
    'Only TLS_RSA_WITH_RC4_128_MD5
    'GetClientHello = Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & CLIENT_RANDOM _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H2) & Chr$(&H0) & Chr$(&H4) & Chr$(&H1) & Chr$(&H0) _
        & Chr$(&H0) & Chr$(Len(SNI)) & SNI
    'Only TLS_RSA_WITH_RC4_128_SHA
    'GetClientHello = Chr$(&H1) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1) _
        & CLIENT_RANDOM _
        & Chr$(&H0) & Chr$(&H0) & Chr$(&H2) & Chr$(&H0) & Chr$(&H5) & Chr$(&H1) & Chr$(&H0) _
        & Chr$(&H0) & Chr$(Len(SNI)) & SNI
    Mid$(GetClientHello, 3, 1) = Chr$(Int((Len(GetClientHello) - 4) / 256))
    Mid$(GetClientHello, 4, 1) = Chr$((Len(GetClientHello) - 4) Mod 256)
'    GetClientHello = AddRecHeader(GetClientHello)
End Function
'Private Function GetRequest() As String
'    Dim sTmp As String
'    sTmp = "GET /" & PAGENAME & vbCrLf '& " HTTP/1.1" & vbCrLf _
'        & "Host: " & SERVER & vbCrLf _
'        & "User-Agent: SSLClient/1.0 (Windows; U; Windows NT 6.0;)" & vbCrLf _
'        & "Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8" & vbCrLf _
'        & "Accept-Language: en-us,en;q=0.5" & vbCrLf _
'        & "Accept -Encoding: gzip , deflate" & vbCrLf _
'        & "Accept-Charset: ISO-8859-1,utf-8;q=0.7,*;q=0.7" & vbCrLf _
'        & "Connection: Keep -Alive" & vbCrLf & vbCrLf
'    If Port = 443 Then
''        Debug.Print sTmp
'        SecureSession.InBuffer = sTmp
'        Call SecureSession.TLSEncrypt
'        GetRequest = AddRecHeader(SecureSession.OutBuffer)
'    Else
'        GetRequest = sTmp
'    End If
'End Function

Private Function GetUnixTime() As String
    Dim M%, N%
    Dim temp$
    GetUnixTime = String$(4, vbNull)
    temp$ = Hex(DateDiff("s", "01/01/1970", Date$))
    M% = 1
    For N% = 1 To Len(temp$) Step 2
        Mid$(GetUnixTime, M%, 1) = Chr$(Val("&H" & Mid$(temp$, N%, 2)))
        M% = M% + 1
    Next N%
End Function

Private Function VerifySignature(sBuff As String) As Boolean
    Dim CertData As String
    Dim CertKey As String
    Dim SigData As String
    Dim sSearch As String
    Dim lPntr1 As Long
    Dim lPntr2 As Long
    Dim lPntr3 As Long
    Dim llen As Long

    'This routine returns the first part of the certificate only (w/o signature)
    lPntr1 = InStr(sBuff, Chr$(&H30) & Chr$(&H82)) + 4
    If lPntr1 < 5 Then Exit Function
    VerifySignature = True
    lPntr2 = InStr(lPntr1, sBuff, Chr$(&H30) & Chr$(&H82))
    llen = Asc(Mid$(sBuff, lPntr2 + 2, 1)) * 256 + Asc(Mid$(sBuff, lPntr2 + 3, 1)) + 4
    CertData = Mid$(sBuff, lPntr2, llen)
    lPntr2 = lPntr2 + llen
    SecureSession.InBuffer = CertData
    Dim slen As Long
    
    If lPntr1 > 0 Then

        RSA_KEY_LEN = 256
        sSearch = RSA_KEY_256V3
        slen = 7

    Else
        lPntr1 = InStr(sBuff, RSA_KEY_128V3)
        If lPntr1 > 0 Then
            RSA_KEY_LEN = 128
            sSearch = RSA_KEY_128V3
            slen = 5
        Else
            VerifySignature = False
            Exit Function
        End If
    End If
    
    'Save Server Public Key
    Dim istart As Integer
    istart = lPntr1
    
    lPntr1 = InStr(istart, sBuff, sSearch)


    If lPntr1 > 0 Then
        lPntr1 = lPntr1 + Len(sSearch) + slen
        PUBLIC_KEY = StrReverse(Mid$(sBuff, lPntr1, RSA_KEY_LEN))
   '     Call DebugPrintHex("Server Public Key", PUBLIC_KEY)
        'Find next certificate Public Key
        lPntr1 = InStr(lPntr1, sBuff, sSearch)
        If lPntr1 > 0 Then 'Use key from next Cert in chain
            CertKey = StrReverse(Mid$(sBuff, lPntr1 + Len(sSearch) + slen, RSA_KEY_LEN))
        Else 'Use Server Public key
            CertKey = PUBLIC_KEY
        End If
    Else
        VerifySignature = False
        Exit Function
    End If
'    Do Until lPntr1 = 0
'        lPntr1 = lPntr1 + Len(sSearch)
'        lPntr3 = lPntr1
'        lPntr1 = InStr(lPntr1, sBuff, sSearch)
'    Loop
'    CertKey = StrReverse(Mid$(sBuff, lPntr3, RSA_KEY_LEN))
  '  Call DebugPrintHex("CA Public Key", CertKey)

  
    SecureSession.PubKey = CertKey
    If Not SecureSession.SetPubKey Then
        VerifySignature = False
        Exit Function
    End If
    If RSA_KEY_LEN = 256 Then
        sSearch = Chr$(&H3) & Chr$(&H82) '& Chr$(&H1) & Chr$(&H1) & Chr$(&H0)
        slen = slen + 3
    ElseIf RSA_KEY_LEN = 128 Then
        sSearch = Chr$(&H3) & Chr$(&H81) '& Chr$(&H81) & Chr$(&H0)
        slen = slen + 2
    End If
    lPntr3 = InStr(lPntr2, sBuff, sSearch)
    If lPntr3 > 0 Then
        lPntr3 = lPntr3 + Len(sSearch) + slen
        SigData = StrReverse(Mid$(sBuff, lPntr3, RSA_KEY_LEN))
    Else
        VerifySignature = False
        Exit Function
    End If
  '  Call DebugPrintHex("Signature", SigData)
    SecureSession.Signature = SigData
    VerifySignature = SecureSession.VerifySig
End Function
Private Sub Class_Initialize()
    
    Set wClient = New ISocket
    Set ISocket = wClient
    ResetSSLVariables
     
End Sub

Private Sub Class_Terminate()
    Layer = -2
    wClient_DataArriving
    Set SecureSession = Nothing
    Set wClient = Nothing
End Sub

Private Function ISocket_Accept(Optional Handle As Long = 0&) As Boolean
    ISocket_Accept = Me.Accept(Handle)
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

Private Function ISocket_Decline(Optional Handle As Long = 0&) As Boolean
    ISocket_Decline = Me.Decline(Handle)
End Function

Private Sub ISocket_Disconnect()
    Me.Disconnect
End Sub

Private Property Get ISocket_Handle() As Long
    ISocket_Handle = Me.Handle
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

Private Property Get ISocket_LocalIPs() As Collection
    Set ISocket_LocalIPs = Me.LocalIPs
End Property

Private Property Let ISocket_Port(ByVal RHS As Long)
    Me.Port = RHS
End Property

Private Property Get ISocket_Port() As Long
    ISocket_Port = Me.Port
End Property

Private Function ISocket_Read() As String
    ISocket_Read = Me.Read
End Function

Private Function ISocket_ReadBytes(Data() As Byte) As Boolean
    ISocket_ReadBytes = Me.ReadBytes(Data)
End Function

Private Function ISocket_ReadNextPart() As Long
    ISocket_ReadNextPart = Me.ReadNextPart
End Function

Private Function ISocket_ReadString(Data As String) As Boolean
    ISocket_ReadString = ISocket.ReadString(Data)
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

Private Sub ISocket_Send(ByVal Data As String)
    Me.Send Data
End Sub

Private Function ISocket_SendBytes(Data() As Byte) As Boolean
    ISocket_SendBytes = Me.SendBytes(Data)
End Function

Private Function ISocket_SendNextPart() As Long
    ISocket_SendNextPart = Me.SendNextPart
End Function

Private Property Let ISocket_SendRate(ByVal RHS As Long)
    Me.SendRate = RHS
End Property

Private Property Get ISocket_SendRate() As Long
    ISocket_SendRate = Me.SendRate
End Property

Private Function ISocket_SendString(Data As String) As Boolean
    ISocket_SendString = Me.SendString(Data)
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


Private Sub SecureSession_Error(ByVal Number As Long, Description As String, ByVal Source As String)
    Debug.Print "Error: " & Number & " " & Description
    
    RaiseEvent Error(Number, Source, Description)
End Sub

Private Sub wClient_Connected()

    If Not SSL Then
        RaiseEvent Connected

    Else
    Debug.Print "Connected"
        'Send Client Hello (Unencrypted)
        Dim strSend As String
        Layer = 0
        hSESSION_MD5 = 0
        hSESSION_SHA1 = 0
        strSend = GetClientHello
        
        hSESSION_MD5 = SecureSession.MD5_AddToHash(strSend, hSESSION_MD5)
        hSESSION_SHA1 = SecureSession.SHA1_AddToHash(strSend, hSESSION_SHA1)
        SecureSession.msgType = Chr$(TLS1_RT_HANDSHAKE) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1)
        
        SecureSession.IncSendSeqNum (True)
        
        wClient.SendString AddRecHeader(strSend)

        RaiseEvent Connected
    End If

End Sub

Private Sub wClient_Connection(Handle As Long)
    RaiseEvent Connection(Handle)
End Sub

Private Sub wClient_DataArriving()

    If Not SSL Then

        RaiseEvent DataArriving

    Else
        On Error GoTo errcatch

        'This routine is re-entrant, hence the next 3 variable must be static
        Static InBuff As New Stream
        Static Header() As Byte
        Static SSLLength As Integer

        If Layer = -2 Then
            Erase Header
            Set InBuff = Nothing
        Else
            Dim bData() As Byte

            Dim SSLHandShakeLength As Integer

            Dim SessionLen As Long
            Dim sCipher As String

            Dim strSHAMD5 As String

            If wClient.ReadBytes(bData) Then

                InBuff.Concat bData

            End If

            Do
                If SSLLength = 0 And InBuff.Length >= 5 Then 'New record
                    Header() = InBuff.Partial(0, 5)

                    SSLLength = Val("&H" & Padding(2, Hex(Header(4)), "0") & _
                                           Padding(2, Hex(Header(5)), "0"))
                    InBuff.Push 5
                End If

                If InBuff.Length >= SSLLength And SSLLength > 0 Then

                    SecureSession.sInBuffer.Length = SSLLength
                    SecureSession.sInBuffer.Placeat InBuff.Partial(0, SSLLength)

                    InBuff.Push SSLLength
                    SSLLength = 0

                    If Header(1) = TLS1_RT_HANDSHAKE Then

                        'record complete - Process it
                        Select Case Layer
                            Case 0 'this layer processes the Server Hello
                                Dim sTemp As New Stream

                                sTemp.Clone SecureSession.sInBuffer

                                hSESSION_MD5 = SecureSession.MD5_AddToHash(SecureSession.InBuffer, hSESSION_MD5)
                                hSESSION_SHA1 = SecureSession.SHA1_AddToHash(SecureSession.InBuffer, hSESSION_SHA1)

                                Do

                                    SSLHandShakeLength = Val("&H" & Padding(2, Hex(sTemp.Partial(1, 1)(1)), "0") & _
                                                                       Padding(2, Hex(sTemp.Partial(2, 1)(1)), "0") & _
                                                                       Padding(2, Hex(sTemp.Partial(3, 1)(1)), "0"))

                                    SSLHandShakeLength = SSLHandShakeLength + 4

                                    Select Case sTemp.Partial(0, 1)(1)
                                        Case TLS1_MT_SERVER_HELLO '1
                                            SERVER_RANDOM = ToString(sTemp.Partial(6, 32))
                                            SessionLen = sTemp.Partial(38, 1)(1)
                                            SESSION_ID = ToString(sTemp.Partial(39, SessionLen))
                                            SecureSession.CipherToUse = ToString(sTemp.Partial(39 + SessionLen, 2))

                                            If sTemp.Length > SSLHandShakeLength Then
                                                sTemp.Push SSLHandShakeLength
                                                SSLHandShakeLength = 0
                                            End If

                                        Case TLS1_MT_CERTIFICATE '2

                                            
 
                                            
                                            Dim certslen As Long
                                            Dim certlen As Long
                                            certslen = Val("&H" & Padding(2, Hex(sTemp.Partial(4, 1)(1)), "0") & _
                                                                Padding(2, Hex(sTemp.Partial(5, 1)(1)), "0") & _
                                                                Padding(2, Hex(sTemp.Partial(6, 1)(1)), "0")) - 3

                                            certlen = Val("&H" & Padding(2, Hex(sTemp.Partial(7, 1)(1)), "0") & _
                                                                Padding(2, Hex(sTemp.Partial(8, 1)(1)), "0") & _
                                                                Padding(2, Hex(sTemp.Partial(9, 1)(1)), "0"))
'
''                                            Debug.Print certslen; certlen
''

                                          ' WriteFile AppPath & "cert.cer", ToString(sTemp.Partial(7, certslen))
''
                                           ' RunProcess "rundll32.exe", "cryptext.dll,CryptExtOpenCER " & AppPath & "cert.cer"

                                            Set cert = New Certificate
                                            cert.LoadCertificate Stream(sTemp.Partial(10, certslen))

'                                            Dim fields As NTNodes10.Collection
'                                            Set fields = Cert.fields
'                                            Dim num As Integer
'                                            For num = 1 To fields.count
'                                                Debug.Print fields.Key(num) & " " & Cert.Names(fields.Key(num)) & " " & Cert.Terms(fields.Key(num))
'                                            Next
'
'                                            Debug.Print "Version " & Cert.HexStream(Cert.fields(Version))
'                                            Debug.Print
'                                            Debug.Print "SerialNumber " & Cert.HexStream(Cert.fields(SerialNumber))
'                                            Debug.Print
'                                            Debug.Print "Algorithm " & Cert.HexStream(Cert.fields(Algorithm))
'                                            Debug.Print
'                                            Debug.Print "Issuer " & Cert.HexStream(Cert.fields(Issuer))
'                                            Debug.Print
'                                            Debug.Print "Validity " & Cert.HexStream(Cert.fields(Validity))
'                                            Debug.Print
'                                            Debug.Print "Subject " & Cert.HexStream(Cert.fields(Subject))
'                                            Debug.Print
'                                            Debug.Print "PublicKey " & Cert.HexStream(Cert.fields(PublicKey))
'                                            Debug.Print
'                                            Debug.Print "Extensions " & Cert.HexStream(Cert.fields(Extensions))
'                                            Debug.Print
'                                            Debug.Print "SignatureAlgorithm " & Cert.HexStream(Cert.fields(SignatureAlgorithm))
'                                            Debug.Print
'                                            Debug.Print "Signature " & Cert.HexStream(Cert.fields(Signature))
'                                            Debug.Print
                                            
'
'                                           If Not VerifySignature(ToString(sTemp.Partial(10, certlen))) Then   'If Not VerifySignature(ToString(sTemp.Partial(10, certlen))) Then
'                                                Debug.Print "Unable to verify signature!"
'                                            End If

                                           ' RSA_KEY_LEN = 256
''
                                            SecureSession.InBuffer = ToString(sTemp.Partial(10, certslen)) '
                                          '  SecureSession.InBuffer = ToString(sTemp.Partial(certlen + 10, (certslen + 10) - (certlen + 10))) 'CertData


                                            
                                            PUBLIC_KEY = StrReverse(ToString(cert.Fields(PublicKeyBlock)))
                                         '   PUBLIC_KEY = StrReverse(Mid$(sBuff, lPntr1)) ', RSA_KEY_LEN))
                                          ' ' CertKey = PUBLIC_KEY
                                          
                                            SecureSession.PubKey = PUBLIC_KEY '
                                           ' SecureSession.PubKey = ToString(sTemp.Partial(0, certlen))  'CertKey

                                            'Debug.Print SecureSession.sInBuffer.Length

                                          '  SecureSession.SetPubKey

                                            'SecureSession.Signature = StrReverse(ToString(cert.Fields(Signature)))
                                            SecureSession.Signature = ToString(sTemp.Partial(10, certlen))
'
'                                            If Not VerifySignature(StrReverse(ToString(cert.Fields(Signature)))) Then
'                                                Debug.Print "Unable to verify signature!"
'                                            End If

                                           If Not VerifySignature(ToString(sTemp.Partial(10, certlen))) Then   'If Not VerifySignature(ToString(sTemp.Partial(10, certlen))) Then
                                                Debug.Print "Unable to verify signature!"
                                            End If


                                            If sTemp.Length > SSLHandShakeLength Then
                                                sTemp.Push SSLHandShakeLength
                                                SSLHandShakeLength = 0
                                            End If

                                        Case TLS1_MT_SERVER_DONE '14

                                            Dim strSend As Stream
                                            Set strSend = Stream(GetClientExchange)

                                            Layer = Layer + 1
                                            
                                            SecureSession.IncSendSeqNum
                                           
                                            wClient.SendBytes strSend.Partial

                                            If sTemp.Length > SSLHandShakeLength Then
                                                sTemp.Push SSLHandShakeLength
                                                SSLHandShakeLength = 0

                                            End If
                                    End Select

                                Loop While ((sTemp.Length > 0) And (SSLHandShakeLength = 0)) And Connected And (Layer < 1)

                                Set sTemp = Nothing
                            Case 1 'this layer processes the Server Key Exchange record

                                Call SecureSession.TLSDecrypt
                                SecureSession.IncRecvSeqNum
                                'Save Decrypted Session Hash
                                strSHAMD5 = SecureSession.OutBuffer

                                SecureSession.sInBuffer.Reset
                                SecureSession.sInBuffer.Concat ToBytes(SecureSession.MD5_GetHash(hSESSION_MD5))
                                SecureSession.sInBuffer.Concat ToBytes(SecureSession.SHA1_GetHash(hSESSION_SHA1))

                                Call SecureSession.Hash_SHAMD5(True)
                                If strSHAMD5 <> SecureSession.InBuffer Then
                                    Debug.Print "Session Hash Appears to be Compromised!"
                                End If
                                SecureSession.msgType = Chr$(TLS1_RT_APPLICATION_DATA) & Chr$(VERSION_MAJOR) & Chr$(VERSION_MINOR_1)
                                Layer = Layer + 1
                                'If Not wClient.State(States.Reading) Then
                                    RaiseEvent SendComplete
                            '    End If

                        End Select

                    ElseIf Header(1) = TLS1_RT_CHANGE_CIPHER_SPEC Then

                        'Cipher Change Spec; Do nothing
                        Debug.Print "Change Cipher Spec Received!"
                      '  SecureSession.IncSendSeqNum (True) 'Reset Send Sequence Number
                        SecureSession.IncRecvSeqNum (True) 'Reset Recv Sequence Number

                    ElseIf Header(1) = TLS1_RT_ALERT Then 'Problem?

                        Select Case SecureSession.sInBuffer.Partial(0, 1)(1)
                            Case 1
                                If SecureSession.sInBuffer.Partial(1, 1)(1) > 0 Then
                                    Debug.Print "ALERT WARNING " & SecureSession.ErrorCode(SecureSession.sInBuffer.Partial(1, 1)(1))
                                End If
                            Case 2
                                Debug.Print "FATAL ALERT " & SecureSession.ErrorCode(SecureSession.sInBuffer.Partial(1, 1)(1))
                                Call Disconnect
                        End Select
                    ElseIf Header(1) = TLS1_RT_APPLICATION_DATA Then

                        If Layer > 1 Then
                           Call SecureSession.TLSDecrypt
                            SecureSession.IncRecvSeqNum

                            If SecureSession.sOutBuffer.Length > 0 Then
                                wClient.AutoRatedEvent AutoRated.ReadEvent

                                RaiseEvent DataArriving
                            End If

                        End If

                    End If
                End If

            Loop While ((InBuff.Length > 0) And (SSLLength = 0)) And (Not (SSLLength = 0 And InBuff.Length < 6)) And Connected

        End If

    End If

Exit Sub
errcatch:
    If Err Then
        RaiseEvent Error(Err.Number, Err.Source, Err.Description)
        Err.Clear

    End If

End Sub

Private Sub wClient_Disconnected()
    Debug.Print "Disconnected"
    RaiseEvent Disconnected
End Sub

Private Sub wClient_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    Debug.Print "Error(" & Number & ", " & Description & ")"
    RaiseEvent Error(Number, Source, Description)
End Sub

Private Sub wClient_SendComplete()
    Debug.Print "SendComplete"
    wClient.AutoRatedEvent AutoRated.SendEvent
    If SSL Then
        If Layer > 1 Then
            RaiseEvent SendComplete
        End If
    Else
        RaiseEvent SendComplete
    End If
End Sub

Public Function Accept(Optional Handle As Long = 0&) As Boolean
    Accept = ISocket.Accept(Handle)
End Function

Public Property Let AutoRate(ByVal RHS As Boolean)
    ISocket.AutoRate = RHS
End Property

Public Property Get AutoRate() As Boolean
    AutoRate = ISocket.AutoRate
End Property

Public Function Connect(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
'    SSL = True
'
'    Dim txt As String
'    txt = ReadFile("C:\Temp\Temp.cer")
'
'    Debug.Print VerifySignature(txt)
'    Debug.Print HexStream(SecureSession.InBuffer)
'    Debug.Print HexStream(SecureSession.Signature)

    If ISocket.Connected Then ISocket.Disconnect
    
    Connect = ISocket.Connect(nHost, nPort)
End Function

Public Property Get Connected() As Boolean
    Connected = ISocket.Connected
End Property

Public Function Decline(Optional Handle As Long = 0&) As Boolean
    Decline = ISocket.Decline(Handle)
End Function

Public Sub Disconnect()
    ISocket.Disconnect
End Sub

Public Property Get Handle() As Long
    Handle = ISocket.Handle
End Property

Public Property Let Host(ByVal RHS As String)
    ISocket.Host = RHS
End Property

Public Property Get Host() As String
    Host = ISocket.Host
End Property

Public Function Listen(Optional ByVal nHost As String = "", Optional ByVal nPort As Long = 0&) As Boolean
    Listen = ISocket.Listen(nHost, nPort)
End Function

Public Property Get Listening() As Boolean
    Listening = ISocket.Listening
End Property

Public Function LocalHost() As String
    LocalHost = ISocket.LocalHost
End Function

Public Function LocalIP() As String
    LocalIP = ISocket.LocalIP
End Function

Public Function LocalIPs() As Collection
    Set LocalIPs = ISocket.LocalIPs
End Function

Public Property Let Port(ByVal RHS As Long)
    ISocket.Port = RHS
End Property

Public Property Get Port() As Long
    Port = ISocket.Port
End Property

Public Function Read() As String
    If Not SSL Then
        ISocket.ReadString Read
    Else

        wClient.AutoRatedEvent AutoRated.readCalls

        Read = SecureSession.OutBuffer

        wClient.AutoRatedEvent AutoRated.ReadWidth

    End If
End Function

Public Function ReadBytes(Data() As Byte) As Boolean
    If Not SSL Then
        ReadBytes = ISocket.ReadBytes(Data)
    Else
        wClient.AutoRatedEvent AutoRated.readCalls

        Erase Data

        Data = SecureSession.sOutBuffer.Partial
        ReadBytes = (SecureSession.sOutBuffer.Length > 0)

        wClient.AutoRatedEvent AutoRated.ReadWidth

    End If
End Function

Public Function ReadNextPart() As Long
    ReadNextPart = ISocket.ReadNextPart
End Function

Public Function ReadString(Data As String) As Boolean
    If Not SSL Then
        ReadString = ISocket.ReadString(Data)
    Else

        wClient.AutoRatedEvent AutoRated.readCalls

        Data = SecureSession.OutBuffer
        ReadString = (SecureSession.sOutBuffer.Length > 0)

        wClient.AutoRatedEvent AutoRated.ReadWidth
    End If
End Function

Public Function Ready() As Boolean
    Ready = ISocket.Ready
End Function

Public Property Let RecvRate(ByVal RHS As Long)
    ISocket.RecvRate = RHS
End Property

Public Property Get RecvRate() As Long
    RecvRate = ISocket.RecvRate
End Property

Public Function RemoteIP() As String
    RemoteIP = ISocket.RemoteIP
End Function

Private Sub ResetSSLVariables()

    CLIENT_RANDOM = ""
    SERVER_RANDOM = ""
    SESSION_ID = ""
    SESSION_KEY = ""
    RSA_KEY_256V3 = ""
    RSA_KEY_256V3 = Chr(&H30) + Chr(&H82) '+ Chr(&H1) + Chr(&HA) + Chr(&H2) + Chr(&H82) + Chr(&H1) + Chr(&H1) + Chr(&H0)
    RSA_KEY_128V3 = ""
    RSA_KEY_128V3 = Chr(&H30) + Chr(&H81) '+ Chr(&H89) + Chr(&H2) + Chr(&H81) + Chr(&H81) + Chr(&H0)
    RSA_KEY_LEN = 0
    PUBLIC_KEY = ""
    PRE_MASTER_KEY = ""
    MASTER_KEY = ""

    hSESSION_MD5 = 0
    hSESSION_SHA1 = 0
    hCOPY_MD5 = 0
    hCOPY_SHA1 = 0

    Layer = 0

    If SSL Then
        Set SecureSession = Nothing
        Set SecureSession = New Crypto
    Else
        Set SecureSession = Nothing
    End If
    If Not cert Is Nothing Then Set cert = Nothing

End Sub

Public Function ResolveIP(ByVal Host As String) As String
    ResolveIP = ISocket.ResolveIP(Host)
End Function

Public Sub Send(ByVal Data As String)
    If Not SSL Then
        wClient.SendBytes modMemory.Convert(Data)
    Else
        wClient.AutoRatedEvent AutoRated.SendCalls

        SecureSession.InBuffer = Data
        Call SecureSession.TLSEncrypt
        Data = AddRecHeader(SecureSession.OutBuffer)
        wClient.SendBytes modMemory.Convert(Data)
        SecureSession.IncSendSeqNum

        wClient.AutoRatedEvent AutoRated.SendWidth
    End If
End Sub

Public Function SendBytes(Data() As Byte) As Boolean
    If Not SSL Then
        SendBytes = wClient.SendBytes(Data)
    Else
        wClient.AutoRatedEvent AutoRated.SendCalls

        SecureSession.InBuffer = modMemory.Convert(Data)
        Call SecureSession.TLSEncrypt
        Data = modMemory.Convert(AddRecHeader(SecureSession.OutBuffer))
        SendBytes = wClient.SendBytes(Data)
        SecureSession.IncSendSeqNum

        wClient.AutoRatedEvent AutoRated.SendWidth
    End If
End Function

Public Function SendNextPart() As Long
    SendNextPart = ISocket.SendNextPart
End Function

Public Property Let SendRate(ByVal RHS As Long)
    ISocket.SendRate = RHS
End Property

Public Property Get SendRate() As Long
    SendRate = ISocket.SendRate
End Property

Public Function SendString(Data As String) As Boolean
    If Not SSL Then
        SendString = wClient.SendBytes(modMemory.Convert(Data))
    Else
        wClient.AutoRatedEvent AutoRated.SendCalls

        SecureSession.InBuffer = Data
        Call SecureSession.TLSEncrypt
        Data = AddRecHeader(SecureSession.OutBuffer)
        SendString = wClient.SendBytes(modMemory.Convert(Data))
        SecureSession.IncSendSeqNum

        wClient.AutoRatedEvent AutoRated.SendWidth

    End If

End Function

Public Property Get State(bBit As Byte) As Boolean
    State = ISocket.State(bBit)
End Property

Public Property Let timeout(ByVal RHS As Long)
    ISocket.timeout = RHS
End Property

Public Property Get timeout() As Long
    timeout = ISocket.timeout
End Property

Public Property Let Transmission(ByVal RHS As Direction)
    ISocket.Transmission = RHS
End Property

Public Property Get Transmission() As Direction
    Transmission = ISocket.Transmission
End Property

Public Function WhoAmI(ByVal sock As Long) As String
    WhoAmI = ISocket.WhoAmI(sock)
End Function

Public Function WhosIs(ByVal sock As Long) As String
    WhosIs = ISocket.WhosIs(sock)
End Function


Public Property Get SSL() As Boolean
    SSL = (Not (SecureSession Is Nothing))
End Property

Public Property Let SSL(ByVal RHS As Boolean)
    If RHS Then
        If (SecureSession Is Nothing) Then
            Set SecureSession = New Crypto

        End If
        ResetSSLVariables
    ElseIf (Not RHS) Then
        If (Not (SecureSession Is Nothing)) Then
            Set SecureSession = Nothing
        End If
    End If
End Property


