VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmIdent 
   BorderStyle     =   0  'None
   Caption         =   "Ident Deamon"
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2640
   ControlBox      =   0   'False
   Icon            =   "frmIdent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIdent.frx":0442
   ScaleHeight     =   570
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckRoute 
      Index           =   0
      Left            =   2070
      Tag             =   "ROUTE"
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBroadcast 
      Left            =   1560
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1065
      Top             =   60
   End
   Begin MSWinsockLib.Winsock sckListen 
      Index           =   0
      Left            =   570
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckIdent 
      Index           =   0
      Left            =   75
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmIdent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Public IdleTimeout As Long
'IdleTimeout - The amount of seconds to wait before a connection is closed.

Public WinsOnlySys As Boolean
'WinsOnlySys - Return all versions of windows as one value 'WINS' when true.

Public IncludeComp As Boolean
'IncludeComp - Include the computer name in the user name when this is true.

Public StandsAlone As Boolean
'StandsAlone - When true do not allow interactions with itself on a network.

Public StandOnNeed As Boolean
'StandOnNeed - When true, listens only when ProcessNeed conditions are meet.

Public ProcessNeed As Collection
'ProcessNeed - A list of exe names or titles that while found listen occurs.

Public AdapterAddr As Collection
'AdapterAddr - A definition list of IPs or value 'every' which to listen on.

Public UserNameSID As Boolean
'UserNameSID - When true, respond the username accounts security identifyer.

Public ServicePort As Collection


Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Public Function UnsignedToLong(Value As Double) As Long
 '
 ' This function takes a Double containing a value in the*
 ' range of an unsigned Long and returns a Long that you*
 ' can pass to an API that requires an unsigned Long
 '
 If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    
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

Public Function IP4ToIP2(ByVal str As String) As String

    Dim b1 As Byte
    Dim B2 As Byte
    Dim b3 As Byte
    Dim b4 As Byte

    b1 = CByte(RemoveNextArg(str, "."))
    B2 = CByte(RemoveNextArg(str, "."))
    b3 = CByte(RemoveNextArg(str, "."))
    b4 = CByte(RemoveNextArg(str, "."))

    
    IP4ToIP2 = IntegerToUnsigned(val("&H" & Padding(2, Hex(b1), "0") & Padding(2, Hex(B2), "0"))) & ", " & IntegerToUnsigned(val("&H" & Padding(2, Hex(b3), "0") & Padding(2, Hex(b4), "0")))

End Function

Public Function IP2ToIP4(ByVal str As String) As String

    Dim i1 As Integer
    Dim i2 As Integer
    i1 = UnsignedToInteger(RemoveNextArg(str, ","))
    i2 = UnsignedToInteger(RemoveNextArg(str, ","))
    
    IP2ToIP4 = val("&H" & Left(Padding(4, Hex(i1), "0"), 2)) & "." & val("&H" & Right(Padding(4, Hex(i1), "0"), 2)) & "." & val("&H" & Left(Padding(4, Hex(i2), "0"), 2)) & "." & val("&H" & Right(Padding(4, Hex(i2), "0"), 2))

End Function

'Public Function IP4ToIP2(ByVal str As String) As String
'    Dim d1 As Long
'    Dim d2 As Long
'    Dim d3 As Long
'    Dim d4 As Long
'    Dim b1 As Long
'    Dim B2 As Long
'
'    str = Replace(str, ".", " , ")
'    d1 = CLng(NextArg(str, ","))
'    d2 = CLng(NextArg(RemoveArg(str, ","), ","))
'    d3 = CLng(NextArg(RemoveArg(RemoveArg(str, ","), ","), ","))
'    d4 = CLng(NextArg(RemoveArg(RemoveArg(RemoveArg(str, ","), ","), ","), ","))
'
'    If d1 > d2 Then Swap d1, d2
'    If d3 > d4 Then Swap d3, d4
'    If d2 > d4 Then
'        Swap d2, d4
'    Else
'        Swap d2, d3
'    End If
'    If d3 > d4 Then Swap d3, d4
'    If d1 > d2 Then Swap d1, d2
'    If d1 > d3 Then
'        Swap d1, d3
'        Swap d2, d4
'    End If
'    IP4ToIP2 = ((d1 * 256) + d2) & " , " & ((d3 * 256) + d4)
'End Function
'
'Public Function IP2ToIP4(ByVal str As String) As String
'    Dim d1 As Long
'    Dim d2 As Long
'    Dim d3 As Long
'    Dim d4 As Long
'    d3 = CLng(NextArg(str, ","))
'    d4 = CLng(NextArg(RemoveArg(str, ","), ","))
'    d1 = CLng(CLng(d3) Mod CLng(256))
'    d2 = CLng(CLng(d4) Mod CLng(256))
'    d3 = CLng(CLng(d3 - d1) / CLng(256))
'    d4 = CLng(CLng(d4 - d2) / CLng(256))
'    IP2ToIP4 = d1 & "." & d2 & "." & d3 & "." & d4
'End Function

Public Function ExistsInSetting(ByRef Settings As Collection, ByVal str As String) As Boolean
    On Error Resume Next
    Dim tmp As String
    tmp = Settings(str)
    If (Err.Number = 0) Then
        ExistsInSetting = True
    Else
        Err.Clear
        ExistsInSetting = False
    End If
    On Error GoTo 0
End Function

Private Function ProcessInNeed(ByVal str As String) As Boolean
    Dim tmp As Variant
    For Each tmp In ProcessNeed
        If LCase(Left(str, Len(tmp))) = LCase(tmp) Then
            ProcessInNeed = True
            Exit Function
        End If
    Next
End Function

Private Sub DefaultINI()
    IdleTimeout = 180
    WinsOnlySys = False
    IncludeComp = False
    StandsAlone = False
    StandOnNeed = False
End Sub

Private Function SocketExists(ByRef sckControl, ByVal Index As Integer) As Boolean
    On Error Resume Next
    Dim obj As Control
    Set obj = sckControl(Index)
    If Not (Err.Number = 340) Then
        SocketExists = True
    Else
        Err.Clear
        SocketExists = False
    End If
    Set obj = Nothing
    On Error GoTo 0
End Function

Private Function NewSocket(ByRef sckControl) As Integer
    Dim ptr As Integer
    If sckControl.UBound + 1 = 1 Then
        If sckControl(0).State = StateConstants.sckClosed Then
            ptr = 0
        Else
            ptr = 1
        End If
    Else
        ptr = sckControl.UBound + 1
    End If
    If ptr > 0 Then
        Load sckControl(ptr)
    End If
    NewSocket = ptr
End Function

Private Sub EndSocket(ByRef sckControl, ByVal Index As Integer)
    sckControl(Index).Tag = ""
    sckControl(Index).Close
    If Index > 0 Then
        Unload sckControl(Index)
    End If
End Sub

Private Sub ClearSockets(ByRef sckControl)
    If (sckControl.Count > 0) Then
        Dim cnt As Integer
        For cnt = 0 To (sckControl.Count - 1)
            sckControl(cnt).Tag = ""
            sckControl(cnt).Close
            If cnt > 0 Then Unload sckControl(cnt)
        Next
    End If
    Debug.Print "Listen Disable"
End Sub

Private Sub SetListens()
    Dim ip As Variant
    Dim IP2 As Variant
    Dim ptr As Integer
    Dim addr As Collection
    Set addr = New Collection
    
    For Each ip In AdapterAddr
        If Trim(LCase(ip)) = "every" Then
            For Each IP2 In GetPortIP()
                If Not ExistsInSetting(addr, IP2) Then
                    addr.Add IP2, IP2
                End If
            Next
        Else
            If Not ExistsInSetting(addr, ip) Then
                addr.Add ip, ip
            End If
        End If
    Next
    
    If sckListen.Count > 0 Then
        For ptr = sckListen.LBound To sckListen.UBound
            If Not ExistsInSetting(addr, sckListen(ptr).localIP) Then
                EndSocket sckListen, ptr
            End If
        Next
    End If

    For Each ip In addr
        ptr = ListenIndex(ip)
        If ptr = -1 Then ptr = NewSocket(sckListen)
        If Not (sckListen(ptr).State = StateConstants.sckListening) Then
            sckListen(ptr).Close
            On Error Resume Next
            Debug.Print "Listening: " & ip & " 113"
            sckListen(ptr).Bind 113, ip
            sckListen(ptr).listen
            If (Err.Number = 10048) Then
                Err.Clear
            End If
            On Error GoTo 0
            sckListen(ptr).Tag = ip
        End If
    Next

    Do Until addr.Count = 0
        addr.Remove 1
    Loop
    Set addr = Nothing
  
    Dim listen As Winsock
    For Each listen In sckListen
        If Not (listen.Tag = "") Then
            listen.Tag = ""
            SendBroadcast "IDENT : " & IP4ToIP2(listen.localIP)
        End If
    Next

End Sub

Private Sub Form_Load()

    sckIdent(0).Tag = "0|"
    sckBroadcast.Tag = "0|-1"
    frmIdent.Tag = "False"
    
    Set ProcessNeed = New Collection
    Set AdapterAddr = New Collection
    Set ServicePort = New Collection
    
    DefaultINI
    
    LoadINI

    If Not StandOnNeed Then
        SetListens
    End If

    Timer1.Enabled = True
    
End Sub

Public Sub SendBroadcast(ByRef Data As Variant)
    If Not StandsAlone Then
        If sckBroadcast.State = StateConstants.sckOpen Then
            sckBroadcast.SendData Data
            Debug.Print "sckBroadcast.SendData " & Data
        Else
            sckBroadcast.Close
        End If
    End If
End Sub
Public Sub Shutdown()
    Timer1.Enabled = False
    frmIdent.Tag = "False"

    If Not StandsAlone Then
        sckBroadcast.Close
    End If

    ClearSockets sckListen
    ClearSockets sckRoute
    ClearSockets sckIdent
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 1 Then Shutdown
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Shutdown

    Do Until ServicePort.Count = 0
        ServicePort.Remove 1
    Loop
    Set ServicePort = Nothing

    Do Until ProcessNeed.Count = 0
        ProcessNeed.Remove 1
    Loop
    Set ProcessNeed = Nothing

    Do Until AdapterAddr.Count = 0
        AdapterAddr.Remove 1
    Loop
    Set AdapterAddr = Nothing

    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Public Function RouteIndex(ByVal ip As String, Optional ByVal Ports As String = "ROUTE") As Long
    Dim Route As Winsock
    For Each Route In sckRoute
        If Route.RemoteHost = ip And RouteData(Route.Index) = Ports Then
            RouteIndex = Route.Index
            Exit Function
        End If
    Next
    RouteIndex = -1
End Function

Public Function ListenIndex(ByVal ip As String) As Long
    Dim listen As Winsock
    For Each listen In sckListen
        If listen.localIP = ip Then
            ListenIndex = listen.Index
            Exit Function
        End If
    Next
    ListenIndex = -1
End Function

Public Function RouteExists(ByVal ip As String, Optional ByVal Ports As String = "ROUTE") As Boolean
    Dim Route As Winsock
    For Each Route In sckRoute
        If Route.RemoteHost = ip And RouteData(Route.Index) = Ports Then
            RouteExists = True
            Exit Function
        End If
    Next
    RouteExists = False
End Function

Public Function ListenExists(ByVal ip As String) As Boolean
    Dim listen As Winsock
    For Each listen In sckListen
        If (listen.localIP = ip) Then
            ListenExists = True
            Exit Function
        End If
    Next
    ListenExists = False
End Function

Public Function IdentExists(ByVal Index As Integer) As Boolean
    Dim Ident As Winsock
    For Each Ident In sckIdent
        If Ident.Index = Index Then
            IdentExists = True
            Exit Function
        End If
    Next
    IdentExists = False
End Function

Public Sub AddIdentRoute(ByVal ip As String, Optional ByVal Ports As String = "ROUTE")
    Dim cnt As Long
    If (Not ListenExists(ip)) Then
        If (Not RouteExists(ip, Ports)) Then
            cnt = NewSocket(sckRoute)
            sckRoute(cnt).RemoteHost = ip
            sckRoute(cnt).RemotePort = 113
            RouteData(cnt) = Ports
            Debug.Print "ADDING ROUTE: " & ip
        End If
        cnt = RouteIndex(ip, Ports)
        If (sckRoute(cnt).State = StateConstants.sckError) Or _
            (sckRoute(cnt).State = StateConstants.sckClosing) Then
            sckRoute(cnt).Close
        End If
        If (sckRoute(cnt).State = StateConstants.sckClosed) Then
            sckRoute(cnt).Connect
        End If
    End If
End Sub

Private Sub ReplenishRoutes()
    Dim Route As Winsock
    For Each Route In sckRoute
        If (Route.State = StateConstants.sckError) Or _
            (Route.State = StateConstants.sckClosing) Then
            Route.Close
        End If
        If (Route.State = StateConstants.sckClosed) Then
            Route.Connect
        End If
    Next
End Sub

Private Sub sckBroadcast_Close()
    sckBroadcast.Close
End Sub

Private Sub sckBroadcast_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckBroadcast.Close
End Sub

Private Sub sckBroadcast_DataArrival(ByVal bytesTotal As Long)
    If sckBroadcast.State = StateConstants.sckOpen Then
        Dim inLine As String
        sckBroadcast.GetData inLine, vbUnicode, bytesTotal
        If (NextArg(inLine, ":") = "IDENT") And (Not (RemoveArg(inLine, ":") = "")) Then
            Dim ip As String
            ip = IP2ToIP4(RemoveArg(inLine, ":"))
            AddIdentRoute ip
        ElseIf (NextArg(inLine, ":") = "IDENT") And (RemoveArg(inLine, ":") = "") Then
            ReplenishRoutes
            Dim listen As Winsock
            For Each listen In sckListen
                If listen.State = StateConstants.sckListening Then
                    SendBroadcast "IDENT : " & IP4ToIP2(listen.localIP)
                End If
            Next
        End If
    End If
End Sub

Private Sub sckIdent_Close(Index As Integer)
    EndSocket sckIdent, Index
End Sub

Private Sub sckIdent_Connect(Index As Integer)
    If RouteExists(sckIdent(Index).RemoteHostIP) Then
        IdentState(Index) = -1
    End If
End Sub

Private Sub sckIdent_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    EndSocket sckIdent, Index
End Sub

Private Sub sckIdent_SendComplete(Index As Integer)
    If IdentState(Index) = -5 Then
        EndSocket sckIdent, Index
    ElseIf (IdentState(Index) = -4) Then
        IdentState(Index) = -Timer
    End If
End Sub

Private Sub sckRoute_Close(Index As Integer)
    If RouteData(Index) = "ROUTE" Then
        sckRoute(Index).Close
    Else
        EndSocket sckRoute, Index
    End If
End Sub

Private Sub sckListen_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckListen(Index).Close
End Sub

Private Sub sckRoute_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If RouteData(Index) = "ROUTE" Then
        sckRoute(Index).Close
    Else
        EndSocket sckRoute, Index
    End If
End Sub

Private Sub sckRoute_Connect(Index As Integer)
    If RouteData(Index) = "ROUTE" Then
        Dim listen As Winsock
        For Each listen In sckListen
            sckRoute(Index).SendData "IDENT : " & listen.localIP & vbLf
        Next
    Else
        sckRoute(Index).SendData RouteData(Index) & vbLf
        
    End If
End Sub

Private Sub sckRoute_SendComplete(Index As Integer)
    If RouteData(Index) = "ROUTE" Then
        Dim Ident As Winsock
        For Each Ident In sckIdent
            If IdentData(Ident.Index) = NextArg(IdentData(Ident.Index), ":") Then
                IdentState(Ident.Index) = -Timer
            End If
        Next
    Else
        RouteState(Index) = -Timer
    End If
End Sub

Private Sub sckRoute_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim inData As String
    Dim inLine As String
    If RouteData(Index) = "ROUTE" Then
    
        If CurrentIndex > Index Then
            CurrentDirection = -1
        ElseIf CurrentIndex < Index Then
            CurrentDirection = 1
        End If
                
        If sckRoute(Index).State = StateConstants.sckConnected Then
            sckRoute(Index).GetData inData, vbUnicode, bytesTotal
            Dim Ident As Winsock
            For Each Ident In sckIdent
                If (IdentData(Ident.Index) = NextArg(inData, ":")) Then
                    If (IdentState(Ident.Index) < -3) Then IdentState(Ident.Index) = -4
                    If sckIdent(Ident.Index).State = StateConstants.sckConnected Then
                        CurrentIndex = Index
                        sckIdent(Ident.Index).SendData inData
                        Debug.Print "sckIdent(" & Ident.Index & ").SendData " & inData
                    End If
                End If
            Next
        End If
    Else

        sckRoute(Index).GetData inData, vbUnicode, bytesTotal
        Log sckRoute(Index).RemoteHostIP & " : " & RouteData(Index) & " RESPONSE " & inData
        RouteState(Index) = inData

    
    End If
End Sub


Private Sub sckIdent_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    Dim inData As String
    Dim inLine As String

    Dim cnt As Long
    Dim localIP As String
    Dim userProcess As Long
    Dim userAccount As String
    Dim portOfClient As String
    Dim portToServer As String
    Dim Route As Winsock
        
    sckIdent(Index).GetData inData, vbUnicode, bytesTotal
    If (IdentState(Index) > 0) Then IdentState(Index) = -3
    inData = Replace(inData, vbCr, "")

    Do Until (inData = "")
        inLine = RemoveNextArg(inData, vbLf)
        Debug.Print "sckIdent(" & Index & ").GetData " & inLine
        
        IdentData(Index) = inLine
            
        If (IsNumeric(NextArg(inLine, ",")) And IsNumeric(NextArg(RemoveArg(inLine, ","), ":"))) And (RemoveArg(inLine, ":") = "") Then
            
            portOfClient = NextArg(inLine, ",")
            portToServer = NextArg(RemoveArg(inLine, ","), ":")
    
            userProcess = GetProcessIDByBothPorts(CLng(portOfClient), CLng(portToServer), localIP)
            If (Not (userProcess = 0)) And (localIP = sckIdent(Index).localIP) Then
                If StandOnNeed And StandsAlone Then
                    If Not ProcessInNeed(LCase(GetFileName(ProcessPathByPID(userProcess)))) Then
                        userProcess = ProcessNeed(LCase(userProcess))
                    End If
                End If
                If (Not (userProcess = 0)) Then
                    userAccount = GetUserByProcessID(userProcess, UserNameSID)
                    If Not IncludeComp Then
                        If InStr(userAccount, "\") > 0 Then
                            userAccount = RemoveArg(userAccount, "\")
                        End If
                    End If
                    If Trim(userAccount) = "" Then
                        userAccount = GetProcessName(userProcess)
                        If Trim(userAccount) = "" Then userAccount = "Unknown"
                    End If
    
                    If WinsOnlySys Then
                        userAccount = "WINS : " & userAccount
                    Else
                        If is9x Then
                            userAccount = "WIN9X : " & userAccount
                        ElseIf isWinXP Then
                            userAccount = "WINXP : " & userAccount
                        ElseIf isNT Then
                            userAccount = "WINNT : " & userAccount
                        ElseIf System64Bit Then
                            userAccount = "WIN64 : " & userAccount
                        Else
                            Select Case Win32Ver
                                Case UnknownOS
                                    userAccount = "WINS : " & userAccount
                                Case Else
                                    userAccount = "WIN32 : " & userAccount
                            End Select
                        End If
                    End If
                Else
                    userAccount = "WINS : Unknwon"
                End If
                sckIdent(Index).SendData IdentData(Index) & " : USERID : " & userAccount & vbLf
                Debug.Print "sckIdent(" & Index & ").SendData " & IdentData(Index) & " : USERID : " & userAccount
            Else
                Dim UsingRoutes As Boolean
                If (Not StandsAlone) Then
                
                    If (Not RouteExists(sckIdent(Index).RemoteHostIP)) And (sckRoute.Count > 0) Then
    
                        If ((CurrentIndex = sckRoute.LBound) And (CurrentDirection = -1)) Or _
                             ((CurrentIndex = sckRoute.UBound) And (CurrentDirection = 1)) Then
                            CurrentDirection = -CurrentDirection
                        End If
                        For cnt = CurrentIndex To IIf(CurrentDirection = -1, sckRoute.LBound, sckRoute.UBound) Step CurrentDirection
                            If ((Not (sckRoute(cnt).localIP = sckIdent(Index).RemoteHostIP)) And (sckRoute(cnt).State = StateConstants.sckConnected)) Then
                                sckRoute(cnt).SendData inLine & vbLf
                                Debug.Print "sckRoute(" & cnt & ").SendData " & inLine
                                UsingRoutes = True
                            End If
                        Next
                        For cnt = IIf(CurrentDirection = -1, sckRoute.UBound, sckRoute.LBound) To CurrentIndex Step CurrentDirection
                            If ((Not (sckRoute(cnt).localIP = sckIdent(Index).RemoteHostIP)) And (sckRoute(cnt).State = StateConstants.sckConnected)) And (Not (CurrentIndex = cnt)) Then
                                sckRoute(cnt).SendData inLine & vbLf
                                Debug.Print "sckRoute(" & cnt & ").SendData " & inLine
                                UsingRoutes = True
                            End If
                        Next
                            
                    End If
                    
                    If IdentState(Index) <= -3 Then IdentState(Index) = -Timer
                End If
                If Not UsingRoutes Then
                    IdentState(Index) = -5
                    sckIdent(Index).SendData NextArg(IdentData(Index), ":") & " : ERROR : INVALID-PORT" & vbLf
                    Debug.Print "sckIdent(" & Index & ").SendData " & NextArg(IdentData(Index), ":") & " : ERROR : INVALID-PORT"
                End If
                
            End If
    
        ElseIf (NextArg(inLine, ":") = "IDENT") And (Not (RemoveArg(inLine, ":") = "")) Then
            If Not StandsAlone Then
                IdentState(Index) = -1
                AddIdentRoute RemoveArg(inLine, ":")
            End If
                 
        Else
            IdentState(Index) = -5
            sckIdent(Index).SendData NextArg(IdentData(Index), ":") & " : ERROR : INVALID-PORT" & vbLf
            Debug.Print "sckIdent(" & Index & ").SendData " & NextArg(IdentData(Index), ":") & " : ERROR : INVALID-PORT"
        End If
    Loop

End Sub
Private Function IndividualsSafety() As Boolean
    Dim cnt As Long
    Dim ready As Boolean
    Dim tmr As Single
    tmr = Timer
    
    Do Until ready Or ((Timer - tmr) > 30)
        If sckIdent.Count > 0 Then
            ready = True
            For cnt = sckIdent.LBound To sckIdent.UBound
                ready = ready And ((Not sckIdent(cnt).State = StateConstants.sckConnecting) _
                            Or (Not sckIdent(cnt).State = StateConstants.sckConnectionPending))
            Next
        Else
            ready = True
        End If
    Loop
    
    IndividualsSafety = ready
    
End Function

Private Sub sckListen_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'Memory intrusion can appear detected from the outside
    'so we use a a few on error statements to directionals
    'and we repeat it just to wear out and cover any cases
    On Local Error GoTo -1
    On Error GoTo -1
    On Local Error GoTo 0
    On Error GoTo 0
    On Local Error Resume Next
    On Error Resume Next
    
    Dim cnt As Long
    
    On Local Error GoTo -1
    On Error GoTo -1
    On Local Error GoTo 0
    On Error GoTo 0
    On Local Error Resume Next
    On Error Resume Next
    
    cnt = -1
    
    On Local Error GoTo -1
    On Error GoTo -1
    On Local Error GoTo 0
    On Error GoTo 0
    On Local Error Resume Next
    On Error Resume Next
    
    On Local Error GoTo exitout5
    On Error GoTo exitout5
    GoTo passit
    
exitout3:
    If Err Then
        Err.Clear
        EndSocket sckIdent, cnt
    End If
    Resume Next
    
exitout4:
    On Local Error GoTo exitout3
    On Error GoTo exitout3
    Resume
    
exitout5:
    On Local Error GoTo exitout4
    On Error GoTo exitout4
    Resume Next
    
passit:

    SendBroadcast "IDENT"
    
    If IndividualsSafety() Then
        
        cnt = NewSocket(sckIdent)
        IdentState(cnt) = CStr(Timer)
        IdentData(cnt) = IP4ToIP2(sckListen(Index).RemoteHostIP)
    
        If IndividualsSafety() Then
    
            sckIdent(cnt).Accept requestID
    
            If Not IndividualsSafety() Then EndSocket sckIdent, cnt

        Else
            EndSocket sckIdent, cnt
        End If
        
    End If
End Sub
Private Property Get IdentState(ByVal Index As Integer) As Single
    Dim tmp As String
    tmp = NextArg(sckIdent(Index).Tag, "|")
    If Not IsNumeric(tmp) Then
        IdentState = 0
    Else
        IdentState = CSng(tmp)
    End If
End Property
Private Property Let IdentState(ByVal Index As Integer, ByVal newVal As Single)
    sckIdent(Index).Tag = newVal & "|" & RemoveArg(sckIdent(Index).Tag, "|")
End Property

Private Property Get IdentData(ByVal Index As Integer) As String
    IdentData = RemoveArg(sckIdent(Index).Tag, "|")
End Property
Private Property Let IdentData(ByVal Index As Integer, ByVal newVal As String)
    sckIdent(Index).Tag = NextArg(sckIdent(Index).Tag, "|") & "|" & newVal
End Property

Private Property Get RouteState(ByVal Index As Integer) As Single
    Dim tmp As String
    tmp = NextArg(sckRoute(Index).Tag, "|")
    If Not IsNumeric(tmp) Then
        RouteState = 0
    Else
        RouteState = CSng(tmp)
    End If
End Property
Private Property Let RouteState(ByVal Index As Integer, ByVal newVal As Single)
    sckRoute(Index).Tag = newVal & "|" & RemoveArg(sckRoute(Index).Tag, "|")
End Property

Private Property Get RouteData(ByVal Index As Integer) As String
    RouteData = RemoveArg(sckRoute(Index).Tag, "|")
End Property
Private Property Let RouteData(ByVal Index As Integer, ByVal newVal As String)
    sckRoute(Index).Tag = NextArg(sckRoute(Index).Tag, "|") & "|" & newVal
End Property

Private Property Get CurrentIndex() As Long
    Dim tmp As Long
    tmp = NextArg(sckBroadcast.Tag, "|")
    If Not IsNumeric(tmp) Then
        CurrentIndex = 0
    Else
        CurrentIndex = CSng(tmp)
    End If
End Property
Private Property Let CurrentIndex(ByVal newVal As Long)
    sckBroadcast.Tag = newVal & "|" & RemoveArg(sckBroadcast.Tag, "|")
End Property

Private Property Get CurrentDirection() As Long
    Dim tmp As Long
    tmp = RemoveArg(sckBroadcast.Tag, "|")
    If Not IsNumeric(tmp) Then
        CurrentDirection = 0
    Else
        CurrentDirection = CSng(tmp)
    End If
End Property
Private Property Let CurrentDirection(ByVal newVal As Long)
    sckBroadcast.Tag = NextArg(sckBroadcast.Tag, "|") & "|" & newVal
End Property

Private Sub Timer1_Timer()
    
    If StandOnNeed Then
        If frmIdent.Tag = "Both" Then
            frmIdent.Tag = "False"
            ClearSockets sckListen
        End If
        Dim needFound As Boolean
        Dim processName As Variant
        For Each processName In ProcessNeed
            needFound = needFound Or (modProcess.ProcessRunning(processName) > 0)
        Next
        If needFound And (frmIdent.Tag = "False") Then
            frmIdent.Tag = "True"
            SetListens
        ElseIf (Not needFound) And (frmIdent.Tag = "True") Then
            frmIdent.Tag = "False"
            ClearSockets sckListen
        End If
    ElseIf Not StandOnNeed Then
        If (frmIdent.Tag = "True") Then
            frmIdent.Tag = "False"
            ClearSockets sckListen
        End If
    
        If (frmIdent.Tag <> "Both") Then
            frmIdent.Tag = "Both"
            SetListens
        End If
    End If
        
    If sckBroadcast.State = StateConstants.sckClosed Then
        Dim localIP As String
        If Not StandsAlone Then
            localIP = sckBroadcast.localIP
            sckBroadcast.protocol = sckUDPProtocol
            sckBroadcast.LocalPort = 113
            sckBroadcast.RemotePort = 113
            sckBroadcast.RemoteHost = "255.255.255.255"
            On Error Resume Next
            sckBroadcast.Bind 113, localIP
            If (Err.Number = 10048) Then
                Err.Clear
            End If
            On Error GoTo 0
            If sckBroadcast.State = StateConstants.sckOpen Then
                SendBroadcast "IDENT"
            End If
        End If
    End If
    
    Dim user As String
   
    If (sckIdent.Count > 0) Then
        Dim Ident As Winsock
        For Each Ident In sckIdent
            If IsNumeric(IdentState(Ident.Index)) Then
                If ((Timer - IdentState(Ident.Index)) >= IdleTimeout) And (IdentState(Ident.Index) > 0) Then
                    EndSocket sckIdent, Ident.Index
                ElseIf ((Timer - (-IdentState(Ident.Index))) >= IdleTimeout) And (IdentState(Ident.Index) < -5) Then
                    If (Ident.State = StateConstants.sckConnected) Then
                        If (Not (IdentData(Ident.Index) = "")) Then
                            IdentState(Ident.Index) = -5
                            Ident.SendData NextArg(IdentData(Ident.Index), ":") & " : ERROR : INVALID-PORT" & vbLf
                            Debug.Print "sckIdent(" & Ident.Index & ").SendData " & NextArg(IdentData(Ident.Index), ":") & " : ERROR : INVALID-PORT"
                        Else
                            EndSocket sckIdent, Ident.Index
                        End If
                    Else
                        EndSocket sckIdent, Ident.Index
                    End If
                End If
            End If
        Next
    End If
    If (sckRoute.Count > 0) Then
        Dim Route As Winsock
        For Each Route In sckRoute
            If Not RouteData(Route.Index) = "ROUTE" Then
                If IsNumeric(RouteState(Route.Index)) Then
                    If ((Timer - RouteState(Route.Index)) >= IdleTimeout) And (RouteState(Route.Index) > 0) Then
                        EndSocket sckRoute, Route.Index
                    ElseIf ((Timer - (-RouteState(Route.Index))) >= IdleTimeout) And (RouteState(Route.Index) < -5) Then
                        If (Route.State = StateConstants.sckConnected) Then
                            If (Not (RouteData(Route.Index) = "")) Then
                                RouteState(Route.Index) = -5
                                Route.SendData NextArg(RouteData(Route.Index), ":") & " : ERROR : INVALID-PORT" & vbLf
                                Debug.Print "sckRoute(" & Route.Index & ").SendData " & NextArg(RouteData(Route.Index), ":") & " : ERROR : INVALID-PORT"
                            Else
                                EndSocket sckRoute, Route.Index
                            End If
                        Else
                            EndSocket sckRoute, Route.Index
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If ServicePort.Count > 0 Then
        Dim ip As String
        Dim pp As String
        
        Dim sp As Variant
        
        Dim p As Variant
        Dim ips As Variant
        Dim rml As String
        
        For Each ips In GetPortIP
            For Each sp In ServicePort
                For Each p In GetAllPortsByLocalIP(CStr(ips))
                    If NextArg(CStr(p), ",") = CStr(sp) Then
                        ip = CStr(GetRemoteIPByBothPorts(CLng(NextArg(p, ",")), CLng(RemoveArg(p, ","))))
                        pp = RemoveArg(p, ",") & " , " & NextArg(CStr(p), ",")
                        If InStr(Timer1.Tag, ip & " : " & pp & vbLf) = 0 Then
                            Timer1.Tag = Timer1.Tag & ip & " : " & pp & vbLf
                            AddIdentRoute ip, pp
                        Else
                            rml = rml & ip & " : " & pp & vbLf
                        End If
                    End If
                Next
            Next
        Next
        
        pp = Timer1.Tag
        Do While pp <> ""
            If InStr(pp, NextArg(rml, vbLf) & vbLf) = 0 Then
                If RouteExists(NextArg(rml, vbLf), RemoveArg(rml, vbLf)) Then
                    pp = Replace(pp, NextArg(rml, vbLf) & vbLf, "")
                Else
                    Timer1.Tag = Replace(Timer1.Tag, NextArg(rml, vbLf) & vbLf, "")
                End If
                RemoveNextArg rml, vbLf
            Else
                Timer1.Tag = Timer1.Tag & NextArg(rml, vbLf) & vbLf
                RemoveNextArg pp, vbLf
            End If
        Loop
    End If
        
    
End Sub


