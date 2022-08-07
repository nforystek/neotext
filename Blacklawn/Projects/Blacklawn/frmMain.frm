VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Blacklawn"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2565
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":0E42
   MousePointer    =   1  'Arrow
   ScaleHeight     =   885
   ScaleWidth      =   2565
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   1545
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1020
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   495
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private px As Integer
Private py As Integer
Private pZ As Integer

Private inPort As Long
Private pAtTime As Single

Private pRecording As Boolean
Private pIsPlayback As Boolean
Private pMultiPlayer As Boolean

Private pMyScoreStrand As String
Private pPlayerMoveStates As String

Private cControls As String
Private pToggler(1 To 16) As Boolean

Private Sock1Data As String
Private Sock2Data As String
Private Sock3Data As String

    Public Property Get AtTime() As Double
    AtTime = pAtTime
End Property

Public Property Get MouseX() As Integer
    MouseX = px
End Property
Public Property Let MouseX(ByVal NewVal As Integer)
    px = NewVal
End Property

Public Property Get MouseY() As Integer
    MouseY = py
End Property
Public Property Let MouseY(ByVal NewVal As Integer)
    py = NewVal
End Property

Public Property Get MouseZ() As Integer
    MouseZ = pZ
End Property
Public Property Let MouseZ(ByVal NewVal As Integer)
    pZ = NewVal
End Property
Public Sub ClearRecord()
    UserData = ""
    WarpData = ""
    ViewData = ""
End Sub

Public Property Get RecordSize() As Long
    RecordSize = LOF(FilmFileNum)
End Property

Public Property Get ToggleString() As String
    Dim cnt As Long
    For cnt = LBound(pToggler) To UBound(pToggler)
        ToggleString = ToggleString & Trim(CStr(-CInt(CBool(pToggler(cnt)))))
    Next
    If String(Len(ToggleString), "0") = ToggleString Then ToggleString = ""
End Property

Public Property Get Toggler(ByVal Index As Integer) As Boolean
    Toggler = pToggler(Index)
End Property
Public Property Let Toggler(ByVal Index As Integer, ByVal NewVal As Boolean)
    pToggler(Index) = NewVal
End Property

Public Property Get Recording() As Boolean
    Recording = pRecording
End Property
Public Property Let Recording(ByVal NewVal As Boolean)
    pRecording = NewVal
End Property

Public Property Get IsPlayback() As Boolean
    IsPlayback = pIsPlayback
End Property
Public Property Let IsPlayback(ByVal NewVal As Boolean)
    pIsPlayback = NewVal
End Property

Public Property Get PlayerMoveStates() As String
    PlayerMoveStates = pPlayerMoveStates
End Property
Public Property Let PlayerMoveStates(ByVal NewVal As String)
    pPlayerMoveStates = NewVal
End Property

Public Property Get MyScoreStrand() As String
    MyScoreStrand = pMyScoreStrand
End Property
Public Property Let MyScoreStrand(ByVal NewVal As String)
    pMyScoreStrand = NewVal
End Property

Public Sub AddControls(ByVal str As String)
    On Error GoTo stoprecord
    
    str = Trim(CStr(CStr(CDbl(Timer)) & IIf(str = ",,,", "", "," & str) & "|"))
    
    Put #FilmFileNum, , str
    
    Exit Sub
stoprecord:
    Err.Clear
    AddMessage "Error recording, stopping."
    StopFilm
End Sub
 
Public Sub StreamControls()
    On Error GoTo multiplayererror
    
    If Sock2Data = "" Then
        Sock2Data = CStr(Round(Player.Rotation, 2)) & "," & CStr(Round(Player.CameraAngle, 2)) & "," & CStr(Round(Player.CameraPitch, 2)) & "," & _
                    CStr(Round(Player.Object.Origin.X, 2)) & "," & CStr(Round(Player.Object.Origin.Y, 2)) & "," & CStr(Round(Player.Object.Origin.z, 2)) & "," & _
                    CStr(Round(Partner.Rotation, 2)) & "," & _
                    CStr(Round(Partner.Object.Origin.X, 2)) & "," & CStr(Round(Partner.Object.Origin.Y, 2)) & "," & CStr(Round(Partner.Object.Origin.z, 2)) & "," & _
                    CStr(Player.Texture) & "," & CStr(CInt(Player.Trails)) & "," & CStr(Player.Model) & "," & _
                    CInt(pToggler(12)) & "," & CInt(pToggler(13)) & "," & CInt(pToggler(14)) & vbCrLf
                    
        Winsock2.SendData Sock2Data
        Debug.Print "SOCK2: " & Sock2Data
    End If
    
    
    Exit Sub
multiplayererror:
    Err.Clear
    AddMessage "Multiplayer error, disconnecting."
    Disconnect
End Sub

Public Sub NextControl()
    On Error GoTo stoprecord
    
    Dim inControl As String
    Dim tmp As String
    
    inControl = GetFilmNextControlData
    
    Do While (Left(inControl, 1) = "B")
        tmp = Mid(inControl, 2)
        inControl = GetFilmNextControlData
        AddBeacon RemoveNextArg(tmp, ","), RemoveNextArg(tmp, ","), RemoveNextArg(tmp, ",")
    Loop

    If Not (inControl = "") Then

        pAtTime = RemoveNextArg(inControl, ",")
        
        Partner.Object.Origin.X = RemoveNextArg(inControl, ",")
        Partner.Object.Origin.Y = RemoveNextArg(inControl, ",")
        Partner.Object.Origin.z = RemoveNextArg(inControl, ",")
        Partner.Rotation = RemoveNextArg(inControl, ",")
        
        tmp = RemoveNextArg(inControl, ",")
        px = IIf(tmp = "", 0, tmp)
        tmp = RemoveNextArg(inControl, ",")
        py = IIf(tmp = "", 0, tmp)
        tmp = RemoveNextArg(inControl, ",")
        pZ = IIf(tmp = "", 0, tmp)
        If inControl = "" Then inControl = String(UBound(pToggler), "0")
        Dim cnt As Long
        cnt = 1
        Do While Not (inControl = "")
            pToggler(cnt) = CBool(Left(inControl, 1))
            inControl = Mid(inControl, 2)
            cnt = cnt + 1
        Loop
    Else
        StopFilm
    End If
    
    Exit Sub
stoprecord:
    Err.Clear
    AddMessage "Error recording, stopping."
    StopFilm
End Sub

Public Property Get Multiplayer() As Boolean
    Multiplayer = pMultiPlayer
End Property
Public Property Let Multiplayer(ByVal NewVal As Boolean)
    pMultiPlayer = NewVal
End Property

Private Sub Winsock1_Connect()
    Winsock1.SendData BlacklawnVer & Player.name & vbCrLf
    Debug.Print "SOCK1: " & BlacklawnVer & Player.name & vbCrLf
End Sub


Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Disconnect
End Sub

Private Sub Winsock2_Close()
    Disconnect
End Sub


Private Sub Winsock2_Connect()
    If Winsock1.Tag = StateConstants.sckConnecting Then
    
        Winsock3.Connect Winsock1.RemoteHostIP, (inPort + 2)
    End If
End Sub

Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Disconnect
End Sub

Private Sub Winsock2_SendComplete()
    Sock2Data = ""
End Sub

Private Sub Winsock3_Close()
    Disconnect
End Sub

Private Sub Winsock3_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Disconnect
End Sub

Public Sub Disconnect()
    If Not Winsock1.Tag = StateConstants.sckClosed Then
        Winsock1.Tag = StateConstants.sckClosed
        Winsock2.Tag = StateConstants.sckClosed
        Winsock3.Tag = StateConstants.sckClosed
        Winsock1.Close
        Winsock2.Close
        Winsock3.Close
        Winsock1.LocalPort = 0
        Winsock2.LocalPort = 0
        Winsock3.LocalPort = 0
        AddMessage "Connection Terminated."
        Multiplayer = False
        If ClearScore Then
            ClearUserData
            LoadUserData
        End If
        pPlayerMoveStates = ""
        FlagPlayers
        PlayersDone
    End If
End Sub

Public Function Connect(ByVal Server As String, ByVal name As String) As Boolean
    On Error GoTo multiplayererror
    
    inPort = 28732
    If InStr(Server, ":") > 0 Then inPort = CLng(RemoveArg(Server, ":"))
    Server = RemoveNextArg(Server, ":")
    Winsock1.Tag = StateConstants.sckConnecting
    Winsock2.Tag = StateConstants.sckConnecting
    Winsock3.Tag = StateConstants.sckConnecting
    Winsock1.Close
    Winsock2.Close
    Winsock3.Close
    Winsock1.LocalPort = 0
    Winsock2.LocalPort = 0
    Winsock3.LocalPort = 0
    Player.name = name
    Winsock1.Connect Server, inPort
    
    Exit Function
multiplayererror:
    Err.Clear
    AddMessage "Multiplayer error, disconnecting."
    Disconnect
End Function

Public Function Listen(ByVal name As String, ByVal port As Long)
    
    Player.name = name
    Winsock1.Tag = StateConstants.sckClosed
    Winsock2.Tag = StateConstants.sckClosed
    Winsock3.Tag = StateConstants.sckClosed
    Winsock1.Close
    Winsock2.Close
    Winsock3.Close
    Winsock1.LocalPort = 0
    Winsock2.LocalPort = 0
    Winsock3.LocalPort = 0
    Winsock1.Bind port, Winsock1.LocalIP
    Winsock1.Listen
    
End Function

Private Sub Winsock1_Close()
    Disconnect
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo multiplayererror
    
    Dim Data As String
    Dim Sock1Data As String

    Winsock1.GetData Sock1Data, , bytesTotal
    If InStr(Sock1Data, vbCrLf) > 0 Then
        Data = RemoveNextArg(Sock1Data, vbCrLf)
        If Not (Data = "") Then
            Debug.Print "SOCK1> " & Data & vbCrLf
            Dim line As String
            Select Case Winsock1.Tag
                Case StateConstants.sckConnecting
                    Select Case Left(Data, 1)
                        Case "0"
                            Select Case RemoveNextArg(Mid(Data, 2), vbCrLf)
                                Case "LIMIT"
                                    AddMessage "Servers at player limits."
                                Case "EXIST"
                                    AddMessage "Invalid or existing name."
                                Case "NOVER"
                                    AddMessage "Invalid client version: " & BlacklawnVer
                                Case Else
                                    AddMessage "Unknown connection clause."
                            End Select
                            Disconnect
                        Case Else
                            Winsock1.Tag = StateConstants.sckConnecting
                            Winsock2.Connect Winsock1.RemoteHostIP, (inPort + 1)
                            If Not (Data = "-") Then
                                AddMessage "Welcome to: " & Mid(Data, 2)
                            Else
                                AddMessage "Welcome to: " & Winsock1.RemoteHostIP
                            End If
                    End Select
                Case StateConstants.sckConnected
                    Do Until (Data = "")
                        line = RemoveNextArg(Data, vbCrLf)
                        Select Case Left(line, 1)
                            Case "0"
                                'display message
                                AddMessage Mid(line, 2)
                                TalkMessage Mid(line, 2)
                            Case "1"
                                'send scores
                                Winsock1.SendData "2" & pMyScoreStrand & vbCrLf
                                Debug.Print "SOCK1: " & "2" & pMyScoreStrand & vbCrLf
                            Case "2"
                                'display scores
                                line = Mid(line, 2)
                                Do While Not (line = "")
                                    AddMessage RemoveNextArg(line, "|")
                                Loop
                        End Select
                    Loop
            End Select
        End If
    End If
    
    Exit Sub
multiplayererror:
    Err.Clear
    AddMessage "Multiplayer error, disconnecting."
    Disconnect
End Sub

Private Sub Winsock3_Connect()
    Winsock1.Tag = StateConstants.sckConnected
    Multiplayer = True
    If ClearScore Then
        SaveUserData
        ClearUserData
    End If
    AddMessage "Connection Established."
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo multiplayererror
    
    Dim allPlayers As String
    Dim playerData As String
    Dim pN As String
    Dim p As Integer

    Dim newPlayer As MyPlayer
    ReDim newPlayer.Spots(0 To 50) As D3DVECTOR
    Dim newPartner As MyPlayer
    ReDim newPartner.Spots(0 To 0) As D3DVECTOR
    newPartner.Model = ShipP
    
    Dim Data As String

    Winsock3.PeekData Data, , bytesTotal
    If InStr(Data, vbCrLf) > 0 Then
    
        Winsock3.GetData Sock3Data, , bytesTotal
        If InStr(Sock3Data, vbCrLf) > 0 Then
            Data = RemoveNextArg(Sock3Data, vbCrLf)
            If Not (Data = "") Then
                pPlayerMoveStates = Data
                
                FlagPlayers
                
                Do Until (Data = "")
                    playerData = RemoveNextArg(Data, "|")
                    If Not (playerData = "") Then
                        pN = RemoveNextArg(playerData, ",")
                        newPlayer.name = pN
                        newPlayer.Flag = False
                        
                        newPlayer.Rotation = CSng(RemoveNextArg(playerData, ","))
                        newPlayer.CameraAngle = CSng(RemoveNextArg(playerData, ","))
                        newPlayer.CameraPitch = CSng(RemoveNextArg(playerData, ","))
            
                        newPlayer.Object.Origin.X = CSng(RemoveNextArg(playerData, ","))
                        newPlayer.Object.Origin.Y = CSng(RemoveNextArg(playerData, ","))
                        newPlayer.Object.Origin.z = CSng(RemoveNextArg(playerData, ","))
                        
                        newPartner.Rotation = CSng(RemoveNextArg(playerData, ","))
                        newPartner.Object.Origin.X = CSng(RemoveNextArg(playerData, ","))
                        newPartner.Object.Origin.Y = CSng(RemoveNextArg(playerData, ","))
                        newPartner.Object.Origin.z = CSng(RemoveNextArg(playerData, ","))
                        
                        newPlayer.Texture = CByte(RemoveNextArg(playerData, ","))
                        newPlayer.Trails = CBool(RemoveNextArg(playerData, ","))
                        newPlayer.Model = CByte(RemoveNextArg(playerData, ","))
                    
                        newPlayer.FlapLock = CBool(RemoveNextArg(playerData, ","))
                        newPlayer.LeftFlap = CBool(RemoveNextArg(playerData, ","))
                        newPlayer.RightFlap = CBool(RemoveNextArg(playerData, ","))
                        
                        p = PlayerExists(pN)
                        If p = 0 Then
                            LoadPlayer newPlayer
                            p = PlayerExists(pN)
                            Partners(p) = newPartner
                        Else
                            Players(p) = newPlayer
                            Partners(p) = newPartner
                        End If
                        
                        RenderShip Partners(p)
                        RenderShip Players(p)
                        
                    End If
                Loop
                
            End If
        End If
    End If
    
    Exit Sub
multiplayererror:
    Err.Clear
End Sub

Public Function Speak(ByVal Text As String)
    On Error GoTo multiplayererror

    Winsock1.SendData "0" & Text & vbCrLf
    Debug.Print "SOCK1: " & "0" & Text & vbCrLf

    Exit Function
multiplayererror:
    Err.Clear
    AddMessage "Multiplayer error, disconnecting."
    Disconnect
End Function

Public Function Scores()
    On Error GoTo multiplayererror

    Winsock1.SendData "1" & vbCrLf
    Debug.Print "SOCK1: " & "1" & vbCrLf
    
    Exit Function
multiplayererror:
    Err.Clear
    AddMessage "Multiplayer error, disconnecting."
    Disconnect
End Function

Private Sub Form_Click()
    TrapMouse = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.Multiplayer Then
        Disconnect
    ElseIf frmMain.Recording Or frmMain.IsPlayback Then
        StopFilm
    End If
    StopGame = True
End Sub
