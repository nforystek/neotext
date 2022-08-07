VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmBlkLServer 
   BorderStyle     =   0  'None
   Caption         =   "Blacklawn Server"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ControlBox      =   0   'False
   Icon            =   "frmBlkLServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2220
      Top             =   75
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   90
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Index           =   0
      Left            =   1680
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Index           =   0
      Left            =   1140
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   615
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBlkLServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Connections As Long
Private Patterns As Collection
Private Scok1Data As String
Private Sock2Data As String
Private Sock3Data As String

Private Function NameIndex(ByVal Name As String) As Long
    Dim cnt As Long
    Dim tmp As clsPattern
    If Patterns.Count > 0 Then
        For cnt = 1 To Patterns.Count
            Set tmp = Patterns(cnt)
            If tmp.Name = Name Then
                NameIndex = cnt
                Exit Function
            End If
        Next
    End If
    NameIndex = 0
End Function

Private Sub AcceptClient(ByVal requestID As Long)
    Dim cnt As Long
    Dim Index As Long
    Index = Connections
    For cnt = Winsock4.UBound To Winsock4.LBound Step -1
        If Not (Winsock2(cnt).State = StateConstants.sckConnected) Then
            Index = cnt
            Exit For
        End If
    Next
    If Not (Index = 0) Then
        Load Winsock2(Index)
        Load Winsock3(Index)
        Load Winsock4(Index)
    End If
    Winsock2(Index).Tag = StateConstants.sckConnecting
    Winsock3(Index).Tag = StateConstants.sckConnecting
    Winsock4(Index).Tag = StateConstants.sckListening
    Winsock2(Index).Close
    Winsock3(Index).Close
    Winsock4(Index).Close
    Winsock2(Index).Accept requestID
    Connections = Connections + 1
End Sub

Private Sub RemoveClient(ByVal Index As Long)
    Winsock2(Index).Close
    Winsock3(Index).Close
    Winsock4(Index).Close
    Dim cnt As Long
    Dim tmp As clsPattern
    For cnt = 1 To Patterns.Count
        Set tmp = Patterns(cnt)
        If tmp.Name = Winsock3(Index).Tag Then
            Patterns.Remove cnt
            Exit For
        End If
        Set tmp = Nothing
    Next
    Winsock2(Index).Tag = StateConstants.sckClosed
    Winsock3(Index).Tag = StateConstants.sckClosed
    Winsock4(Index).Tag = StateConstants.sckClosed
    If (Not (Index = 0)) Then
        Unload Winsock2(Index)
        Unload Winsock3(Index)
        Unload Winsock4(Index)
    End If
    Connections = Connections - 1
End Sub

Private Sub Form_Initialize()
    Set Patterns = New Collection
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    Winsock1.Bind ListenPort
    Winsock1.Listen
End Sub

Private Sub Form_Terminate()
    Set Patterns = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    On Error GoTo -1
    On Error GoTo 0
    On Error Resume Next
    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo rollback
    
    Dim tmp As clsPattern
    Dim str As String
    Dim Ret As String
    Dim sck As Long
    Dim ptr As Long
    Dim cnt As Long
    Dim sta As Boolean
    If (Connections > 0) Then
        If (Patterns.Count > 0) Then
            For cnt = 1 To Patterns.Count
                Set tmp = Patterns(cnt)
                If (Not (tmp.List = "")) And (tmp.Peek = 0) Then
                    ptr = cnt
                    For sck = Winsock4.LBound To Winsock4.UBound
                        If Winsock4(sck).State = StateConstants.sckConnected Then
                            If tmp.Name = Winsock3(sck).Tag Then
                                Winsock2(sck).SendData "2" & tmp.List & vbCrLf
                            End If
                        End If
                    Next
                    tmp.List = ""
                    Patterns.Add tmp, , cnt
                    Patterns.Remove cnt
                End If
                Set tmp = Nothing
            Next
        End If
        Dim strand As String
        For sck = Winsock4.LBound To Winsock4.UBound
            If Winsock4(sck).State = StateConstants.sckConnected Then
                strand = ""
                If (Patterns.Count > 0) Then
                    For cnt = 1 To Patterns.Count
                        Set tmp = Patterns(cnt)
                        If Not (tmp.Name = Winsock3(sck).Tag) And (Not (tmp.Pass = "")) Then
                            strand = strand & tmp.Name & "," & tmp.Pass & "|"
                        End If
                        Set tmp = Nothing
                    Next
                    If Winsock4(sck).Tag = "" Then
                        Winsock4(sck).Tag = strand & "|" & vbCrLf
                        Winsock4(sck).SendData Winsock4(sck).Tag
                    End If
                End If
            End If
        Next
    End If

    Exit Sub
rollback:
    Err.Clear
End Sub

Private Sub Winsock2_Close(Index As Integer)
    RemoveClient Index
End Sub
Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RemoveClient Index
End Sub

Private Sub Winsock3_Close(Index As Integer)
    RemoveClient Index
End Sub

Private Sub Winsock3_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Winsock2(Index).Tag = StateConstants.sckResolvingHost And Winsock3(Index).RemoteHost = Winsock2(Index).RemoteHost Then
        Winsock3(Index).Close
        Winsock3(Index).Accept requestID
    End If
End Sub

Private Sub Winsock3_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RemoveClient Index
End Sub

Private Sub Winsock4_Close(Index As Integer)
    RemoveClient Index
End Sub

Private Sub Winsock4_Connect(Index As Integer)
    Winsock4(Index).Tag = ""
End Sub

Private Sub Winsock4_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    If Winsock4(Index).Tag = StateConstants.sckConnecting And Winsock4(Index).RemoteHost = Winsock2(Index).RemoteHost Then
        Winsock4(Index).Close
        Winsock4(Index).Accept requestID
        Winsock4(Index).Tag = ""
    End If
End Sub

Private Sub Winsock4_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    RemoveClient Index
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    AcceptClient requestID
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Winsock1.Close
    Winsock1.Bind ListenPort
    Winsock1.Listen
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo -1
    On Error GoTo 0
    On Error Resume Next
    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo rollback

    Dim Data As String
    
    Winsock2(Index).GetData Sock2Data, , bytesTotal
    If InStr(Sock2Data, vbCrLf) > 0 Then
        Data = RemoveNextArg(Sock2Data, vbCrLf)
        If Not (Data = "") Then
            Select Case Winsock2(Index).Tag
                Case StateConstants.sckConnecting
                    Data = Replace(Data, vbCrLf, "")
                    If (Connections > ClientCount) Then
                        Winsock2(Index).SendData "0LIMIT" & vbCrLf
                    ElseIf (NameIndex(Data) > 0) Or (Not IsAlphaNumeric(Data)) Or (Len(Data) > 64) Then
                        Winsock2(Index).SendData "0EXIST" & vbCrLf
                    ElseIf Not (Left(Data, 2) = "v5") Then
                        Winsock2(Index).SendData "0NOVER" & vbCrLf
                    Else
                    
                        Winsock3(Index).Tag = Mid(Data, 3)

                        Winsock4(Index).Tag = StateConstants.sckConnecting
                        Winsock3(Index).Bind (ListenPort + 1), Winsock2(Index).LocalIP
                        Winsock3(Index).Listen
                        Winsock4(Index).Bind (ListenPort + 2), Winsock2(Index).LocalIP
                        Winsock4(Index).Listen
                        
                        Winsock2(Index).Tag = StateConstants.sckResolvingHost
                        Winsock2(Index).SendData "-" & DisplayMsg & vbCrLf
                    End If
                Case StateConstants.sckConnected
                    Dim cnt As Long
                    Dim ptr As Long
                    Dim sck As Long
                    Dim tmp As clsPattern
                    Select Case Left(Data, 1)
                        Case "0"
                            'broadcast message
                            If (Patterns.Count > 0) Then
                                For cnt = 1 To Patterns.Count
                                    Set tmp = Patterns(cnt)
                                    If (tmp.Name = Winsock3(Index).Tag) Then
                                        ptr = cnt 'from this client
                                        For sck = Winsock4.LBound To Winsock4.UBound
                                            If Winsock4(sck).State = StateConstants.sckConnected Then
                                                If Not (tmp.Name = Winsock3(sck).Tag) Then
                                                    Winsock2(sck).SendData "0" & tmp.Name & " says: " & Mid(Data, 2) & vbCrLf
                                                Else
                                                    Winsock2(sck).SendData "0" & "You say: " & Mid(Data, 2) & vbCrLf
                                                End If
                                            End If
                                        Next
                                    End If
                                    Set tmp = Nothing
                                Next
                            End If
                        Case "1"
                            'round up scores
                            If (Patterns.Count > 0) Then
                                For cnt = 1 To Patterns.Count
                                    Set tmp = Patterns(cnt)
                                    If (tmp.Name = Winsock3(Index).Tag) Then
                                        ptr = cnt 'for this client
                                        tmp.Peek = (Connections - 1)
                                        tmp.List = ""
                                        Patterns.Add tmp, , cnt
                                        Patterns.Remove cnt
                                        For sck = Winsock4.LBound To Winsock4.UBound
                                            If Winsock4(sck).State = StateConstants.sckConnected Then
                                                If Not (tmp.Name = Winsock3(sck).Tag) Then
                                                    Winsock2(sck).SendData "1" & vbCrLf
                                                End If
                                            End If
                                        Next
                                    End If
                                    Set tmp = Nothing
                                Next
                            End If
                        Case "2"
                            'receiving score
                            If (Patterns.Count > 0) Then
                                For cnt = 1 To Patterns.Count
                                    Set tmp = Patterns(cnt)
                                    If (tmp.Peek > 0) Then
                                        ptr = cnt 'for this client
                                        If Not (tmp.Name = Winsock3(Index).Tag) Then
                                            tmp.Peek = tmp.Peek - 1
                                            tmp.List = tmp.List & Winsock3(Index).Tag & "'s Scores: " & Mid(Data, 2) & "|"
                                            Patterns.Add tmp, , cnt
                                            Patterns.Remove cnt
                                        End If
                                    End If
                                    Set tmp = Nothing
                                Next
                            End If
                    End Select
            End Select
        End If
    End If

    Exit Sub
rollback:
    Err.Clear
End Sub

Private Sub Winsock3_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error GoTo -1
    On Error GoTo 0
    On Error Resume Next
    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo rollback
    
    Dim Data As String
    Dim pc As clsPattern
    Dim cnt As Long
    Dim pack As String
    Dim sck As Long
        
    Winsock3(Index).GetData Sock3Data, , bytesTotal
    If InStr(Sock3Data, vbCrLf) > 0 Then
        Data = RemoveNextArg(Sock3Data, vbCrLf)
        If Not (Data = "") Then
            Select Case Winsock2(Index).Tag
                Case StateConstants.sckResolvingHost
                    Winsock2(Index).Tag = StateConstants.sckConnected
                    Set pc = New clsPattern
                    pc.Name = Winsock3(Index).Tag
                    pc.Pass = Data
                    Patterns.Add pc
                    Set pc = Nothing
                Case StateConstants.sckConnected
                    If (Patterns.Count > 0) Then
                        For cnt = 1 To Patterns.Count
                            Set pc = Patterns(cnt)
                            If (pc.Name = Winsock3(Index).Tag) Then
                                pc.Pass = Data
                                Patterns.Add pc, , cnt
                                Patterns.Remove cnt
                            End If
                            Set pc = Nothing
                        Next
                    End If
            End Select
        End If
    End If

    Exit Sub
rollback:
    Err.Clear
End Sub

Private Sub Winsock4_SendComplete(Index As Integer)
    Winsock4(Index).Tag = ""
End Sub
