VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Private WithEvents Socket As NTAdvFTP61.Socket
Attribute Socket.VB_VarHelpID = -1
Private WithEvents TimeOutTimer As NTSchedule20.Timer
Attribute TimeOutTimer.VB_VarHelpID = -1
Private TimedOut As Boolean
Private MsgState As Integer

Private SockConnected As Boolean
Private Received250 As Boolean
Private ReceivedError As String

Private sUsername As String
Private sPassword As String

Private sFromEMail As String
Private sFromName As String
Private sToEMail As String
Private sToName As String

Private sSubjectText As String
Private sMessageType As String
Private sMessageData As String
Private sExplicitSSL As Boolean
Private sAttachment As String

Private sServer As String
Private sPort As Long

Public Property Get Receiver() As String
    Receiver = IIf((sToEMail = ""), "", """" & IIf((sToName = ""), sToEMail, sToName) & """ <" & sToEMail & ">")
End Property
Public Property Let Receiver(ByVal NewVal As String)

    sToEMail = RemoveQuotedArg(NewVal, "<", ">")
    sToName = RemoveQuotedArg(NewVal)
    
    If Trim(sToEMail) = "" Then
        sToEMail = Trim(NewVal)
    End If
    
    If Trim(sToName) = "" Then
        If Trim(NewVal) = "" Then
            sToName = sToEMail
        Else
            sToName = Trim(NewVal)
        End If
    End If

End Property


Public Property Get Sender() As String
    Sender = IIf((sFromEMail = ""), "", """" & IIf((sFromName = ""), sFromEMail, sFromName) & """ <" & sFromEMail & ">")
End Property
Public Property Let Sender(ByVal NewVal As String)
    
    sFromEMail = RemoveQuotedArg(NewVal, "<", ">")
    sFromName = RemoveQuotedArg(NewVal)
    
    If Trim(sFromEMail) = "" Then
        sFromEMail = Trim(NewVal)
    End If
    
    If Trim(sFromName) = "" Then
        If Trim(NewVal) = "" Then
            sFromName = sFromEMail
        Else
            sFromName = Trim(NewVal)
        End If
    End If
End Property

Public Property Get Username() As String
    Username = sUsername
End Property
Public Property Let Username(ByVal RHS As String)
    sUsername = RHS
End Property

Public Property Get Password() As String
    Password = sPassword
End Property
Public Property Let Password(ByVal RHS As String)
    sPassword = RHS
End Property

Public Property Get SubjectText() As String
    SubjectText = sSubjectText
End Property
Public Property Let SubjectText(ByVal NewVal As String)
    sSubjectText = NewVal
End Property
Public Property Get MessageType() As String
    MessageType = sMessageType
End Property
Public Property Let MessageType(ByVal NewVal As String)
    sMessageType = NewVal
End Property
Public Property Get MessageData() As String
    MessageData = sMessageData
End Property
Public Property Let MessageData(ByVal NewVal As String)
    sMessageData = NewVal
End Property

Public Property Get Attachment() As String
    Attachment = sAttachment
End Property
Public Property Let Attachment(ByVal NewVal As String)
    sAttachment = NewVal
End Property
Public Property Get Server() As String
    Server = sServer
End Property
Public Property Let Server(ByVal NewVal As String)
    If InStr(NewVal, ":") > 0 Then
        sServer = RemoveNextArg(NewVal, ":")
        If IsNumeric(NewVal) Then
            Port = CLng(NewVal)
        Else
            sServer = sServer & ":" & NewVal
        End If
    Else
        sServer = NewVal
    End If
End Property
Public Property Get Port() As Long
    Port = sPort
End Property
Public Property Let Port(ByVal NewVal As Long)
    sPort = IIf((NewVal > 0), NewVal, 25)
End Property

Public Property Get timeout() As Long
    timeout = (TimeOutTimer.Interval / 1000)
End Property
Public Property Let timeout(ByVal NewVal As Long)
    If (NewVal > 0) And (NewVal < 600) Then
        TimeOutTimer.Interval = (NewVal * 1000)
    End If
End Property

Private Function Continue()

End Function

Public Sub Deliver()
    MsgState = 1
    
    If sServer = "" Then
        Err.Raise 8, "Deliver", "Server property not set."
    End If
    If sToEMail = "" Then
        Err.Raise 8, "Deliver", "Receiver property not set."
    End If
    If sFromEMail = "" Then
        Err.Raise 8, "Deliver", "From property not set."
    End If

    If sExplicitSSL Then
        Socket.SSL = True
    End If
    Socket.Connect Socket.resolveip(sServer), sPort


End Sub

Public Sub Cancel()
    If MsgState = 1 Then
        MsgState = 2
    End If
End Sub

Private Sub SocketSend(ByVal outData As String)
    If MsgState = 1 Then
        Socket.sendstring outData
        Debug.Print outData
        WaitForMessage250
    End If
End Sub
Private Sub SocketSSL(ByVal outData As String)
    If MsgState = 1 Then

        Socket.sendstring outData
        Debug.Print outData
        WaitForMessage250
    End If
End Sub
Private Sub WaitForConnect()
    
    TimedOut = False
    TimeOutTimer.Enabled = True
    
    Do Until (TimedOut Or SockConnected Or ReceivedError <> "" Or (MsgState > 1)) Or (Not TimeOutTimer.Enabled)
        DoTasks
    Loop
    
    TimeOutTimer.Enabled = False
    If (MsgState = 2) Then
        Err.Raise 8, "Cancel", "Delivery canceled."
    End If
    If TimedOut Then
        Err.Raise 8, "Deliver", "The server" & IIf((sServer = ""), "", ", " & sServer & ",") & " did  response has timed out."
    End If
    If ReceivedError <> "" Then
        Err.Raise 8, "Deliver", ReceivedError
    End If
    
End Sub

Private Sub WaitForMessage250()
    
    TimedOut = False
    TimeOutTimer.Interval = 20000
    TimeOutTimer.Enabled = True
    
    Do Until (TimedOut Or Received250 Or ReceivedError <> "" Or (MsgState > 1)) Or (Not Socket.Connected) Or (Not TimeOutTimer.Enabled)
        DoTasks
    Loop
    
    TimeOutTimer.Enabled = False
    If (MsgState = 2) Then
        Err.Raise 8, "Cancel", "Delivery canceled."
    End If
    If TimedOut Then
        Err.Raise 8, "Deliver", "The server" & IIf((sServer = ""), "", ", " & sServer & ",") & " did  response has timed out."
    End If
    If ReceivedError <> "" Then
        Err.Raise 8, "Deliver", ReceivedError
    End If
    
    Received250 = False
    
End Sub

Private Sub SendMimeAttachement(strPath As String)
    Dim l As Long, i As Long, FileIn As Long
    Dim strTemp As String
    Dim b As Integer
    Dim Base64Tab As Variant
    Dim bIn(3) As Byte
    Dim s As Variant
    
    'Base64Tab holds the encode tab
    Base64Tab = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "+", "/")
    
    FileIn = FreeFile
    Open strPath For Binary As FileIn
    
    'Preparing the Mime Header
    strTemp = vbCrLf & "--NextMimePart" & vbNewLine
    strTemp = strTemp & "Content-Type: application/octet-stream; name=" & Chr(34) & GetFileName(strPath) & Chr(34) & vbNewLine
    strTemp = strTemp & "Content-Transfer-Encoding: base64" & vbNewLine
    strTemp = strTemp & "Content-Disposition: attachment; filename=" & Chr(34) & GetFileName(strPath) & Chr(34) & vbNewLine
    Socket.Send strTemp & vbCrLf
    
    l = LOF(FileIn) - (LOF(FileIn) Mod 3)
    
    For i = 1 To l Step 3
        'Read three bytes
        Get FileIn, , bIn(0)
        Get FileIn, , bIn(1)
        Get FileIn, , bIn(2)
        
        'Always wait until there're more then 64 characters
        If Len(s) > 64 Then
            Do
            Loop Until Socket.Connected
            DoTasks
            s = s & vbCrLf
            Socket.Send s
            s = ""
        End If
        
        'Calc Base64-encoded char
        b = (bIn(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
        
        'the character s holds the encoded chars
        s = s & Base64Tab(b)
        
        b = ((bIn(0) And &H3) * 16) Or ((bIn(1) \ 16) And &HF)
        s = s & Base64Tab(b)
        b = ((bIn(1) And &HF) * 4) Or ((bIn(2) \ 64) And &H3)
        s = s & Base64Tab(b)
        b = bIn(2) And &H3F
        s = s & Base64Tab(b)
    Next i
    
    If Not (LOF(FileIn) Mod 3 = 0) Then 'check if there is something left
        For i = 1 To (LOF(FileIn) Mod 3) 'Reads the number of bytes left
            Get FileIn, , bIn(i - 1)
        Next i
        If (LOF(FileIn) Mod 3) = 2 Then 'If there are only 2 chars left
            b = (bIn(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            b = ((bIn(0) And &H3) * 16) Or ((bIn(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            b = ((bIn(1) And &HF) * 4) Or ((bIn(2) \ 64) And &H3)
            s = s & Base64Tab(b)
            s = s & "="
        Else 'If there is only one char left
            b = (bIn(0) \ 4) And &H3F 'right shift 2 bits (&H3F=111111b)
            s = s & Base64Tab(b)
            b = ((bIn(1) And &H3) * 16) Or ((bIn(1) \ 16) And &HF)
            s = s & Base64Tab(b)
            s = s & "=="
        End If
    End If
    
    If s <> "" Then 'Send the characters left
        s = s & vbCrLf
        Socket.Send s
    End If
    
    'Send the last part of the MIME Body
    Socket.Send vbCrLf & "--NextMimePart--" & vbCrLf
    Close FileIn
End Sub



Private Sub Class_Initialize()

    Set Socket = New NTAdvFTP61.Socket
    Set TimeOutTimer = New NTSchedule20.Timer
    TimeOutTimer.Interval = 20000
    
    MsgState = 0
    SockConnected = False
    Received250 = False
    
    sMessageType = "text/plain"
    sPort = 25
End Sub

Public Property Get ImplicitSSL() As Boolean
    ImplicitSSL = Socket.SSL
End Property
Public Property Let ImplicitSSL(ByVal RHS As Boolean)

    Socket.SSL = RHS
End Property

Private Sub Class_Terminate()
    MsgState = 0
    TimeOutTimer.Enabled = False
    Set TimeOutTimer = Nothing
    Socket.Disconnect
    Set Socket = Nothing
    
    Debug.Print "OK"
End Sub


Private Sub Socket_Connected()
    MsgState = MsgState + 1
    

End Sub

Private Sub Socket_DataArriving()
    Dim Data As String
    Dim line As String
    
    Socket.ReadString Data
    Debug.Print Data
    Do Until Data = ""
        line = RemoveNextArg(Data, vbCrLf)
        Select Case Left(line, 3)
            Case "220", "HEL"
                SockConnected = True
                Received250 = True
            Case "250"
                Received250 = True
            Case "354"
                Received250 = True
            Case "221"
                Received250 = True
            Case "334"
                Debug.Print Decode64(Mid(line, 5))
                Received250 = True
            Case "235"
                Received250 = True
        End Select
        If Left(line, 1) = "5" Then
            ReceivedError = line
        End If
    Loop
    
    Socket_SendComplete
End Sub

Private Sub Socket_Disconnected()
    Socket.Disconnect
    SockConnected = False
                Received250 = True
    MsgState = 0
End Sub

Private Sub Socket_SendComplete()
comeagain:
    
    MsgState = MsgState + 1
    If MsgState = 3 Then
        
        Socket.sendstring "HELO " & Socket.resolveip(Socket.localhost) & vbCrLf
           ' Socket.SendNextPart
        
    ElseIf MsgState = 4 Then
        
        If sExplicitSSL Then
            Socket.SSL = True

            Socket.sendstring "STARTTLS" & vbCrLf
               ' Socket.SendNextPart
            
        Else
            GoTo comeagain
        End If
    ElseIf MsgState = 5 Then
        If sExplicitSSL Then
            Socket.SSL = True '.ExplicitSSL
        Else
            GoTo comeagain
        End If
    ElseIf MsgState = 6 Then
    If sUsername <> "" Or sUsername <> "" Then
            Socket.sendstring "AUTH LOGIN" & vbCrLf
               ' Socket.SendNextPart
        Else
            GoTo comeagain
            
                    End If
    ElseIf MsgState = 7 Then
        
           'WaitForConnect

            If sUsername <> "" Then
                Socket.sendstring Encode64(sUsername) & vbCrLf
                 '   Socket.SendNextPart
            Else
                GoTo comeagain
            End If

    ElseIf MsgState = 8 Then
            If sPassword <> "" Then
                Socket.sendstring Encode64(sPassword) & vbCrLf
                 '   Socket.SendNextPart
            Else
            GoTo comeagain
            End If
    ElseIf MsgState = 9 Then
        Socket.sendstring "MAIL FROM: <" & sFromEMail & ">" & vbCrLf
           ' Socket.SendNextPart
    ElseIf MsgState = 10 Then
        Socket.sendstring "RCPT TO: <" & sToEMail & ">" & vbCrLf
          '  Socket.SendNextPart
    ElseIf MsgState = 11 Then
        Socket.sendstring "DATA" & vbCrLf
         '   Socket.SendNextPart
    ElseIf MsgState = 12 Then
            Dim fullMsg As String
            fullMsg = fullMsg & "Date: " & Now & vbCrLf
            fullMsg = fullMsg & "From: " & Sender & vbCrLf
            fullMsg = fullMsg & "To: " & Receiver & vbCrLf
            fullMsg = fullMsg & "Subject: " & SubjectText & vbCrLf
            'fullMsg = fullMsg & "X-Mailer: Neotext SMTP Object " & AppVersion(True) & vbCrLf
            fullMsg = fullMsg & "Mime-Version: 1.0" & vbCrLf
            If Attachment <> "" Then
                fullMsg = fullMsg & "Content-Type: multipart/mixed; boundary=" & Chr(34) & "NextMimePart" & Chr(34) & vbCrLf & vbCrLf
                fullMsg = fullMsg & "This is a multi-part message in MIME format." & vbCrLf
                fullMsg = fullMsg & "--NextMimePart" & vbCrLf
            End If
            fullMsg = fullMsg & "Content-type: " & MessageType & ";" & vbCrLf
            fullMsg = fullMsg & vbCrLf & MessageData
                    
            Socket.Send fullMsg
              '  Socket.SendNextPart
            Debug.Print fullMsg
    ElseIf MsgState = 13 Then
            If Attachment <> "" Then
                SendMimeAttachement Attachment
            Else
                GoTo comeagain
                
            End If
    ElseIf MsgState = 14 Then
           Socket.sendstring vbCrLf & "." & vbCrLf
             '  Socket.SendNextPart
    ElseIf MsgState = 15 Then
            Socket.sendstring "QUIT" & vbCrLf
           '     Socket.SendNextPart
    ElseIf MsgState = 16 Then
    Socket.Disconnect
    End If



End Sub

Private Sub TimeOutTimer_OnTicking()
    TimedOut = True
End Sub

