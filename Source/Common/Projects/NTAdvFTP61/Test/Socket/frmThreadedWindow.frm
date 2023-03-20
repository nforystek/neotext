VERSION 5.00
Begin VB.Form frmThreadedWindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmThreadWindow"
   ClientHeight    =   1845
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1485
      Top             =   210
   End
   Begin VB.TextBox Text1 
      Height          =   1275
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   510
      Width           =   4545
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Send"
      Height          =   372
      Left            =   2820
      TabIndex        =   3
      Top             =   75
      Width           =   612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   372
      Left            =   3540
      TabIndex        =   2
      Top             =   75
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   372
      Left            =   2100
      TabIndex        =   1
      Top             =   75
      Width           =   612
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmThreadedWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Code for the form frmThreadedWindow.
Public ThreadedWindow As ThreadedWindow

Public WithEvents Sock As NTAdvFTP61.socket
Attribute Sock.VB_VarHelpID = -1
Public WithEvents Acpt As NTAdvFTP61.socket
Attribute Acpt.VB_VarHelpID = -1

Private CommLayer As Integer
Private Parity As Integer
Private StopBit As Integer

Private Sub Acpt_Connected()
    Sock_Connected
End Sub

Private Sub Acpt_Connection(ByRef Handle As Long)
    Sock_Connection Handle
End Sub

Private Sub Acpt_DataArriving()
    Sock_DataArriving
End Sub

Private Sub Acpt_Disconnected()
    Sock_Disconnected
End Sub

Private Sub Acpt_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    Sock_Error Number, Source, Description
End Sub

Private Sub Acpt_SendComplete()
    Sock_SendComplete
End Sub

Private Sub Command1_Click()
    If Me.Tag = "Listen" Then
        Sock.Disconnect
        Sock.Listen Sock.LocalHost, 300
        If Not Acpt Is Nothing Then
            Acpt.Disconnect
            Set Acpt = Nothing
        End If
    Else
        Sock.Disconnect
        Sock.Connect Sock.LocalHost, 300
    End If
End Sub

Private Sub Command2_Click()
    If Me.Tag = "Listen" Then
        If Not Acpt Is Nothing Then
            Acpt.Disconnect
        Else
            Sock.Disconnect
        End If
    Else
        Sock.Disconnect
    End If
End Sub

Private Sub Sock_Connected()
    DebugSocket Me.Tag & "_Connected"
End Sub

Private Sub Sock_Connection(ByRef Handle As Long)
    DebugSocket Me.Tag & "_Connection(" & Handle & ")"
    '### ACCEPT STAY LISTEN
    Set Acpt = New socket
    Acpt.Accept Handle
    
    '### DECLINE STAY LISTEN
    'Set Acpt = New Socketq
    'Acpt.Decline
    
    '### ACCEPT w/NO LISTEN
    'Sock.Accept Handle
  
    '### DECLINE w/NO LISTEN
    'Sock.Decline Handle
End Sub

Private Sub Sock_Disconnected()
    DebugSocket Me.Tag & "_Disconnected"
End Sub

Private Sub Command3_Click()
    If Me.Tag = "Listen" Then
        SendOut "This is a sentence being sent to the cleint."
    Else
        SendOut "This is a sentence being sent to the server."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call ThreadedWindow.Closing
   Set ThreadedWindow = Nothing
End Sub

Private Sub Sock_Error(ByVal Number As Long, ByVal Source As String, ByVal Message As String)
    DebugSocket Me.Tag & "_Error " & Number & " " & Message
End Sub

Private Sub Sock_SendComplete()
    DebugSocket Me.Tag & "_SendComplete"
End Sub

Private Sub DebugSocket(ByVal txt As String)
    Text1.Text = Text1.Text & txt & vbCrLf
    Debug.Print vbTab & txt
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub Sock_DataArriving()
'    Dim data As String
'    If Not Acpt Is Nothing Then
'        data = Acpt.read
'    Else
'        data = Sock.read
'    End If
'    DebugSocket Me.Tag & "_DataArriving " & data
End Sub

'Private Sub SendOut(ByVal Text As String)
'    If Not Acpt Is Nothing Then
'        Acpt.sendstring Text
'    Else
'        Sock.sendstring Text
'    End If
'End Sub


Private Function SendOut(Optional ByVal Text As String = "") As String
    Static out As String
    If Text <> "" Then 'supplied text then build up out
    'buffer as byte headed records, size of following data
        Do While Text <> ""
            If Len(Text) > 255 Then
                out = out & Chr(255) & Left(Text, 255)
                Text = Mid(Text, 256)
            Else
                out = out & Chr(Len(Text)) & Text
                Text = ""
            End If
        Loop
    ElseIf out <> "" Then
        'no parameter, so requesting any next byte to send
        SendOut = Left(out, 1)
        out = Mid(out, 2)
    Else 'no text, and no out data, assuming request for info,
    'sock.send will immediate return for nullstring, no send
        SendOut = ""
    End If
End Function

Private Sub Timer1_Timer()
    Dim use As socket
    If Not Acpt Is Nothing Then
        Set use = Acpt
    Else
        Set use = Sock
    End If
    
    Static data As String
    Dim bdat() As Byte
    
    If use.Connected Then
        Dim bl As Byte
        
        use.ReadBytes bdat
        data = data & Convert(bdat)
        
        'data = data & use.Read
        If data <> "" Then
            bl = Asc(Left(data, 1))
        ElseIf data = "" Then
            'no data record
            bdat = Convert(SendOut)
            use.Sendbytes bdat
            Erase bdat
            
            'use.Send SendOut   'do any send
        End If

        If (Len(Left(data, bl + 1)) = (bl + 1)) And (bl > 0) Then
            'buffer contains amount of data for record, or more
            data = Mid(data, 2)
            DebugSocket Me.Tag & "_DataArriving " & Left(data, bl)
            data = Mid(data, bl + 1)
        Else 'buffer is under the amount of data for record
        
            
            use.ReadBytes bdat
            data = data & Convert(bdat)
            Erase bdat
            
            'data = data & use.Read
        End If

    End If
    Set use = Nothing
End Sub

Public Function Convert(Info)
    'slow method of converting byte
    'array to string nd vice versa
    Dim N As Long
    Dim out() As Byte
    Dim ret As String
    Select Case TypeName(Info)
        Case "String"
            If Len(Info) > 0 Then
                ReDim out(0 To Len(Info) - 1) As Byte
                For N = 0 To Len(Info) - 1
                    out(N) = Asc(Mid(Info, N + 1, 1))
                Next
            Else
                ReDim out(-1 To -1) As Byte
            End If
            Convert = out
        Case "Byte()"
            If (ArraySize(Info) > 0) Then
                For N = LBound(Info) To UBound(Info)
                    ret = ret & Chr(Info(N))
                Next
            End If
            Convert = ret
    End Select
End Function
