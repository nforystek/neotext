VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   555
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   1155
      Left            =   3540
      TabIndex        =   5
      Top             =   540
      Width           =   2235
   End
   Begin VB.CommandButton Command5 
      Caption         =   "AUTH PASS"
      Height          =   375
      Left            =   900
      TabIndex        =   4
      Top             =   1800
      Width           =   1395
   End
   Begin VB.CommandButton Command4 
      Caption         =   "AUTH LOGIN"
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   1380
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HELP"
      Height          =   375
      Left            =   900
      TabIndex        =   2
      Top             =   960
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   900
      TabIndex        =   1
      Top             =   2460
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   900
      TabIndex        =   0
      Top             =   300
      Width           =   1395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents socket As NTAdvFTP61.socket
Attribute socket.VB_VarHelpID = -1
Private smtp As New NTSMTP23.EMail
    
Private Sub Command1_Click()
    socket.SSL = True
  '  socket.Connect socket.ResolveIP("smtp.gmail.com"), 587
 '   socket.Connect socket.ResolveIP("smtp.gmail.com"), 465
    
    
   ' socket.Connect socket.ResolveIP("mail.smtp2go.com"), 80
  ''
'   socket.Connect socket.ResolveIP("mail.neotext.org"), 587
'socket.Connect socket.ResolveIP("mail.neotext.org"), 587
   
   socket.Connect socket.ResolveIP("ftp.neotext.org"), 990
End Sub

Private Sub Command2_Click()
    socket.Disconnect
End Sub

Private Sub Command3_Click()
    socket.Send "HELO " & socket.ResolveIP(socket.LocalHost) & vbCrLf


End Sub

Private Sub Command4_Click()

    socket.SendString "AUTH LOGIN" & vbCrLf

End Sub

Private Sub Command5_Click()
socket.SendString Encode64("noreply@neotext.org") & vbCrLf

End Sub

Private Sub Command6_Click()


    smtp.Receiver = "9524579224@tmomail.net"
    smtp.Sender = "nforystek@neotext.org"
    smtp.SubjectText = "Test"
    smtp.MessageData = "Test"
    smtp.Port = 25
    smtp.Server = "mail.smtp2go.com"
   ' smtp.Username = "noreply@neotext.org"
   ' smtp.Password = "UWf%2h7f"
  '  smtp.ImplicitSSL = True
     smtp.Deliver
     

End Sub

Private Sub Command7_Click()
    socket.SendString Encode64("UWf%2h7f") & vbCrLf
End Sub

Private Sub Form_Load()


    Set socket = New NTAdvFTP61.socket
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set socket = Nothing
End Sub

Private Sub socket_Certificate(Handle As Long)
Debug.Print Handle

    'socket.decline Handle
    
End Sub

Private Sub socket_Connected()
    Debug.Print "socket_Connected"

End Sub

Private Sub socket_Connection(Handle As Long)
    Debug.Print "socket_Connection(" & Handle & ")"
End Sub

Private Sub socket_DataArriving()
    Dim tmp As String
    tmp = socket.Read
    Debug.Print "socket_DataArriving( " & tmp & ")"
    If InStr(tmp, "Ready to start TLS") > 0 Then
                    socket.SSL = True
                End If
End Sub

Private Sub socket_Disconnected()
    Debug.Print "socket_Disconnected"
    socket.SSL = False
End Sub

Private Sub socket_Error(ByVal Number As Long, ByVal Source As String, ByVal Description As String)
    Debug.Print "socket_Error(" & Number & ", " & Source & ", " & Description & ")"
End Sub

Private Sub socket_SendComplete()
    Debug.Print "socket_SendComplete"
End Sub
