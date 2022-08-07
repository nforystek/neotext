VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   600
      Left            =   2685
      TabIndex        =   1
      Top             =   510
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   660
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim ncode
    Dim objMail As NTSMTP23.EMail
    Dim objconstants
    Dim objsmtpserver



Private Sub Command1_Click()

        On Error GoTo catcherr
    

    Set ncode = CreateObject("NTCipher10.NCode")

    Set objMail = New NTSMTP23.EMail
    
    With objMail
        .Sender = "<nforystek@neotext.org>"
        .Receiver = "<9524579224@tmomail.net>"
        
        .Server = "mail.smtp2go.com"
        .Port = 25
        '.Username = "nforystek@neotext.org"
        '.password = "ZFdeWuNhyRxY"
       '' .Username = "smstextrelay"
       ' .Password = "N2gydXJmOHl1NDh6"
        
         .SubjectText = Now
        .MessageData = "Is the time"
      '  .explicitssl = True
        .Deliver
        
    End With
'
'    objMail.Clear
'
'    objMail.FromAddress = "sosouix@gmail.com"
'    objMail.FromName = "Nicholas Forystek"
'
'    objMail.Subject = "SENSOR ALERT AT " & Now
'    objMail.BodyPlainText = "SENSOR ALERT AT " & Now
'    objMail.Priority = 1
'
'    objMail.Encoding = 0
'
'
'    objMail.AddTo "9524579224@tmomail.net", ""

        
   ' objsmtpserver.Clear
    
   ' objsmtpserver.LogFile = "C:\SmtpLog.txt"
    

   ' objsmtpserver.SetSecure 465


    'objsmtpserver.Connect "smtp.gmail.com", "sosouix@gmail.com", ncode.decryptstring("E9C7E9CACEEDC7C3C0E9", "sosouix@gmail.com", True)
    
    'objsmtpserver.Send objMail
    
    'objsmtpserver.Disconnect


    Exit Sub
catcherr:
    Debug.Print Err.Description
End Sub

Private Sub Command2_Click()
    
    Set objMail = Nothing
    Set objconstants = Nothing
    Set objsmtpserver = Nothing
    Set ncode = Nothing
    Debug.Print "DONE"
End Sub
