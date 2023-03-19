VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Dim client As NTCipher10.tractor
Dim server As NTCipher10.tractor
Dim temp As String


Private Sub Command1_Click()
    temp = client.Initial
End Sub

Private Sub Command2_Click()
    Debug.Print server.Checksum(client.seeding)
    Debug.Print client.Checksum(server.seeding)
    
  '  Debug.Print client.Checksum(client.Seed, server.para)
  '  Debug.Print server.Checksum(server.para, server.Seed)

    
End Sub

Private Sub Command3_Click()
    Set client = Nothing
    Set client = New NTCipher10.tractor
    
End Sub

Private Sub Form_Load()

    Dim p As New PoolID
    Do While True
        DoEvents
        Debug.Print p.Generate
    Loop
    
'
'    Dim uu As New NTCipher10.UUCode
'    uu.UUEncode "C:\Documents and Settings\Nickels\My Documents\IMG_20200421_0001.pdf"
'    uu.UUEncode "C:\Documents and Settings\Nickels\My Documents\IMG_20200421_0002.pdf"
'    uu.UUEncode "C:\Documents and Settings\Nickels\My Documents\Proposed Order of Document.pdf"
'End

    Set client = New NTCipher10.tractor
    Set server = New NTCipher10.tractor
    client.TimeSync = "time.nist.gov"
    server.TimeSync = "time.nist.gov"
    client.Initiate
    server.Initiate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set client = Nothing
    Set server = Nothing

End Sub
