VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThreadedWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Code for a MultiUse ThreadedWindow class.
Private mMainApp As MainApp
Private mfrm As New frmThreadedWindow

Public Sub Initialize(ByVal ma As MainApp, ByVal Listen As Boolean)
   Set mMainApp = ma
   Set mfrm.ThreadedWindow = Me
   mfrm.Caption = mfrm.Caption & " (" & App.ThreadID & ")"
   mfrm.Top = ma.Form.Top
   mfrm.Show
   If Listen Then
        mfrm.Tag = "Listen"
        mfrm.Label1.Caption = "Listen"
        Set mfrm.Sock = New NTAdvFTP61.socket
        mfrm.Sock.Listen mfrm.Sock.LocalHost, 300
        mfrm.Left = ma.Form.Left - mfrm.Width
   Else
        mfrm.Tag = "Client"
        mfrm.Label1.Caption = "Client"
        Set mfrm.Sock = New NTAdvFTP61.socket
        mfrm.Sock.Connect mfrm.Sock.LocalHost, 300
        mfrm.Left = ma.Form.Left + ma.Form.Width
   End If
End Sub

Friend Sub Closing()
    mfrm.Sock.Disconnect
    Set mfrm.Sock = Nothing
    If Not mfrm.Acpt Is Nothing Then
        mfrm.Acpt.Disconnect
        Set mfrm.Acpt = Nothing
    End If
    
   Set mfrm = Nothing
End Sub
