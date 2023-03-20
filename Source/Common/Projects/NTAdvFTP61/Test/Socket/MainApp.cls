VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Code for a MainApp class.
Private mfrmMTMain As New frmMTMain

Private Sub Class_Initialize()
   Set mfrmMTMain.MainApp = Me
   mfrmMTMain.Caption = mfrmMTMain.Caption & " (" & App.ThreadID & ")"
   mfrmMTMain.Show
   Refresh
End Sub

Friend Property Get Form() As Form
    Set Form = mfrmMTMain
End Property
Friend Sub Refresh()
   mfrmMTMain.Label1.Caption = Forms.Count
End Sub
Friend Sub Closing()
   Set mfrmMTMain = Nothing
End Sub

Private Sub Class_Terminate()
   ' Clean up the hidden window.
   Call FreeProcessWindow
End Sub

