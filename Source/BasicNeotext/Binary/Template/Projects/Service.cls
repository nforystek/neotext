VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'make our cache object controller and available events
Public WithEvents Controller As BasicService.Controller
Attribute Controller.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Controller = New BasicService.Controller 'lbind
End Sub

Private Sub Class_Terminate()
    Set Controller = Nothing 'properlly set to nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    Success = True 'always set success to your worries
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    Success = True 'passed it back acknowledge success
End Sub

Private Sub Controller_StartService(Success As Boolean)
    Success = True 'this can occur solo and cause main
End Sub

Private Sub Controller_StopService()
    StopService 'call our premaid of sub main reversed
End Sub
