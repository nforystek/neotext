VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public WithEvents ServiceTimer As NTSchedule20.Timer
Attribute ServiceTimer.VB_VarHelpID = -1
Public WithEvents Controller As NTService20.Controller
Attribute Controller.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Controller = New NTService20.Controller
    Set ServiceTimer = New NTSchedule20.Timer
    ServiceTimer.Interval = 50
End Sub

Private Sub Class_Terminate()
    Set ServiceTimer = Nothing
    Set Controller = Nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    Success = True
    Operations.Load
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    Success = True
    Operations.Clear
End Sub

Private Sub Controller_StartService(Success As Boolean)
    Success = True
End Sub

Private Sub Controller_StopService()
    StopService
End Sub

Private Sub ServiceTimer_OnTicking()
    ProcessMessage
End Sub
