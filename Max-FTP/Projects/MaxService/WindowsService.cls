VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public WithEvents Controller As NTService20.Controller
Attribute Controller.VB_VarHelpID = -1
Public WithEvents ServiceTimer As NTSchedule20.Timer
Attribute ServiceTimer.VB_VarHelpID = -1

Private Sub Class_Initialize()
    On Error GoTo failstartup
    
    Set Controller = New NTService20.Controller
    If (Command = "") Then
        Set ServiceTimer = New NTSchedule20.Timer
        ServiceTimer.Interval = 40
    End If
    
    Exit Sub
failstartup:
    Err.Clear
    Set ServiceTimer = Nothing

End Sub

Private Sub Class_Terminate()
    If Not ServiceTimer Is Nothing Then
        ServiceTimer.Enabled = False
        Set ServiceTimer = Nothing
    End If
    Set Controller = Nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    LoadSchedule
    PauseSchedules False
    Success = True
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    PauseSchedules True
    UnloadSchedules
    Success = True
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



