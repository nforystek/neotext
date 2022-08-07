VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaxService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Public WithEvents Controller As NTService20.Controller
Attribute Controller.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set Controller = New NTService20.Controller
End Sub

Private Sub Class_Terminate()
    Set Controller = Nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    LoadSchedule
    Success = True
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    UnloadSchedules
    Success = True
End Sub

Private Sub Controller_StartService(Success As Boolean)
    Success = True
End Sub

Private Sub Controller_StopService()
    StopService
End Sub



