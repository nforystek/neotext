VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Text

Public WithEvents Controller As NTService20.Controller
Attribute Controller.VB_VarHelpID = -1

Private Sub Class_Initialize()
    On Error Resume Next
    Set Controller = New NTService20.Controller
    If Err Then
        Err.Clear
        Set Controller = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    Set Controller = Nothing
End Sub

Private Sub Controller_ContinueService(Success As Boolean)
    Success = True
End Sub

Private Sub Controller_PauseService(Success As Boolean)
    Success = True
End Sub

Private Sub Controller_StartService(Success As Boolean)
    Success = True
End Sub

Private Sub Controller_StopService()
    StopService
End Sub
