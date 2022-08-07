VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pID As Long
Private pName As String
Private pProcedure As String
Private pParameters As String

Public WithEvents rmTimer As NTSchedule20.Schedule
Attribute rmTimer.VB_VarHelpID = -1

Public Property Get ID() As Long
    ID = pID
End Property
Public Property Let ID(ByVal newval As Long)
    pID = newval
End Property

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal NewValue As String)
    pName = NewValue
End Property

Public Property Get Procedure() As String
    Procedure = pProcedure
End Property
Public Property Let Procedure(ByVal NewValue As String)
    pProcedure = NewValue
End Property

Public Property Get Parameters() As String
    Parameters = pParameters
End Property
Public Property Let Parameters(ByVal NewValue As String)
    pParameters = NewValue
End Property

Public Property Get Enabled() As Boolean
    Enabled = rmTimer.Enabled
End Property
Public Property Let Enabled(ByVal newval As Boolean)
    rmTimer.Enabled = newval
End Property

Public Property Get ScheduleType() As NTSchedule20.ScheduleTypes
    ScheduleType = rmTimer.ScheduleType
End Property
Public Property Let ScheduleType(ByVal newval As NTSchedule20.ScheduleTypes)
    rmTimer.ScheduleType = newval
End Property

Public Property Get ExecuteDate() As String
    ExecuteDate = rmTimer.ExecuteDate
End Property
Public Property Let ExecuteDate(ByVal newval As String)
    rmTimer.ExecuteDate = newval
End Property

Public Property Get ExecuteTime() As String
    ExecuteTime = rmTimer.ExecuteTime
End Property
Public Property Let ExecuteTime(ByVal newval As String)
    rmTimer.ExecuteTime = newval
End Property

Public Property Get IncrementType() As NTSchedule20.IncrementTypes
    IncrementType = rmTimer.IncrementType
End Property
Public Property Let IncrementType(ByVal newval As NTSchedule20.IncrementTypes)
    rmTimer.IncrementType = newval
End Property

Public Property Get IncrementInterval() As Long
    IncrementInterval = rmTimer.IncrementInterval
End Property
Public Property Let IncrementInterval(ByVal newval As Long)
    rmTimer.IncrementInterval = newval
End Property
Public Sub StopOperation()
    'rmTimer.Enabled = False
End Sub
Public Sub StartOperation()
    Select Case dbSettings.GetSetting("Language")
        Case "JScript"
            frmService.RunProcedure pProcedure & "(" & pParameters & ");"
        Case "VBScript"
            frmService.RunProcedure pProcedure & " " & pParameters
    End Select
End Sub

Private Sub Class_Initialize()
    Set rmTimer = New NTSchedule20.Schedule
End Sub

Private Sub Class_Terminate()
    Set rmTimer = Nothing
End Sub

Private Sub rmTimer_ScheduledEvent()
    If (dbSettings Is Nothing) Then Exit Sub
    dbSettings.Message ServiceFileName, "startoperation:" & ID
End Sub
