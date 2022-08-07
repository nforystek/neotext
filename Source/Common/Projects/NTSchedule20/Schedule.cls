VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary


Private WithEvents IncTimer As Timer
Attribute IncTimer.VB_VarHelpID = -1

Public Enum ScheduleTypes
    Manual = 0
    Increment = 1
    SetTime = 2
End Enum

Public Enum IncrementTypes
    Minute = 0
    Hour = 1
    Day = 2
End Enum

Public Event ScheduledEvent()

Private sName As String
Private sEnabled As Boolean

Private sScheduleType As ScheduleTypes
Private sIncrementType As IncrementTypes
Private sIncrementInterval As Integer
Private sExecuteDate As String
Private sExecuteTime As String
Private sLastEventTime As String

Friend Property Get hwnd() As Long
    hwnd = IncTimer.hwnd
End Property

Public Property Get Enabled() As Boolean
    Enabled = sEnabled
End Property
Public Property Let Enabled(ByVal newVal As Boolean)
    sEnabled = newVal
    IncTimer.Enabled = newVal
End Property

Public Property Get Name() As String
    Name = sName
End Property
Public Property Let Name(ByVal NewValue As String)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    sName = NewValue
End Property

Public Property Get ScheduleType() As ScheduleTypes
    ScheduleType = sScheduleType
End Property
Public Property Let ScheduleType(ByVal NewValue As ScheduleTypes)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    sScheduleType = NewValue
End Property

Public Property Get IncrementType() As IncrementTypes
    IncrementType = sIncrementType
End Property
Public Property Let IncrementType(ByVal NewValue As IncrementTypes)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    sIncrementType = NewValue
End Property

Public Property Get IncrementInterval() As Integer
    IncrementInterval = sIncrementInterval
End Property
Public Property Let IncrementInterval(ByVal NewValue As Integer)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    If NewValue < 1 Then
        Err.Raise vbObjectError, App.EXEName, "This value can not be set to let then 1."
    End If
    sIncrementInterval = NewValue
End Property

Public Property Get ExecuteDate() As String
    ExecuteDate = sExecuteDate
End Property
Public Property Let ExecuteDate(ByVal NewValue As String)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    sExecuteDate = NewValue
End Property

Public Property Get ExecuteTime() As String
    ExecuteTime = sExecuteTime
End Property
Public Property Let ExecuteTime(ByVal NewValue As String)
    If Enabled Then
        Err.Raise vbObjectError, App.EXEName, "You can not set this value while the schedule is enabled."
    End If
    sExecuteTime = NewValue
End Property

Public Property Get LastEventTime() As String
    LastEventTime = sLastEventTime
End Property

Private Sub Class_Initialize()
    sScheduleType = ScheduleTypes.Manual
    sIncrementType = IncrementTypes.Minute
    sIncrementInterval = 1
    Set IncTimer = New Timer
    IncTimer.Interval = 20
    Enabled = False
End Sub
Private Sub Class_Terminate()
    Enabled = False
    Set IncTimer = Nothing
End Sub

Private Sub RunEvent()
    sLastEventTime = Now
    RaiseEvent ScheduledEvent
End Sub

Private Sub IncTimer_OnTicking()
    Dim sExecuteDateTime As String
    If IsDate(sExecuteDate) And IsDate(sExecuteTime) Then
        sExecuteDateTime = sExecuteDate & " " & sExecuteTime
    ElseIf IsDate(sExecuteTime) Then
        sExecuteDateTime = VBA.Date & " " & sExecuteTime
    ElseIf IsDate(sExecuteDate) Then
        sExecuteDateTime = sExecuteDate & " " & VBA.time
    End If

    If IsDate(sExecuteDateTime) Then
        Select Case sScheduleType
            Case ScheduleTypes.Increment
                If IsIncrementEqual(Now, sExecuteDateTime) Then
                    If Not EqualDate(Now, sLastEventTime) Then
                        RunEvent
                    End If
                End If
            Case ScheduleTypes.SetTime
                If EqualDate(Now, sExecuteDateTime) Then
                    If Not EqualDate(Now, sLastEventTime) Then
                        RunEvent
                    End If
                End If
        End Select
    End If
End Sub

Private Function EqualDate(ByVal Date1 As String, ByVal Date2 As String) As Boolean
    If Not IsDate(Date2) Then
        EqualDate = False
    Else
        EqualDate = ((Year(Date1) = Year(Date2)) And (Month(Date1) = Month(Date2)) And (VBA.Day(Date1) = VBA.Day(Date2))) And (VBA.Hour(Date1) = VBA.Hour(Date2)) And (VBA.Minute(Date1) = VBA.Minute(Date2))
    End If
End Function

Private Function IsIncrementEqual(ByVal Date1 As String, ByVal Date2 As String) As Boolean
    Dim nDiff As Long
    nDiff = DateDiff("n", Date2, Date1)
    If nDiff = 0 Then
        IsIncrementEqual = True
    ElseIf nDiff > 0 Then
        Select Case sIncrementType
            Case IncrementTypes.Minute
                IsIncrementEqual = (((nDiff * 60) Mod (sIncrementInterval * 60)) = 0)
            Case IncrementTypes.Hour
                IsIncrementEqual = (((nDiff * 60) Mod (sIncrementInterval * 3600)) = 0)
            Case IncrementTypes.Day
                IsIncrementEqual = (((nDiff * 60) Mod (sIncrementInterval * 86400)) = 0)
            Case Else
                IsIncrementEqual = False
        End Select
    Else
        IsIncrementEqual = False
    End If
End Function
