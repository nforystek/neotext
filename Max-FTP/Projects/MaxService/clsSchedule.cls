VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Sid As Long
Private sName As String

Private sAction As Boolean
Private sCancel As Boolean

Private Operations As NTNodes10.Collection

Private WithEvents oRemindMe As NTSchedule20.Schedule
Attribute oRemindMe.VB_VarHelpID = -1

Public Property Get RemindMe() As NTSchedule20.Schedule
    Set RemindMe = oRemindMe
End Property

Private Sub Class_Initialize()
    Set Operations = New NTNodes10.Collection
    Set oRemindMe = New NTSchedule20.Schedule
End Sub

Private Sub Class_Terminate()
    ClearOperations
    Set oRemindMe = Nothing
    Set Operations = Nothing
End Sub

Private Sub oRemindMe_ScheduledEvent()
    If IsServiceFormStarted Then
        MessageQueueAdd ServiceFileName, "/runschedule " & Trim(Sid)
    End If
End Sub

Public Property Get ID() As Long
    ID = Sid
End Property
Public Property Let ID(ByVal newValue As Long)
    Sid = newValue
End Property
Public Property Get Name() As String
    Name = sName
End Property
Public Property Let Name(ByVal newValue As String)
    sName = newValue
End Property

Public Property Get Cancel() As Boolean
    Cancel = sCancel
End Property
Public Property Let Cancel(ByVal newval As Boolean)
    sCancel = newval
End Property

Public Sub RunSchedule()
    If Not sAction Then
        sCancel = False
        If Operations.Count > 0 Then
            RemindMe.Enabled = False
            sAction = True
            Dim cnt As Long
            For cnt = 1 To Operations.Count
                GetOperation(cnt).RunOperation
                If sCancel Then Exit For
            Next
            sAction = False
            RemindMe.Enabled = IsServiceFormStarted
        End If
    End If
End Sub

Public Sub StopSchedule()
    RemindMe.Enabled = False
    If sAction Then
        sCancel = True
        If Operations.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To Operations.Count
                GetOperation(cnt).StopOperation
            Next
        End If
        sAction = False
    End If
End Sub

Public Function AddOperation() As Operation
    Dim obj  As New Operation
    Operations.Add obj
    Set obj = Nothing
    Set AddOperation = Operations(Operations.Count)
End Function

Public Function GetOperation(ByVal Index As Long) As Operation
    Set GetOperation = Operations(Index)
End Function

Public Sub ClearOperations()
    StopSchedule
    If Operations.Count > 0 Then
        Do Until Operations.Count = 0
            Operations.Remove 1
        Loop
    End If
End Sub

Public Property Get OperationCount() As Long
    OperationCount = Operations.Count
End Property


