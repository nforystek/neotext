VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Driver"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private eventCallBacks() As Long
Private classCallBacks As Callbacks
Private countCallBacks As Long

Public Event Callback()

Public Sub AddCallBack(ByVal lpProcAddress As Long)
    'this will be called to define any sub as a event
    'from outside of this dll, by supplyinh its address
    countCallBacks = countCallBacks + 1
    If classCallBacks Is Nothing Then
        Set classCallBacks = New Callbacks
        ReDim eventCallBacks(1 To countCallBacks) As Long
        'we are using the class Callbacks to make a procedure to
        'call that is dynamic, or created on demand in otherwords
        memCallBacks = ObjectPointers(classCallBacks, 3, 0, 0, 1)
    Else
        ReDim Preserve eventCallBacks(1 To countCallBacks) As Long
    End If
    eventCallBacks(countCallBacks) = lpProcAddress
End Sub


Private Sub TrapCallback(ByVal CallbackEvent As Long)
    On Error Resume Next
    On Error GoTo errhandle:
    
    If Not classCallBacks Is Nothing Then
        If eventCallBacks(CallbackEvent) <> 0 Then
            'swap out the dummy class procedure to reflect the address call back we need
            RtlMoveMemory ByVal memCallBacks(0).Pointer, eventCallBacks(CallbackEvent), 4
            classCallBacks.ProcedureDefined1 'thus upon calling the user defiend sub runs
            RtlMoveMemory ByVal memCallBacks(0).Pointer, memCallBacks(0).Address, 4
        End If
    End If
    
    Exit Sub
    On Error GoTo 0
errhandle: 'technically we don't want this, assuming it's their procedure
'then we throw the exception terminally to goto 0 like VB6 would halt on.
    RtlMoveMemory ByVal memCallBacks(0).Pointer, memCallBacks(0).Address, 4
    If Err.Number <> -1 Then
        Err.Raise Err.Number, Err.Source, Err.Description
        Resume
    End If
End Sub
Friend Sub RaiseCallBack()
    If countCallBacks > 0 Then
        'here is what the frmEvents actually uses to call the user defined event
        'at the default it is merely running a timer event, but potentially your
        'form is as custom as you will it to be for whatever system or app events.
        On Error GoTo handlemore:
        GoTo handleless:
handlemore:
        On Error Resume Next
        Err.Raise Err.Number, Err.Source, Err.Description
        On Error GoTo 0
handleless:

        If countCallBacks > 0 Then
            Dim cnt As Long
            For cnt = 1 To countCallBacks
                TrapCallback cnt
            Next
        End If
        RaiseEvent Callback
    End If
End Sub

Private Sub Class_Initialize()
    Load frmEvents
    Set frmEvents.Parent = Me
End Sub

Private Sub Class_Terminate()
    Unload frmEvents
End Sub
