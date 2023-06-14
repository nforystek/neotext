VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OnEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pKey As String

Private pEventType As Long

Private pApplyTo As Orbit

Private pConstraint As VBA.Collection

Private pCode As String

Public Event BeforeEvent(ByRef cancel As Boolean)

Public Property Get Key() As String
    Key = pKey
End Property
Public Property Let Key(ByVal RHS As String)
    If RHS <> "" Then
        If pKey <> RHS And pKey <> "" Then If All.Exists(pKey) Then All.Remove pKey
        If All.Exists(RHS) Then
            All.Remove RHS
            All.Add Me, RHS
        End If
        pKey = RHS
    End If
End Property

Public Property Get EventType() As Long
    EventType = pEventType
End Property
Public Property Let EventType(ByVal RHS As Long)
    pEventType = RHS
End Property

Public Property Get ApplyTo() As Object
    Set ApplyTo = pApplyTo
End Property
Public Property Set ApplyTo(ByRef RHS As Object)
    Set pApplyTo = All(RHS)
End Property

Public Property Get Constraint() As VBA.Collection
    Set Constraint = pConstraint
End Property

Public Property Let Constraint(ByVal RHS)
    If pConstraint Is Nothing Then
        Set pConstraint = New Collection
    Else
        Do Until pConstraint.Count = 0
            pConstraint.Remove 1
        Loop
    End If
    Dim inVal As String
    Do Until RHS = ""
        inVal = RemoveNextArg(RHS, ",")
        inVal = UCase(Left(inVal, 1)) & LCase(Mid(inVal, 2))
        pConstraint.Add inVal, inVal
    Loop
End Property

Public Property Get Code() As String
    Code = pCode
End Property
Public Property Let Code(ByVal RHS As String)
    pCode = RHS
End Property

Public Property Let ToString(ByVal RHS As String)

End Property
Public Property Get ToString() As String

End Property
Public Function RunEvent(ByVal ConditionResult As Boolean) As Boolean
    Static Toggle1 As Boolean
    Static Toggle2 As Boolean
    
    If ConditionResult Then
        If Toggle1 = False Then Toggle2 = False
        If (Not Toggle1) Then
            Toggle1 = True
            RunEvent = True
        End If
    End If
    
    If (Not ConditionResult) Then
        If Toggle2 = False Then Toggle1 = False
        If Not Toggle2 Then
            Toggle2 = True
            RunEvent = False
        End If
    End If
    
    If RunEvent Then
        Dim cancel As Boolean
        RaiseEvent BeforeEvent(cancel)
        RunEvent = Not cancel
        If RunEvent Then frmMain.RunEvent pCode
    End If
End Function

Private Sub Class_Initialize()
    Stats_OnEvent_Count = Stats_OnEvent_Count + 1
    Key = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    Stats_OnEvent_Count = Stats_OnEvent_Count - 1
    
    If Not pConstraint Is Nothing Then
        Do Until pConstraint.Count = 0
            pConstraint.Remove 1
        Loop
    End If
    Set pConstraint = Nothing
    
    Set pApplyTo = Nothing
End Sub
