VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private pForm As MotionTypes

Private pKey As String

Private pSerialize As Boolean
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Public Property Get Clone() As Motion
    Set Clone = New Motion
    With Clone
        .Key = pKey
        .Form = pForm
        .Serialize = pSerialize
    End With
End Property

Public Property Get Key() As String
Attribute Key.VB_UserMemId = 0
Attribute Key.VB_MemberFlags = "600"
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

Public Property Get Form() As Long
    Form = pForm
End Property
Public Property Let Form(ByRef RHS As Long)
    pForm = RHS
End Property


Private Sub Class_Initialize()
    Stats_Motion_Count = Stats_Motion_Count + 1
    pKey = "K" & ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    Stats_Motion_Count = Stats_Motion_Count - 1
End Sub
