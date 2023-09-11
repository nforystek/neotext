VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Option Compare Binary
'the inner hidden for each object for the collection, making use of this
'allows for the collection to have mixed types of variable in one list
'such as object and primitives like a long, it also extents the key and index
'property to the user when in for each or accessing the variable from the
'collection, including the user defined interfaces no different via defaults.
'exposed this object for the most part acts as if hidden, and is also private.

Private propKey As Variant
Private propType As Boolean
Private propObject As Object
Private propVariant As Variant
Private propCollection As Object

Public Property Get Value()
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
    If propType Then
        Set Value = propObject
    Else
        Value = propVariant
    End If
End Property

Friend Property Set Value(ByVal RHS As IUnknown)
    propType = True
    Set propObject = RHS
    propVariant = Empty
End Property

Friend Property Let Value(ByVal RHS)
    propType = False
    Set propObject = Nothing
    propVariant = RHS
End Property

Public Property Get Key() As Variant
    Key = propKey
End Property
Friend Property Let Key(ByVal RHS As Variant)
    propKey = RHS
End Property

Public Property Get Index() As Long
    Dim Check As Object
    For Each Check In Parent
        Index = Index + 1
        If ObjPtr(Check) = ObjPtr(Me) Then
            Exit For
        End If
    Next
End Property

Friend Property Get Parent() As Object
    Set Parent = propCollection
End Property
Friend Property Set Parent(ByRef RHS As Object)
    Set propCollection = RHS
End Property

Private Sub Class_Terminate()
    Set propCollection = Nothing
    Set propObject = Nothing
    propVariant = Empty
End Sub

