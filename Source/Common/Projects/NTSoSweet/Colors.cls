VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private List As VBA.Collection

Public Property Get Basename() As String
    Basename = "IListing"
End Property

Public Property Get Typename() As String
    Typename = "Colors"
End Property

Public Property Get Count() As Long
    Count = List.Count
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = List.[_NewEnum]
End Function

Public Property Get Item(IdxKey As Variant) As Object
Attribute Item.VB_UserMemId = 0
    Set Item = List(IdxKey)
End Property

Public Sub Add(obj As Object, Optional IdxKey As Variant)
    If Not obj Is Nothing Then
        If obj.Basename = "IListing" Or obj.Typename = "IListing" Then
            Do Until obj.Count = 0
                If Not IsMissing(IdxKey) Then
                    List.Add obj.Item(1), IdxKey
                Else
                    List.Add obj.Item(1)
                End If
                obj.Remove 1
            Loop
        Else
            If Not IsMissing(IdxKey) Then
                List.Add obj, IdxKey
            Else
                List.Add obj
            End If
        End If
    End If
End Sub

Public Sub Remove(IdxKey As Variant)
    List.Remove IdxKey
End Sub

Public Function Exists(IdxKey As Variant) As Boolean
    Dim obj As Object
    For Each obj In List
        If ObjPtr(obj) = ObjPtr(IdxKey) Then
            Exists = True
            Exit Function
        ElseIf VarPtr(obj) = VarPtr(IdxKey) Then
            Exists = True
            Exit Function
        ElseIf obj = IdxKey Then
            Exists = True
            Exit Function
        ElseIf obj Is IdxKey Then
            Exists = True
            Exit Function
        End If
    Next
End Function

Public Sub Clear()
    ClearCollection List, True
End Sub

Private Sub Class_Initialize()
    Set List = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection List, True
End Sub
