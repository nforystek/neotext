VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Option Compare Binary

Private xSweet As Sweet
Private xIdentity As Long
Private xRawvalue As String
Private xEndvalue As String

Private List As VBA.Collection

Public Property Get Basename() As String
    Basename = "IListing"
End Property

Public Property Get Typename() As String
    Typename = "Rains"
End Property

Public Property Get Sweet() As Sweet
    Set Sweet = xSweet
End Property

Public Property Set Sweet(ByRef newVal As Sweet)
    Set xSweet = newVal
End Property

Public Property Get Identity() As Long
    Identity = xIdentity
End Property

Public Property Let Identity(ByVal newVal As Long)
    xIdentity = newVal
End Property

Public Static Property Get Rawvalue() As String
    Rawvalue = xRawvalue
End Property

Public Property Let Rawvalue(ByVal newVal As String)
    xRawvalue = newVal
End Property

Public Property Get Endvalue() As String
    Endvalue = xEndvalue
End Property

Public Property Let Endvalue(ByVal newVal As String)
    xEndvalue = newVal
End Property

Private Sub Class_Initialize()
    Set List = New VBA.Collection
    xIdentity = GetNewIdentity
End Sub

Private Sub Class_Terminate()
    ClearCollection List, True
    Set xSweet = Nothing
End Sub

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

Public Sub Add(Obj As Object, Optional IdxKey As Variant)
    If Not Obj Is Nothing Then
        If Obj.Basename = "IListing" Or Obj.Typename = "IListing" Then
            Do Until Obj.Count = 0
                If Not IsMissing(IdxKey) Then
                    List.Add Obj.Item(1), IdxKey
                Else
                    List.Add Obj.Item(1)
                End If
                Obj.Remove 1
            Loop
        Else
            If Not IsMissing(IdxKey) Then
                List.Add Obj, IdxKey
            Else
                List.Add Obj
            End If
        End If
    End If
End Sub

Public Sub Remove(IdxKey As Variant)
    List.Remove IdxKey
End Sub

Public Function Exists(IdxKey As Variant) As Boolean
    Dim Obj As Object
    For Each Obj In List
        If ObjPtr(Obj) = ObjPtr(IdxKey) Then
            Exists = True
            Exit Function
        ElseIf VarPtr(Obj) = VarPtr(IdxKey) Then
            Exists = True
            Exit Function
        ElseIf Obj = IdxKey Then
            Exists = True
            Exit Function
        ElseIf Obj Is IdxKey Then
            Exists = True
            Exit Function
        End If
    Next
End Function

Public Sub Clear()
    ClearCollection List, True
End Sub
