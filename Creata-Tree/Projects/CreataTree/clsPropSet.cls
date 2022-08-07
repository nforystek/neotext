VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPropSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private cCol As Collection

Public Property Get Count() As Long
    Count = cCol.Count
End Property

Public Property Get Item(ByVal cKey As Variant) As clsProp
    Set Item = cCol.Item(SafeKey(cKey))
End Property
Public Property Get Items() As Collection
    Set Items = cCol
End Property

Public Function Exists(ByVal cKey As Variant) As Boolean
    If cCol.Count > 0 Then
        Exists = True
        Dim Test As clsProp
        On Error Resume Next
            Set Test = cCol(SafeKey(cKey))
            If Err Then
                Err.Clear
                Exists = False
            End If
        On Error GoTo 0
    Else
        Exists = False
    End If
End Function

Public Sub Add(ByVal cKey As String, ByVal cValue As Variant)
    Dim nProp As New clsProp
    nProp.SetProperty cKey, cValue
    cCol.Add nProp, nProp.Key
    Set nProp = Nothing
End Sub

Public Sub Remove(ByVal cKey As Variant)
    cCol.Remove SafeKey(cKey)
End Sub

Private Sub Class_Initialize()
    Set cCol = New Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection cCol, True, True
    Set cCol = Nothing
End Sub


