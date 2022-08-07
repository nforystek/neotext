VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemSet"
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

Friend Property Get Item(ByVal cKey As Variant) As clsItem
    Set Item = cCol.Item(SafeKey(cKey))
End Property
Public Property Get Items() As Collection
    Set Items = cCol
End Property
Friend Function Exists(ByVal cKey As Variant) As Boolean
    If cCol.Count > 0 Then
        Exists = True
        Dim Test As clsItem
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

Friend Function Add(ByVal nXML As String) As String
    Dim nItem As New clsItem
    nItem.XMLText = nXML
    cCol.Add nItem, nItem.Key
    Add = nItem.Key
    Set nItem = Nothing
End Function

Friend Sub Remove(ByVal cKey As Variant)
    cCol.Remove SafeKey(cKey)
End Sub

Private Sub Class_Initialize()
    Set cCol = New Collection
End Sub

Private Sub Class_Terminate()
    ClearCollection cCol, True, True
    Set cCol = Nothing
End Sub


