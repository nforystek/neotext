VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN

Private lptrList As VBA.Collection
Private propList As VBA.Collection

Private Sub Class_Initialize()
    Set lptrList = New VBA.Collection
    Set propList = New VBA.Collection
End Sub

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = lptrList.[_NewEnum]
End Function

Public Sub Add(VarObj, Optional Key, Optional Before, Optional After)
    Dim itm As New Item
    If Not IsMissing(Key) Then
        If Exists(Key) Then Err.Raise 8, App.Title, "Key already exists."
        itm.Key = Key
    End If
    If IsObject(VarObj) Then
        Set itm.Value = VarObj
    Else
        itm.Value = VarObj
    End If
    Set itm.Parent = Me
    If IsMissing(Key) Then
        If (IsMissing(After)) And (IsMissing(Before)) Then
            propList.Add itm
            lptrList.Add itm.Value
        ElseIf (Not (IsMissing(After))) And (IsMissing(Before)) Then
            propList.Add itm, , , After
            lptrList.Add itm.Value, , , After
        ElseIf (IsMissing(After)) And (Not (IsMissing(Before))) Then
            propList.Add itm, , Before
            lptrList.Add itm.Value, , Before
        ElseIf (Not (IsMissing(After))) And (Not (IsMissing(Before))) Then
            propList.Add itm, , Before, After
            lptrList.Add itm.Value, , Before, After
        End If
    Else
        If (IsMissing(After)) And (IsMissing(Before)) Then
            propList.Add itm, Key
            lptrList.Add itm.Value, Key
        ElseIf (Not (IsMissing(After))) And (IsMissing(Before)) Then
            propList.Add itm, Key, , After
            lptrList.Add itm.Value, Key, , After
        ElseIf (IsMissing(After)) And (Not (IsMissing(Before))) Then
            propList.Add itm, Key, Before
            lptrList.Add itm.Value, Key, Before
        ElseIf (Not (IsMissing(After))) And (Not (IsMissing(Before))) Then
            propList.Add itm, Key, Before, After
            lptrList.Add itm.Value, Key, Before, After
        End If
    End If
End Sub

Public Sub Remove(Optional KeyIndex)
    Dim idx
    If Not IsMissing(KeyIndex) Then
        If Not Exists(KeyIndex) Then Err.Raise 8, App.Title, "Key does not exist."
        If Not IsNumeric(KeyIndex) Then
            idx = KeyIndex
        Else
            idx = GetKeyIndex(KeyIndex)
        End If
    ElseIf propList.Count = 0 Then
        Err.Raise 8, App.Title, "Collection is empty."
    Else
        idx = 1
    End If
    Dim ref As Item
    Set ref = propList(idx)
    propList.Remove idx
    lptrList.Remove idx
    If IsObject(ref.Value) Then
        Set ref.Value = Nothing
    End If
    Set ref.Parent = Nothing
    Set ref = Nothing
End Sub

Public Property Get Exists(ByVal KeyIndex) As Boolean
    If IsNumeric(KeyIndex) Then
        Exists = (GetIndexKey(KeyIndex) <> "")
    Else
        Exists = (GetKeyIndex(KeyIndex) <> 0)
    End If
End Property

Friend Property Get GetKeyIndex(ByVal Key)
    Dim itm As Object
    Dim idx
    For Each itm In propList
        idx = idx + 1
        If itm.Key = Key Then
            GetKeyIndex = idx
            Exit Property
        End If
    Next
End Property

Public Property Get Count() As Long
    Count = propList.Count
End Property
Friend Property Get GetIndexKey(ByVal Index)
    If Index > 0 And Index <= propList.Count Then
        GetIndexKey = propList(Index).Key
    Else
        GetIndexKey = ""
    End If
End Property

Public Property Get Item(ByVal KeyIndex)
    Dim idx
    If IsNumeric(KeyIndex) Then
        idx = KeyIndex
    Else
        idx = GetKeyIndex(KeyIndex)
    End If
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    If IsObject(propList(idx).Value) Then
        Set Item = propList(idx).Value
    Else
        Item = propList(idx).Value
    End If
End Property

Public Property Set Item(ByVal KeyIndex, ByRef Value)
    Dim idx
    If IsNumeric(KeyIndex) Then
        idx = KeyIndex
    Else
        idx = GetKeyIndex(KeyIndex)
    End If
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    Dim itm As Item
    Set itm = propList(idx)
    Set itm.Value = Value
End Property


Public Property Let Item(ByVal KeyIndex, ByVal Value)
    Dim idx
    If IsNumeric(KeyIndex) Then
        idx = KeyIndex
    Else
        idx = GetKeyIndex(KeyIndex)
    End If
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    Dim itm As Item
    Set itm = propList(idx)
    itm.Value = Value
End Property

Public Sub Clear()
    Dim ref As Item
    Do Until propList.Count = 0
        Set ref = propList(1)
        propList.Remove 1
        lptrList.Remove 1
        If IsObject(ref.Value) Then
            Set ref.Value = Nothing
        End If
        Set ref.Parent = Nothing
        Set ref = Nothing
    Loop
End Sub

Private Sub Class_Terminate()
    Clear
    Set propList = Nothing
    Set lptrList = Nothing
End Sub
