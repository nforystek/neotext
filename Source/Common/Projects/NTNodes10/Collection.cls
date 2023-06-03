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

Option Compare Binary

'An attempt to be no different the VBA.Collection interface use except fot the following additions:
'Exists property (accepts Key or Index), Key/Index properties (reverse lookup) Clear/Clone methods.

Private lptrList As VBA.Collection
Private propList As VBA.Collection

Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Function NewEnum() As IUnknown ' _
Foreach iteration enumerator.
Attribute NewEnum.VB_Description = "Foreach iteration enumerator."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = lptrList.[_NewEnum]
End Function

Public Sub Add(Anything, Optional Key, Optional Before, Optional After) ' _
Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation.
Attribute Add.VB_Description = "Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation."
    Dim itm As New Item
    If Not IsMissing(Key) Then
        If Exists(Key) Then Err.Raise 8, App.Title, "Key already exists."
        itm.Key = Key
    End If
    If VBA.IsObject(Anything) Then
        Set itm.Value = Anything
    Else
        itm.Value = Anything
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

Public Sub Remove(Optional KeyOrIndex) ' _
Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed.
Attribute Remove.VB_Description = "Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed."
    If Not propList Is Nothing Then
        Dim idx
        If propList.Count = 0 Then
            Err.Raise 8, App.Title, "Collection is empty."
        Else
            If Not IsMissing(KeyOrIndex) Then
                idx = GetKeyIndex(KeyOrIndex)
                If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
            Else
                idx = 1
            End If
            If VBA.TypeName(idx) = "Empty" Then
                Err.Raise 8, App.Title, "Key does not exist."
            End If
        End If
        Dim ref As Item
        Set ref = propList(idx)
        propList.Remove idx
        lptrList.Remove idx
        If VBA.IsObject(ref.Value) Then
            Set ref.Value = Nothing
        End If
        Set ref.Parent = Nothing
        Set ref = Nothing
    End If
End Sub

Public Property Get Exists(ByVal KeyOrIndex) As Boolean ' _
Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false.
Attribute Exists.VB_Description = "Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false."
    If Not propList Is Nothing Then
    Dim idx
    idx = GetKeyIndex(KeyOrIndex)
    Exists = (idx > 0 And idx <= propList.Count)
    End If
End Property

Friend Property Get GetKeyIndex(ByVal Key)
    If IsNumeric(Key) Then
        GetKeyIndex = Key
    Else
        Dim itm As Object
        Dim idx
        For Each itm In propList
            idx = idx + 1
            If itm.Key = Key Then
                GetKeyIndex = idx
                Exit Property
            End If
        Next
    End If
End Property

Public Property Get Count() As Long ' _
Numerical value indicating of how many items the collection contains.
Attribute Count.VB_Description = "Numerical value indicating of how many items the collection contains."
    If Not propList Is Nothing Then
        Count = propList.Count
    End If
End Property

Public Property Get Key(ByVal Index) ' _
Gets the key for the item in the collection specified by the Index parameter.
    Dim idx
    idx = GetKeyIndex(Index)
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    If VBA.IsObject(propList(idx).Key) Then
        Set Key = propList(idx).Key
    Else
        Key = propList(idx).Key
    End If
End Property

Public Property Get Item(ByVal KeyOrIndex) ' _
Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item.
Attribute Item.VB_Description = "Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item."
Attribute Item.VB_UserMemId = 0
    Dim idx
    idx = GetKeyIndex(KeyOrIndex)
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    If VBA.IsObject(propList(idx).Value) Then
        Set Item = propList(idx).Value
    Else
        Item = propList(idx).Value
    End If
End Property

Public Property Set Item(ByVal KeyOrIndex, ByRef Anything)
    Dim idx
    idx = GetKeyIndex(KeyOrIndex)
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    Dim itm As Item
    Set itm = propList(idx)
    Set itm.Value = Anything
End Property

Public Property Let Item(ByVal KeyOrIndex, ByRef Anything)
    Dim idx
    idx = GetKeyIndex(KeyOrIndex)
    If Not (idx > 0 And idx <= propList.Count) Then Err.Raise 8, App.Title, "Index does not exist."
    Dim itm As Item
    Set itm = propList(idx)
    itm.Value = Anything
End Property

Public Sub Clear() ' _
Clears all items with-in the collection, and the Count property becomes zero.
Attribute Clear.VB_Description = "Clears all items with-in the collection, and the Count property becomes zero."
    Dim ref As Item
    If Not propList Is Nothing Then
        Do Until propList.Count = 0
            Set ref = propList(1)
            propList.Remove 1
            lptrList.Remove 1
            If VBA.IsObject(ref.Value) Then
                Set ref.Value = Nothing
            End If
            Set ref.Parent = Nothing
            Set ref = Nothing
        Loop
    End If
End Sub

Friend Sub AddItem(ByVal itm As Item)
    If VBA.TypeName(itm.Key) = "Empty" Then
        propList.Add itm
        lptrList.Add itm.Value
    Else
        propList.Add itm, itm.Key
        lptrList.Add itm.Value, itm.Key
    End If
End Sub

Public Function Clone() As NTNodes10.Collection ' _
Creates a duplicate collection object of the calling object doubling all the primitive data types with-in, for object items new reference pointers are created, unless they have a Clone() function present.
    Dim ncc As New NTNodes10.Collection
    Dim itm As Item
    Dim nxt As Item
    Dim test As Object
    For Each itm In propList
        Set nxt = New Item
        nxt.Key = itm.Key
        If VBA.IsObject(itm.Value) Then
            On Error Resume Next
            Set test = itm.Value.Clone
            If Err Then
                Err.Clear
                vbaObjSetAddref nxt.Value, ObjPtr(itm.Value)
            Else
                Set nxt.Value = test
            End If
            Set test = Nothing
        Else
            nxt.Value = itm.Value
        End If
        ncc.AddItem nxt
    Next
    vbaObjSetAddref Clone, ObjPtr(ncc)
    Set ncc = Nothing
    Set nxt = Nothing
End Function

Private Sub Class_Initialize()
    Set propList = New VBA.Collection
    Set lptrList = New VBA.Collection
End Sub

Private Sub Class_Terminate()
    Set propList = Nothing
    Set lptrList = Nothing
End Sub
