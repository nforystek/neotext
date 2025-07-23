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

Implements NTNodes10.Collection

Private pSerialize As Boolean
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property
Public Property Get Serialize() As Boolean
Attribute Serialize.VB_MemberFlags = "40"
    Serialize = True
End Property

Public Function NewEnum() As IUnknown ' _
Foreach iteration enumerator.
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = Collection.[NewEnum]
End Function

Public Sub Add(Anything As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)  ' _
Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation.
    Collection.Add Anything, Key, Before, After
End Sub

Public Sub Remove(Optional KeyOrIndex As Variant) ' _
Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed.
    Collection.Remove KeyOrIndex
End Sub

Public Property Get Exists(ByVal KeyOrIndex As Variant) As Boolean ' _
Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false.
    Exists = Collection.Exists(KeyOrIndex)
End Property


Public Property Get count() As Long ' _
Numerical value indicating of how many items the collection contains.
    count = Collection.count
End Property

Public Property Get Key(ByVal Index As Variant) ' _
Gets the key for the item in the collection specified by the Index parameter.
    Key = Collection.Key(Index)
End Property

Public Property Get Item(ByVal KeyOrIndex As Variant) ' _
Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item.
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "240"
    On Error Resume Next
    
 '   temp = Collection.Item(KeyOrIndex)
    If VBA.TypeName(Collection.Item(KeyOrIndex)) = "String" Then
        Item = Collection.Item(KeyOrIndex)
    Else
        Set Item = Collection.Item(KeyOrIndex)
    End If
    
 '   Set Item = Collection.Item(KeyOrIndex)
    If Err.number <> 0 Then
        Dim number As Long
        Dim description As String
        Dim source As String
        source = Err.source
        description = Err.description
        number = Err.number
        On Error GoTo 0
        On Error GoTo -1
        Err.Raise number, source, description
    Else
        On Error GoTo 0
    End If
End Property

Public Property Set Item(ByVal KeyOrIndex As Variant, Anything)
    Set Collection.Item(KeyOrIndex) = Anything
End Property

Public Property Let Item(ByVal KeyOrIndex As Variant, Anything)
    Collection.Item(KeyOrIndex) = Anything
End Property

Public Sub Clear() ' _
Clears all items with-in the collection, and the Count property becomes zero.
    'Dim tmpObj As Object
    Do Until count = 0
        'tmpObj = Collection.Item(1)
        Collection.Remove 1
        'Set tmpObj = Nothing
    Loop
End Sub

Private Sub Class_Initialize()
    Stats_Collection_Count = Stats_Collection_Count + 1
    vbaObjSetAddref Collection, ObjPtr(New NTNodes10.Collection)
End Sub

Private Sub Class_Terminate()
    Clear
    vbaObjSet Collection, ObjPtr(Nothing)
    Stats_Collection_Count = Stats_Collection_Count - 1
End Sub

Private Sub Collection_Add(Anything As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
    Collection.Add Anything, Key, Before, After
End Sub

Private Sub Collection_Clear()
    Collection.Clear
End Sub

Private Function Collection_Clone() As NTNodes10.Collection
    Set Collection_Clone = Collection.Clone
End Function

Private Property Get Collection_Count() As Long
    Collection_Count = Collection.count
End Property

Private Property Get Collection_Exists(ByVal KeyOrIndex As Variant) As Boolean
    Collection_Exists = Collection.Exists(KeyOrIndex)
End Property

Private Property Set Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
    Set Collection.Item(KeyOrIndex) = RHS
End Property

Private Property Let Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
    Collection.Item(KeyOrIndex) = RHS
End Property

Private Property Get Collection_Item(ByVal KeyOrIndex As Variant) As Variant
    Set Collection_Item = Collection.Item(KeyOrIndex)
End Property

Private Property Get Collection_Key(ByVal Index As Variant) As Variant
    Collection_Key = Collection.Key(Index)
End Property

Private Function Collection_NewEnum() As stdole.IUnknown
    Set Collection_NewEnum = Collection.[NewEnum]
End Function

Private Sub Collection_Remove(Optional KeyOrIndex As Variant)
    Collection.Remove KeyOrIndex
End Sub






