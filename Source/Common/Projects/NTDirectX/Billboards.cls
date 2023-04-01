VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Billboards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements NTNodes10.Collection

Private this As NTNodes10.Collection
Private pParent As Molecule
Private pSerialize As Boolean
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property
Public Property Get Serialize() As Boolean
    Serialize = pSerialize
End Property

'Friend Property Get Parent() As Molecule
'    Set Parent = pParent
'End Property
'Friend Property Set Parent(ByRef RHS As Molecule)
'    Set pParent = RHS
'    Dim m As Molecule
'    For Each m In this
'        Set m.Parent = pParent
'    Next
'End Property

Public Function NewEnum() As IUnknown ' _
Foreach iteration enumerator.
Attribute NewEnum.VB_Description = "Foreach iteration enumerator."
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.[NewEnum]
End Function

Public Sub Add(ByRef Anything As Billboard, Optional Key, Optional Before, Optional After) ' _
Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation.
Attribute Add.VB_Description = "Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation."
    this.Add Anything, Key, Before, After
'    If Not Anything Is Nothing Then Set Anything.Parent = pParent
End Sub

Public Sub Remove(Optional KeyOrIndex) ' _
Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed.
Attribute Remove.VB_Description = "Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed."
    this.Remove KeyOrIndex
End Sub

Public Property Get Exists(ByVal KeyOrIndex) As Boolean ' _
Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false.
Attribute Exists.VB_Description = "Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false."
    Exists = this.Exists(KeyOrIndex)
End Property

Public Property Get Count() As Long ' _
Numerical value indicating of how many items the collection contains.
Attribute Count.VB_Description = "Numerical value indicating of how many items the collection contains."
    Count = this.Count
End Property

Public Property Get Key(ByVal Index) ' _
Gets the key for the item in the collection specified by the Index parameter.
Attribute Key.VB_Description = "Gets the key for the item in the collection specified by the Index parameter."
    Key = this.Key(Index)
End Property

Public Property Get Item(ByVal KeyOrIndex) As Billboard ' _
Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item.
Attribute Item.VB_Description = "Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item."
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "600"
    Set Item = this.Item(KeyOrIndex)
End Property

Public Property Set Item(ByVal KeyOrIndex, ByRef Anything As Billboard)
    Set this.Item(KeyOrIndex) = Anything
End Property

Public Sub Clear() ' _
Clears all items with-in the collection, and the Count property becomes zero.
Attribute Clear.VB_Description = "Clears all items with-in the collection, and the Count property becomes zero."
    this.Clear
End Sub

Private Sub Class_Initialize()
    vbaObjSetAddref this, ObjPtr(New NTNodes10.Collection)
End Sub

Private Sub Class_Terminate()
    Set pParent = Nothing
    Set this = Nothing
End Sub

Private Sub Collection_Add(Anything As Variant, Optional Key As Variant, Optional Before As Variant, Optional After As Variant)
    this.Add Anything, Key, Before, After
End Sub

Private Sub Collection_Clear()
    this.Clear
End Sub

Private Function Collection_Clone() As NTNodes10.Collection
    Set Collection_Clone = this.Clone
End Function

Private Property Get Collection_Count() As Long
    Collection_Count = this.Count
End Property

Private Property Get Collection_Exists(ByVal KeyOrIndex As Variant) As Boolean
    Collection_Exists = this.Exists(KeyOrIndex)
End Property

Private Property Set Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
    Set this.Item(KeyOrIndex) = RHS
End Property

Private Property Let Collection_Item(ByVal KeyOrIndex As Variant, RHS As Variant)
    this.Item(KeyOrIndex) = RHS
End Property

Private Property Get Collection_Item(ByVal KeyOrIndex As Variant) As Variant
    Set Collection_Item = this.Item(KeyOrIndex)
End Property

Private Property Get Collection_Key(ByVal Index As Variant) As Variant
    Collection_Key = this.Key(Index)
End Property

Private Function Collection_NewEnum() As stdole.IUnknown
    Set Collection_NewEnum = this.[NewEnum]
End Function

Private Sub Collection_Remove(Optional KeyOrIndex As Variant)
    this.Remove KeyOrIndex
End Sub


