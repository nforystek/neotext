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

Private this As New NTNodes10.Collection
Public Serialize As Boolean
Private pParent As Molecule

Friend Property Get Parent() As Molecule
    Set Parent = pParent
End Property
Friend Property Set Parent(ByRef RHS As Molecule)
    Set pParent = RHS
    Dim m As Molecule
    For Each m In this
        Set m.Parent = pParent
    Next
End Property

Function NewEnum() As IUnknown ' _
Foreach iteration enumerator.
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.[NewEnum]
End Function

Public Sub Add(ByRef Anything As Billboard, Optional Key, Optional Before, Optional After) ' _
Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation.
    this.Add Anything, Key, Before, After
    If Not Anything Is Nothing Then Set Anything.Parent = pParent
End Sub

Public Sub Remove(Optional KeyOrIndex) ' _
Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed.
    this.Remove KeyOrIndex
End Sub

Public Property Get Exists(ByVal KeyOrIndex) As Boolean ' _
Returns true if the given parameter for the property is a valid key or index contained in the collection, otherwise false.
    Exists = this.Exists(KeyOrIndex)
End Property

Public Property Get Count() As Long ' _
Numerical value indicating of how many items the collection contains.
    Count = this.Count
End Property

Public Property Get Key(ByVal Index) ' _
Gets the key for the item in the collection specified by the Index parameter.
    Key = this.Key(Index)
End Property

Public Property Get Item(ByVal KeyOrIndex) As Billboard ' _
Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item.
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = this.Item(KeyOrIndex)
End Property

Public Property Set Item(ByVal KeyOrIndex, ByRef Anything As Billboard)
    Set this.Item(KeyOrIndex) = Anything
End Property

Public Sub Clear() ' _
Clears all items with-in the collection, and the Count property becomes zero.
    this.Clear
End Sub

Private Sub Class_Initialize()
    Set this = New NTNodes10.Collection
    Serialize = True
End Sub

Private Sub Class_Terminate()
    Set pParent = Nothing
    Set this = Nothing
End Sub


