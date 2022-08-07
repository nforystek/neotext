VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'basis object that sets anything in a 3d scope apart from amother has no relevance collective, but has a stars collection
Private this As NTNodes10.Collection
Private pSerialize As Boolean
Private pCurrent As Variant

Public Property Get Current()
    If this.Count > 0 Then
        Set Current = this.Item(pCurrent)
    End If
End Property
Public Property Let Current(ByVal KeyOrIndex)
    If this.Exists(KeyOrIndex) Then
        pCurrent = KeyOrIndex
    End If
End Property
Public Property Get Serialize() As Boolean
    Serialize = pSerialize
End Property
Public Property Let Serialize(ByVal RHS As Boolean)
    pSerialize = RHS
End Property

Function NewEnum() As IUnknown ' _
Foreach iteration enumerator.
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = this.[NewEnum]
End Function

Public Sub Add(ByRef Anything As Spirit, Optional Key, Optional Before, Optional After) ' _
Add a item to the collection, accepts anything.  Optional Key value to identify by, must not be numerical only.  Before or After are to specify a index or key in the collection at which the new item will be in added at in relation.
    this.Add Anything, Key, Before, After
    pCurrent = this.Count
End Sub

Public Sub Remove(Optional KeyOrIndex) ' _
Removes an item from the collection, by the optional parameter as key or index if it exists, if it does not exist, the first item is removed.
    this.Remove KeyOrIndex
    If Not this.Exists(pCurrent) Then pCurrent = this.Count
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

Public Property Get Item(ByVal KeyOrIndex) As Spirit ' _
Gets or sets the item at the key or index supplied as the properties parameter, allowing anything, and disposes of the existing item.
    Set Item = this.Item(KeyOrIndex)
End Property

Public Property Set Item(ByVal KeyOrIndex, ByRef Anything As Spirit)
    Set this.Item(KeyOrIndex) = Anything
End Property

Public Sub Clear() ' _
Clears all items with-in the collection, and the Count property becomes zero.
    this.Clear
End Sub

Public Function Clone() As NTNodes10.Collection ' _
Creates a duplicate collection object of the calling object doubling all the primitive data types with-in, for object items new reference pointers are created, unless they have a Clone() function present.
    Set Clone = this.Clone
End Function


Private Sub Class_Initialize()
    Set this = New NTNodes10.Collection
End Sub

Private Sub Class_Terminate()
    Set this = Nothing
End Sub


