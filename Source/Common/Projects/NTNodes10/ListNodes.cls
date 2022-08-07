VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Binary

Implements INode

Private List As ListType

Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" _
    (dstObject As Any, ByVal srcObjPtr As Long) As Long

Private Sub Class_Initialize()
    Set INode = New INode
End Sub

Private Sub Class_Terminate()
    Clear
    vbaObjSet INode, ObjPtr(Me)
End Sub

Private Sub INode_Append()
    Me.Append
End Sub

Private Sub INode_Backward()
    Me.Backward
End Sub

Private Sub INode_Clear()
    Me.Clear
End Sub

Private Property Get INode_Count() As Long
    INode_Count = Me.Count
End Property

Private Sub INode_Delete()
    Me.Delete
End Sub

Private Function INode_Valid() As Boolean
    INode_Valid = Me.Valid
End Function

Private Sub INode_Forward()
    Me.Forward
End Sub

Private Sub INode_Insert()
    Me.Insert
End Sub

Private Property Get INode_IsObject() As Boolean
    INode_IsObject = Me.IsObject
End Property

Private Property Set INode_Object(RHS As Object)
    Set Me.Object = RHS
End Property

Private Property Get INode_Object() As Object
    Set INode_Object = Me.Object
End Property

Private Property Get INode_Point() As Long
    INode_Point = Me.Point
End Property

Private Sub INode_Remove()
    Me.Remove
End Sub

Private Property Get INode_TypeName() As String
    INode_TypeName = Me.TypeName
End Property

Private Property Let INode_Value(ByVal RHS As Long)
    Me.Value = RHS
End Property

Private Property Get INode_Value() As Long
    INode_Value = Me.Value
End Property

Public Function CreateList() As ListNodes ' _
Creates a new node list, and returns the Node object identifying that list.  Node objects are not creatable by themselves, you must use this function to make create them.
Attribute CreateList.VB_Description = "Creates a new node list, and returns the Node object identifying that list.  Node objects are not creatable by themselves, you must use this function to make create them."
    Set CreateList = New ListNodes
End Function

Public Property Get Count() As Long ' _
Returns the amount of total nodes taht exist with in the node list.
Attribute Count.VB_Description = "Returns the amount of total nodes taht exist with in the node list."
    Count = Abs(List.Total)
End Property

Public Property Get Check() As Long ' _
Refers to the address of the Prior pointer node of the list, this changes with the pointer in movements relative to the First and Final.
Attribute Check.VB_Description = "Refers to the address of the Prior pointer node of the list, this changes with the pointer in movements relative to the First and Final."
    Check = Abs(List.Check)
End Property

Public Property Get First() As Long ' _
The first node is the lists theoretical very first node, this is used to obtain logical reference and may not always reflect first added node to diminish itteration of the entire node list.
Attribute First.VB_Description = "The first node is the lists theoretical very first node, this is used to obtain logical reference and may not always reflect first added node to diminish itteration of the entire node list."
    First = Abs(List.First)
End Property

Public Property Get Final() As Long ' _
The final node is the lists theoretical very last node, this is used to obtain logical reference and may not always reflect last added node to diminish itteration of the entire node list.
Attribute Final.VB_Description = "The final node is the lists theoretical very last node, this is used to obtain logical reference and may not always reflect last added node to diminish itteration of the entire node list."
    Final = Abs(List.Final)
End Property

Public Property Get Point() As Long ' _
Refers to the address of the current Pointer node of the list, this changes in movements relative to the First and Final.
Attribute Point.VB_Description = "Refers to the address of the current Pointer node of the list, this changes in movements relative to the First and Final."
    Point = Abs(List.Point)
End Property

Public Function Valid() As Boolean ' _
Checks the existance of a node or node list with the specified lsit.
Attribute Valid.VB_Description = "Checks the existance of a node or node list with the specified lsit."
    Valid = IsValidNode(List.Point)
End Function

Public Sub Clear() ' _
Destroies all nodes with in the specified node list, unlike the Clear function in the Node object, this will also set the node to Nothing.
Attribute Clear.VB_Description = "Destroies all nodes with in the specified node list, unlike the Clear function in the Node object, this will also set the node to Nothing."
    DisposeOfAll List
End Sub

Public Sub Remove() ' _
Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point.
Attribute Remove.VB_Description = "Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point."
    DelFirstNode List
End Sub

Public Sub Append() ' _
Adds a node to the end of the list, the point location reflects the new node after Append is called.
Attribute Append.VB_Description = "Adds a node to the end of the list, the point location reflects the new node after Append is called."
    AddLastNode List
End Sub

Public Sub Delete() ' _
Deletes the node that is located at the current point in the list.
Attribute Delete.VB_Description = "Deletes the node that is located at the current point in the list."
    AddDelMiddleNode List, False
End Sub

Public Sub Insert() ' _
Inserts a node at the current point in the list.  Nodes can be added to the end by using the Append function.
Attribute Insert.VB_Description = "Inserts a node at the current point in the list.  Nodes can be added to the end by using the Append function."
    AddDelMiddleNode List, True
End Sub

Public Sub Forward() ' _
Moves to the next node in the list.  This function projects along with backward if it is called infrequent to backward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.
Attribute Forward.VB_Description = "Moves to the next node in the list.  This function projects along with backward if it is called infrequent to backward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists."
    MoveNode List, False
End Sub

Public Sub Backward() ' _
Moves to the prior node in the list.  This function projects along with forward if it is called infrequent to forward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.
Attribute Backward.VB_Description = "Moves to the prior node in the list.  This function projects along with forward if it is called infrequent to forward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists."
    MoveNode List, True
End Sub

Public Property Get IsObject() As Boolean ' _
Gets whether or not the current node at Point in the List is set to a object data type.
Attribute IsObject.VB_Description = "Gets whether or not the current node at Point in the List is set to a object data type."
    IsObject = INode.IsObject()
End Property

Public Property Get TypeName() As String ' _
Gets the type name of the data of the current node at Point in the list.
Attribute TypeName.VB_Description = "Gets the type name of the data of the current node at Point in the list."
    TypeName = INode.TypeName()
End Property

Public Property Get Value() As Long ' _
Sets or gets the primitive data type to this node as a Long numerical.  Can not be used with the Object property.
Attribute Value.VB_Description = "Sets or gets the primitive data type to this node as a Long numerical.  Can not be used with the Object property."
    Value = INode.Value
End Property
Public Property Let Value(ByVal RHS As Long) ' _
Sets or gets the primitive data type to this node as a Long numerical.  Can not be used with the Object property.
Attribute Value.VB_Description = "Sets or gets the primitive data type to this node as a Long numerical.  Can not be used with the Object property."
    INode.Value = RHS
End Property
Public Property Get Object() As Object ' _
Gets or sets the currnet node at Point as an object.  Can not be used with Value property.
Attribute Object.VB_Description = "Gets or sets the currnet node at Point as an object.  Can not be used with Value property."
    Set Object = INode.Object
End Property
Public Property Set Object(ByRef RHS As Object) ' _
Gets or sets the currnet node at Point as an object.  Can not be used with Value property.
Attribute Object.VB_Description = "Gets or sets the currnet node at Point as an object.  Can not be used with Value property."
    Set INode.Object = RHS
End Property


