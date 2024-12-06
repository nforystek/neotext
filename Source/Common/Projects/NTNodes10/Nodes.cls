VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True





Option Explicit

Option Compare Binary



Implements INode



Private List As ListType



Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long



Private Sub Class_Initialize()

    Set INode = New INode

End Sub



Private Sub Class_Terminate()

    INode.Clear

    vbaObjSet INode, ObjPtr(Me)

End Sub



Private Sub INode_Append()

    Me.Append

End Sub



Private Sub INode_Backward()

    Me.Backward

End Sub



Private Property Get INode_Check() As Long

    INode_Check = Me.Check

End Property



Private Sub INode_Clear()

    Me.Clear

End Sub



Private Property Get INode_Count() As Long

    INode_Count = Me.Count

End Property



Private Sub INode_Delete()

    Me.Delete

End Sub



Private Function INode_Exists() As Boolean

    INode_Exists = Me.Exists

End Function



Private Property Get INode_Final() As Long

    INode_Final = Me.Final

End Property



Private Property Get INode_First() As Long

    INode_First = Me.First

End Property



Private Sub INode_Forward()

    Me.Forward

End Sub



Private Property Get INode_Handle() As Long

    INode_Handle = Me.Handle

End Property



Private Sub INode_Insert()

    Me.Insert

End Sub



Private Property Get INode_IsObject() As Boolean

    INode_IsObject = Me.IsObject

End Property



Private Sub INode_Load(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False)

    Me.Load FileName, Clear

End Sub



Private Property Set INode_Object(ByVal RHS As Object)

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



Private Sub INode_Save(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False)

    Me.Save FileName, Clear

End Sub



Private Property Get INode_TypeName() As String

    INode_TypeName = Me.TypeName

End Property



Private Property Let INode_Value(ByVal RHS As Variant)

    Me.Value = RHS

End Property



Private Property Get INode_Value() As Variant

    INode_Value = Me.Value

End Property



Public Sub Append() ' _
Adds a node to the end of the list, the point location reflects the new node after Append is called.


    INode.Append

End Sub



Public Sub Backward() ' _
Moves to the prior node in the list.  This function projects along with forward if it is called infrequent to forward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.


    INode.Backward

End Sub



Public Property Get Check() As Long ' _
Refers to the address of the Prior pointer node of the list, this changes with the pointer in movements relative to the First and Final.


    Check = INode.Check

End Property



Public Sub Clear() ' _
Destroies all nodes with in the specified node list, unlike the Clear function in the Node object, this will also set the node to Nothing.


    INode.Clear

End Sub



Public Property Get Count() As Long ' _
Returns the amount of total nodes taht exist with in the node list.


    Count = INode.Count

End Property



Public Sub Delete() ' _
Deletes the node that is located at the current point in the list.


    INode.Delete

End Sub



Public Function Exists() As Boolean ' _
Checks the existance of a node or node list with the specified lsit.


    Exists = INode.Exists

End Function



Public Property Get Final() As Long ' _
The final node is the lists theoretical very last node, this is used to obtain logical reference and may not always reflect last added node to diminish itteration of the entire node list.


    Final = INode.Final

End Property



Public Property Get First() As Long ' _
The first node is the lists theoretical very first node, this is used to obtain logical reference and may not always reflect first added node to diminish itteration of the entire node list.


    First = INode.First

End Property



Public Sub Forward() ' _
Moves to the next node in the list.  This function projects along with backward if it is called infrequent to backward.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.


    INode.Forward

End Sub



Public Property Get Handle() As Long ' _
Returns the handle of the list, that is the memory location and identifyer for the list of nodes.


    Handle = INode.Handle

End Property



Public Sub Insert() ' _
Inserts a node at the current point in the list.  Nodes can be added to the end by using the Append function.


    INode.Insert

End Sub



Public Property Get IsObject() As Boolean ' _
Gets whether or not the current node at Point in the List is set to a object data type.


    IsObject = INode.IsObject()

End Property



Public Sub Load(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Load a list of nodes from a file that was saved by the use of the save function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal before loading.


    INode.Load FileName, Clear

End Sub



Public Property Set Object(ByVal RHS As Object) ' _
Sets or gets the variable or object to this node.


    Set INode.Object = RHS

End Property



Public Property Get Object() As Object ' _
Sets or gets the variable or object to this node.


    Set Object = INode.Object

End Property



Public Property Get Point() As Long ' _
Refers to the address of the current Pointer node of the list, this changes in movements relative to the First and Final.


    Point = INode.Point

End Property



Public Sub Remove() ' _
Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point.


    INode.Remove

End Sub



Public Sub Save(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Saves a list of nodes to a file that can be loaded later by the load function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal after saving, the default


    INode.Save FileName, Clear

End Sub



Public Property Get TypeName() As String ' _
Gets the type name of the data of the current node at Point in the list.


    TypeName = INode.TypeName()

End Property



Public Property Let Value(ByVal RHS As Variant) ' _
Sets or gets the variable or object to this node.


    INode.Value = RHS

End Property



Public Property Get Value() As Variant ' _
Sets or gets the variable or object to this node.


    Value = INode.Value

End Property







