VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True





Option Explicit



Private List As ListType



Public Sub Append() ' _
Adds a node to the end of the list, the point location reflects the new node after Append is called.


    AddToLastNode List

End Sub



Public Sub Backward() ' _
Moves to the prior node in the list.  This function projects along with forth if it is called infrequent to forth. At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.


    MoveNode List, True

End Sub



Public Property Get Check() As Long ' _
Refers to the address of the Prior pointer node of the list, this changes with the pointer in movements relative to the First and Final.


    Check = Abs(List.Check)

End Property



Public Sub Clear() ' _
Clears all nodes in the list, this does not set the object itself to Nothing, unlike the DisposeList function of the Nodes object.


    DisposeOfAll List

End Sub



Public Property Get Count() As Long ' _
Returns the amount of total nodes taht exist with in the node list.


    Count = Abs(List.Total)

End Property



Public Sub Delete() ' _
Deletes a node at the current position of the lists Point node.


    AddDelMiddleNode List, False

End Sub



Public Function Exists() As Boolean ' _
Checks the existance of nodes in the list, the count property may be checked against zero for a similar indicator.


    Exists = IsValidNode(List)

End Function



Public Property Get Final() As Long ' _
The final node is the lists theoretical very last node, this is used to obtain logical reference and may not always reflect last added node to diminish itteration of the entire node list.


    Final = Abs(List.Final)

End Property



Public Property Get First() As Long ' _
The first node is the lists theoretical very first node, this is used to obtain logical reference and may not always reflect first added node to diminish itteration of the entire node list.


    First = Abs(List.First)

End Property



Public Sub Forward() ' _
Moves to the next node in the list.  This function projects along with prior if it is called infrequent to prior.  At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists.


    MoveNode List, False

End Sub



Public Property Get handle() As Long ' _
Returns the handle of the list, that is the memory location and identifyer for the list of nodes.


    handle = VarPtr(List)

End Property



Public Sub Insert() ' _
Inserts a node into the list at the current Point in the list.


    AddDelMiddleNode List, True

End Sub



Public Property Get IsObject() As Boolean ' _
Gets whether or not the current node at Point in the List is set to a object data type.


    IsObject = modBase.IsObject(List)

End Property



Public Sub Load(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Load a list of nodes from a file that was saved by the use of the save function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal before loading.


    If FileName = "" Then FileName = App.path & "\Nodes.bin"

    LoadNodes List, FileName, Clear

End Sub



Public Property Set Object(ByVal RHS As Object) ' _
Gets or sets the currnet node at Point as an object.


    Dim Node As NodeType

    Node = GetNode(List.Point)

    Set NodeObject(VarPtr(Node.Value)) = RHS

    SetNode List.Point, Node

End Property



Public Property Get Object() As Object ' _
Gets or sets the currnet node at Point as an object.


    Dim Node As NodeType

    Node = GetNode(List.Point)

    Set Object = NodeObject(VarPtr(Node.Value))

End Property



Public Property Get Point() As Long ' _
Refers to the address of the current Pointer node of the list, this changes in movements relative to the First and Final.


    Point = Abs(List.Point)

End Property



Public Sub Remove() ' _
Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point.


    DelFirstNode List

End Sub



Public Sub Save(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Saves a list of nodes to a file that can be loaded later by the load function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal after saving, the default


    If FileName = "" Then FileName = App.path & "\Nodes.bin"

    SaveNodes List, FileName, Clear

End Sub



Public Property Get TypeName() As String ' _
Gets the type name of the data of the current node at Point in the list.


    TypeName = modBase.TypeName(List)

End Property



Public Property Let Value(ByVal RHS As Variant) ' _
Sets or gets the variable or object to this node.


    Dim Node As NodeType

    Node = GetNode(List.Point)

    Node.Value = RHS

    SetNode List.Point, Node

End Property



Public Property Get Value() As Variant ' _
Sets or gets the variable or object to this node.


    Dim Node As NodeType

    Node = GetNode(List.Point)

    Value = Node.Value

End Property





