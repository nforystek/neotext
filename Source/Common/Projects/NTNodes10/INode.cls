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

Private LastMove As Integer

Public Function Append() As Long ' _
Adds a node to the end of the list, the point location reflects the new node after Append is called.
Attribute Append.VB_Description = "Adds a node to the end of the list, the point location reflects the new node after Append is called."

   Append = AddToLastNode(List)
    
End Function


Public Sub Backward() ' _
Moves to the prior node in the list.
Attribute Backward.VB_Description = "Moves to the prior node in the list.  This function projects along with forth if it is called infrequent to forth. At first calls it is projecting reversal not resulting in movement to diminish iteration of full lists."

    MoveNode List, True
    MoveNode List, True
    
    If List.Total < 0 Then
        MoveNode List, True
        MoveNode List, True
    Else
        MoveNode List, False
    End If

#If TestNodes = -1 Then
    Debug.Print "Point: " & List.Point
#End If
End Sub


Public Sub Forward() ' _
Moves to the next node in the list.

    MoveNode List, (List.Total < 0)

#If TestNodes = -1 Then
    Debug.Print "Point: " & List.Point
#End If

End Sub



Public Function BeginOfList() As Boolean ' _
Checks whether the point is at the beginning of the list or not.
Attribute BeginOfList.VB_Description = "Checks whether the point is at the beginning of the list or not."
    BeginOfList = modBase.BOL(List)
End Function


Public Sub Clear() ' _
Clears all nodes in the list, this does not set the object itself to Nothing, unlike the DisposeList function of the Nodes object.
Attribute Clear.VB_Description = "Clears all nodes in the list, this does not set the object itself to Nothing, unlike the DisposeList function of the Nodes object."

    DisposeOfAll List

End Sub


Public Property Get Count() As Long ' _
Returns the amount of total nodes taht exist with in the node list.
Attribute Count.VB_Description = "Returns the amount of total nodes taht exist with in the node list."

    Count = Abs(List.Total)
    
End Property

#If VBIDE = -1 Then
'Public Property Get Track() As Long
'    Track = List.Track
'End Property
'Public Property Get Total() As Long
'    Total = List.Total
'End Property
Public Property Get Count2() As Long ' _
Returns the amount of total nodes taht exist with in the node list.
Attribute Count2.VB_Description = "Returns the amount of total nodes taht exist with in the node list."

    Count2 = sec.Count
    
End Property
#End If


Public Function Delete() As Long ' _
Deletes a node at the current position of the lists Point node.
Attribute Delete.VB_Description = "Deletes a node at the current position of the lists Point node."

    Delete = AddDelMiddleNode(List, False)

End Function

Public Function EndOfList() As Boolean ' _
Checks whether the point is at the end of the list or not.
Attribute EndOfList.VB_Description = "Checks whether the point is at the end of the list or not."
    EndOfList = modBase.EOL(List)
End Function

Public Function Exists() As Boolean ' _
Checks the existance of nodes in the list, the count property may be checked against zero for a similar indicator.
Attribute Exists.VB_Description = "Checks the existance of nodes in the list, the count property may be checked against zero for a similar indicator."

    Exists = IsValidList(List)

End Function


Public Property Get Final() As Long ' _
Refers to the address of the final node is the list.
Attribute Final.VB_Description = "Refers to the address of the final node is the list."

    If List.Total < 0 Or List.Point = List.Final Then
        Final = Abs(List.First)
    Else
        Final = Abs(List.Final)
    End If


End Property


Public Property Get First() As Long ' _
Refers to the address of the first node is the list.
Attribute First.VB_Description = "Refers to the address of the first node is the list."

    If List.Total < 0 Or List.Point = List.First Then
        First = Abs(List.Final)
    Else
        First = Abs(List.First)
    End If

    
End Property


Public Property Get Handle() As Long ' _
Returns the handle of the list, that is the memory location and identifyer for the list of nodes.
Attribute Handle.VB_Description = "Returns the handle of the list, that is the memory location and identifyer for the list of nodes."

    Handle = VarPtr(List)

End Property


Public Function Insert() As Long ' _
Inserts a node into the list at the current Point in the list.
Attribute Insert.VB_Description = "Inserts a node into the list at the current Point in the list."

    Insert = AddDelMiddleNode(List, True)

End Function

Public Function Inverse() As Boolean
    Inverse = (List.Total < 0)
End Function

Public Property Get IsObject() As Boolean ' _
Gets whether or not the current node at Point in the List is set to a object data type.
Attribute IsObject.VB_Description = "Gets whether or not the current node at Point in the List is set to a object data type."

    IsObject = modBase.IsObject(List)

End Property


Public Sub Load(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Load a list of nodes from a file that was saved by the use of the save function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal before loading.
Attribute Load.VB_Description = "Load a list of nodes from a file that was saved by the use of the save function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal before loading."

    If FileName = "" Then FileName = App.Path & "\Nodes.bin"

    LoadNodes List, FileName, Clear

End Sub


Public Property Set Object(ByVal RHS As Object) ' _
Gets or sets the currnet node at Point as an object.
Attribute Object.VB_Description = "Gets or sets the currnet node at Point as an object."

    Set NodeObject(VarPtr(Register(List.Point, 0))) = RHS
    
End Property


Public Property Get Object() As Object

    Set Object = NodeObject(VarPtr(Register(List.Point, 0)))

End Property


Public Property Get Point() As Long ' _
Refers to the address of the node at the current point in the list.
Attribute Point.VB_Description = "Refers to the address of the node at the current point in the list."

    Point = Abs(List.Point)

End Property

Public Property Get prior() As Long ' _
Refers to the address of the Prior pointer node of the list.
Attribute prior.VB_Description = "Refers to the address of the Prior pointer node of the list."

    prior = Abs(List.prior)
   
    
End Property

Public Function Remove() As Long ' _
Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point.
Attribute Remove.VB_Description = "Removes a node from the list at the very beginning of the list.  You can use the delete function to remove at the current location point."

'Swap List.First, List.Final


    Remove = DelFirstNode(List)

 '   MoveNode List, False
'Swap List.Point, List.Prior
End Function


Public Sub Save(Optional ByVal FileName As String, Optional ByVal Clear As Boolean = False) ' _
Saves a list of nodes to a file that can be loaded later by the load function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal after saving, the default
Attribute Save.VB_Description = "Saves a list of nodes to a file that can be loaded later by the load function.  Optionally specify a file name, or by default Nodes.bin is used in the application path.  Clear flag may also be used to force the nodes to disposal after saving, the default"

    If FileName = "" Then FileName = App.Path & "\Nodes.bin"

    SaveNodes List, FileName, Clear

End Sub


Public Property Get TypeName() As String ' _
Gets the type name of the data of the current node at Point in the list.
Attribute TypeName.VB_Description = "Gets the type name of the data of the current node at Point in the list."

    TypeName = modBase.TypeName(List)

End Property



Public Property Let Value(ByVal RHS As Variant) ' _
Sets or gets the variable or object to this node.
Attribute Value.VB_Description = "Sets or gets the variable or object to this node."

    Register(List.Point, 0) = RHS

End Property


Public Property Get Value() As Variant

    Value = Register(List.Point, 0)

End Property



