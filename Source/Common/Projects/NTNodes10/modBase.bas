Attribute VB_Name = "modBase"

#Const modBase = -1
Option Explicit

Option Private Module

Public Type ListType
    First As Long 'first node of the list
    Point As Long 'current node in the list
    Final As Long 'last node of the list
    Check As Long 'node after this one
    Total As Long 'total count of nodes
End Type

Public Type NodeType
    Value As Long 'value of the node
    Check As Long 'node after this one
End Type

'memory allocating flags
Private Const GMEM_FIXED = &H0
Private Const GPTR = &H40

'the following functions are to allocate the memory of the nodes, you can change local to global for it
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub rtlMovMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
'Private Declare Sub rtlMovObjRef Lib "kernel32" Alias "RtlMoveMemory" (xDest As Object, ByRef xSource As Long, ByVal nbytes As Long)
Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
'Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Private Declare Sub PutLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, num As Any, Optional ByVal Size As Long = 4&)
Private Declare Sub GetLong Lib "kernel32" Alias "RtlMoveMemory" (num As Any, ByVal ptr As Long, Optional ByVal Size As Long = 4&)
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)


Public Sub Main()
    
End Sub

Public Sub SaveNodes(ByRef List As ListType, ByVal SourceFileName As String, Optional ByVal Clear As Boolean = False)
    Dim hFile As Long 'kill any file that is there already
    If (Dir(SourceFileName) <> "") Then Kill SourceFileName
    hFile = FreeFile 'assign and open file record size long
    Open SourceFileName For Binary As #hFile Len = 4
    Do Until List.Point = List.First 'move to the
        MoveNode List, False 'beginning of list
    Loop
    Dim Totalsize As Long
    Dim Val As Long
    Dim fSize As Long
    Dim Node As NodeType
    fSize = Abs(List.Total) 'save the total size in temp
    Put #hFile, Totalsize + 1, fSize& 'put total
    Totalsize = Totalsize + 4 'at the frst record
    Do Until fSize = 0 'until last node
        Node = GetNode(List.Point) 'get the node
        Val = Node.Value 'save the nodes value
        Put #hFile, Totalsize + 1, Val&
        fSize = fSize - 1 'deincrement count
        Totalsize = Totalsize + 4 'location ptr
        MoveNode List, False 'nove to next node
    Loop 'close the file
    Close #hFile
    If Clear Then
        DisposeOfAll List 'clear all current node
        List.Total = 0 'get the first record, count
    End If
End Sub

Public Sub LoadNodes(ByRef List As ListType, ByVal SourceFileName As String, Optional ByVal Clear As Boolean = False)
    Dim hFile As Long 'freefile and opens
    hFile = FreeFile 'it record size long
    Open SourceFileName For Binary As #hFile Len = 4
    Dim Totalsize As Long
    Dim Val As Long
    Dim fSize As Long
    Dim Node As NodeType
    If Clear Then
        DisposeOfAll List 'clear all current node
        List.Total = 0 'get the first record, count
    End If
    Get #hFile, Totalsize + 1, fSize&
    Totalsize = Totalsize + 4 'keep loc track
    Do Until Abs(List.Total) >= fSize  'until past eof
        AddToLastNode List 'add a new node to list
        Node = GetNode(List.Final) 'get the node
        Get #hFile, Totalsize + 1, Val& 'from file
        Node.Value = Val 'set the nodes value
        SetNode List.Final, Node 'commit the value
        Totalsize = Totalsize + 4 'loation track
    Loop 'close the file
    Close #hFile
End Sub

Public Function GetNode(ByRef Addr As Long) As NodeType
    If Addr = 0 Then Exit Function 'exit if no addr
    Dim Node As NodeType 'get at alloc addr from memory
    rtlMovMem ByVal VarPtr(Node), ByVal Addr, 8&
    GetNode = Node
End Function

Public Sub SetNode(ByRef Addr As Long, ByRef Node As NodeType)
    If Addr = 0 Then Exit Sub 'exit if no addr
    Dim NodePtr As Long 'store in var for api byref
    NodePtr = VarPtr(Node) 'set at alloc addr in memory
    rtlMovMem ByVal Addr, ByVal NodePtr, 8&
End Sub

Private Function PtrVar(ByVal lPtr As Long) As Variant
    Dim lZero As Long
    Dim newObj As Variant
    rtlMovMem newObj, lPtr, 4&
    PtrVar = newObj
    rtlMovMem newObj, lZero, 4&
End Function

Public Function GetVariant(ByRef Addr As Long) As Variant
    If Addr = 0 Then Exit Function 'exit if no addr
    Dim Node As NodeType
    Node = GetNode(Addr)
    Dim Var1 As Long
    Var1 = Node.Value
    If LocalSize(Var1) = 4 Then
        GetVariant = PtrVar(Node.Value)
        rtlMovMem GetVariant, ByVal Node.Value, LenB(GetVariant)
    End If
End Function

Public Sub SetVariant(ByRef Addr As Long, ByRef Var As Variant)
    If Addr = 0 Then Exit Sub 'exit if no addr
    Dim Node As NodeType
    Node = GetNode(Addr)
    If Not LocalSize(Node.Value) = 4 Then
        Node.Value = LocalAlloc(GMEM_FIXED And VarPtr(Var), LenB(Var)) 'allocate 12 bytes and set to mid struct
        If Not Node.Value = LocalLock(Node.Value) Then Err.Raise 8, App.Title, "Memory lock error."
    End If
    rtlMovMem ByVal Node.Value, Var, LenB(Var)
    SetNode Addr, Node
End Sub

Public Property Get NodeObject(ByRef Addr As Long) As Object

    Set NodeObject = Nothing
    Dim optr As Long
    Dim Zero As Long
    GetLong optr, Addr
    Dim newObj As Object
    RtlMoveMemory newObj, optr, 4&
    Set NodeObject = newObj
    RtlMoveMemory newObj, Zero, 4&
End Property

Public Property Set NodeObject(ByRef Addr As Long, RHS As Object)
    Dim Obj As Object
    Set Obj = Nothing
    vbaObjSetAddref ObjPtr(Obj), ObjPtr(RHS)
    PutLong Addr, ObjPtr(RHS)
    Set RHS = Obj
End Property

Public Property Get IsObject(ByRef List As ListType) As Boolean ' _

    Dim Node As NodeType
    Node = GetNode(List.Point)
    IsObject = VBA.IsObject(PtrVar(Node.Check))
    If LocalSize(Node.Value) <> 0 Then
        IsObject = VBA.IsObject(NodeObject(VarPtr(Node.Value)))
    End If
End Property

Public Property Get TypeName(ByRef List As ListType) As String ' _
Gets the type name of the data of the current node at Point in the list.
Attribute TypeName.VB_Description = "Gets the type name of the data of the current node at Point in the list."
  '  TypeName = modBase.TypeName(List)
    Dim Node As NodeType
    Node = GetNode(List.Point)
    TypeName = VBA.TypeName(PtrVar(Node.Check))
    If LocalSize(Node.Value) <> 0 Then
        TypeName = VBA.TypeName(NodeObject(VarPtr(Node.Value)))
    End If
End Property


'Private Function PtrObj(ByVal lPtr As Long) As Object
'    Dim lZero As Long
'    Dim newObj As Object
'    rtlMovObjRef newObj, lPtr, 4&
'    Set PtrObj = newObj
'    rtlMovObjRef newObj, lZero, 4&
'End Function
'
'Private Function IsPtrObj(ByVal lPtr As Long) As Boolean
'    Dim lZero As Long
'    Dim newObj
'    rtlMovMem newObj, lPtr, 4&
'    IsPtrObj = VBA.IsObject(newObj)
'    rtlMovMem newObj, lZero, 4&
'End Function

Public Function IsValidNode(ByRef List As ListType) As Boolean
    IsValidNode = (List.Point <> 0) And (List.First <> 0) And (List.Final <> 0)
    'no current no list when we have all zero values for these
End Function

Private Sub Swap(ByRef n1, ByRef n2, Optional ByRef n3 = Empty, Optional ByRef n4 = Empty)
    Dim n5 'variable to use for swap
    If VBA.TypeName(n3) = "Empty" Then
        n5 = n1 'do a simple
        n1 = n2 'two variable
        n2 = n5 'swap
    ElseIf VBA.TypeName(n4) = "Empty" Then
        n5 = n1 'do a more
        n1 = n2 'complex
        n2 = n3 'three var
        n3 = n5 'swap
    Else
        n5 = n1 'do a more
        n1 = n2 'complex
        n2 = n3 'three var
        n3 = n4 'swap
        n4 = n5
    End If
End Sub

Public Sub AddDelMiddleNode(ByRef List As ListType, ByVal AddOrDel As Boolean)
    If (AddOrDel And ((List.Point = _
        List.Final) Or Not IsValidNode(List))) Then
        'list is at final already just delete
        AddToLastNode List
    ElseIf (Not AddOrDel) And (List.Point = List.First) Then
        'list is at first already just delete
        DelFirstNode List
    ElseIf IsValidNode(List) Then
        Dim lFirst As Long
        Dim lFinal As Long
        If AddOrDel Then
            lFinal = List.Final
            lFirst = List.First
            Swap List.Point, List.First
            Swap List.Check, List.Final
            AddToLastNode List
            List.First = lFirst
        Else
            lFinal = List.Final
            lFirst = List.First
            MoveNode List, False
            DelFirstNode List
        End If
    End If
End Sub

Public Sub AddToLastNode(ByRef List As ListType)
    Dim Node As NodeType
    If Not IsValidNode(List) Then 'no list exists
        List.Point = LocalAlloc(GMEM_FIXED And VarPtr(List), 8)  'allocate 12 bytes and set to mid struct
        If Not (List.Point) = LocalLock(List.Point) Then Err.Raise 8, App.Title, "Memory lock error."
        'initial node values
        List.Check = List.Point
        Node.Check = List.Point
        'initial list values
        List.First = List.Point
        List.Final = List.Point
        'commit the node
        SetNode List.Point, Node
    Else
        Node = GetNode(List.Point)
        List.Check = Node.Check
        Node.Check = LocalAlloc(GMEM_FIXED And List.Point, 8)  'allocate 12 bytes and set to mid struct
        If Not (Node.Check) = LocalLock(Node.Check) Then Err.Raise 8, App.Title, "Memory lock error."
        'commit the final nodes Check
        SetNode List.Point, Node
        'set the current nodes prior
        'and the initial list values
        List.Point = Node.Check
        List.Final = List.Point
        Node.Check = List.Check
        'commit the current node
        SetNode List.Point, Node
    End If
    If List.Total > 0 Then
        List.Total = Abs(List.Total) + 1
    Else
        List.Total = -(Abs(List.Total) + 1)
    End If
End Sub

Public Sub DelFirstNode(ByRef List As ListType)
    Dim Node As NodeType
    Dim Point As Long
    If IsValidNode(List) Then
        'get the first node in list
        Node = GetNode(List.Point)
        Point = Node.Check 'save impass
        If (List.First = List.Final) Then 'eol
            'unlock and release memory
            If LocalSize(Node.Value) = 4 Then
                LocalUnlock Node.Value
                LocalFree Node.Value
            End If
            LocalUnlock List.Point
            LocalFree List.Point
            'set list values none
            List.Point = 0
            List.First = 0
            List.Final = 0
        Else
            If LocalSize(Node.Value) = 4 Then
                LocalUnlock Node.Value
                LocalFree Node.Value
            End If
            If Point <> 0 Then
                'arrange first and final
                Node = GetNode(List.Check)
                Node.Check = Point
                SetNode List.Check, Node
                'set the new Check of
                Node = GetNode(Point)
                Swap List.Final, Node.Check
            End If
            'unlock and release memory
            LocalUnlock List.Point
            LocalFree List.Point
            'set list to retained value
            List.First = Point
            List.Point = Point
        End If
        If List.Total > 0 Then
            List.Total = Abs(List.Total) - 1
        Else
            List.Total = -(Abs(List.Total) - 1)
        End If
    End If
End Sub

Public Sub MoveNode(ByRef List As ListType, ByVal Reverse As Boolean)
    If List.Point = 0 Then Exit Sub
    If ((Reverse And List.Total > 0) Xor (Not Reverse And List.Total < 0)) Then
        Swap List.Check, List.Final
        Swap List.Point, List.Final
        Swap List.First, List.Check
        List.Total = -List.Total
    Else
        Swap List.Check, List.Point
        If Reverse Then 'preform a reverse
            List.Check = GetNode(List.Point).Check
        Else 'else preform a normal forward
            List.Point = GetNode(List.Check).Check
        End If
    End If
End Sub

Public Sub DisposeOfAll(ByRef List As ListType)
    Do While IsValidNode(List) 'until we are done
        DelFirstNode List 'remove the first node
    Loop
End Sub




'Option Explicit
'
'Option Compare Binary
'
'Option Private Module
'
'Public Type ListType
'    First As Long 'first node of the list
'    Point As Long 'current node in the list
'    Final As Long 'last node of the list
'    Check As Long 'node after this one
'    Total As Long 'total count of nodes
'End Type
'
'Public Type NodeType
'    Value As Long 'value of the node
'    Check As Long 'node after this one
'End Type
'
''memory allocating flags
'Private Const GMEM_FIXED = &H0
'Private Const GPTR = &H40
'
''the following functions are to allocate the memory of the nodes, you can change local to global for it
'Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
'Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
'Private Declare Sub RtlMoveMemoryAny Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Any, ByRef xSource As Any, ByVal nbytes As Long)
'Private Declare Sub RtlMoveMemoryObj Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Private Declare Sub RtlMovMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
'Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
'Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
'Private Declare Sub PutLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal Ptr As Long, num As Any, Optional ByVal Size As Long = 4&)
'Private Declare Sub GetLong Lib "kernel32" Alias "RtlMoveMemory" (num As Any, ByVal Ptr As Long, Optional ByVal Size As Long = 4&)
'
'
'Public Sub Main()
'
'End Sub
'
'Public Sub SaveNodes(ByRef List As ListType, ByVal SourceFileName As String, Optional ByVal Clear As Boolean = False)
'    Dim hFile As Long 'kill any file that is there already
'    If (Dir(SourceFileName) <> "") Then Kill SourceFileName
'    hFile = FreeFile 'assign and open file record size long
'    Open SourceFileName For Binary As #hFile Len = 4
'    Do Until List.Point = List.First 'move to the
'        MoveNode List, False 'beginning of list
'    Loop
'    Dim Totalsize As Long
'    Dim Val As Long
'    Dim fsize As Long
'    Dim Node As NodeType
'    fsize = List.Total 'save the total size in temp
'    Put #hFile, Totalsize + 1, fsize& 'put total
'    Totalsize = Totalsize + 4 'at the frst record
'    Do Until fsize = 0 'until last node
'        Node = GetNode(List.Point) 'get the node
'        Val = Node.Value 'save the nodes value
'        Put #hFile, Totalsize + 1, Val&
'        fsize = fsize - 1 'deincrement count
'        Totalsize = Totalsize + 4 'location ptr
'        MoveNode List, False 'nove to next node
'    Loop 'close the file
'    Close #hFile
'    If Clear Then
'        DisposeOfAll List 'clear all current node
'        List.Total = 0 'get the first record, count
'    End If
'End Sub
'
'Public Sub LoadNodes(ByRef List As ListType, ByVal SourceFileName As String, Optional ByVal Clear As Boolean = False)
'    Dim hFile As Long 'freefile and opens
'    hFile = FreeFile 'it record size long
'    Open SourceFileName For Binary As #hFile Len = 4
'    Dim Totalsize As Long
'    Dim Val As Long
'    Dim fsize As Long
'    Dim Node As NodeType
'    If Clear Then
'        DisposeOfAll List 'clear all current node
'        List.Total = 0 'get the first record, count
'    End If
'    Get #hFile, Totalsize + 1, fsize&
'    Totalsize = Totalsize + 4 'keep loc track
'    Do Until List.Total >= fsize 'until past eof
'        AddLastNode List 'add a new node to list
'        Node = GetNode(List.Final) 'get the node
'        Get #hFile, Totalsize + 1, Val& 'from file
'        Node.Value = Val 'set the nodes value
'        SetNode List.Final, Node 'commit the value
'        Totalsize = Totalsize + 4 'loation track
'    Loop 'close the file
'    Close #hFile
'End Sub
'
'Public Function GetNode(ByRef Addr As Long) As NodeType
'    If Addr = 0 Then Exit Function 'exit if no addr
'    Dim Node As NodeType 'get at alloc addr from memory
'    RtlMoveMemory ByVal VarPtr(Node), ByVal (Addr - 2), 8
'    GetNode = Node
'End Function
'
'Public Sub SetNode(ByRef Addr As Long, ByRef Node As NodeType)
'    If Addr = 0 Then Exit Sub 'exit if no addr
'    Dim NodePtr As Long 'store in var for api byref
'    NodePtr = VarPtr(Node) 'set at alloc addr in memory
'    RtlMoveMemory ByVal (Addr - 2), ByVal NodePtr, 8
'End Sub
'
'
'
'Public Function IsValidNode(ByRef List As ListType) As Boolean
'    IsValidNode = (List.Point <> 0) And (List.First <> 0) And (List.Final <> 0)
'    'no current no list when we have all zero values for these
'End Function
'
'Private Sub Swap(ByRef var1 As Variant, ByRef var2 As Variant, Optional ByRef Var3 As Variant, Optional ByRef Var4 As Variant, Optional ByRef Var5 As Variant, Optional ByRef Var6 As Variant)
'    Dim Var0 As Variant
'    If IsMissing(Var3) Then
'        Var0 = var1
'        var1 = var2
'        var2 = Var0
'    ElseIf IsMissing(Var4) Then
'        Var0 = var1
'        var1 = var2
'        var2 = Var3
'        Var3 = Var0
'    ElseIf IsMissing(Var5) Then
'        Var0 = var1
'        var1 = var2
'        var2 = Var3
'        Var3 = Var4
'        Var4 = Var0
'    ElseIf IsMissing(Var6) Then
'        Var0 = var1
'        var1 = var2
'        var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var0
'    Else
'        Var0 = var1
'        var1 = var2
'        var2 = Var3
'        Var3 = Var4
'        Var4 = Var5
'        Var5 = Var6
'        Var6 = Var0
'    End If
'End Sub
'
'Public Sub AddDelMiddleNode(ByRef List As ListType, ByVal AddOrDel As Boolean)
'    If (AddOrDel And ((List.Point = _
'        List.Final) Or Not IsValidNode(List))) Then
'        'list is at final already just delete
'        AddLastNode List
'    ElseIf (Not AddOrDel) And (List.Point = List.First) Then
'        'list is at first already just delete
'        DelFirstNode List
'    ElseIf IsValidNode(List) Then
'        Dim lFirst As Long
'        Dim lFinal As Long
'        If AddOrDel Then
'            'If Not (List.Final = List.Point) Then
'                MoveNode List, (List.Check = List.First)
'            'End If 'call add node
'            AddLastNode List
'        Else
'            'If Not (List.First = List.Point) Then
'                MoveNode List, (List.Check = List.Final)
'            'End If 'call delete node
'            DelFirstNode List
'        End If
'    End If
'End Sub
'
'Public Sub AddLastNode(ByRef List As ListType)
'    Dim Node As NodeType
'    If Not IsValidNode(List) Then 'no list exists
'        List.Point = LocalAlloc(GMEM_FIXED, 8) + 2 'allocate 12 bytes and set to mid struct
'        If Not (List.Point - 2) = LocalLock(List.Point - 2) Then Err.Raise 8, App.Title, "Memory lock error."
'        'initial node values
'        List.Check = List.Point
'        Node.Check = List.Point
'        'initial list values
'        List.First = List.Point
'        List.Final = List.Point
'        'commit the node
'        SetNode List.Point, Node
'    Else
'        Node = GetNode(List.Point) 'get the last node of list
'        List.Check = Node.Check
'        Node.Check = LocalAlloc(GMEM_FIXED, 8) + 2  'allocate 12 bytes and set to mid struct
'        If Not (Node.Check - 2) = LocalLock(Node.Check - 2) Then Err.Raise 8, App.Title, "Memory lock error."
'        'commit the final nodes Check
'        SetNode List.Point, Node
'        'set the current nodes Check
'        'and the initial list values
'        List.Point = Node.Check
'        List.Final = List.Point
'        Node.Check = List.Check
'        'commit the current node
'        SetNode List.Point, Node
'    End If
'    If List.Total > 0 Then
'        List.Total = Abs(List.Total) + 1
'    Else
'        List.Total = -(Abs(List.Total) + 1)
'    End If
'End Sub
'
'Public Sub DelFirstNode(ByRef List As ListType)
'    Dim Node As NodeType
'    Dim Point As Long
'    If IsValidNode(List) Then
'        'get the first node in list
'        Node = GetNode(List.Point)
'        Point = Node.Check 'save impass
'        If (List.First = List.Final) Then 'eol
'            'unlock and release memory
'            If LocalSize(Node.Value) = 4 Then
'                LocalUnlock Node.Value
'                LocalFree Node.Value
'            ElseIf IsObject(List) Then
''                If TypeName(List) = "Nodes" Then
''                    NodeObject(List.Point).Clear
''                End If
'                Set NodeObject(List.Point) = Nothing
'            End If
'            LocalUnlock List.Point - 2
'            LocalFree List.Point - 2
'            'set list values none
'            List.Point = 0
'            List.First = 0
'            List.Final = 0
'        Else
'            If LocalSize(Node.Value) = 4 Then
'                LocalUnlock Node.Value
'                LocalFree Node.Value
'            ElseIf IsObject(List) Then
''                If TypeName(List) = "Nodes" Then
''                    NodeObject(List.Point).Clear
''                End If
'                Set NodeObject(List.Point) = Nothing
'            End If
'            If Point <> 0 Then
'                'arrange first and final
'                Node = GetNode(List.Check)
'                Node.Check = Point
'                SetNode List.Check, Node
'                'set the new Check of
'                Node = GetNode(Point)
'                Swap List.Final, Node.Check
'            End If
'            'unlock and release memory
'            LocalUnlock List.Point - 2
'            LocalFree List.Point - 2
'            'set list to retained value
'            List.First = Point
'            List.Point = Point
'        End If
'        If List.Total > 0 Then
'            List.Total = Abs(List.Total) - 1
'        Else
'            List.Total = -(Abs(List.Total) - 1)
'        End If
'    End If
'End Sub
'
'Public Sub MoveNode(ByRef List As ListType, ByVal Reverse As Boolean)
'    If List.Point = 0 Then Exit Sub
'    Dim Back As NodeType
'    Dim Node As NodeType
'    If Reverse Then 'preform a reverse
'        Node = GetNode(List.Check) 'get the Check as node
'        Back = GetNode(List.Point) 'and the point as back
'        List.Check = ShiftNode(List, List.Point, Node.Check, Reverse)
'    Else 'else preform a normal forward
'        Node = GetNode(List.Point) 'get the point as node
'        Back = GetNode(List.Check) 'and the Check as back
'        List.Point = ShiftNode(List, Back.Check, Node.Check, Reverse)
'    End If
'End Sub
'
'Private Function ShiftNode(ByRef List As ListType, ByRef hNode As Long, ByVal hBack As Long, ByRef Reverse As Boolean) As Long
'    Dim Back As NodeType
'    Dim Node As NodeType
'    ShiftNode = hNode
'    Node = GetNode(hBack)
'    Back = GetNode(hNode)
'    Swap List.Check, List.Point
'    If (Not (((List.Point = List.Final) Xor (List.First = List.Check)) And _
'        ((List.Point = List.First) Xor (List.Final = List.Check)))) Then
'        If Reverse Then
'           Swap Node.Check, Back.Check
'            Swap List.Check, List.Point, Node.Check
'            List.Check = ShiftNode(List, Node.Check, Back.Check, Not Reverse)
'            Swap Node.Check, Back.Check
'        Else
'            Swap Back.Check, Node.Check, List.Point
'            Swap Node.Check, Back.Check
'        End If
'    ElseIf Not Reverse Then
'        Swap List.Point, Node.Check, List.Check
'    End If
'    Swap List.Point, hNode, Back.Check
'    Swap hBack, ShiftNode
'End Function
'
'Public Sub DisposeOfAll(ByRef List As ListType)
'    Do While IsValidNode(List) 'until we are done
'        DelFirstNode List 'remove the first node
'    Loop
'End Sub
'
'Public Property Get NodeObject(ByRef Addr As Long) As Object
'    Set NodeObject = Nothing
'    Dim optr As Long
'    Dim Zero As Long
'    GetLong optr, Addr
'    Dim newObj As Object
'    RtlMoveMemoryObj newObj, optr, 4&
'    Set NodeObject = newObj
'    RtlMoveMemoryObj newObj, Zero, 4&
'End Property
'
'Public Property Set NodeObject(ByRef Addr As Long, RHS As Object)
'    Dim Obj As Object
'    Set Obj = Nothing
'    vbaObjSetAddref ObjPtr(Obj), ObjPtr(RHS)
'    PutLong Addr, ObjPtr(RHS)
'    Set RHS = Obj
'End Property
'
'Public Property Get IsObject(ByRef List As ListType) As Boolean ' _
'
'    Dim Node As NodeType
'    Node = GetNode(List.Point)
'    IsObject = VBA.IsObject(PtrVar(Node.Check))
'    If LocalSize(Node.Value) <> 0 Then
'        IsObject = VBA.IsObject(NodeObject(VarPtr(Node.Value)))
'    End If
'End Property
'
'Public Property Get TypeName(ByRef List As ListType) As String ' _
'Gets the type name of the data of the current node at Point in the list.
'  '  TypeName = modBase.TypeName(List)
'    Dim Node As NodeType
'    Node = GetNode(List.Point)
'    TypeName = VBA.TypeName(PtrVar(Node.Check))
'    If LocalSize(Node.Value) <> 0 Then
'        TypeName = VBA.TypeName(NodeObject(VarPtr(Node.Value)))
'    End If
'End Property
'
'Private Function PtrVar(ByVal lPtr As Long)
'    Dim lZero As Long
'    Dim newObj
'    RtlMovMem newObj, lPtr, 4&
'    PtrVar = newObj
'    RtlMovMem newObj, lZero, 4&
'End Function

