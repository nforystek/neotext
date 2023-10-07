Attribute VB_Name = "modBase"

#Const modBase = -1
Option Explicit

Option Private Module

Public Type ListType
    First As Long 'first node of the list
    Point As Long 'current node in the list
    Final As Long 'last node of the list
    Prior As Long 'node after this one
    

    Total As Long 'total count of nodes
    Track As Long 'an offset with total
End Type

Public Type NodeType
    Value As Long
    Prior As Long
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

#If VBIDE = -1 Then
Public sec As Object
Public DebugLate As String
Public Sub Setup()

    Set sec = CreateObject("AddrPool.Addresses")
    CleanUp
End Sub
Public Sub CleanUp()
    Dim Addr As Variant
    For Each Addr In sec
       ' Debug.Print "Freeing: " & CLng(Addr)
        sec.Free CLng(Addr)
    Next
End Sub

#End If

Public Sub Main()
#If VBIDE = -1 Then
    Setup
#If TestNodes = -1 Then
    Form1.Show
#End If
#End If
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
    fSize = Abs(List.Total) 'save the total size in temp
    Put #hFile, Totalsize + 1, fSize& 'put total
    Totalsize = Totalsize + 4 'at the frst record
    Do Until fSize = 0 'until last node
        Node = GetNode(List.Point) 'get the node
        Val = Register(List.Point, Val) 'save the nodes value
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
    If Clear Then
        DisposeOfAll List 'clear all current node
        List.Total = 0 'get the first record, count
    End If
    Get #hFile, Totalsize + 1, fSize&
    Totalsize = Totalsize + 4 'keep loc track
    Do Until Abs(List.Total) >= fSize  'until past eof
        AddToLastNode List 'add a new node to list
        Get #hFile, Totalsize + 1, Val& 'from file
        Register(List.Final, 0) = Val 'set the nodes value
        Totalsize = Totalsize + 4 'loation track
    Loop 'close the file
    Close #hFile
End Sub


Public Function GetNode(ByRef Addr As Long) As NodeType
    If Addr = 0 Then Exit Function 'exit if no addr
#If VBIDE = -1 Then
    GetNode.Value = Register(Addr, 0)
    GetNode.Prior = Register(Addr, 4)
#Else
    Dim Node As NodeType 'get at alloc addr from memory
    rtlMovMem ByVal VarPtr(Node), ByVal Addr, 8&
    GetNode = Node
#End If
End Function

Public Sub SetNode(ByRef Addr As Long, ByRef Node As NodeType)
    If Addr = 0 Then Exit Sub 'exit if no addr
#If VBIDE = -1 Then
    Register(Addr, 0) = Node.Value
    Register(Addr, 4) = Node.Prior
#Else
    Dim NodePtr As Long 'store in var for api byref
    NodePtr = VarPtr(Node) 'set at alloc addr in memory
    rtlMovMem ByVal Addr, ByVal NodePtr, 8&
#End If
End Sub

Public Property Get Register(ByVal Addr As Long, Optional ByVal offset As Long = 0) As Long
    If Addr <> 0 Then
#If VBIDE = -1 Then
        Register = sec.Register(Abs(Addr), offset)
#Else
        GetLong Register, Abs(Addr) + offset
#End If
    End If
End Property
Public Property Let Register(ByVal Addr As Long, Optional ByVal offset As Long = 0, ByVal RHS As Long)
    If Addr <> 0 Then
#If VBIDE = -1 Then
        sec.Register(Abs(Addr), offset) = RHS
#Else
        PutLong Abs(Addr) + offset, RHS
#End If
    End If
End Property

Public Function GetVariant(ByRef Addr As Long) As Variant
    If Addr = 0 Then Exit Function 'exit if no addr

    Dim Var1 As Long
    Var1 = Register(Addr, 0)
#If VBIDE = -1 Then
    If sec.Size(Var1) = 4 Then
        GetVariant = PtrVar(Var1)
        sec.rtlMovMem GetVariant, False, Var1, True, LenB(GetVariant)
    End If
#Else
    If LocalSize(Var1) = 4 Then
        GetVariant = PtrVar(Var1)
        rtlMovMem GetVariant, ByVal Var1, LenB(GetVariant)
    End If
#End If

End Function

Public Sub SetVariant(ByRef Addr As Long, ByRef Var As Variant)
    If Addr = 0 Then Exit Sub 'exit if no addr

    Dim Var1 As Long
    Var1 = Register(Addr, 0)
#If VBIDE = -1 Then
    If Not sec.Size(Var1) = 4 Then
        Var1 = sec.Alloc(GMEM_FIXED And VarPtr(Var), LenB(Var)) 'allocate 12 bytes and set to mid struct
        If Not Var1 = sec.Lock(Var1) Then Err.Raise 8, App.Title, "Memory lock error."
    End If
    sec.rtlMovMem Var1, True, Var, False, LenB(Var)
#Else
    If Not LocalSize(Var1) = 4 Then
        Var1 = LocalAlloc(GMEM_FIXED And VarPtr(Var), LenB(Var)) 'allocate 12 bytes and set to mid struct
        If Not Var1 = LocalLock(Var1) Then Err.Raise 8, App.Title, "Memory lock error."
    End If
    rtlMovMem ByVal Var1, Var, LenB(Var)
#End If
    Register(Addr, 0) = Var1
End Sub

Public Property Get NodeObject(ByRef Addr As Long) As Object
    Set NodeObject = Nothing
    Dim optr As Long
    Dim Zero As Long
#If VBIDE = -1 Then
    sec.GetLong optr, Addr
    Dim newObj As Object
    sec.RtlMoveMemory newObj, optr, 4&
    Set NodeObject = newObj
    sec.RtlMoveMemory newObj, Zero, 4&
#Else
    GetLong optr, Addr
    Dim newObj As Object
    RtlMoveMemory newObj, optr, 4&
    Set NodeObject = newObj
    RtlMoveMemory newObj, Zero, 4&
#End If

End Property

Public Property Set NodeObject(ByRef Addr As Long, RHS As Object)
    Dim Obj As Object
    Set Obj = Nothing
#If VBIDE = -1 Then
    sec.vbaObjSetAddref ObjPtr(Obj), ObjPtr(RHS)
    sec.PutLong Addr, ObjPtr(RHS)
#Else
    vbaObjSetAddref ObjPtr(Obj), ObjPtr(RHS)
    PutLong Addr, ObjPtr(RHS)
#End If

    Set RHS = Obj
End Property


Public Property Get TypeName(ByRef List As ListType) As String ' _
Gets the type name of the data of the current node at Point in the list.
  '  TypeName = modBase.TypeName(List)
    Dim Var1 As Long

    TypeName = VBA.TypeName(PtrVar(Register(List.Point, 4)))
    Var1 = Register(List.Point, 0)
#If VBIDE = -1 Then
    If sec.Size(Var1) <> 0 Then
        TypeName = VBA.TypeName(NodeObject(VarPtr(Var1)))
    End If
#Else
    If LocalSize(Var1) <> 0 Then
        TypeName = VBA.TypeName(NodeObject(VarPtr(Var1)))
    End If
#End If
End Property


Private Function PtrVar(ByVal lPtr As Long) As Variant
    Dim lZero As Long
    Dim newObj As Variant
#If VBIDE = -1 Then
    sec.rtlMovMem newObj, False, lPtr, False, 4&
    PtrVar = newObj
    sec.rtlMovMem newObj, False, lZero, False, 4&
#Else
    rtlMovMem newObj, lPtr, 4&
    PtrVar = newObj
    rtlMovMem newObj, lZero, 4&
#End If

End Function


Public Property Get IsObject(ByRef List As ListType) As Boolean
    Dim Var1 As Long
    IsObject = VBA.IsObject(PtrVar(Register(List.Point, 4)))
    Var1 = Register(List.Point, 0)
#If VBIDE = -1 Then
    If sec.Size(Var1) <> 0 Then
        IsObject = VBA.IsObject(NodeObject(VarPtr(Var1)))
    End If
#Else
    If LocalSize(Var1) <> 0 Then
        IsObject = VBA.IsObject(NodeObject(VarPtr(Var1)))
    End If
#End If
End Property

Public Function Total(ByRef List As ListType) As Long
    With List
        Total = Abs(.Track + .Total - .Track + .Total + -.Total)
    End With
End Function

Public Function IsValidList(ByRef List As ListType) As Boolean
    IsValidList = ((List.Point <> 0) And (List.First <> 0) And (List.Final <> 0))
    'no current no list when we have all zero values for these
End Function

Public Function BOL(ByRef List As ListType) As Boolean
    BOL = (List.Point = List.First) Or (List.First = 0)
    'BOL =((List.Point = List.First) And (List.Total > 0)) Or ((List.Point = List.Final) And (List.Total < 0))
End Function

Public Function EOL(ByRef List As ListType) As Boolean
    EOL = (List.Point = List.Final) Or (List.Final = 0)
    'EOL = ((List.Point = List.Final) And (List.Total > 0)) Or ((List.Point = List.First) And (List.Total < 0))
End Function


Public Function AddDelMiddleNode(ByRef List As ListType, ByVal AddOrDel As Boolean) As Long
    If (AddOrDel And (EOL(List) Or (Total(List) = 0))) Then
        'list is at final already just delete
        AddDelMiddleNode = AddToLastNode(List)
    ElseIf ((Not AddOrDel) And BOL(List)) Then
        'list is at first already just delete
        AddDelMiddleNode = DelFirstNode(List)
    ElseIf IsValidList(List) Or List.Total <> 0 Then
        Dim lNode As Long
        If AddOrDel Then
           ' lNode = List.First
           ' Swap List.Point, List.First
          '  Swap List.Prior, List.Final
       '     MoveNode List, True
            AddDelMiddleNode = AddToLastNode(List)
         '   List.First = lNode
        Else
         '   MoveNode List, True
            AddDelMiddleNode = DelFirstNode(List)



        End If

    End If
End Function

Public Function AddToLastNode(ByRef List As ListType) As Long

  '  Dim Point As Long
    Dim Node As NodeType
'    List.Total = (-List.Total)
'    List.Track = (-List.Track)
    Swap List.First, List.Final
        
    Swap List.Point, List.Prior
    If IsValidList(List) Then
        'List.Prior = Register(List.Point, 4)
        Node = GetNode(List.Point)
        List.Prior = Node.Prior
    End If

#If VBIDE = -1 Then
    Node.Prior = sec.Alloc(0, 8)
    If (Not ((Node.Prior = sec.Freeze(Node.Prior)))) Then Err.Raise 8, App.Title, "Memory lock error."
#If TestNodes = -1 Then
    Debug.Print "Alloc: " & Node.Prior
#End If
#Else
    Node.Prior = LocalAlloc(GMEM_FIXED, 8)
    If (Not ((Node.Prior = LocalLock(ByVal Node.Prior)))) Then Err.Raise 8, App.Title, "Memory lock error."
#End If
    AddToLastNode = Node.Prior

    If IsValidList(List) Then

        SetNode List.Point, Node
        'Register(List.Point, 4) = Point

    Else
        List.Prior = Node.Prior
        List.First = Node.Prior
    End If

    List.Point = Node.Prior
    Node.Prior = List.Prior
    List.Final = List.Point
    'commit the current node

    SetNode List.Point, Node
    'Register(List.Point, 4) = Point

    Swap List.Point, List.Prior
    Swap List.First, List.Final
        
    List.Track = (-(List.Total - Abs(List.Track)))
    List.Total = (List.Total + IIf(List.Total > 0, 1, -1))

End Function

Public Function DelFirstNode(ByRef List As ListType) As Long

    If IsValidList(List) Then

        Dim Revert As Boolean
        List.Total = (-List.Total)
        Revert = (List.Track < 0)
        List.Track = Abs(List.Track)

        Swap List.First, List.Final

        Dim Point As Long
        Dim Node As NodeType
        'get the first node in list
        Node = GetNode(List.Point)
        'Point = Register(List.Point, 0)
    #If VBIDE = -1 Then
        If (sec.Size(Node.Value) = 4) Then
            sec.UnFreeze Node.Value
            sec.Free Node.Value
        End If
    #Else
        If (LocalSize(Node.Value) = 4) Then
            LocalUnlock Node.Value
            LocalFree Node.Value
        End If
    #End If
        'Point = Register(List.Point, 4)

        Point = Node.Prior
        If (Point <> 0) Then
            SetNode List.Prior, Node
            'Register(List.Prior, 4) = Point
            Node = GetNode(Point)
            List.Final = Node.Prior
            'List.Final = Register(Point, 4)

        End If

        DelFirstNode = List.Point
    #If VBIDE = -1 Then
        sec.UnFreeze List.Point
        sec.Free List.Point
    #If TestNodes = -1 Then
        Debug.Print "Free: " & List.Point
    #End If
    #Else
        LocalUnlock List.Point
        LocalFree List.Point
    #End If

        'set list to retained value
        If (Total(List) = 1) Then
            List.Point = 0
            List.First = 0
            List.Final = 0
            List.Track = 0
        Else
            List.First = Point
            List.Point = Point
        End If

        Swap List.First, List.Final

        List.Track = ((-List.Track) + IIf(List.Track > 0, 1, -1))
        List.Total = (List.Total + IIf(List.Total > 0, -1, 1))
        List.Track = Abs(List.Track) * IIf(Revert, -1, 1)

    End If
End Function

Public Sub MoveNode(ByRef List As ListType, ByVal Reverse As Boolean)
    If (List.Point = 0) Then Exit Sub
    If (Reverse And (List.Total > 0)) Or ((Not Reverse) And (List.Total < 0)) Then
        Swap List.Prior, List.Final
        Swap List.Point, List.Final
        Swap List.First, List.Prior
        List.Total = -List.Total
    Else
        Swap List.Prior, List.Point
        'List.Point = Register(List.Prior, 4)
        List.Point = GetNode(List.Prior).Prior
    End If

End Sub

Public Sub DisposeOfAll(ByRef List As ListType)
    Do While IsValidList(List) 'until we are done
        DelFirstNode List 'remove the first node
    Loop
    #If VBIDE = -1 Then
        List.Total = 0
    #End If
End Sub




'Public Function AddDelMiddleNode(ByRef List As ListType, ByVal AddOrDel As Boolean) As Long
'    If (AddOrDel And (EOL(List) Or (Total(List) = 0))) Then
'        'list is at final already just delete
'
'        AddDelMiddleNode = AddToLastNode(List)
'
'    ElseIf ((Not AddOrDel) And BOL(List)) Then
'        'list is at first already just delete
'
'        AddDelMiddleNode = DelFirstNode(List)
'
'    ElseIf IsValidList(List) Or List.Total <> 0 Then
'
'        If AddOrDel Then
'
'            'todo: use the short cut instead of looping
'            Do Until EOL(List)
'                MoveNode List, (List.Total < 0)
'            Loop
'           ' lFinal = List.Final
''            Swap List.Point, List.First
''            Swap List.prior, List.Final
'
'           ' MoveNode List, True
'            AddDelMiddleNode = AddToLastNode(List)
'
'           ' List.Final = lFinal
'
'        Else
'
'            'todo: use the short cut instead of looping
'            Do Until BOL(List)
'                MoveNode List, (List.Total < 0)
'            Loop
'            'MoveNode List, False
'            'Swap List.Final, List.First
'            AddDelMiddleNode = DelFirstNode(List)
'
''            Swap List.Point, List.First
''            Swap List.prior, List.Final
'        End If
'    End If
'End Function
'
'Public Function AddToLastNode(ByRef List As ListType) As Long
'
'  '  Dim Point As Long
'    Dim Node As NodeType
''    List.Total = (-List.Total)
''    List.Track = (-List.Track)
'
'    Swap List.Point, List.Prior
'    If IsValidList(List) Then
'        'List.Prior = Register(List.Point, 4)
'        Node = GetNode(List.Point)
'        List.Prior = Node.Prior
'    End If
'
'#If VBIDE = -1 Then
'    Node.Prior = sec.Alloc(0, 8)
'    If (Not ((Node.Prior = sec.Freeze(Node.Prior)))) Then Err.Raise 8, App.Title, "Memory lock error."
'#If TestNodes = -1 Then
'    Debug.Print "Alloc: " & Node.Prior
'#End If
'#Else
'    Node.Prior = LocalAlloc(GMEM_FIXED, 8)
'    If (Not ((Node.Prior = LocalLock(ByVal Node.Prior)))) Then Err.Raise 8, App.Title, "Memory lock error."
'#End If
'    AddToLastNode = Node.Prior
'
'    If IsValidList(List) Then
'
'        SetNode List.Point, Node
'        'Register(List.Point, 4) = Point
'
'    Else
'        List.Prior = Node.Prior
'        List.First = Node.Prior
'    End If
'
'    List.Point = Node.Prior
'    Node.Prior = List.Prior
'    List.Final = List.Point
'    'commit the current node
'
'    SetNode List.Point, Node
'    'Register(List.Point, 4) = Point
'
'    Swap List.Point, List.Prior
'
'    List.Track = (-(List.Total - Abs(List.Track)))
'    List.Total = (List.Total + IIf(List.Total > 0, 1, -1))
'
'End Function
'
'Public Function DelFirstNode(ByRef List As ListType) As Long
'
'    If IsValidList(List) Then
'
'        Dim Revert As Boolean
'        List.Total = (-List.Total)
'        Revert = (List.Track < 0)
'        List.Track = Abs(List.Track)
'
'        Swap List.First, List.Final
'
'        Dim Point As Long
'        Dim Node As NodeType
'        'get the first node in list
'        Node = GetNode(List.Point)
'        'Point = Register(List.Point, 0)
'    #If VBIDE = -1 Then
'        If (sec.Size(Node.Value) = 4) Then
'            sec.UnFreeze Node.Value
'            sec.Free Node.Value
'        End If
'    #Else
'        If (LocalSize(Node.Value) = 4) Then
'            LocalUnlock Node.Value
'            LocalFree Node.Value
'        End If
'    #End If
'        'Point = Register(List.Point, 4)
'
'        Point = Node.Prior
'        If (Point <> 0) Then
'            SetNode List.Prior, Node
'            'Register(List.Prior, 4) = Point
'            Node = GetNode(Point)
'            List.Final = Node.Prior
'            'List.Final = Register(Point, 4)
'
'        End If
'
'        DelFirstNode = List.Point
'    #If VBIDE = -1 Then
'        sec.UnFreeze List.Point
'        sec.Free List.Point
'    #If TestNodes = -1 Then
'        Debug.Print "Free: " & List.Point
'    #End If
'    #Else
'        LocalUnlock List.Point
'        LocalFree List.Point
'    #End If
'
'        'set list to retained value
'        If (Total(List) = 1) Then
'            List.Point = 0
'            List.First = 0
'            List.Final = 0
'            List.Track = 0
'        Else
'            List.First = Point
'            List.Point = Point
'        End If
'
'        Swap List.First, List.Final
'
'        List.Track = ((-List.Track) + IIf(List.Track > 0, 1, -1))
'        List.Total = (List.Total + IIf(List.Total > 0, -1, 1))
'        List.Track = Abs(List.Track) * IIf(Revert, -1, 1)
'
'    End If
'End Function
'
'Public Sub MoveNode(ByRef List As ListType, ByVal Reverse As Boolean)
'    If (List.Point = 0) Then Exit Sub
'    If (Reverse And (List.Total > 0)) Or ((Not Reverse) And (List.Total < 0)) Then
'        Swap List.Prior, List.Final
'        Swap List.Point, List.Final
'        Swap List.First, List.Prior
'        List.Total = -List.Total
'    Else
'        Swap List.Prior, List.Point
'        'List.Point = Register(List.Prior, 4)
'        List.Point = GetNode(List.Prior).Prior
'    End If
'
'End Sub
'
'Public Sub DisposeOfAll(ByRef List As ListType)
'    Do While IsValidList(List) 'until we are done
'        DelFirstNode List 'remove the first node
'    Loop
'    #If VBIDE = -1 Then
'        List.Total = 0
'    #End If
'End Sub
