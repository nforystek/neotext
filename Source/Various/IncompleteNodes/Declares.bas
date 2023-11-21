Attribute VB_Name = "Declares"

'y = mx + b
'b = y + mx
'm = y2 - y1 / x2 - x1



 'y = ax2 + bx + c
 
Option Explicit
Option Compare Binary

Public Type ListType
    First As Long 'first node of the list
    Point As Long 'current node in the list
    check As Long 'node after this one
    Final As Long 'last node of the list
    Total As Long 'total count of nodes
End Type

Public Type NodeType
    Value As Long 'value of the node
    check As Long 'node after this one
End Type

Public Type Addresses
    A As Long 'first
    B As Long
    
    C As Long 'point
    D As Long
    
    E As Long 'check
    F As Long

    G As Long 'final
    H As Long
    
    I As Long 'Total
    J As Long
End Type

Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_DISCARDED = &H4000
Public Const LMEM_FIXED = &H0
Public Const LMEM_INVALID_HANDLE = &H8000
Public Const LMEM_LOCKM = &HFF
Public Const LMEM_MODIFY = &H80
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10
Public Const LMEM_NODISCARD = &H20
Public Const LMEM_VALID_FLAGS = &HF72
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = (LMEM_FIXED + LMEM_ZEROINIT)


Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Declare Function LocalShrink Lib "kernel32" (ByVal hMem As Long, ByVal cbNewSize As Long) As Long
Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKM = &HFF
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long



Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long


Public Const VK_ESCAPE = &H1B

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Sub rtlMovMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
'Private Declare Sub rtlMovObjRef Lib "kernel32" Alias "RtlMoveMemory" (xDest As Object, ByRef xSource As Long, ByVal nbytes As Long)

Public Declare Sub PutLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, num As Any, Optional ByVal Size As Long = 4&)
Public Declare Sub GetLong Lib "kernel32" Alias "RtlMoveMemory" (num As Any, ByVal ptr As Long, Optional ByVal Size As Long = 4&)
Public Declare Sub GetMem8 Lib "MSVBVM60.DLL" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub GetMem4 Lib "MSVBVM60.DLL" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub PutMem4 Lib "MSVBVM60.DLL" (ByVal Addr As Long, ByVal newVal As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)

Public Declare Sub CopyBytes Lib "MSVBVM60.DLL" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByRef Dest As Any, ByVal src As Any)

Public Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Quit As Integer

Public Const LongQuad As Currency = 1073741824

Private Const Zero As Long = 0


Public DebugLate As String
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Public Function GetCaption(ByVal hwnd As Long) As String
    Dim txt As String
    Dim lSize As Long
    txt = String(255, Chr(0))
    lSize = Len(txt)
    Call GetWindowText(hwnd, txt, lSize)
    GetCaption = Trim(Replace(txt, Chr(0), ""))
End Function

Public Sub Swap(ByRef Var1 As Variant, ByRef Var2 As Variant, Optional ByRef Var3 As Variant, Optional ByRef Var4 As Variant, Optional ByRef Var5 As Variant, Optional ByRef Var6 As Variant, Optional ByRef Var7 As Variant, Optional ByRef Var8 As Variant, Optional ByRef Var9 As Variant, Optional ByRef Var0 As Variant)
    If IsMissing(Var3) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var0
    ElseIf IsMissing(Var4) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var0
    ElseIf IsMissing(Var5) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var0
    ElseIf IsMissing(Var6) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var5
        Var5 = Var0
    ElseIf IsMissing(Var7) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var5
        Var5 = Var6
        Var6 = Var0
    ElseIf IsMissing(Var8) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var5
        Var5 = Var6
        Var6 = Var7
        Var7 = Var0
    ElseIf IsMissing(Var9) Then
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var5
        Var5 = Var6
        Var6 = Var7
        Var7 = Var8
        Var8 = Var0
    Else
        Var0 = Var1
        Var1 = Var2
        Var2 = Var3
        Var3 = Var4
        Var4 = Var5
        Var5 = Var6
        Var6 = Var7
        Var7 = Var8
        Var8 = Var9
        Var9 = Var0
    End If
End Sub


Public Function Padding(ByVal txt As String, ByVal nLen As Long) As String
    If nLen - Len(Trim(txt)) > 0 Then
        Padding = Trim(txt) & String(nLen - Len(Trim(txt)), " ")
    Else
        Padding = Trim(txt) & " "
    End If
End Function

Public Sub DebugPrint(ByRef Addrs As Addresses, Optional ByVal PrintToo As Boolean = True, Optional ByVal NewLine As Boolean = True)
    With Addrs
        Dim directs As String
        directs = DebugLate & IIf(.A > .B, ">", "<") & IIf(.B > .C, ">", "<") & IIf(.C > .D, ">", "<")
        directs = directs & _
                "a" & Padding(.A, 11) & "b" & Padding(.B, 11) & "c" & Padding(.C, 11) & "d" & Padding(.D, 11) & _
                "e" & Padding(.E, 11) & "f" & Padding(.F, 11) & "g" & Padding(.G, 11) & "h" & Padding(.H, 11) & _
                "i" & Padding(.I, 11) & "j" & Padding(.J, 11) & "l" & Padding(Lowest(Addrs), 11) & "u" & Padding(Highest(Addrs), 11)
        If PrintToo Then Form1.PrintMessage directs, NewLine
        If Not NewLine Then
            Debug.Print directs;
        Else
            Debug.Print directs
        End If
        DebugLate = ""
    End With
    If PrintToo Then Form1.PrintPrograms
End Sub



Public Sub DebugFlush(Optional ByVal PrintToo As Boolean = True)
    With Addrs
        If Left(DebugLate, 1) = "-" Then
            Do While InStr(DebugLate, "!") > 0
                DebugLate = Left(DebugLate, InStr(DebugLate, "!") - 2) & Mid(DebugLate, (InStr(InStr(DebugLate, "!") + 1, DebugLate, " ") + 1))
            Loop
        End If
       ' If PrintToo Then Form1.PrintMessage DebugLate, True

        Debug.Print DebugLate
        DebugLate = ""
    End With
   'If PrintToo Then Form1.PrintPrograms
End Sub

Public Function GetNode(ByRef Addr As Long) As NodeType
    If Addr = 0 Then Exit Function 'exit if no addr
    Dim Node As NodeType 'get at alloc addr from memory
'#If VBIDE = -1 Then
'    sec.rtlMovMem VarPtr(Node), True, Addr, True, 8&
'#Else
    rtlMovMem ByVal VarPtr(Node), ByVal Addr, 8&
'#End If
    GetNode.check = Node.check
    GetNode.Value = Node.Value
    GetNode = Node
End Function

Public Sub SetNode(ByRef Addr As Long, ByRef Node As NodeType)
    If Addr = 0 Then Exit Sub 'exit if no addr
    Dim NodePtr As Long 'store in var for api byref
    NodePtr = VarPtr(Node) 'set at alloc addr in memory
'#If VBIDE = -1 Then
'    sec.rtlMovMem Addr, True, NodePtr, True, 8&
'#Else
    rtlMovMem ByVal Addr, ByVal NodePtr, 8&
'#End If
End Sub

Public Property Get Register(ByVal Addr As Long, Optional ByVal OffSet As Long = 0) As Long
    If Addr <> 0 Then
#If VBIDE = -1 Then
        Register = sec.Register(Abs(Addr), OffSet)
#Else
        GetLong Register, Abs(Addr) + OffSet
#End If
    End If
End Property
Public Property Let Register(ByVal Addr As Long, Optional ByVal OffSet As Long = 0, ByVal RHS As Long)
    If Addr <> 0 Then
#If VBIDE = -1 Then
        sec.Register(Abs(Addr), OffSet) = RHS
#Else
        PutLong Abs(Addr) + OffSet, RHS
#End If
    End If
End Property

Public Function GetVariant(ByRef Addr As Long) As Variant
    If Addr = 0 Then Exit Function 'exit if no addr
    Dim Node As NodeType
    Node = GetNode(Addr)
    Dim Var1 As Long
    Var1 = Node.Value
#If VBIDE = -1 Then
    If sec.Size(Var1) = 4 Then
        GetVariant = PtrVar(Node.Value)
        sec.rtlMovMem GetVariant, False, Node.Value, True, LenB(GetVariant)
    End If
#Else
    If LocalSize(Var1) = 4 Then
        GetVariant = PtrVar(Node.Value)
        rtlMovMem GetVariant, ByVal Node.Value, LenB(GetVariant)
    End If
#End If

End Function

Public Sub SetVariant(ByRef Addr As Long, ByRef Var As Variant)
    If Addr = 0 Then Exit Sub 'exit if no addr
    Dim Node As NodeType
    Node = GetNode(Addr)
#If VBIDE = -1 Then
    If Not sec.Size(Node.Value) = 4 Then
        Node.Value = sec.Alloc(GMEM_FIXED And VarPtr(Var), LenB(Var)) 'allocate 12 bytes and set to mid struct
        If Not Node.Value = sec.Lock(Node.Value) Then Err.Raise 8, App.Title, "Memory lock error."
    End If
    sec.rtlMovMem Node.Value, True, Var, False, LenB(Var)
#Else
    If Not LocalSize(Node.Value) = 4 Then
        Node.Value = LocalAlloc(GMEM_FIXED And VarPtr(Var), LenB(Var)) 'allocate 12 bytes and set to mid struct
        If Not Node.Value = LocalLock(Node.Value) Then Err.Raise 8, App.Title, "Memory lock error."
    End If
    rtlMovMem ByVal Node.Value, Var, LenB(Var)
#End If

    SetNode Addr, Node
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
    Dim Node As NodeType
    Node = GetNode(List.Point)
    TypeName = VBA.TypeName(PtrVar(Node.check))
#If VBIDE = -1 Then
    If sec.Size(Node.Value) <> 0 Then
        TypeName = VBA.TypeName(NodeObject(VarPtr(Node.Value)))
    End If
#Else
    If LocalSize(Node.Value) <> 0 Then
        TypeName = VBA.TypeName(NodeObject(VarPtr(Node.Value)))
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
    Dim Node As NodeType
    Node = GetNode(List.Point)
    IsObject = VBA.IsObject(PtrVar(Node.check))
#If VBIDE = -1 Then
    If sec.Size(Node.Value) <> 0 Then
        IsObject = VBA.IsObject(NodeObject(VarPtr(Node.Value)))
    End If
#Else
    If LocalSize(Node.Value) <> 0 Then
        IsObject = VBA.IsObject(NodeObject(VarPtr(Node.Value)))
    End If
#End If
End Property





'Private Property Get Value(ByVal Addr As Long, Optional ByVal OffSet As Long = 0) As Long
'#If VBIDE = -1 Then
'    Value = sec.Register(Abs(Addr), OffSet)
'#Else
'    If Addr > 0 Then
'        GetLong Value, Addr + OffSet
'    End If
'#End If
'End Property
'Private Property Let Value(ByVal Addr As Long, Optional ByVal OffSet As Long = 0, ByVal RHS As Long)
'#If VBIDE = -1 Then
'    sec.Register(Abs(Addr), OffSet) = RHS
'#Else
'    If Addr > 0 Then
'        PutLong Addr + OffSet, RHS
'    End If
'#End If
'End Property
'
'
'Public Property Get Object(ByRef Addr As Long) As Object
'    Set Object = Nothing
'    Dim optr As Long
'    Dim Zero As Long
'#If VBIDE = -1 Then
'    optr = sec.Register(Addr)
'#Else
'    If Addr > 0 Then
'        GetLong optr, Addr
'    End If
'#End If
'    If optr > 0 Then
'        Dim newObj As Object
'        RtlMoveMemory newObj, optr, 4&
'        Set Object = newObj
'        RtlMoveMemory newObj, Zero, 4&
'    End If
'End Property
'
'Public Property Set Object(ByRef Addr As Long, RHS As Object)
'    Dim Obj As Object
'    Set Obj = Nothing
'    vbaObjSetAddref ObjPtr(Obj), ObjPtr(RHS)
'#If VBIDE = -1 Then
'    sec.Register(Addr) = ObjPtr(RHS)
'#Else
'    If Addr > 0 Then
'        PutLong Addr, ObjPtr(RHS)
'    End If
'#End If
'    Set RHS = Obj
'End Property

'
'Public Property Get IsObject(ByRef Addr8Bytes As Long) As Boolean ' _
'Gets whether or not the current node at B in the List is set to a object not equal to nothing.
'    Dim tmp As Object
'    Set tmp = Object(Addr8Bytes)
'    If VBA.IsObject(tmp) Then
'        IsObject = Not (tmp Is Nothing)
'    Else
'        IsObject = False
'    End If
'End Property


Public Function Toggler(ByVal Value As Long) As Long
    Toggler = ((-CInt(CBool(Value)) + -1) + -CInt(Not CBool(-Value + -1)))
End Function

Public Function Quitting() As Boolean
    Quitting = Quit <> GetKeyState(VK_ESCAPE)
End Function
 
Public Sub CleanUp()

    Dim Addr As Variant
    For Each Addr In sec

        sec.Free CLng(Addr)
    Next
    Set sec = Nothing
    
End Sub

Public Sub Setup()
    
    Quit = GetKeyState(VK_ESCAPE)
#If VBIDE = -1 Then
    Set sec = CreateObject("AddrPool.Addresses")
    Dim Addr As Variant
    For Each Addr In sec
        Debug.Print "Freeing: " & CLng(Addr)
        
        sec.Free CLng(Addr)
    Next
#End If
End Sub



Public Function LargeOf(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant
    If IsMissing(V3) Then
        If V1 > V2 Then
            LargeOf = V1
        ElseIf V2 <> 0 Then
            LargeOf = V2
        Else
            LargeOf = V1
        End If
    ElseIf IsMissing(V4) Then
        If V2 > V3 And V2 > V1 Then
            LargeOf = V2
        ElseIf V1 > V3 And V1 > V2 Then
            LargeOf = V1
        ElseIf V3 <> 0 Then
            LargeOf = V3
        Else
            LargeOf = LargeOf(V1, V2)
        End If
    Else
        If V2 > V3 And V2 > V1 And V2 > V4 Then
            LargeOf = V2
        ElseIf V1 > V3 And V1 > V2 And V1 > V4 Then
            LargeOf = V1
        ElseIf V3 > V1 And V3 > V2 And V3 > V4 Then
            LargeOf = V3
        ElseIf V4 <> 0 Then
            LargeOf = V4
        Else
            LargeOf = LargeOf(V1, V2, V3)
        End If
    End If
End Function

Public Function LeastOf(ByVal V1 As Variant, ByVal V2 As Variant, Optional ByVal V3 As Variant, Optional ByVal V4 As Variant) As Variant

    If IsMissing(V3) Then
        If (V2 = 0) And (Not (V2 = 0)) Then
            LeastOf = V1
        ElseIf (V1 = 0) And (Not (V1 = 0)) Then
            LeastOf = V2
        Else
            If (V1 > V2) Then
                LeastOf = V2
            Else
                LeastOf = V1
            End If
        End If
    ElseIf IsMissing(V4) Then
        If Not (V1 = 0 And V2 = 0 And V3 = 0) Then
            If V3 = 0 Then
                LeastOf = LeastOf(V1, V2)
            ElseIf V2 = 0 Then
                LeastOf = LeastOf(V1, V3)
            ElseIf V1 = 0 Then
                LeastOf = LeastOf(V3, V2)
            Else
                If V2 < V3 And V2 < V1 Then
                    LeastOf = V2
                ElseIf V1 < V3 And V1 < V2 Then
                    LeastOf = V1
                Else
                    LeastOf = V3
                End If
            End If
        End If
    Else
        If Not (V1 = 0 And V2 = 0 And V3 = 0 And V4 = 0) Then
            If V3 = 0 Then
                LeastOf = LeastOf(V1, V2, V4)
            ElseIf V3 = 0 Then
                LeastOf = LeastOf(V1, V2, V4)
            ElseIf V2 = 0 Then
                LeastOf = LeastOf(V1, V3, V4)
            ElseIf V1 = 0 Then
                LeastOf = LeastOf(V3, V2, V4)

            Else
                If ((V2 < V3) And (V2 < V1) And (V2 < V4)) Then
                    LeastOf = V2
                ElseIf ((V1 < V3) And (V1 < V2) And (V1 < V4)) Then
                    LeastOf = V1
                ElseIf ((V3 < V1) And (V3 < V2) And (V3 < V4)) Then
                    LeastOf = V3
                Else
                    LeastOf = V4
                End If
            End If
        End If
    End If
End Function


