






Attribute VB_Name = "modBase"
#Const modBase = -1
Option Explicit

Option Private Module

Public Enum NodeRelation
    CircularLinkedList = 1
    AngulingLinkedTree = 2
End Enum

Public Type Memory
    Pointer As Long
    Address As Long
End Type

Public Type NodeType
    Start As Long
    Point As Long
    Final As Long
    Value As Long
End Type

Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GPTR = &H40
Private Const GHND = &H42

Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long

Private Declare Sub CopyMemoryNode Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As NodeType, ByVal xSource As Long, ByVal nbytes As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByVal xSource As Long, ByVal nbytes As Long)

Private Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)


Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long

'Public Methods() As Memory
'Public Simples() As Memory
'Public Complex() As Memory

'Public Sub Main()
''    Simples = ObjectPointers(Framing, 1, 2, 0, 6)
''    Complex = ObjectPointers(Framing, 2, 2, 0, 6)
''    Methods = ObjectPointers(Framing, 3, 2, 0, 6)
'End Sub

'Public Function ObjectPointers(ObjectClass As Object, ByVal SectionOf As Long, Optional ByVal SimpleCount As Long, Optional ByVal ComplexCount As Long, Optional ByVal MethodCount As Long) As Memory()
'    Dim FPS() As Memory
'    Dim OBJ1 As Long
'    OBJ1 = ObjPtr(ObjectClass)
'    Dim VTable As Long
'    RtlMoveMemory VTable, ByVal OBJ1, 4
'    Dim PTX As Long
'    Dim cnt As Long
'    Select Case SectionOf
'        Case 1 'public simple data types
'            If SimpleCount > 0 Then
'                ReDim FPS(SimpleCount - 1)
'                For cnt = 0 To SimpleCount - 1
'                    PTX = VTable + 28 + (cnt * 2 * 4)
'                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
'                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
'                Next
'            End If
'        Case 2 'public objects and variants
'            If ComplexCount > 0 Then
'                ReDim FPS(ComplexCount - 1)
'                For cnt = 0 To ComplexCount - 1
'                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (cnt * 3 * 4)
'                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
'                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
'                Next
'            End If
'        Case 3 'public Functions and Subs
'            If MethodCount > 0 Then
'                ReDim FPS(MethodCount - 1)
'                For cnt = 0 To MethodCount - 1
'                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (ComplexCount * 3 * 4) + cnt * 4
'                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
'                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
'                Next
'            End If
'    End Select
'    ObjectPointers = FPS
'End Function

Public Function Padding(ByVal Length As Long, ByVal Value As String, Optional ByVal PadWith As String = " ") As String
    Padding = String(Abs((Length * Len(PadWith)) - (Len(Value) \ Len(PadWith))), PadWith) & Value
End Function

Public Function Toggler(ByVal Value As Long) As Long
    Toggler = (-CInt(CBool(Value)) + -1) + -CInt(Not CBool(-Value + -1))
End Function

Private Function GetNode(ByVal Addr As Long) As NodeType
    Dim n As NodeType
    CopyMemoryNode n, Addr - 6, 16
    GetNode = n
End Function

Private Sub SetNode(ByVal Addr As Long, ByRef Node As NodeType)
    Dim tmpAddr As Long
    tmpAddr = VarPtr(Node)
    RtlMoveMemory ByVal Addr - 6, ByVal tmpAddr, 16
End Sub

Public Sub GetValue(ByRef My As NodeType)
    RtlMoveMemory My.Value, ByVal My.Point + 6, 4&
End Sub

Public Sub SetValue(ByRef My As NodeType)
    RtlMoveMemory ByVal My.Point + 6, My.Value, 4&
End Sub

Public Function IsValidNode(ByRef My As NodeType) As Boolean
    Dim tmpCheck As Long
    If My.Start <> 0 Then RtlMoveMemory tmpCheck, ByVal My.Start, 4&
    IsValidNode = (tmpCheck <> 0 And My.Final <> 0)
End Function

Public Sub AddDelMiddleNode(ByRef My As NodeType, ByVal AddOrDel As Boolean)
    Dim n1 As NodeType
    Dim n2 As NodeType
    Dim n3 As NodeType
    Dim n4 As NodeType
    Dim n5 As NodeType
    Dim tmp As Long
    Dim tmp2 As Long
    Dim Forth As Long
    Dim Prior As Long
    
    If My.Start <> 0 Then RtlMoveMemory tmp, ByVal My.Start, 4&
    If (Not (tmp = 0 Or My.Final = 0)) Then
        If AddOrDel And (My.Start = My.Point) Then
            AddToLastNode My
        ElseIf (Not AddOrDel) And (My.Start = My.Point) Then
            DelFirstNode My
        Else

            n1 = GetNode(tmp)
            n3 = GetNode(My.Point)
            Prior = GetLastAddr(My)
            Forth = GetNextAddr(My)
            If Prior <> 0 Then n2 = GetNode(Prior)
            If Forth <> 0 Then n4 = GetNode(Forth)
            n5 = GetNode(My.Final)
            n5.Point = tmp
            n3.Point = 0
            n2.Point = My.Point
            n4.Start = 0
            n4.Final = 0
            n3.Start = HiWord(Prior - My.Point)
            n3.Final = LoWord(Prior - My.Point)
            n1.Start = HiWord(tmp - My.Final)
            n1.Final = LoWord(tmp - My.Final)
                
            If AddOrDel Then
            
                SetNode tmp, n1
                If Prior <> 0 Then SetNode Prior, n2
                SetNode My.Point, n3
                If Forth <> 0 Then SetNode Forth, n4
                SetNode My.Final, n5
                RtlMoveMemory ByVal My.Start, tmp, 4&
            
                AddToLastNode My
            Else
            
                SetNode tmp, n3
                If Prior <> 0 Then SetNode Prior, n2
                SetNode My.Point, n1
                If Forth <> 0 Then SetNode Forth, n4
                SetNode My.Final, n5
                RtlMoveMemory ByVal My.Start, tmp, 4&

                SetNextAddr My

                DelFirstNode My
            End If
        End If
    ElseIf AddOrDel Then
        AddToLastNode My
    End If
End Sub

Public Sub AddToLastNode(ByRef My As NodeType)
    Dim n As NodeType
    Dim tmp As Long
    If My.Start <> 0 Then RtlMoveMemory tmp, ByVal My.Start, 4&
    If (tmp = 0) And (My.Final = 0) Then
        n.Start = 0
        n.Final = 0
        My.Start = GlobalAlloc(GMEM_FIXED Or GPTR, 4)
        If Not My.Start = GlobalLock(My.Start) Then Err.Raise 8, App.Title, "Memory lock error."
        tmp = GlobalAlloc(GMEM_FIXED And My.Start, 16) + 6
        If Not (tmp - 6) = GlobalLock(tmp - 6) Then Err.Raise 8, App.Title, "Memory lock error."
        RtlMoveMemory ByVal My.Start, tmp, 4&
        n.Point = 0
    Else
        If Not (My.Final = 0) Then n = GetNode(My.Final)
        tmp = GlobalAlloc(GMEM_FIXED And My.Final, 16) + 6
        If Not (tmp - 6) = GlobalLock(tmp - 6) Then Err.Raise 8, App.Title, "Memory lock error."
        n.Point = tmp
        If (My.Final = 0) Then My.Final = My.Point
        SetNode My.Final, n
        n = GetNode(My.Final)
        n.Start = HiWord((tmp - My.Final))
        n.Final = LoWord((tmp - My.Final))
        n.Point = 0
    End If
    SetNode tmp, n
    My.Final = tmp
    My.Point = tmp
End Sub

Public Sub DelFirstNode(ByRef My As NodeType)
    Dim n As NodeType
    If My.Start <> 0 Then RtlMoveMemory My.Point, ByVal My.Start, 4&
    If My.Point = 0 Or My.Final = 0 Then Exit Sub
    n = GetNode(My.Point)
    Dim tmp As Long
    tmp = n.Point
    RtlMoveMemory ByVal My.Start, tmp, 4&
    If tmp <> 0 Then
        n = GetNode(tmp)
        n.Start = 0
        n.Final = 0
        SetNode tmp, n
    End If
    If My.Point = My.Final Or My.Start = My.Final Then
        GlobalUnlock My.Point - 6
        GlobalFree My.Point - 6
        GlobalUnlock My.Start
        GlobalFree My.Start
        RtlMoveMemory ByVal My.Start, 0&, 4&
        My.Start = 0
        My.Final = 0
    Else
        GlobalUnlock My.Point - 6
        GlobalFree My.Point - 6
    End If
    My.Point = tmp
End Sub

Public Function GetLastAddr(ByRef My As NodeType) As Long
    If IsValidNode(My) Then
        Dim n As NodeType
        n = GetNode(My.Point)
        Dim tmp As Long
        HiWord(tmp) = n.Start
        LoWord(tmp) = n.Final
        If tmp = 0 Then
            GetLastAddr = My.Final
        Else
            GetLastAddr = My.Point - tmp
        End If
    End If
End Function

Public Function GetNextAddr(ByRef My As NodeType) As Long
    If IsValidNode(My) Then
        Dim n As NodeType
        n = GetNode(My.Point)
        Dim tmp As Long
        HiWord(tmp) = n.Start
        LoWord(tmp) = n.Final
        If n.Point = 0 Then
            RtlMoveMemory GetNextAddr, ByVal My.Start, 4&
        ElseIf n.Point > tmp Then
            GetNextAddr = n.Point
        Else
            GetNextAddr = My.Point - tmp
        End If
    End If
End Function

Public Sub SetLastAddr(ByRef My As NodeType)
    If IsValidNode(My) Then
        Dim n As NodeType
        n = GetNode(My.Point)
        Dim tmp As Long
        HiWord(tmp) = n.Start
        LoWord(tmp) = n.Final
        If tmp = 0 Then
            My.Point = My.Final
        Else
            My.Point = My.Point - tmp
        End If
        RtlMoveMemory ByVal My.Start, My.Point, 4&
    End If
End Sub

Public Sub SetNextAddr(ByRef My As NodeType)
    If IsValidNode(My) Then
        Dim n As NodeType
        n = GetNode(My.Point)
        Dim tmp As Long
        HiWord(tmp) = n.Start
        LoWord(tmp) = n.Final
        If n.Point = 0 Then
            RtlMoveMemory My.Point, ByVal My.Start, 4&
        ElseIf n.Point > tmp Then
            My.Point = n.Point
        Else
            My.Point = My.Point - tmp
        End If
    End If
End Sub

Public Sub DisposeOfAll(ByRef My As NodeType)
    Do While IsValidNode(My)
        DelFirstNode My
    Loop
End Sub

Public Property Get LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And &HFFFF&)
End Property

Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Long)
   lThis = lThis And Not &HFFFF& Or lLoWord
End Property

Public Property Get HiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      HiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Property

Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Long)
   If (lHiWord And &H8000&) = &H8000& Then
      lThis = lThis And Not &HFFFF0000 Or ((lHiWord And &H7FFF&) * &H10000) Or &H80000000
   Else
      lThis = lThis And Not &HFFFF0000 Or (lHiWord * &H10000)
   End If
End Property








