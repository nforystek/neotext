Attribute VB_Name = "modCallbacks"
#Const modCallbacks = -1
#Const modService = -1
Option Explicit

Option Compare Binary

Public Type Memory
    Pointer As Long
    Address As Long
End Type

Public memCallBacks() As Memory

Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub Main()
    'This DLL is designed to allow users to create events that link to simple
    'sub procedures and doesn't require them on a form or class withevents...
    'default I've written it to fire an event on a timer as a simple example.
End Sub


'this function is originally derived form a piece on pscode.com titled "OBJECT METHODS/PROPERTIES POINTERS"
Public Function ObjectPointers(ObjectClass As Object, ByVal SectionOf As Long, Optional ByVal SimpleCount As Long, Optional ByVal ComplexCount As Long, Optional ByVal MethodCount As Long) As Memory()
    Dim FPS() As Memory
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(ObjectClass)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4
    Dim PTX As Long
    Dim cnt As Long
    'public simple data types section, mainly number types
    If SimpleCount > 0 And SectionOf = 1 Or SectionOf = 0 Then
        ReDim FPS(SimpleCount - 1)
        For cnt = 0 To SimpleCount - 1
            PTX = VTable + 28 + (cnt * 2 * 4)
            RtlMoveMemory FPS(cnt).Pointer, PTX, 4
            RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
        Next
    End If
    'public objects and variants, more then just a pointer
    If ComplexCount > 0 And SectionOf = 2 Or SectionOf = 0 Then
        ReDim FPS((SimpleCount + ComplexCount) - 1)
        For cnt = 0 To ComplexCount - 1
            PTX = VTable + 28 + (SimpleCount * 2 * 4) + (cnt * 3 * 4)
            RtlMoveMemory FPS(cnt).Pointer, PTX, 4
            RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
        Next
    End If
    'public Functions and Subs, this does work for property get/let
    If MethodCount > 0 And SectionOf = 3 Or SectionOf = 0 Then
        ReDim FPS((SimpleCount + ComplexCount + MethodCount) - 1)
        For cnt = 0 To MethodCount - 1
            PTX = VTable + 28 + (SimpleCount * 2 * 4) + (ComplexCount * 3 * 4) + cnt * 4
            RtlMoveMemory FPS(cnt).Pointer, PTX, 4
            RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
        Next
    End If
    ObjectPointers = FPS
End Function



