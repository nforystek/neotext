Attribute VB_Name = "modCallbacks"
#Const modCallbacks = -1
#Const modService = -1
Option Explicit
'TOP DOWN

Option Compare Binary

Public Type Memory
    Pointer As Long
    Address As Long
End Type

Public memCallBacks() As Memory

Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function ObjectPointers(ObjectClass As Object, ByVal SectionOf As Long, Optional ByVal SimpleCount As Long, Optional ByVal ComplexCount As Long, Optional ByVal MethodCount As Long) As Memory()
    Dim FPS() As Memory
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(ObjectClass)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4
    Dim PTX As Long
    Dim cnt As Long
    'Select Case SectionOf
   '     Case 1 'public simple data types
            If SimpleCount > 0 Then
                ReDim FPS(SimpleCount - 1)
                For cnt = 0 To SimpleCount - 1
                    PTX = VTable + 28 + (cnt * 2 * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next
            End If
  '      Case 2 'public objects and variants
            If ComplexCount > 0 Then
                ReDim FPS((SimpleCount + ComplexCount) - 1)
                For cnt = 0 To ComplexCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (cnt * 3 * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next
            End If
 '       Case 3 'public Functions and Subs
            If MethodCount > 0 Then
                ReDim FPS((SimpleCount + ComplexCount + MethodCount) - 1)
                For cnt = 0 To MethodCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (ComplexCount * 3 * 4) + cnt * 4
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next
            End If
 '   End Select
    ObjectPointers = FPS
End Function


