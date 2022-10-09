VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN
Option Compare Binary

Public Type Partition
    Offset As Long
    Length As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long

Private Pointer As Long

Friend Property Get Address(Optional ByVal Offset As Long = 0)
    Address = Pointer + 4 + Offset
End Property
Private Sub Class_Initialize()
    Dim lSize As Long: lSize = 0
    Pointer = HeapAlloc(GetProcessHeap, 0, 4)
    CopyMemory ByVal Pointer, lSize, 4
End Sub

Private Sub Class_Terminate()
    HeapFree GetProcessHeap, 0, Pointer
End Sub

Public Sub Reset()
    Class_Terminate
    Class_Initialize
End Sub

Public Property Get Size() As Long
    CopyMemory Size, ByVal Pointer, 4
End Property

Public Property Let Size(ByVal newVal As Long)
    If ((newVal <> Size) And (newVal > -1)) Then
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, newVal + 4)
        CopyMemory ByVal Pointer, newVal, 4
    ElseIf (newVal = 0) Then
        Reset
    End If
End Property

Public Sub Push(ByVal Offset As Long)
    Dim lSize As Long
    lSize = Size
    If (Offset < lSize) Then
        CopyMemory ByVal (Pointer + 4), ByVal (Pointer + 4) + Offset, (-lSize + -Offset + (lSize * 2))
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (lSize - Offset) + 4)
        CopyMemory ByVal Pointer, (lSize - Offset), 4
    ElseIf (Offset = lSize) Then
        Reset
    End If
End Sub

Public Sub Concat(ByRef Value() As Byte)
    Dim nSize As Long
    nSize = UBound(Value) + -CInt(Not CBool(-LBound(Value)))
    If (nSize > 0) Then
        Dim lSize As Long
        CopyMemory lSize, ByVal Pointer, 4
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (nSize + lSize) + 4)
        CopyMemory ByVal Pointer, (nSize + lSize), 4
        CopyMemory ByVal (Pointer + 4) + lSize, ByVal VarPtr(Value(LBound(Value))), nSize
    End If
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte()
    Dim lSize As Long
    Dim Ret() As Byte
    CopyMemory lSize, ByVal Pointer, 4
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    If (Width > 0) Then
        ReDim Ret(1 To Width) As Byte
        CopyMemory ByVal VarPtr(Ret(LBound(Ret))), ByVal Pointer + 4 + Offset, Width
        ReDim Preserve Ret(1 To Width) As Byte
    Else
        ReDim Preserve Ret(-1 To -1) As Byte
    End If
    Partial = Ret
End Function

