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
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Binary

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

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
    RtlMoveMemory ByVal Pointer, lSize, 4
End Sub

Private Sub Class_Terminate()
    HeapFree GetProcessHeap, 0, Pointer
End Sub

Public Sub Reset()
    Class_Terminate
    Class_Initialize
End Sub

Public Property Get Length() As Long
    RtlMoveMemory Length, ByVal Pointer, 4
End Property

Public Property Let Length(ByVal RHS As Long)
    If (RHS > -1) Then
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, RHS + 4)
        RtlMoveMemory ByVal Pointer, RHS, 4
    ElseIf (RHS = 0) Then
        Reset
    End If
End Property

Public Sub Push(ByVal Offset As Long)
    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        RtlMoveMemory ByVal (Pointer + 4), ByVal (Pointer + 4) + Offset, (-lSize + -Offset + (lSize * 2))
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (lSize - Offset) + 4)
        RtlMoveMemory ByVal Pointer, (lSize - Offset), 4
    ElseIf (Offset = lSize) Then
        Reset
    End If
End Sub

Public Sub Pull(ByVal Offset As Long)
     Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        RtlMoveMemory ByVal (Pointer + 4) + Offset, ByVal (Pointer + 4), (-lSize + -Offset + (lSize * 2))
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (lSize - Offset) + 4)
        RtlMoveMemory ByVal Pointer, (lSize - Offset), 4
    ElseIf (Offset = lSize) Then
        Reset
    End If
End Sub

Public Sub Concat(ByRef RHS() As Byte)
    Dim nSize As Long
    nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
    If (nSize > 0) Then
        Dim lSize As Long
        RtlMoveMemory lSize, ByVal Pointer, 4
        Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (nSize + lSize) + 4)
        RtlMoveMemory ByVal Pointer, (nSize + lSize), 4
        RtlMoveMemory ByVal (Pointer + 4) + lSize, ByVal VarPtr(RHS(LBound(RHS))), nSize
    End If
End Sub


Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
    Reset
    Concat Origin.Partial(Offset, Width)
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte()
    Dim lSize As Long
    RtlMoveMemory lSize, ByVal Pointer, 4
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    If (Width > 0) Then
        Dim Ret() As Byte
        ReDim Ret(1 To Width) As Byte
        RtlMoveMemory ByVal VarPtr(Ret(LBound(Ret))), ByVal Pointer + 4 + Offset, Width
        Partial = Ret
        Erase Ret
    End If
End Function

Public Sub Prepend(ByRef RHS() As Byte)
    If Length = 0 Then
        Concat RHS
    Else
        Dim nSize As Long
        nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
        Length = Length + (nSize * 2)
        Pull nSize
        Placeat RHS, 0, nSize
    End If
End Sub

Public Function Placeat(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
    Dim lSize As Long
    RtlMoveMemory lSize, ByVal Pointer, 4
    If Width = -1 Then Width = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    If (Width > 0) Then
        RtlMoveMemory ByVal Pointer + 4 + Offset, ByVal VarPtr(RHS(LBound(RHS))), Width
    End If
End Function
