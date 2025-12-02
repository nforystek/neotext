VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Binary

Option Base 1
Public Enum MemoryScopes
    Heap = 0
    Locking = 1
    [Local] = 2
    [Global] = 4
End Enum

Private Enum MemoryFunc
    Free = 0
    Alloc = 1
    ReAlloc = 2
End Enum

Private ScopeAs As MemoryScopes
Private Pointer As Long

Public Function Address(Optional ByVal Offset As Long = 0) As Long ' _
Returned the stream objects memory pointer, with any optional Offset specified added to it.
    RtlMoveMemory ByVal VarPtr(Address), (Pointer + Offset), 4
End Function

Private Sub Class_Initialize()
    ScopeAs = MemoryScopes.Heap
End Sub

Private Sub Class_Terminate()
    Reset ScopeAs
End Sub

Private Sub MemoryModify(Optional ByVal Size As Long = 0)
    If (Pointer <> 0) Then
        If ((ScopeAs And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                If HeapUnlock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory unlock failure."
            End If
        End If
    End If
    
    If (Size = 0) Then
        If (Pointer <> 0) And (Not (Length = 0)) Then
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                LocalFree Pointer
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                GlobalFree Pointer
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                HeapFree GetProcessHeap, 0, ByVal Pointer
            End If
            Pointer = 0
        End If
    Else
        If (Pointer = 0) And (Length = 0) Then
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalAlloc(modMemory.LMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalAlloc(modMemory.GMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                Pointer = HeapAlloc(GetProcessHeap, 0, Size)
            End If
        Else
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalReAlloc(Pointer, Size, modMemory.LMEM_MOVEABLE Or modMemory.lPtr)
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalReAlloc(Pointer, Size, modMemory.GMEM_MOVEABLE Or modMemory.GPTR)
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                Pointer = HeapReAlloc(GetProcessHeap, 0, ByVal Pointer, Size)
            End If
        End If
    End If

    If (Pointer <> 0) Then
        If ((ScopeAs And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                If HeapLock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory lock mismatch."
            End If
        End If
    End If
End Sub


Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
Destroys the callie object to be exactly the same as the Origin specified, with optional Offset bytes left out to the left and only of Width numebr of bytes.
Attribute Clone.VB_Description = "Destroys the callie object to be exactly the same as the Origin specified, with optional Offset bytes left out to the left and only of Width numebr of bytes."
    Reset ScopeAs
    Concat Origin.Partial(Offset, Width)
End Sub

Public Sub Concat(ByRef RHS() As Byte) ' _
Concatenation of byte array RHS to the end of the stream object adding as nessisary.
Attribute Concat.VB_Description = "Concatenation of byte array RHS to the end of the stream object adding as nessisary."
    Dim nSize As Long
    nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
    If (nSize > 0) Then
        Dim lSize As Long
        lSize = Length
        MemoryModify (nSize + lSize)
        RtlMoveMemory ByVal Pointer + lSize, ByVal VarPtr(RHS(LBound(RHS))), nSize
    End If
End Sub

Public Property Get Length() As Long ' _
Gets or sets the length of the stream object, does not change the memory scope.
Attribute Length.VB_Description = "Gets or sets the length of the stream object, does not change the memory scope."
    If Pointer <> 0 Then
        If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
            Length = LocalSize(Pointer)
        ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
            Length = GlobalSize(Pointer)
        ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
            Length = HeapSize(GetProcessHeap, 0, ByVal Pointer)
        End If
    Else
        Length = 0
    End If
End Property

Public Property Let Length(ByVal RHS As Long) ' _
Gets or sets the length of the stream object, does not change the memory scope.
    If (RHS > 0) Then
        MemoryModify RHS
    ElseIf (RHS = 0) Then
        Reset ScopeAs
    End If
End Property


Public Sub Placeat(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
Places contents of memory byte array RHS into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted.
    Dim lSize As Long
    lSize = Length
    If Width = -1 Then Width = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    If (Width > 0) And (Pointer <> 0) Then
        RtlMoveMemory ByVal Pointer + Offset, ByVal VarPtr(RHS(LBound(RHS))), Width
    ElseIf (Width > 0) Then
        Concat RHS
        Length = Width
    End If
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte() ' _
Returns a byte array of the contents exlcuded of OFfset number of bytes on the left and including of Width amount of bytes or max length.
Attribute Partial.VB_Description = "Returns a byte array of the contents exlcuded of OFfset number of bytes on the left and including of Width amount of bytes or max length."
    Dim lSize As Long
    lSize = Length
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    Dim ret() As Byte
    If (Width > 0) And (Pointer <> 0) Then
        ReDim ret(1 To Width) As Byte
        RtlMoveMemory ByVal VarPtr(ret(LBound(ret))), ByVal Pointer + Offset, Width
    Else
        ReDim ret(-1 To -1) As Byte
    End If
    Partial = ret
    Erase ret
End Function

Public Sub Prepend(ByRef RHS() As Byte) ' _
Places the byte array RHS to the left of the stream object, adding memory as nessisary ina left to right order.  Opposite of Concat()
Attribute Prepend.VB_Description = "Places the byte array RHS to the left of the stream object, adding memory as nessisary ina left to right order.  Opposite of Concat()"
    If (Length = 0) Or (Pointer = 0) Then
        Concat RHS
    Else
        Dim nSize As Long
        nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
        Length = Length + (nSize * 2)
        Pull nSize
        Placeat RHS, 0, nSize
    End If
End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Moves the contents of the memory by it's container viewport shifting to the left, truncating contents on the right, does not change the size, nor initialize newly formed memory.
Attribute Pull.VB_Description = "Moves the contents of the memory by it's container viewport shifting to the left, truncating contents on the right, does not change the size, nor initialize newly formed memory."
    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer + Offset, ByVal Pointer, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset ScopeAs
    End If
End Sub

Public Sub Push(ByVal Offset As Long) ' _
Moves the contents of the memory by it's container viewport shifting to the right, truncating contents on the left, does not change the size, nor initialize newly formed memory.
Attribute Push.VB_Description = "Moves the contents of the memory by it's container viewport shifting to the right, truncating contents on the left, does not change the size, nor initialize newly formed memory."
    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer, ByVal Pointer + Offset, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset ScopeAs
    End If
End Sub

Public Sub Reset(Optional ByRef Scopes As MemoryScopes) ' _
Resets the stream object optionally initialized to the specified Scopes.
Attribute Reset.VB_Description = "Resets the stream object optionally initialized to the specified Scopes."
    MemoryModify
    ScopeAs = Scopes
End Sub

