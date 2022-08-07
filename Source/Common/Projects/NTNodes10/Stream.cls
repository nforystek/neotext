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

Implements IStream

Private Enum MemoryFunc
    Free = 0
    Alloc = 1
    Realloc = 2
End Enum

Private ScopeAs As MemoryScopes
Private Pointer As Long

Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Function Address(Optional ByVal Offset As Long = 0) As Long ' _
Retrieves the address numberical value where the memory this object is allocated at.
Attribute Address.VB_Description = "Retrieves the address numberical value where the memory this object is allocated at."
    RtlMoveMemory ByVal VarPtr(Address), (Pointer + Offset), 4
End Function

Private Sub Class_Initialize()
    vbaObjSet Me, ObjPtr(IStream)
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
                Pointer = LocalAlloc(LMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalAlloc(modStrands.GMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((ScopeAs And MemoryScopes.Heap) = MemoryScopes.Heap) Then
                Pointer = HeapAlloc(GetProcessHeap, 0, Size)
            End If
        Else
            If ((ScopeAs And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalReAlloc(Pointer, Size, LMEM_MOVEABLE Or lPtr)
            ElseIf ((ScopeAs And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalReAlloc(Pointer, Size, modStrands.GMEM_MOVEABLE Or modStrands.GPTR)
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
Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, but does not change the scope of the object cloning the Origin, to do so, use Reset before cloning.
    Reset ScopeAs
    If Origin.Length > 0 Then Concat Origin.Partial(Offset, Width)
End Sub

Public Sub Post(ByRef RHS() As Byte) ' _
Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Concat()
    Concat RHS
End Sub

Public Sub Concat(ByRef RHS() As Byte) ' _
Appends RHS bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by RHS. Psudonym to Post()

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
Gets the amount of bytes with in this instance of the stream object.
Attribute Length.VB_Description = "Gets or sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes."

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
Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes.
    If (RHS > 0) Then
        MemoryModify RHS
    ElseIf (RHS = 0) Then
        Reset ScopeAs
    End If
End Property



Public Sub PlaceAt(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) ' _
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
Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve.
Attribute Partial.VB_Description = "Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve."

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
Places RHS bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in RHS.

    If (Length = 0) Or (Pointer = 0) Then
        Concat RHS
    Else
        Dim nSize As Long
        nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
        Length = Length + (nSize * 2)
        Pull nSize
        PlaceAt RHS, 0, nSize
    End If
End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left.

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
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
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
Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use.
    MemoryModify
    ScopeAs = Scopes
End Sub

Public Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Long ' _
Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width.
    Dim lPos As Long
    Dim bPart() As Byte
    bPart = Partial(Offset, Width)
    For lPos = LBound(bPart) To UBound(bPart)
        If bPart(lPos) = Data Then Pass = Pass + 1
    Next
End Function

Public Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Long ' _
Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width.
    If Ballot > 0 And Width <> 0 And Length > 0 Then
        Dim lPos As Long
        Dim bPart() As Byte
        If Ballot = 0 Then Exit Function
        bPart = Partial(Offset, Width)
        For lPos = LBound(bPart) To UBound(bPart)
            If bPart(lPos) = Data Then
                Poll = Poll + 1
                If Poll = Ballot Then
                    Poll = lPos - LBound(bPart)
                    Exit Function
                End If
                
            End If
        Next
        If Width > 0 Then
            Poll = Width
        Else
            Poll = Length
        End If
    End If
End Function

Public Sub Pinch(ByVal Offset As Long, ByVal Width As Long) ' _
Removes a section of the stream that is in the middle having portions of the stream before and after it resulting as continuous, least the pinched Width.
    Dim lSize As Long
    lSize = Length
    If Offset < lSize And Offset > 0 And Offset + Width < lSize And Width > 0 Then
        MemoryModify (lSize + (Width * 2))
        RtlMoveMemory ByVal Pointer + Offset, ByVal Pointer + Offset + Width, (-lSize + -Width + (lSize * 2))
        MemoryModify (lSize - Width)
    End If
End Sub

Public Sub Pyramid(ByRef SubStream As Stream, ByVal Offset As Long, ByVal Width As Long) ' _
Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream.
    Dim lSize As Long
    lSize = Length
    If Offset >= 0 And Offset + Width <= lSize Then
        Dim lSize2 As Long
        lSize2 = SubStream.Length
        If lSize2 > 0 Then
            If (Width - lSize2) > 0 Then
                MemoryModify (lSize + Width + lSize2)
                RtlMoveMemory ByVal Pointer + Offset, ByVal SubStream.Address, lSize2
                RtlMoveMemory ByVal Pointer + (Offset + lSize2), ByVal Pointer + (Offset + lSize2) + (Width - lSize2), lSize
                MemoryModify (lSize - (Width - lSize2))
            ElseIf (Width - lSize2) < 0 Then
                MemoryModify (lSize + Width + lSize2) + (Width - lSize2)
                RtlMoveMemory ByVal Pointer + (Offset + lSize2), ByVal Pointer + (Offset + lSize2) + (Width - lSize2), lSize
                RtlMoveMemory ByVal Pointer + Offset, ByVal SubStream.Address, lSize2
                MemoryModify (lSize - (Width - lSize2))
            Else
                RtlMoveMemory ByVal Pointer + Offset, ByVal SubStream.Address, lSize2
            End If
        End If
    End If
End Sub

Public Function Peek(Optional ByVal Offset As Long = 0) As Byte ' _
Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset.
    If Offset >= 0 And Offset < Length Then
        RtlMoveMemory ByVal VarPtr(Peek), ByVal Pointer + Offset, 1
    End If
End Function

Private Function IStream_Address(Optional ByVal Offset As Long = 0&) As Long
    IStream_Address = Me.Address(Offset)
End Function

Private Sub IStream_Clone(Origin As Stream, Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&)
    Me.Clone Origin, Offset, Width
End Sub

Private Sub IStream_Concat(RHS() As Byte)
    Me.Concat RHS
End Sub

Private Property Let IStream_Length(ByVal RHS As Long)
    Me.Length = RHS
End Property

Private Property Get IStream_Length() As Long
    IStream_Length = Me.Length
End Property

Private Function IStream_Partial(Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&) As Byte()
    IStream_Partial = Me.Partial(Offset, Width)
End Function

Private Function IStream_Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&) As Long
    IStream_Pass = Me.Pass(Data, Offset, Width)
End Function

Private Sub IStream_Pinch(ByVal Offset As Long, ByVal Width As Long)
    Me.Pinch Offset, Width
End Sub

Private Sub IStream_PlaceAt(RHS() As Byte, Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&)
    Me.PlaceAt RHS, Offset, Width
End Sub

Private Function IStream_Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&) As Long
    IStream_Poll = Me.Poll(Data, Ballot, Offset, Width)
End Function

Private Function IStream_Peek(Optional ByVal Offset As Long = 0&) As Byte
    IStream_Peek = Me.Peek(Offset)
End Function

Private Sub IStream_Post(RHS() As Byte)
    Me.Post RHS
End Sub

Private Sub IStream_Prepend(RHS() As Byte)
    Me.Prepend RHS
End Sub

Private Sub IStream_Pull(ByVal Offset As Long)
    Me.Pull Offset
End Sub

Private Sub IStream_Push(ByVal Offset As Long)
    Me.Push Offset
End Sub

Private Sub IStream_Pyramid(SubStream As IStream, ByVal Offset As Long, ByVal Width As Long)
    Me.Pyramid SubStream, Offset, Width
End Sub

Private Sub IStream_Reset(Optional Scopes As MemoryScopes)
    Me.Reset ScopeAs
End Sub

