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

Public Enum MemoryScopes
    heap = 0
    Locking = 1
    [Local] = 2
    [Global] = 4
End Enum

Private pScope As MemoryScopes
Private Pointer As Long

Public Property Get Address(Optional ByVal Offset As Long = 0) As Long ' _
Retrieves the address numberical value where the memory this object is allocated at.
Attribute Address.VB_Description = "Retrieves the address numberical value where the memory this object is allocated at."
    RtlMoveMemory ByVal VarPtr(Address), (Pointer + Offset), 4
End Property

Public Property Get Scope() As MemoryScopes ' _
Retrieves the type of memory functionally this stream is utilizing, Heap, Locking and/or Local or Global defined by the Reset method.
Attribute Scope.VB_Description = "Retrieves the type of memory functionally this stream is utilizing, Heap, Locking and/or Local or Global defined by the Reset method."
    Scope = pScope
End Property

Private Sub Class_Initialize()
    pScope = MemoryScopes.heap
    'pScope = MemoryScopes.[Local] Or MemoryScopes.Locking
End Sub

Private Sub Class_Terminate()
    Reset pScope
End Sub

Private Sub MemoryModify(Optional ByVal Size As Long = 0)

    If (Pointer <> 0) Then
        If ((Scope And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((Scope And MemoryScopes.Local) = MemoryScopes.Local) Then
                If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
            ElseIf ((Scope And MemoryScopes.Global) = MemoryScopes.Global) Then
                If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
            ElseIf ((Scope And MemoryScopes.heap) = MemoryScopes.heap) Then
                If HeapUnlock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory unlock failure."
            End If
        End If
    End If
    
    If (Size = 0) Then
        If (Pointer <> 0) And (Not (Length = 0)) Then
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                LocalFree Pointer
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                GlobalFree Pointer
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                HeapFree GetProcessHeap, 0, ByVal Pointer
            End If
            Pointer = 0
        End If
    Else
        If (Pointer = 0) And (Length = 0) Then
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalAlloc(LMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalAlloc(modStrands.GMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                Pointer = HeapAlloc(GetProcessHeap, 0, Size)
            End If
        Else
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalReAlloc(Pointer, Size, LMEM_MOVEABLE Or lPtr)
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalReAlloc(Pointer, Size, modStrands.GMEM_MOVEABLE Or modStrands.GPTR)
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                Pointer = HeapReAlloc(GetProcessHeap, 0, ByVal Pointer, Size)
            End If
        End If
    End If

    If (Pointer <> 0) Then
        If ((Scope And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((Scope And MemoryScopes.Local) = MemoryScopes.Local) Then
                If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
            ElseIf ((Scope And MemoryScopes.Global) = MemoryScopes.Global) Then
                If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
            ElseIf ((Scope And MemoryScopes.heap) = MemoryScopes.heap) Then
                If HeapLock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory lock mismatch."
            End If
        End If
    End If

End Sub

Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, or same as a portion of the Origin strean, optionally defined by the Offset and Width.
Attribute Clone.VB_Description = "Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, or same as a portion of the Origin strean, optionally defined by the Offset and Width."
    Dim lSize As Long
    Reset Origin.Scope
    lSize = Origin.Length
    If Width = -1 Or lSize - Offset < Width Then Width = lSize - Offset
    If Offset < lSize And Offset >= 0 And Offset + Width <= lSize And Width > 0 Then
        Length = Width
        RtlMoveMemory ByVal Pointer, ByVal Origin.Address(Offset), Width
    End If
End Sub

Public Sub Post(ByRef Data As Byte) ' _
Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Concat()
Attribute Post.VB_Description = "Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Concat()"
    Length = Length + 1
    Plot Length - 1, Data
End Sub

Public Sub Concat(ByRef Data() As Byte) ' _
Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Post()
Attribute Concat.VB_Description = "Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Post()"
    Dim nSize As Long
    nSize = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
    If (nSize > 0) Then
        Dim lSize As Long
        lSize = Length
        MemoryModify (nSize + lSize)
        RtlMoveMemory ByVal Pointer + lSize, ByVal VarPtr(Data(LBound(Data))), nSize
'    Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Property Get Length() As Long ' _
Gets the amount of bytes with in this instance of the stream object., Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes.
Attribute Length.VB_Description = "Gets the amount of bytes with in this instance of the stream object., Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes."
    If Pointer <> 0 Then
        If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
            Length = LocalSize(Pointer)
        ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
            Length = GlobalSize(Pointer)
        ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
            Length = HeapSize(GetProcessHeap, 0, ByVal Pointer)
        End If
    Else
        Length = 0
    End If
End Property

Public Property Let Length(ByVal RHS As Long)
    If (RHS > 0) Then
        MemoryModify RHS
    ElseIf (RHS = 0) Then
        Reset pScope
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Property

Public Sub PlaceAt(ByRef Data() As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Places contents of memory byte array Data into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted.
Attribute PlaceAt.VB_Description = "Places contents of memory byte array Data into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted."
    Dim lSize As Long
    lSize = Length
    If Width = -1 Then Width = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
    If ((lSize - Offset) <= Width) Then Width = lSize - Offset
    If (Width > 0) And (Pointer <> 0) And lSize >= (Offset + Width) Then
        RtlMoveMemory ByVal Pointer + Offset, ByVal VarPtr(Data(LBound(Data))), Width
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Byte() ' _
Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve.
Attribute Partial.VB_Description = "Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve."

    Dim lSize As Long
    lSize = Length
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    Dim Ret() As Byte
    If (Width > 0) And (Pointer <> 0) Then
        ReDim Ret(0 To Width - 1) As Byte
        RtlMoveMemory ByVal VarPtr(Ret(LBound(Ret))), ByVal Pointer + Offset, Width
    Else
        ReDim Ret(-1 To -1) As Byte
    End If
    Partial = Ret
    Erase Ret
End Function

Public Sub Prepend(ByRef Data() As Byte) ' _
Places Data bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data.
Attribute Prepend.VB_Description = "Places Data bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data."

    If (Length = 0) Or (Pointer = 0) Then
        Concat Data
    Else
        Dim nSize As Long
        nSize = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
        Length = Length + (nSize * 2)
        Pull nSize
        PlaceAt Data, 0, nSize
    End If
End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left.
Attribute Pull.VB_Description = "Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left."

    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer + Offset, ByVal Pointer, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset pScope
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub


Public Sub Push(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
Attribute Push.VB_Description = "Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right."
    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer, ByVal Pointer + Offset, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset pScope
'    Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub


Public Sub Reset(Optional ByRef ScopeAs As MemoryScopes) ' _
Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use.
Attribute Reset.VB_Description = "Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use."
    MemoryModify
End Sub

Public Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width.
Attribute Pass.VB_Description = "Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width."
    Dim lPos As Long
    Dim bPart() As Byte
    bPart = Partial(Offset, Width)
    For lPos = LBound(bPart) To UBound(bPart)
        If bPart(lPos) = Data Then Pass = Pass + 1
    Next
End Function

Public Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width.
Attribute Poll.VB_Description = "Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width."
    If Ballot > 0 And Width <> 0 And Length > 0 Then
        Dim lPos As Long
        Dim bPart() As Byte
        If Ballot = 0 Then Exit Function
        bPart = Partial(Offset, Width)
        For lPos = LBound(bPart) To UBound(bPart)
            If bPart(lPos) = Data Then
                Poll = Poll + 1
                If Poll = Ballot Then
                    Poll = lPos + LBound(bPart)
                    Exit Function
                End If
                
            End If
        Next
        If Width > 0 Then
            Poll = Width
        Else
            Poll = Length
        End If
 '   Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Function

Public Sub Pinch(ByVal Offset As Long, ByVal Width As Long) ' _
Removes a portion of the stream defined by Offset and Width.
Attribute Pinch.VB_Description = "Removes a portion of the stream defined by Offset and Width."
    Dim lSize As Long
    lSize = Length
    If Offset < lSize And Offset > 0 And Offset + Width < lSize And Width > 0 Then
        Dim n As New Stream
        
        n.Concat Partial(0, Offset)
        n.Concat Partial(Offset + Width, lSize - (Offset + Width))
        Clone n
        
        Set n = Nothing
        
    ElseIf Offset = 0 And Width = lSize Then
        Reset
    ElseIf Offset = 0 Then
        Push Width
        'Length = Length - Width
    ElseIf Offset + Width = lSize Then
        Length = Length - Width
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Sub Pyramid(ByRef SubStream As Stream, ByVal Offset As Long, ByVal Width As Long) ' _
Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream.
Attribute Pyramid.VB_Description = "Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream."
    Dim lSize As Long
    lSize = Length
    If Offset >= 0 And Offset + Width <= lSize Then
        Dim lSize2 As Long
        lSize2 = SubStream.Length
        If lSize2 > 0 Then
            Dim n As New Stream
            If Offset > 0 Then n.Concat Partial(0, Offset)
            n.Concat SubStream.Partial
            If lSize - (Offset + Width) > 0 Then n.Concat Partial(Offset + Width, lSize - (Offset + Width))
           ' If (Offset + Width) > 0 And (Offset + Width) < lSize Then n.Concat Partial(Offset + Width)  ', lSize - (Offset + Width))
            Clone n
            Set n = Nothing
        Else
            Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
        End If
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Function Peek(Optional ByVal Offset As Long = 0) As Byte ' _
Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset.
Attribute Peek.VB_Description = "Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset."
    If Offset >= 0 And Offset < Length Then
        RtlMoveMemory ByVal VarPtr(Peek), ByVal Pointer + Offset, 1
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Function

Public Sub Plot(ByVal Offset As Long, ByVal Data As Byte) ' _
Sets a byte in the data stream immediatly after Offset.
Attribute Plot.VB_Description = "Sets a byte in the data stream immediatly after Offset."

    If Offset >= 0 And Offset < Length Then
        RtlMoveMemory ByVal Pointer + Offset, ByVal VarPtr(Data), 1
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If

End Sub
