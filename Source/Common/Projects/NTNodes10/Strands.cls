VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Compare Binary

Implements Stream

Public Enum Behaviors

    scopeNormal = 0
    ScopeLocale = 4
    ScopeGlobal = -4

    Nothings_Initialized = -7 And 7
End Enum

'Private ScopeAs As MemoryScopes
'Private Pointer As Long 'stream data nothing checks agaiinst it
Private Breaks() As Long 'array ubound is count of lines, value
'0 element is current line number, other elements are locations
'of the line numebrs linefeed character in the stream, by total
'thus element n's value, subtracting n-1's value is a line size
Private Behave As Behaviors
Private PathHandle As Long
Private PathFileName As String

Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Private Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Private Sub Class_Initialize()
    vbaObjSetAddref Stream, ObjPtr(New Stream)
    
    PathFileName = MirrorPath_None
    Behave = Nothings_Initialized
    ReDim Breaks(0 To 0) As Long
    Change scopeNormal
End Sub

Private Sub Class_Terminate()
    Terminate
    vbaObjSet Stream, ObjPtr(Nothing)
End Sub

Private Sub Terminate()
    If Behave <> Nothings_Initialized Then
        If Not PathHandle = 0 Then
            CloseHandle PathHandle
            PathHandle = 0
            If Left(PathFileName, 6) = MirrorPath_Temp Then Kill Mid(PathFileName, 7)
        End If
        PathFileName = MirrorPath_None
        
'        Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.ScopeNormal)
'            Case Behaviors.ScopeGlobal
'                'GlobalUnlock Pointer
'                GlobalFree Pointer
'            Case Behaviors.ScopeLocale
'                'LocalUnlock Pointer
'                LocalFree Pointer
'            Case Else
'                HeapFree GetProcessHeap, 0, Pointer
'        End Select
        ReDim Breaks(0 To 0) As Long
        Behave = Nothings_Initialized
        
    End If
End Sub

Public Property Get Behavior() As Behaviors ' _
Gets or sets where in scope of memory the strands will be allocated, resetting the object to of of the values; heap (normal), locale, or global.
Attribute Behavior.VB_Description = "Gets or sets where in scope of memory the strands will be allocated, resetting the object to of of the values; heap (normal), locale, or global."
    Behavior = Behave
End Property
Public Property Let Behavior(ByVal RHS As Behaviors)
    If Behave <> RHS Then
        Dim pfn As String
        pfn = PathFileName
        Change RHS, pfn
    End If
End Property

Public Property Get Filename() As String ' _
Gets or sets the filename path that the objects to be mirrored of, any text file lines are loaded, and object edits are real time written.
Attribute Filename.VB_Description = "Gets or sets the filename path that the objects to be mirrored of, any text file lines are loaded, and object edits are real time written."
    Filename = PathFileName
End Property
Public Property Let Filename(ByVal RHS As String)
    Dim bh As Behaviors
    bh = Behave
    Change bh, RHS
End Property

Private Sub Change(ByVal Behavior As Behaviors, Optional ByVal MirrorFilePath As String = MirrorPath_None)
    Terminate
    If Behave = Nothings_Initialized Then
        Dim sc As SECURITY_ATTRIBUTES
        Dim lSize As Long
        ReDim Breaks(0 To 0) As Long
        lSize = 0
        Select Case Behavior And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
            Case Behaviors.ScopeGlobal
                Stream.Reset MemoryScopes.Global
            Case Behaviors.ScopeLocale
                Stream.Reset MemoryScopes.Local
            Case Else
                Stream.Reset MemoryScopes.Heap
        End Select
        Behave = Behavior
        Select Case Left(MirrorFilePath, 6)
            Case MirrorPath_Temp
                PathFileName = MirrorPath_Temp & GetTemporaryFile
                PathHandle = CreateFile(Mid(PathFileName, 7), FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY And FILE_FLAG_DELETE_ON_CLOSE, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
            Case MirrorPath_None
                PathHandle = 0
                PathFileName = MirrorPath_None
            Case Else
                PathFileName = MirrorFilePath
                PathHandle = CreateFile(PathFileName, FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
                lSize = Length
                If lSize = 0 Then
                    Dim loorder As Long
                    Dim hiorder As Long
                    loorder = modStrands.GetFileSize(PathHandle, hiorder)
                    Dim nSize As Long
                    LoWord(nSize) = loorder
                    HiWord(nSize) = hiorder
                    If (nSize <> lSize Or lSize = 0) And (nSize > 0) Then
                        Dim Data() As Byte
                        Dim ol As OVERLAPPED
                        ol.Offset = LoWord(0)
                        ol.OffsetHigh = HiWord(0)
                        Dim m As MemInfo
                        Do
                            ReDim Preserve Data(0 To nSize - 1) As Byte
                            m = lArrayToANSI(Data)
                            If hReadFile(PathHandle, ByVal m.MemH, nSize, lSize, ol) = 0 Then nSize = 0
                            If (lSize = nSize) And (lSize > 0) Then
                                nSize = nSize - lSize
                                Data = lArrayOfANSI(m)
                                ReDim Preserve Data(0 To lSize - 1) As Byte
                                lSize = ConcatNoFile(Data, lSize)
                                For loorder = LBound(Data) To UBound(Data)
                                    lSize = lSize + 1
                                    If Data(loorder) = 10 Then
                                        ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
                                        Breaks(UBound(Breaks)) = lSize
                                    End If
                                Next
                                If Data(UBound(Data)) <> 10 Then
                                    lSize = lSize + 1
                                    ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
                                    Breaks(UBound(Breaks)) = lSize
                                End If
                            End If
                            lDestroyANSI m
                        Loop Until nSize = 0
                        Erase Data
                    End If
                End If
        End Select
    Else
        Err.Raise 8, App.Title, "Object already initialized, the Behavior property can not be changed once set for this object."
    End If
End Sub

'###############################################################################
'###############################################################################

Public Function Append(ByRef lines() As Byte) As Long ' _
Adds lines to the end of the strands collection, returning the numer of new lines added.
Attribute Append.VB_Description = "Adds lines to the end of the strands collection, returning the numer of new lines added."
    If Behave <> Nothings_Initialized Then
        Dim llen As Long
        Dim Temp As Long
        Dim line() As Byte
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(lines, line, llen)
            Concat2 line
            ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
            Breaks(UBound(Breaks)) = Stream.Length
            Append = Append + 1
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Sub Clear() ' _
Clears all the contents of the strands collection.
Attribute Clear.VB_Description = "Clears all the contents of the strands collection."
    If Behave <> Nothings_Initialized Then
        ReDim Breaks(0 To 0) As Long
        Breaks(0) = 0
        Reset
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Property Get Count() As Long ' _
Gets the number of lines in the strands collection.
Attribute Count.VB_Description = "Gets the number of lines in the strands collection."
    If Behave <> Nothings_Initialized Then
        Count = UBound(Breaks)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Get Current() As Long ' _
Gets or sets the current line number for which Item, Insert or Delete refers too.
Attribute Current.VB_Description = "Gets or sets the current line number for which Item, Insert or Delete refers too."
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            If (Breaks(0) = 0) Then Breaks(0) = 1
            If (Breaks(0) > Count) Then Breaks(0) = Count
            Current = Breaks(0)
        Else
            Current = 0
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let Current(ByVal RHS As Long)
    If Behave <> Nothings_Initialized Then
    
        If ((RHS <= UBound(Breaks)) And (UBound(Breaks) > 0)) Or (RHS = 0) Then
            Breaks(0) = RHS
        ElseIf RHS > UBound(Breaks) Then
            Breaks(0) = UBound(Breaks)
        ElseIf RHS < 0 Then
            Breaks(0) = IIf(UBound(Breaks) > 0, 1, 0)
        'else
            'Err.Raise 8, App.Title, "Line number invalid or the collection is empty."
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Function Delete() As Byte() ' _
Deletes a line from the collection at the number indicated by the Current property.
Attribute Delete.VB_Description = "Deletes a line from the collection at the number indicated by the Current property."
    If Behave <> Nothings_Initialized Then
        If (Count <= 2) Or (Current = 1) Or (Current = Count) Then
            If (Current = 1) Then
                Delete = Pop()
            Else
                Delete = Remove()
            End If
        Else
            Dim llen As Long
            Dim Temp As Long
            Dim part As Stream
            Temp = Current
            Breaks(0) = 0
            llen = (Breaks(Temp) - Breaks(Temp - 1))
            Delete = Stream.Partial(Breaks(Temp - 1), llen)
            Set part = New Stream
            part.Concat Stream.Partial(Breaks(Temp))
            Stream.Length = Breaks(Temp - 1)
            Concat2 part.Partial
            Set part = Nothing
            Dim cnt As Long
            For cnt = Temp To UBound(Breaks) - 1
                Breaks(cnt) = (Breaks(cnt) - llen) + (Breaks(cnt + 1) - Breaks(cnt))
            Next
            ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) - 1) As Long
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function


Public Property Get First() As Byte() ' _
Gets or sets the bytes of the text to the first line of the strand collection.
Attribute First.VB_Description = "Gets or sets the bytes of the text to the first line of the strand collection."
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then First = Item(1)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let First(ByRef lines() As Byte)
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim llen As Long
            Dim cnt As Long
            Dim line() As Byte
            Do While NextLine(lines, line, llen)
                If (Count > 1) Then
                    For cnt = UBound(Breaks) To 2 Step -1
                        Breaks(cnt) = Breaks(cnt) - (Breaks(1) - llen)
                    Next
                End If
                Push2 Breaks(1)
                Prepend2 line
                Breaks(1) = llen
            Loop
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Function Forward() As Byte() ' _
Increments the Current property to the next line in the strand collection returning the bytes for the text of the line.
Attribute Forward.VB_Description = "Increments the Current property to the next line in the strand collection returning the bytes for the text of the line."
    If Behave <> Nothings_Initialized Then
        If (Breaks(0) < UBound(Breaks)) Then Breaks(0) = Breaks(0) + 1
        Forward = Item(Breaks(0))
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Function Insert(ByRef lines() As Byte) As Long ' _
Adds lines to the strand collection at the line that the Current property indicates, returning the last line number added.
Attribute Insert.VB_Description = "Adds lines to the strand collection at the line that the Current property indicates, returning the last line number added."
    If Behave <> Nothings_Initialized Then
        If (Count = 0) Or (Current = 1) Or (Current = Count) Then
            If (Current = 1) Then
                Insert = Add(lines)
            Else
                Insert = Append(lines)
            End If
        Else
            Dim llen As Long
            Dim Temp As Long
            Dim line() As Byte
            Dim cnt As Long
            Dim part As Stream
            Temp = Current
            Breaks(0) = 0
            Do While NextLine(lines, line, llen)
                Set part = New Stream
                part.Concat Stream.Partial(Breaks(Temp - 1))
                Stream.Length = Breaks(Temp - 1)
                Concat2 line
                Concat2 part.Partial
                Set part = Nothing
                ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) + 1) As Long
                For cnt = UBound(Breaks) To Temp Step -1
                    Breaks(cnt) = Breaks(cnt - 1) + llen
                Next
            Loop
            Breaks(0) = Temp
            Insert = Breaks(0)
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Property Get Offset(Optional ByVal Number As Long = -1&) As Long ' _
Gets the number of bytes existing before the line expressed by the Number argument, if omitted, the Current property is used.
    If Behave <> Nothings_Initialized Then
        If ((Current > 0) And (Current <= UBound(Breaks))) And (Number = -1) Then
            Number = Breaks(0)
        End If
        If (UBound(Breaks) > 0) Then
            If (Number <= UBound(Breaks)) Then
                If (Number > 1) Then
                    Offset = Breaks(Number - 1)
                End If
            End If
      '  Else
           ' Err.Raise 8, App.Title, "The number, or current line number, is invalid or no line exist."
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Get SizeOf(Optional ByVal Number As Long = -1&) As Long ' _
Gets the number of bytes in the line expressed by the Number argument, if omitted, the Current property is used.
    If Behave <> Nothings_Initialized Then
        If ((Current > 0) And (Current <= UBound(Breaks))) And (Number = -1) Then
            Number = Breaks(0)
        End If
        If (UBound(Breaks) > 0) Then
            If (Number <= UBound(Breaks)) Then
                If (Number > 1) Then
                    SizeOf = Breaks(Number) - Breaks(Number - 1)
                ElseIf (Number = 1) Then
                    SizeOf = Breaks(Number)
                End If
            End If
       ' Else
          '  Err.Raise 8, App.Title, "The number, or current line number, is invalid or no line exist."
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Get Item(Optional ByVal Number As Long = -1&) As Byte() ' _
Gets or sets the bytes for the line expressed by the Number argument, if omitted, the Current property is used. Upon setting, the current property is changed to reflect Number.
Attribute Item.VB_Description = "Gets or sets the bytes for the line expressed by the Number argument, if omitted, the Current property is used. Upon setting, the current property is changed to reflect Number."
    If Behave <> Nothings_Initialized Then
        If ((Current > 0) And (Current <= UBound(Breaks))) And (Number = -1) Then
            Number = Breaks(0)
        End If
        If (UBound(Breaks) > 0) Then
            If (Number <= UBound(Breaks)) Then
                If (Number > 1) Then
                    Item = Stream.Partial(Breaks(Number - 1), Breaks(Number) - Breaks(Number - 1))
                ElseIf (Number = 1) Then
                    Item = Stream.Partial(0, Breaks(Number))
                End If
            Else
                ReDim Item(-1 To -1) As Byte
            End If
        Else
           ' Err.Raise 8, App.Title, "The number, or current line number, is invalid or no line exist."
            ReDim Item(-1 To -1) As Byte
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let Item(Optional ByVal Number As Long = -1&, ByRef lines() As Byte)
    If Behave <> Nothings_Initialized Then
        If (Count > 0) And Number >= -1 Then
            If (Number = 0) Then Number = 1
            If (Number > Count) Then Number = Count
        End If
        If Number = -1 Then Number = Current
        If (Count > 0) Then
            Dim Temp As Long
            Temp = Breaks(0)
            Breaks(0) = Number
            Delete
            Breaks(0) = Number
            Insert lines
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Get Last() As Byte() ' _
Gets or sets the bytes for the text of the last line in the strands collection.
Attribute Last.VB_Description = "Gets or sets the bytes for the text of the last line in the strands collection."
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then Last = Item(Count)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let Last(ByRef lines() As Byte)
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim llen As Long
            Dim Temp As Long
            Dim line() As Byte
            Temp = Current
            Breaks(0) = 0
            Do While NextLine(lines, line, llen)
                Stream.Length = Stream.Length - (Breaks(Count) - Breaks(Count - 1))
                Concat2 line
                Breaks(Count) = Breaks(Count - 1) + llen
            Loop
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Get Value() As Byte() ' _
Gets or sets the bytes of the entire strand collections text for which it is representing lines of.
Attribute Value.VB_Description = "Gets or sets the bytes of the entire strand collections text for which it is representing lines of."
    If Behave <> Nothings_Initialized Then
        Value = Stream.Partial
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let Value(ByRef lines() As Byte)
    If Behave <> Nothings_Initialized Then
        Dim line() As Byte
        Clear
        Dim llen As Long
        Dim Temp As Long
        ReDim Preserve Breaks(0 To 0) As Long
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(lines, line, llen)
            Concat2 line
            ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
            Breaks(UBound(Breaks)) = Breaks(UBound(Breaks) - 1) + llen
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Function Previous() As Byte() ' _
Decrements the Current property to the pervious line in the strand collection returning the bytes for the text of the line.
Attribute Previous.VB_Description = "Decrements the Current property to the pervious line in the strand collection returning the bytes for the text of the line."
    If Behave <> Nothings_Initialized Then
        If (Breaks(0) > 1) Then Breaks(0) = Breaks(0) - 1
        Previous = Item(Breaks(0))
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Function Pop() As Byte() ' _
Removes a line at the very beginning of the strand collection, and returns its text in bytes.
Attribute Pop.VB_Description = "Removes a line at the very beginning of the strand collection, and returns its text in bytes."
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim Temp As Long
            Temp = Current
            Pop = Stream.Partial(0, Breaks(1))
            Push2 Breaks(1)
            If UBound(Breaks) > 0 Then
                Breaks(0) = Breaks(1)
                If UBound(Breaks) >= 2 Then
                    Dim cnt As Long
                    For cnt = 1 To UBound(Breaks) - 1
                        Breaks(cnt) = Breaks(cnt + 1) - Breaks(0)
                    Next
                End If
                ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) - 1) As Long
            End If
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Function Add(ByRef lines() As Byte) As Long ' _
Adds lines by bytes to the very beginning of the strand collection returning the number of lines added.
Attribute Add.VB_Description = "Adds lines by bytes to the very beginning of the strand collection returning the number of lines added."
    If Behave <> Nothings_Initialized Then
        Dim cnt As Long
        Dim Temp As Long
        Dim llen As Long
        Dim line() As Byte
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(lines, line, llen)
            Prepend2 line
            ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) + 1) As Long
            Add = Add + 1
            If Count > 1 Then
                For cnt = UBound(Breaks) To 1 Step -1
                    Breaks(cnt) = Breaks(cnt - 1) + llen
                Next
            Else
                Breaks(1) = llen
            End If
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Function Remove() As Byte() ' _
Removes a line at the very end of the strand collection, and returns its text in bytes.
Attribute Remove.VB_Description = "Removes a line at the very end of the strand collection, and returns its text in bytes."
    If Behave <> Nothings_Initialized Then
        If Count > 1 Then
            Remove = Stream.Partial(Breaks(UBound(Breaks) - 1), (Breaks(UBound(Breaks)) - Breaks(UBound(Breaks) - 1)))
            Length = Length - (Breaks(UBound(Breaks)) - Breaks(UBound(Breaks) - 1))
            ReDim Preserve Breaks(0 To UBound(Breaks) - 1) As Long
        ElseIf Count = 1 Then
            Remove = Stream.Partial
            Reset
            ReDim Breaks(0 To 0) As Long
            Breaks(0) = 0
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

'###############################################################################
'###############################################################################

Public Property Get Address(Optional ByVal Offset As Long = 0) As Long ' _
Retrieves the address numberical value where the memory this object is allocated at.
    Address = Stream.Address(Offset)
End Property

Public Property Get Scope() As MemoryScopes ' _
Retrieves the type of memory functionally this stream is utilizing, Heap, Locking and/or Local or Global defined by the Reset method.
    Scope = Stream.Scope
End Property

Public Sub Reset(Optional ScopeAs As MemoryScopes) ' _
Reset the stream object freeing up any allocated memory changing the Length to reflect zero.
Attribute Reset.VB_Description = "Reset the stream object freeing up any allocated memory changing the Length to reflect zero."
    If ScopeAs And MemoryScopes.Local = MemoryScopes.Local Then
        Me.Behavior = ScopeGlobal
    ElseIf ScopeAs And MemoryScopes.Global = MemoryScopes.Global Then
        Me.Behavior = ScopeLocale
    Else
        Me.Behavior = scopeNormal
    End If
    If Behave <> Nothings_Initialized Then
        Stream.Reset ScopeAs
        If PathHandle <> 0 Then
            CloseHandle PathHandle
            PathHandle = 0
            If Left(PathFileName, 6) = MirrorPath_Temp And Mid(PathFileName, 7) <> "" Then Kill Mid(PathFileName, 7)
        End If
        Dim sc As SECURITY_ATTRIBUTES
        Select Case Left(PathFileName, 6)
            Case MirrorPath_Temp
                PathFileName = MirrorPath_Temp & GetTemporaryFile
                PathHandle = CreateFile(Mid(PathFileName, 7), FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY And FILE_FLAG_DELETE_ON_CLOSE, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
            Case MirrorPath_None
                PathHandle = 0
            Case Else
                PathHandle = CreateFile(PathFileName, FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL And FILE_FLAG_WRITE_THROUGH, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
        End Select
        EndOfStream 0
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Property Get Length() As Long ' _
Gets or sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes.
Attribute Length.VB_Description = "Gets or sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes."
    If Behave <> Nothings_Initialized Then
        Length = Stream.Length
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Public Property Let Length(ByVal RHS As Long)
    If Behave <> Nothings_Initialized Then
        If (RHS > -1) Then
            Stream.Length = RHS
            EndOfStream RHS
            RebreakStream
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Property

Private Sub Push2(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
    If Behave <> Nothings_Initialized Then
        Stream.Push Offset
        Dim lSize As Long
        lSize = Length
        If lSize > 0 Then WriteToStream Partial, 0, lSize
        EndOfStream lSize
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Sub Push(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
    Push2 Offset
    RebreakStream
End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left.
Attribute Pull.VB_Description = "Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left."
    If Behave <> Nothings_Initialized Then
        Stream.Pull Offset
        EndOfStream Length
        RebreakStream
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Private Sub Prepend2(ByRef Data() As Byte) ' _
Places Data() bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data().
    If Behave <> Nothings_Initialized Then
        Stream.Prepend Data
        Dim lSize As Long
        lSize = Length
        If lSize > 0 Then WriteToStream Partial, 0, lSize
        EndOfStream lSize
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Sub Prepend(ByRef Data() As Byte) ' _
Places Data() bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data().
Attribute Prepend.VB_Description = "Places Data() bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data()."
    Prepend2 Data
    RebreakStream
End Sub

Public Sub Clone(ByRef Origin As Strands, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, or same as a portion of the Origin strean, optionally defined by the Offset and Width.
    Dim lSize As Long
    Behavior = Origin.Behavior
    Stream.Clone Origin, Offset, Width
    If Length > 0 Then WriteToStream Partial, 0, Width
    EndOfStream Width
    RebreakStream
End Sub

Public Sub Post(ByRef Data As Byte) ' _
Appends Data byte onto the stream at the end of the stream increasing it's size by one. Psudonym to Concat()
    If Behave <> Nothings_Initialized Then
        Stream.Post Data
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Private Sub Concat2(ByRef Data() As Byte) ' _
Appends Data() bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data().
    If Behave <> Nothings_Initialized Then
        Dim nSize As Long
        Dim lSize As Long
        lSize = Length
        nSize = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
        If nSize > 0 Then
            Stream.Concat Data
            WriteToStream Data, lSize, nSize
            EndOfStream (lSize + nSize)
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Sub Concat(ByRef Data() As Byte) ' _
Appends Data() bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data().
Attribute Concat.VB_Description = "Appends Data() bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data()."
    Concat2 Data
    RebreakStream
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Byte() ' _
Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve.
Attribute Partial.VB_Description = "Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve."
    If Behave <> Nothings_Initialized Then
        Partial = Stream.Partial(Offset, Width)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function
Public Sub PlaceAt(ByRef Data() As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Places Width many bytes of Data() in the stream just after Offset amount of bytes, does not alter the stream size, nor place more then Data() amounts of bytes.
Attribute PlaceAt.VB_Description = "Places Width many bytes of Data() in the stream just after Offset amount of bytes, does not alter the stream size, nor place more then Data() amounts of bytes."
    If Behave <> Nothings_Initialized Then
        Stream.PlaceAt Data, Offset, Width
        If Width > 0 Then WriteToStream Data, Offset, Width
        RebreakStream
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width.
    If Behave <> Nothings_Initialized Then
        Pass = Stream.Pass(Data, Offset, Width)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width.
    If Behave <> Nothings_Initialized Then
        Poll = Stream.Poll(Data, Ballot, Offset, Width)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Sub Pinch(ByVal Offset As Long, ByVal Width As Long) ' _
Removes a section of the stream that is in the middle having portions of the stream before and after it resulting as continuous, least the pinched Width.
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        lSize = Length
        Stream.Pinch Offset, Width
        If Length - Offset > 0 Then WriteToStream Partial(Offset, Length - Offset), Offset, Length - Offset
        EndOfStream Length
        RebreakStream
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Sub Pyramid(ByRef SubStream As Strands, ByVal Offset As Long, ByVal Width As Long) ' _
Topples a SubStream, which were initially an exact Partial reflection of this base stream defined by where were Offset and Width, reflecting the smaller substream's changes into this base stream.
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        lSize = Length
        Dim lSize2 As Long
        lSize2 = SubStream.Length
        
        Stream.Pyramid SubStream, Offset, Width
                
        If Length - Offset > 0 Then WriteToStream Partial(Offset, Length - Offset), Offset, Length - Offset
        EndOfStream Length
        RebreakStream

    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

Public Function Peek(Optional ByVal Offset As Long = 0) As Byte ' _
Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset.
    If Behave <> Nothings_Initialized Then
        Peek = Stream.Peek(Offset)
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Public Sub Plot(ByVal Offset As Long, ByVal Data As Byte)   ' _
Sets a byte in the data stream, optionally excluding Offset number of bytes setting the immediate byte after Offset.
    If Behave <> Nothings_Initialized Then
        Stream.Plot Offset, Data
        WriteToStream Partial(Offset, 1), Offset, 1
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Sub

'###############################################################################
'###############################################################################


'###############################################################################
'the below functions are private and used to be adaptable to the above stranding
'functions in such a way they they are convieniently obtaining all possible nums
'that are based on the provided information and return to from the object relate
'###############################################################################
Public Sub RebreakStream()
    Dim cnt As Long
    Dim cnt2 As Long
    Dim Data() As Byte
    Data = Partial()
    cnt2 = 1
    For cnt = LBound(Data) To UBound(Data)
        If Data(cnt) = Asc(vbLf) Then
            If cnt2 > UBound(Breaks) Then ReDim Preserve Breaks(0 To cnt2) As Long
            Breaks(cnt2) = cnt + -CInt(Not CBool(-LBound(Data)))
            cnt2 = cnt2 + 1
        End If
    Next
    If UBound(Breaks) > (cnt2 - 1) Then ReDim Preserve Breaks(0 To (cnt2 - 1)) As Long
    If Breaks(UBound(Breaks)) < Length Then
        ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
        Breaks(UBound(Breaks)) = Length '+ -CInt(Not CBool(-LBound(Data)))
    End If
End Sub

Private Function ConcatNoFile(Data() As Byte, ByVal nSize As Long) As Long
    If Behave <> Nothings_Initialized Then
        If (nSize > 0) And UBound(Data) > -1 Then
            Dim lSize As Long
            lSize = Stream.Length
            Stream.Length = (nSize + lSize)
'            Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
'                Case Behaviors.ScopeGlobal
'                    'If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
'                    Stream.Address = GlobalReAlloc(Stream.Address, (nSize + lSize), modStrands.GMEM_MOVEABLE Or modStrands.GPTR)
'                    'If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
'                Case Behaviors.ScopeLocale
'                    'If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
'                    Stream.Address = LocalReAlloc(Stream.Address, (nSize + lSize), LMEM_MOVEABLE Or lPtr)
'                    'If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
'                Case Else
'                    Stream.Address = HeapReAlloc(GetProcessHeap, 0, Stream.Address, (nSize + lSize))
'            End Select

            RtlMoveMemory ByVal Stream.Address + lSize, ByVal VarPtr(Data(LBound(Data))), nSize
            ConcatNoFile = lSize
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, the Behavior property must be set first to use this object."
    End If
End Function

Private Sub WriteToStream(ByRef Data() As Byte, ByVal Offset As Long, ByVal Width As Long)
    Select Case Left(PathFileName, 6)
        Case MirrorPath_None
        Case Else
            If Width > 0 And (UBound(Data) > -1) Then
                Dim pos As Long
                Dim ol As OVERLAPPED
                pos = Offset
                ol.Offset = LoWord(pos)
                ol.OffsetHigh = HiWord(pos)
                Offset = 0
                Do
                    hWriteFile PathHandle, ByVal nArrayToString(Data), Width, Offset, ol
                    Width = Width - Offset
                    pos = pos + Offset
                    ol.Offset = LoWord(pos)
                    ol.OffsetHigh = HiWord(pos)
                    Offset = Width
                Loop Until Width = 0
            End If
    End Select
End Sub
Private Sub EndOfStream(ByVal Size As Long)
    Select Case Left(PathFileName, 6)
        Case MirrorPath_None
        Case Else
            Dim loorder As Long
            Dim hiorder As Long
            loorder = LoWord(Size)
            hiorder = HiWord(Size)
            If SetFilePointer(PathHandle, loorder, hiorder, 0) = INVALID_SET_FILE_POINTER Then
                Err.Raise 8, App.Title, "Error moving file position."
            End If
            If SetEndOfFile(PathHandle) = 0 Then
                Err.Raise 8, App.Title, "Error allocating file size."
            End If
    End Select
End Sub
Private Function LineLength(ByRef Data() As Byte) As Long
    LineLength = (UBound(Data) + -CInt(Not CBool(-LBound(Data))))
    If Data(UBound(Data)) <> 10 Then
        ReDim Preserve Data(LBound(Data) To UBound(Data) + 1) As Byte
        Data(UBound(Data)) = 10
        LineLength = LineLength + 1
    End If
End Function

Private Function CRLFCheck() As Boolean
    CRLFCheck = True
    If Length > 0 And UBound(Breaks) > 1 Then
        Dim Char() As Byte
        Dim cnt As Long
        For cnt = 1 To UBound(Breaks)
            Char = Partial(Breaks(cnt) - 1, 1)
            If Char(LBound(Char)) <> 10 Then
                Err.Raise 8, App.Title, "Failed CLF checksum."
                Exit Function
            End If
        Next
    End If
End Function

Private Function NextLine(ByRef lines() As Byte, ByRef line() As Byte, ByRef llen As Long) As Boolean
    Static inLines() As String
    Static atLine As Long
    Static AtState As Long
    If AtState = 0 Then ReDim inLines(0 To 0) As String
    Dim str As String
    If atLine > UBound(inLines) And AtState > 0 Then
        atLine = 0
        AtState = 0
        Erase inLines
        Erase line
    Else
        Select Case AtState
            Case 0
                atLine = 0
                Erase inLines
                str = Replace(modStrands.Convert(lines), Chr(13), "")
                inLines = Split(str, Chr(10))
                atLine = LBound(inLines)
                Erase line
                If UBound(inLines) > -1 Then
                    line = modStrands.Convert(inLines(atLine) & vbLf)
                    llen = LineLength(line)
                    NextLine = True
                End If
                AtState = 1
            Case 1
                If (atLine >= LBound(inLines)) And (atLine < UBound(inLines)) Then
                    If atLine <= UBound(inLines) Then
                        line = modStrands.Convert(inLines(atLine) & vbLf)
                        llen = LineLength(line)
                        NextLine = True
                    Else
                        AtState = 0
                    End If
                Else
                    AtState = 0
                End If
        End Select
        atLine = atLine + 1
    End If
End Function

Private Property Get Stream_Address(Optional ByVal Offset As Long = 0&) As Long
    Stream_Address = Me.Address(Offset)
End Property

Private Sub Stream_Clone(Origin As Stream, Optional ByVal Offset As Long = 0&, Optional ByRef Width As Long = -1&)
    Me.Clone Origin, Offset, Width
End Sub

Private Sub Stream_Concat(Data() As Byte)
    Me.Concat Data
End Sub

Private Property Let Stream_Length(ByVal Data As Long)
    Me.Length = Data
End Property

Private Property Get Stream_Length() As Long
    Stream_Length = Me.Length
End Property

Private Function Stream_Partial(Optional ByVal Offset As Long = 0&, Optional ByRef Width As Long = -1&) As Byte()
    Stream_Partial = Me.Partial(Offset, Width)
End Function

Private Function Stream_Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0&, Optional ByRef Width As Long = -1&) As Long
    Stream_Pass = Me.Pass(Data, Offset, Width)
End Function

Private Sub Stream_Pinch(ByVal Offset As Long, ByVal Width As Long)
    Me.Pinch Offset, Width
End Sub

Private Sub Stream_PlaceAt(Data() As Byte, Optional ByVal Offset As Long = 0&, Optional ByRef Width As Long = -1&)
    Me.PlaceAt Data, Offset, Width
End Sub

Private Sub Stream_Plot(ByVal Offset As Long, ByVal Data As Byte)
    Me.Plot Offset, Data
End Sub

Private Function Stream_Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0&, Optional ByRef Width As Long = -1&) As Long
    Stream_Poll = Me.Poll(Data, Ballot, Offset, Width)
End Function

Private Function Stream_Peek(Optional ByVal Offset As Long = 0&) As Byte
    Stream_Peek = Me.Peek(Offset)
End Function

Private Sub Stream_Post(Data As Byte)
    Me.Post Data
End Sub

Private Sub Stream_Prepend(Data() As Byte)
    Me.Prepend Data
End Sub

Private Sub Stream_Pull(ByVal Offset As Long)
    Me.Pull Offset
End Sub

Private Sub Stream_Push(ByVal Offset As Long)
    Me.Push Offset
End Sub

Private Sub Stream_Pyramid(SubStream As Stream, ByVal Offset As Long, ByVal Width As Long)
    Me.Pyramid SubStream, Offset, Width
End Sub

Private Sub Stream_Reset(Optional ScopeAs As MemoryScopes)
    Me.Reset ScopeAs
End Sub

Private Property Get Stream_Scope() As MemoryScopes
    Stream_Scope = Me.Scope
End Property
