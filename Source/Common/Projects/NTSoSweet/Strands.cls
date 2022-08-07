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
'TOP DOWN


Public Enum Behaviors
    StateAccess = 0 'Default behavior of ram construct conductly
    StateMirror = -2 'Coupled behavior of direct to file access

    RemmitSioux = 0 'disserning it's creation is not read nor write
    'memory read implies removing upon revisiting and/or touched
    'with uri file location this is not equivelent MemoryAccess
    CommitFirms = 1 'commits the creating so future access keeps
    'this can be in memoryaccess and/or mirroraccess sequential

    scopeNormal = 0
    ScopeLocale = 4
    ScopeGlobal = -4

    Memory_Remmit = Behaviors.StateAccess And Behaviors.RemmitSioux '0 'virtualmemory and heap Sioux to scope
    Mirror_Remmit = Behaviors.StateMirror And Behaviors.RemmitSioux '-2 'direct file and heap Sioux to scope

    Memory_Commit = Behaviors.StateAccess And Behaviors.CommitFirms ' 1 'virtualmemory and heap Firms to scope
    Mirror_Commit = Behaviors.StateMirror And Behaviors.CommitFirms '-1 'direct file and heap Firms to scope

    Normal_Memory_Remmit = Behaviors.scopeNormal And Behaviors.StateAccess And Behaviors.RemmitSioux ' 0 'virtualmemory and heap Sioux to scope
    Locale_Memory_Remmit = Behaviors.ScopeLocale And Behaviors.StateAccess And Behaviors.RemmitSioux ' 4 'virtualmemory and heap Sioux to scope
    Global_Memory_Remmit = Behaviors.ScopeGlobal And Behaviors.StateAccess And Behaviors.RemmitSioux ' -4 'virtualmemory and heap Sioux to scope

    Normal_Mirror_Remmit = Behaviors.scopeNormal And Behaviors.StateMirror And Behaviors.RemmitSioux '-2 'direct file and heap Sioux to scope
    Locale_Mirror_Remmit = Behaviors.ScopeLocale And Behaviors.StateMirror And Behaviors.RemmitSioux ' 2 'direct file and heap Sioux to scope
    Global_Mirror_Remmit = Behaviors.ScopeGlobal And Behaviors.StateMirror And Behaviors.RemmitSioux '-6 'direct file and heap Sioux to scope

    Normal_Memory_Commit = Behaviors.scopeNormal And Behaviors.StateAccess And Behaviors.CommitFirms ' 1 'virtualmemory and heap Firms to scope
    Locale_Memory_Commit = Behaviors.ScopeLocale And Behaviors.StateAccess And Behaviors.CommitFirms ' 5 'virtualmemory and heap Firms to scope
    Global_Memory_Commit = Behaviors.ScopeGlobal And Behaviors.StateAccess And Behaviors.CommitFirms '-3 'virtualmemory and heap Firms to scope

    Normal_Mirror_Commit = Behaviors.scopeNormal And Behaviors.StateMirror And Behaviors.CommitFirms '-1  'direct file and heap Firms to scope
    Locale_Mirror_Commit = Behaviors.ScopeLocale And Behaviors.StateMirror And Behaviors.CommitFirms '3 'direct file and heap Firms to scope
    Global_Mirror_Commit = Behaviors.ScopeGlobal And Behaviors.StateMirror And Behaviors.CommitFirms '-5  'direct file and heap Firms to scope

    Nothings_Initialized = -7 And 7
End Enum


Private Pointer As Long 'stream data nothing checks agaiinst it
Private Breaks() As Long 'array ubound is count of lines, value
'0 element is current line number, other elements are locations
'of the line numebrs linefeed character in the stream, by total
'thus element n's value, subtracting n-1's value is a line size
Private Behave As Behaviors
Private PathHandle As Long
Private PathFileName As String

Private Sub Class_Initialize()
    Behave = Nothings_Initialized
    ReDim Breaks(0 To 0) As Long
    Behavior scopeNormal
End Sub

Private Sub Class_Terminate()
    If Behave <> Nothings_Initialized Then
        If Not PathHandle = 0 Then
            CloseHandle PathHandle
            PathHandle = 0
            If Left(PathFileName, 6) = MirrorPath_Temp Then Kill Mid(PathFileName, 7)
        End If
        PathFileName = MirrorPath_None
        Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
            Case Behaviors.ScopeGlobal
                GlobalUnlock Pointer
                GlobalFree Pointer
            Case Behaviors.ScopeLocale
                LocalUnlock Pointer
                LocalFree Pointer
            Case Else
                HeapFree GetProcessHeap, 0, Pointer
        End Select
        ReDim Breaks(0 To 0) As Long
        Behave = Nothings_Initialized
    End If
End Sub

Private Sub Terminate()
    If Behave <> Nothings_Initialized Then
        Class_Terminate
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Sub Behavior(ByVal Behavior As Behaviors, Optional ByVal MirrorFilePath As String = MirrorPath_None)
    Class_Terminate
    If Behave = Nothings_Initialized Then
        Dim sc As SECURITY_ATTRIBUTES
        Dim lSize As Long
        ReDim Breaks(0 To 0) As Long
        lSize = 0
        Select Case Behavior And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
            Case Behaviors.ScopeGlobal
                Pointer = ObjPtr(Me)
                Pointer = GlobalAlloc(GMEM_MOVEABLE And Pointer, 4)
                If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
            Case Behaviors.ScopeLocale
                Pointer = ObjPtr(Me)
                Pointer = LocalAlloc(LMEM_MOVEABLE And Pointer, 4)
                If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
            Case Else
                Pointer = HeapAlloc(GetProcessHeap, 0, 4)
        End Select
        RtlMoveMemory ByVal Pointer, lSize, 4
        Behave = Behavior
        Select Case Left(MirrorFilePath, 6)
            Case MirrorPath_Temp
                PathFileName = MirrorPath_Temp & GetTemporaryFile
                PathHandle = CreateFile(Mid(PathFileName, 7), FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY And FILE_FLAG_DELETE_ON_CLOSE And FILE_FLAG_WRITE_THROUGH, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
            Case MirrorPath_None
                PathHandle = 0
                PathFileName = MirrorPath_None
            Case Else
                PathFileName = MirrorFilePath
                PathHandle = CreateFile(PathFileName, FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL And FILE_FLAG_WRITE_THROUGH, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
                lSize = Length
                If lSize = 0 Then
                    Dim loorder As Long
                    Dim hiorder As Long
                    loorder = GetFileSize(PathHandle, hiorder)
                    Dim nSize As Long
                    Dim tmp As String
                    LoWord(nSize) = loorder
                    HiWord(nSize) = hiorder
                    If (nSize <> lSize Or lSize = 0) And (nSize > 0) Then
                        Dim RHS() As Byte
                        Dim ol As OVERLAPPED
                        ol.Offset = LoWord(0)
                        ol.OffsetHigh = HiWord(0)
                        lSize = 0
                        Dim m As MemInfo
                        Do
                            ReDim Preserve RHS(0 To nSize - 1) As Byte
                            m = lArrayToANSI(RHS)
                            hReadFile PathHandle, ByVal m.MemH, nSize, lSize, ol
                            If lSize > 0 Then
                                nSize = nSize - lSize
                                RHS = lArrayOfANSI(m)
                                ReDim Preserve RHS(0 To lSize - 1) As Byte
                                lSize = ConcatNoFile(RHS, lSize)
                                For loorder = LBound(RHS) To UBound(RHS)
                                    lSize = lSize + 1
                                    If RHS(loorder) = 10 Then
                                        ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
                                        Breaks(UBound(Breaks)) = lSize
                                    End If
                                Next
                            End If
                        Loop Until nSize = 0
                    End If
                End If
        End Select
    Else
        Err.Raise 8, App.Title, "Object already initialized, use the Terminate() function to re-call Initialize()."
    End If
End Sub

'###############################################################################
'###############################################################################

Public Function Append(ByRef Lines() As Byte) As Long
    If Behave <> Nothings_Initialized Then
        Dim lLen As Long
        Dim Temp As Long
        Dim Line() As Byte
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(Lines, Line, lLen)
            Concat Line
            ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
            Breaks(UBound(Breaks)) = Length
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Sub Clear()
    If Behave <> Nothings_Initialized Then
        ReDim Breaks(0 To 0) As Long
        Breaks(0) = 0
        Reset
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Property Get Count() As Long
    If Behave <> Nothings_Initialized Then
        Count = UBound(Breaks)
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Let Current(ByVal RHS As Long)
    If Behave <> Nothings_Initialized Then
        If ((RHS <= UBound(Breaks)) And (UBound(Breaks) > 0)) Or (RHS = 0) Then
            Breaks(0) = RHS
        Else
            Err.Raise 8, App.Title, "Line number invalid or the collection is empty."
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Get Current() As Long
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            If (Breaks(0) = 0) Then Breaks(0) = 1
            If (Breaks(0) > Count) Then Breaks(0) = Count
            Current = Breaks(0)
        Else
            Current = 0
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Function Delete() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Count <= 2) Or (Current = 1) Or (Current = Count) Then
            If (Current = 1) Then
                Delete = Pop()
            Else
                Delete = Remove()
            End If
        Else
            Dim lLen As Long
            Dim Temp As Long
            Dim Part As Stream
            Temp = Current
            Breaks(0) = 0
            lLen = (Breaks(Temp) - Breaks(Temp - 1))
            Delete = Partial(Breaks(Temp - 1), lLen)
            Set Part = New Stream
            Part.Concat Partial(Breaks(Temp))
            Length = Breaks(Temp - 1)
            Concat Part.Partial
            Set Part = Nothing
            Dim cnt As Long
            For cnt = Temp To UBound(Breaks) - 1
                Breaks(cnt) = (Breaks(cnt) - lLen) + (Breaks(cnt + 1) - Breaks(cnt))
            Next
            ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) - 1) As Long
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Property Get Filename() As String
    Filename = PathFileName
End Property
Public Property Let Filename(ByVal RHS As String)
    Dim bh As Behaviors
    bh = Behave
    Class_Terminate
    Behavior bh, RHS
End Property

Public Property Let First(ByRef Lines() As Byte)
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim lLen As Long
            Dim cnt As Long
            Dim Line() As Byte
            Do While NextLine(Lines, Line, lLen)
                If (Count > 1) Then
                    For cnt = UBound(Breaks) To 2 Step -1
                        Breaks(cnt) = Breaks(cnt) - (Breaks(1) - lLen)
                    Next
                End If
                Push Breaks(1)
                Prepend Line
                Breaks(1) = lLen
            Loop
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Get First() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then First = Item(1)
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Function Forward() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Breaks(0) < UBound(Breaks)) Then Breaks(0) = Breaks(0) + 1
        Forward = Item(Breaks(0))
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Function Insert(ByRef Lines() As Byte) As Long
    If Behave <> Nothings_Initialized Then
        If (Count = 0) Or (Current = 1) Or (Current = Count) Then
            If (Current = 1) Then
                Insert = Add(Lines)
            Else
                Insert = Append(Lines)
            End If
        Else
            Dim lLen As Long
            Dim Temp As Long
            Dim Line() As Byte
            Dim cnt As Long
            Dim Part As Stream
            Temp = Current
            Breaks(0) = 0
            Do While NextLine(Lines, Line, lLen)
                Set Part = New Stream
                Part.Concat Partial(Breaks(Temp - 1))
                Length = Breaks(Temp - 1)
                Concat Line
                Concat Part.Partial
                Set Part = Nothing
                ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) + 1) As Long
                For cnt = UBound(Breaks) To Temp Step -1
                    Breaks(cnt) = Breaks(cnt - 1) + lLen
                Next
            Loop
            Breaks(0) = Temp
            Insert = Breaks(0)
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Property Let Item(Optional ByVal Number As Long = -1&, ByRef Lines() As Byte)
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
            Insert Lines
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Get Item(Optional ByVal Number As Long = -1&) As Byte()
    If Behave <> Nothings_Initialized Then
        If ((Current > 0) And (Current <= UBound(Breaks))) And (Number = -1) Then
            Number = Breaks(0)
        End If
        If (UBound(Breaks) > 0) Then
            If (Number <= UBound(Breaks)) Then
                If (Number > 1) Then
                    Item = Partial(Breaks(Number - 1), Breaks(Number) - Breaks(Number - 1))
                ElseIf (Number = 1) Then
                    Item = Partial(0, Breaks(Number))
                End If
            End If
        Else
            Err.Raise 8, App.Title, "The number, or current line number, is invalid or no line exist."
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Let Last(ByRef Lines() As Byte)
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim lLen As Long
            Dim Temp As Long
            Dim Line() As Byte
            Temp = Current
            Breaks(0) = 0
            Do While NextLine(Lines, Line, lLen)
                Length = Length - (Breaks(Count) - Breaks(Count - 1))
                Concat Line
                Breaks(Count) = Breaks(Count - 1) + lLen
            Loop
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Get Last() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then Last = Item(Count)
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Let Value(ByRef Lines() As Byte)
    If Behave <> Nothings_Initialized Then
        Clear
        Dim lLen As Long
        Dim Temp As Long
        Dim Line() As Byte
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(Lines, Line, lLen)
            Concat Line
            ReDim Preserve Breaks(0 To UBound(Breaks) + 1) As Long
            Breaks(UBound(Breaks)) = Breaks(UBound(Breaks) - 1) + lLen
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Get Value() As Byte()
    If Behave <> Nothings_Initialized Then
        Value = Partial
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Function Previous() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Breaks(0) > 1) Then Breaks(0) = Breaks(0) - 1
        Previous = Item(Breaks(0))
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Function Pop() As Byte()
    If Behave <> Nothings_Initialized Then
        If (Count > 0) Then
            Dim Temp As Long
            Temp = Current
            Pop = Partial(0, Breaks(1))
            Push Breaks(1)
            Breaks(0) = Breaks(1)
            If UBound(Breaks) >= 2 Then
                Dim cnt As Long
                For cnt = 1 To UBound(Breaks) - 1
                    Breaks(cnt) = Breaks(cnt + 1) - Breaks(0)
                Next
            End If
            ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) - 1) As Long
            Breaks(0) = Temp
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Function Add(ByRef Lines() As Byte) As Long
    If Behave <> Nothings_Initialized Then
        Dim cnt As Long
        Dim Temp As Long
        Dim lLen As Long
        Dim Line() As Byte
        Temp = Current
        Breaks(0) = 0
        Do While NextLine(Lines, Line, lLen)
            Prepend Line
            ReDim Preserve Breaks(LBound(Breaks) To UBound(Breaks) + 1) As Long
            If Count > 1 Then
                For cnt = UBound(Breaks) To 1 Step -1
                    Breaks(cnt) = Breaks(cnt - 1) + lLen
                Next
            Else
                Breaks(1) = lLen
            End If
        Loop
        Breaks(0) = Temp
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Function Remove() As Byte()
    If Behave <> Nothings_Initialized Then
        If Count > 1 Then
            Remove = Partial(Breaks(UBound(Breaks) - 1), (Breaks(UBound(Breaks)) - Breaks(UBound(Breaks) - 1)))
            Length = Length - (Breaks(UBound(Breaks)) - Breaks(UBound(Breaks) - 1))
            ReDim Preserve Breaks(0 To UBound(Breaks) - 1) As Long
        ElseIf Count = 1 Then
            Remove = Partial
            Reset
            ReDim Breaks(0 To 0) As Long
            Breaks(0) = 0
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function


'###############################################################################
'###############################################################################


Public Sub Reset()
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long: lSize = 0
        Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
            Case Behaviors.ScopeGlobal
                If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
                GlobalFree Pointer
                Pointer = GlobalAlloc(GMEM_MOVEABLE And VarPtr(Pointer), 4)
                If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
            Case Behaviors.ScopeLocale
                If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
                LocalFree Pointer
                Pointer = LocalAlloc(LMEM_MOVEABLE And VarPtr(Pointer), 4)
                If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
            Case Else
                HeapFree GetProcessHeap, 0, Pointer
                Pointer = HeapAlloc(GetProcessHeap, 0, 4)
        End Select
        RtlMoveMemory ByVal Pointer, lSize, 4
        If PathHandle <> 0 Then
            CloseHandle PathHandle
            PathHandle = 0
            If Left(PathFileName, 6) = MirrorPath_Temp And Mid(PathFileName, 7) <> "" Then Kill Mid(PathFileName, 7)
        End If
        Dim sc As SECURITY_ATTRIBUTES
        Select Case Left(PathFileName, 6)
            Case MirrorPath_Temp
                PathFileName = MirrorPath_Temp & GetTemporaryFile
                PathHandle = CreateFile(Mid(PathFileName, 7), FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, CREATE_ALWAYS, FILE_ATTRIBUTE_TEMPORARY And FILE_FLAG_DELETE_ON_CLOSE And FILE_FLAG_WRITE_THROUGH, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
            Case MirrorPath_None
                PathHandle = 0
            Case Else
                PathHandle = CreateFile(PathFileName, FILE_GENERIC_WRITE Or FILE_GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, sc, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL And FILE_FLAG_WRITE_THROUGH, ByVal 0&)
                If PathHandle <= 0 Then Err.Raise 8, App.Title, "Error opening file."
        End Select
        EndOfStream 0
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Property Get Length() As Long
    If Behave <> Nothings_Initialized Then
        RtlMoveMemory Length, ByVal Pointer, 4
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Property Let Length(ByVal RHS As Long)
    If Behave <> Nothings_Initialized Then
        If (RHS > -1) Then
            Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
                Case Behaviors.ScopeGlobal
                    If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
                    Pointer = GlobalReAlloc(Pointer, RHS + 4, GMEM_MOVEABLE Or GPTR)
                    If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
                Case Behaviors.ScopeLocale
                    If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
                    Pointer = LocalReAlloc(Pointer, RHS + 4, LMEM_MOVEABLE Or lPtr)
                    If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
                Case Else
                    Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, RHS + 4)
            End Select
            RtlMoveMemory ByVal Pointer, RHS, 4
            EndOfStream RHS
        ElseIf (RHS = 0) Then
            Reset
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Property

Public Sub Push(ByVal Offset As Long)
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        lSize = Length
        If (Offset < lSize) Then
            RtlMoveMemory ByVal (Pointer + 4), ByVal (Pointer + 4) + Offset, (-lSize + -Offset + (lSize * 2))
            Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
                Case Behaviors.ScopeGlobal
                    If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
                    Pointer = GlobalReAlloc(Pointer, (lSize - Offset) + 4, GMEM_MOVEABLE Or GPTR)
                    If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
                Case Behaviors.ScopeLocale
                    If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
                    Pointer = LocalReAlloc(Pointer, (lSize - Offset) + 4, LMEM_MOVEABLE Or lPtr)
                    If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
                Case Else
                    Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (lSize - Offset) + 4)
            End Select
            RtlMoveMemory ByVal Pointer, (lSize - Offset), 4
            WroteToStream Partial, 0, Length
            EndOfStream Length
        ElseIf (Offset = lSize) Then
            Reset
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Sub Pull(ByVal Offset As Long)
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        lSize = Length
        If (Offset < lSize) Then
            RtlMoveMemory ByVal (Pointer + 4) + Offset, ByVal (Pointer + 4), (-lSize + -Offset + (lSize * 2))
            Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
                Case Behaviors.ScopeGlobal
                    If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
                    Pointer = GlobalReAlloc(Pointer, (lSize - Offset) + 4, GMEM_MOVEABLE Or GPTR)
                    If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
                Case Behaviors.ScopeLocale
                    If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
                    Pointer = LocalReAlloc(Pointer, (lSize - Offset) + 4, LMEM_MOVEABLE Or lPtr)
                    If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
                Case Else
                    Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (lSize - Offset) + 4)
            End Select
            RtlMoveMemory ByVal Pointer, (lSize - Offset), 4
            WroteToStream Partial, 0, lSize
            EndOfStream Length
        ElseIf (Offset = lSize) Then
            Reset
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Sub Prepend(ByRef RHS() As Byte)
    If Behave <> Nothings_Initialized Then
        If Length = 0 Then
            Concat RHS
        Else
            Dim nSize As Long
            nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
            Length = Length + (nSize * 2)
            Pull nSize
            Placeat RHS, 0, nSize
            WroteToStream Partial, 0, Length
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Private Function ConcatNoFile(RHS() As Byte, ByVal nSize As Long) As Long
    If Behave <> Nothings_Initialized Then
        If (nSize > 0) And UBound(RHS) > -1 Then
            Dim lSize As Long
            RtlMoveMemory lSize, ByVal Pointer, 4
            Select Case Behave And (Behaviors.ScopeGlobal Or Behaviors.ScopeLocale Or Behaviors.scopeNormal)
                Case Behaviors.ScopeGlobal
                    If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
                    Pointer = GlobalReAlloc(Pointer, (nSize + lSize) + 4, GMEM_MOVEABLE Or GPTR)
                    If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
                Case Behaviors.ScopeLocale
                    If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
                    Pointer = LocalReAlloc(Pointer, (nSize + lSize) + 4, LMEM_MOVEABLE Or lPtr)
                    If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
                Case Else
                    Pointer = HeapReAlloc(GetProcessHeap, 0, Pointer, (nSize + lSize) + 4)
            End Select
            RtlMoveMemory ByVal Pointer, (nSize + lSize), 4
            RtlMoveMemory ByVal (Pointer + 4) + lSize, ByVal VarPtr(RHS(LBound(RHS))), nSize
            ConcatNoFile = lSize
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

Public Sub Concat(ByRef RHS() As Byte)
    If Behave <> Nothings_Initialized Then
        Dim nSize As Long
        nSize = UBound(RHS) + -CInt(Not CBool(-LBound(RHS)))
        If (nSize > 0) Then
            Dim lSize As Long
            lSize = ConcatNoFile(RHS, nSize)
            WroteToStream RHS, lSize, nSize
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte()
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        RtlMoveMemory lSize, ByVal Pointer, 4
        If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
        If (Width > 0) Then
            Dim Ret() As Byte
            ReDim Ret(0 To Width - 1) As Byte
            RtlMoveMemory ByVal VarPtr(Ret(LBound(Ret))), ByVal Pointer + 4 + Offset, Width
            Partial = Ret
            Erase Ret
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function
Public Function Placeat(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
    If Behave <> Nothings_Initialized Then
        Dim lSize As Long
        RtlMoveMemory lSize, ByVal Pointer, 4
        If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
        If (Width > 0) And (UBound(RHS) > -1) Then
            RtlMoveMemory ByVal Pointer + 4 + Offset, ByVal VarPtr(RHS(LBound(RHS))), Width
            WroteToStream RHS, Offset, Width
        End If
    Else
        Err.Raise 8, App.Title, "Object not initialized, use the Initialize() function first, to use this object."
    End If
End Function

'###############################################################################
'###############################################################################


'###############################################################################
'the below functions are private and used to be adaptable to the above stranding
'functions in such a way they they are convieniently obtaining all possible nums
'that are based on the provided information and return to from the object relate
'###############################################################################

Private Sub WroteToStream(ByRef RHS() As Byte, ByVal Offset As Long, ByVal Width As Long)
    Select Case Left(PathFileName, 6)
        Case MirrorPath_None
        Case Else
            If Width > 0 And (UBound(RHS) > -1) Then
                Dim pos As Long
                Dim ol As OVERLAPPED
                pos = Offset
                ol.Offset = LoWord(pos)
                ol.OffsetHigh = HiWord(pos)
                Offset = 0
                Do
                    hWriteFile PathHandle, ByVal nArrayToString(RHS), Width, Offset, ol
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
Private Function LineLength(ByRef RHS() As Byte) As Long
    LineLength = (UBound(RHS) + -CInt(Not CBool(-LBound(RHS))))
    If RHS(UBound(RHS)) <> 10 Then
        ReDim Preserve RHS(LBound(RHS) To UBound(RHS) + 1) As Byte
        RHS(UBound(RHS)) = 10
        LineLength = LineLength + 1
    End If
End Function

Public Function CRLFCheck() As Boolean
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

Private Function NextLine(ByRef Lines() As Byte, ByRef Line() As Byte, ByRef lLen As Long) As Boolean
    Static inLines() As String
    Static atLine As Long
    Static AtState As Long
    If AtState = 0 Then ReDim inLines(0 To 0) As String
    Dim str As String
    If atLine > UBound(inLines) And AtState > 0 Then
        atLine = 0
        AtState = 0
        Erase inLines
        Erase Line
    Else
        Select Case AtState
            Case 0
                atLine = 0
                Erase inLines
                str = Replace(Convert(Lines), Chr(13), "")
                inLines = Split(str, Chr(10))
                atLine = LBound(inLines)
                Erase Line
                Line = Convert(inLines(atLine) & vbCr)
                lLen = LineLength(Line)
                NextLine = True
                AtState = 1
            Case 1
                If (atLine >= LBound(inLines)) And (atLine < UBound(inLines)) Then
                    If atLine <= UBound(inLines) Then
                        Line = Convert(inLines(atLine) & vbCr)
                        lLen = LineLength(Line)
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

''###############################################################################
''below until the private definition section are functions of the main strandings
''manipulation of relevant information from cursor to count and other types edits
''each function demands inputs or calling to be predetermined the with out errors
''such that they do not "give" for inaccurate state of the strands collection set
''###############################################################################
'
'Public Property Let Value(ByRef Lines() As Byte)
''this propeprty gets or sets the whole line collection
''listing accepting multiple lines with in the argument
''all previous lines will be erased as new ones reflect
'    IStrands.Value = Lines
'End Property
'Public Property Get Value() As Byte()
'    Value = IStrands.Value
'End Property
'
'Public Function Add(ByRef Lines() As Byte) As Long
''Add forces the lines in the supplied argument to
''beginning of the line colleciton, moving existing
''lines to higher positions in collection returning
''the total number of lines added to the list, when
''no lines are supplied the function returns a zero
'    Add = IStrands.Add(Lines)
'End Function
'
'Public Function Insert(ByRef Lines() As Byte) As Long
''insert adds the supplied lines in the argument to the
''current cursor positioned location in line collection
''if no lines are added the return value is zero other-
''wise the return value is the numebr of lines inserted
'    Insert = IStrands.Insert(Lines)
'End Function
'
'Public Function Append(ByRef Lines() As Byte) As Long
''append adds the lines of the argument to the end of
''collection increasing the count and ignores current
'    Append = IStrands.Append(Lines)
'End Function
'
'Public Sub Clear()
''before: condditions of strand and stream information
''that are reset to useable with zero count no current
'    IStrands.Clear
'End Sub
'
'Public Property Get Count() As Long
''returns the total amoung of lines in the collection
''this is same as the ubound value of the breaks arry
'    Count = IStrands.Count
'End Property
'
'Public Property Get Current() As Long
''property that allows get or set of the current
''cursor postion of the collection of lines data
''initial current is the cont of list if a count
''is zero then the property raises error failure
'    Current = IStrands.Current
'End Property
'Public Property Let Current(ByVal RHS As Long)
''must be a valid numerical of any numer from 1
''to count and all those inbetween, else errors
'    IStrands.Current = RHS
'End Property
'
'Public Function Delete() As Byte()
''delete extracts the current line's text returning
''it and erasing it if the current line is valid or
''it will error, which is the difference of 3 erase
'    Delete = IStrands.Delete
'End Function
'
'Public Function Pop() As Byte()
''pop extracts and returns the first line in the
''line collection erasing it from the collection
''if no lines exist the function causes an error
'    Pop = IStrands.Pop
'End Function
'
'Public Function Remove() As Byte()
''remove extracts and returns the last line in
''the line collection, erasing from collection
''if no lines exist it will cause an error msg
'    Remove = IStrands.Remove
'End Function
'
'Public Property Let First(ByRef Line() As Byte)
''first property gets or sets the first line in
''the line collection, the line must exist else
''the property will cause an error if count = 0
''this will not accpet multiple lines in supply
''they will be treated as one included linefeed
'    IStrands.First = Line
'End Property
'Public Property Get First() As Byte()
'    First = IStrands.First
'End Property
'
'Public Property Let Item(Optional ByVal Number As Long = -1&, ByRef Line() As Byte)
''the item property gets or sets either the current
''line number in the collection, or the supplied as
''the number argument line number in the collection
''this function, unlike adding functions wont allow
''multiple lines and will cause an error if count=0
''supplying the number does not change the currents
'    IStrands.Item(Number) = Line
'End Property
'Public Property Get Item(Optional ByVal Number As Long = -1&) As Byte()
'    Item = IStrands.Item(Number)
'End Property
'
'Public Property Let Last(ByRef Lines() As Byte)
''this property gets or sets the last line in the
''line collection which like first, must exist as
''well as the collection count >0 else causes err
''this will not accept multiple lines in argument
'    IStrands.Last = Lines
'End Property
'Public Property Get Last() As Byte()
'    Last = IStrands.Last
'End Property
'
'Public Function Forward() As Byte()
''this function moves the current cursor line
''number of the collection inc up in the list
''and returns the information of the line and
''will error when if current line before call
''if at very last line in the line collection
'    Forward = IStrands.Forward
'End Function
'
'Public Function Previous() As Byte()
''this function moves the current cursor line
''number of the collection down a value until
''it is the frist line in the list, any calls
''made if at the first line, will cause error
'    Previous = IStrands.Previous
'End Function
'
'Public Sub Filename(Optional ByVal Temporary As Boolean = True, Optional ByVal FilePath As String = MirrorPath_None)
''this functions sets the level of backing the list by file
''usage, whther none, or temporary (held by state supported
''but not kept) or custom which is not removed by resetting
''and is loaded upon setting the file if and when is custom
'    IStrands.Filename Temporary, FilePath
'End Sub
'
'
''###############################################################################
''Everything below is part of the stream and has no relevence and/or reference to
''strands information above strands, no dependance, any modification also handles
''the direct file manipulation just the same, not overwriting custom in instanced
''loading the file, temporary file use is opened overwritten and is for stability
''###############################################################################
'
'Public Sub Reset()
'    IStrands.Reset
'End Sub
'
'Public Property Get Length() As Long
'    Length = IStrands.Length
'End Property
'
'Public Property Let Length(ByVal RHS As Long)
'    IStrands.Length = RHS
'End Property
'
'Public Sub Prepend(ByRef RHS() As Byte)
'    IStrands.Prepend RHS
'End Sub
'
'Public Function Placeat(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
'    Placeat = IStrands.Placeat(RHS, Offset, Width)
'End Function
'
'Public Sub Pull(ByVal Offset As Long)
'    IStrands.Pull Offset
'End Sub
'
'Public Sub Push(ByVal Offset As Long)
'    IStrands.Push Offset
'End Sub
'
'Public Sub Concat(ByRef RHS() As Byte)
'    IStrands.Concat RHS
'End Sub
'
'Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte()
'    Partial = IStrands.Partial(Offset, Width)
'End Function

