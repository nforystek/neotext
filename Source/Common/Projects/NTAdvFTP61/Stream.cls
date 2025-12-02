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
Option Compare Binary
Option Base 1

Implements IStream

'Private Enum AtomValues
'    AtomUnique = 1 'the actual atom number returned by api
'    AtomOffset = 2 'the numerical count where this atom among others (address defining)
'    AtomWidths = 3 'the length of the data atom (defines partial to next atoms address)
'    AtomFields = 3 'how many fields of this enumerator there are for redim measure
'End Enum
'
'Private AtomInfo() As Long

Private wStream As IStream
Attribute wStream.VB_VarHelpID = -1

'Private Declare Function InitAtomTable Lib "kernel32" (ByVal nSize As Long) As Long
'Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
'Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'Private Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
'Private Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub Class_Initialize()
    Set wStream = New IStream
    Set IStream = wStream
'    InitAtomTable 2 'as a double buffer for incoming and outgoing
'                    'at least two or more buffers in theory is constant
'    ReDim AtomInfo(AtomFields, 1) As Long
'                    'make the information table an array starting out
End Sub

Private Sub Class_Terminate()
'    DeleteAllAtoms
    Set wStream = Nothing
End Sub

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

Private Sub IStream_Placeat(RHS() As Byte, Optional ByVal Offset As Long = 0&, Optional ByVal Width As Long = -1&)
    Me.Placeat RHS, Offset, Width
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

Private Sub IStream_Reset(Optional Scopes As MemoryScopes)
    Me.Reset Scopes
End Sub

'Private Function AtomCount() As Long
'    AtomCount = UBound(AtomInfo, 2)
'End Function
'
'Private Function NewAtom() As Long
'    If (AtomInfo(AtomUnique, 1) <> 0) Or (AtomCount > 1) Then
'        ReDim Preserve AtomInfo(AtomWidths, AtomCount + 1) As Long
'    End If
'    AtomInfo(AtomUnique, AtomCount) = GlobalAddAtom(ObjPtr(Me) & "_" & AtomCount)
'    NewAtom = AtomInfo(AtomUnique, AtomCount)
'End Function
'
'Private Sub DelOffsetAtom(ByVal Atom As Long)
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomUnique, cnt) = Atom Then
'            GlobalDeleteAtom Atom
'            AtomInfo(AtomUnique, cnt) = 0
'        End If
'        If (cnt > 1) Then
'            If (AtomInfo(AtomUnique, cnt - 1) = 0) Then
'                AtomInfo(AtomUnique, cnt - 1) = AtomInfo(AtomUnique, cnt)
'                AtomInfo(AtomOffset, cnt - 1) = AtomInfo(AtomOffset, cnt)
'                AtomInfo(AtomWidths, cnt - 1) = AtomInfo(AtomWidths, cnt)
'                AtomInfo(AtomUnique, cnt) = 0
'            End If
'        End If
'    Next
'    If AtomCount - 1 > 1 Then
'        ReDim Preserve AtomInfo(AtomWidths, AtomCount - 1) As Long
'    Else
'        AtomInfo(AtomUnique, 1) = 0
'    End If
'End Sub
'
'Public Property Get Offsets(ByRef Atom As Long) As Long
'    If Atom = 0 Then Atom = NewAtom
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomUnique, cnt) = Atom Then
'            Offsets = AtomInfo(AtomOffset, cnt)
'            Exit Property
'        End If
'    Next
'End Property
'
'Public Property Let Offsets(ByRef Atom As Long, ByVal RHS As Long)
'    If Atom = 0 Then Atom = NewAtom
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomUnique, cnt) = Atom Then
'            AtomInfo(AtomOffset, cnt) = RHS
'            Exit Property
'        End If
'    Next
'End Property
'
'Public Property Get Widths(ByRef Atom As Long) As Long
'    If Atom = 0 Then Atom = NewAtom
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomUnique, cnt) = Atom Then
'            Widths = AtomInfo(AtomWidths, cnt)
'            Exit Property
'        End If
'    Next
'End Property
'
'Public Property Let Widths(ByRef Atom As Long, ByVal RHS As Long)
'    If Atom = 0 Then Atom = NewAtom
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomUnique, cnt) = Atom Then
'            AtomInfo(AtomWidths, cnt) = RHS
'            Exit Property
'        End If
'    Next
'End Property

Public Function Address(Optional ByVal Offset As Long = 0) As Long
    Address = IStream.Address(Offset)
End Function

Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
    IStream.Clone Origin, Offset, Width
End Sub

Public Sub Concat(ByRef RHS() As Byte)
    IStream.Concat RHS
End Sub

Public Property Let Length(ByVal RHS As Long)
'    ModifyOffsetWidths , IStream.Length - RHS
    IStream.Length = RHS
End Property

Public Property Get Length() As Long
    Length = IStream.Length
End Property

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1) As Byte()
    Partial = IStream.Partial(Offset, Width)
End Function

Public Sub Placeat(ByRef RHS() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Width As Long = -1)
    IStream.Placeat RHS, Offset, Width
End Sub

Public Sub Prepend(ByRef RHS() As Byte)
    Dim llen As Long
    llen = Length
    IStream.Prepend RHS
'    ModifyOffsetWidths -(Length - llen)
End Sub

Public Sub Pull(ByVal Offset As Long)
    IStream.Pull Offset
'    ModifyOffsetWidths -Offset, Offset
End Sub

Public Sub Push(ByVal Offset As Long)
    IStream.Push Offset
'    ModifyOffsetWidths Offset
End Sub

'Private Sub ModifyOffsetWidths(Optional ByVal PredicutOffset As Long = 0, Optional ByVal TruncateWidth As Long = 0)
'    Dim cnt As Long
'    For cnt = 1 To AtomCount
'        If AtomInfo(AtomOffset, cnt) > 0 Then
'            AtomInfo(AtomOffset, cnt) = AtomInfo(AtomOffset, cnt) - PredicutOffset
'        ElseIf AtomInfo(AtomWidths, cnt) > 0 And AtomInfo(AtomOffset, cnt) <= 0 Then
'            AtomInfo(AtomWidths, cnt) = AtomInfo(AtomWidths, cnt) - PredicutOffset
'        End If
'        If AtomInfo(AtomOffset, cnt) < 0 Then
'            AtomInfo(AtomWidths, cnt) = AtomInfo(AtomOffset, cnt) + AtomInfo(AtomWidths, cnt)
'            AtomInfo(AtomOffset, cnt) = 0
'        End If
'        If AtomInfo(AtomWidths, cnt) > Length - AtomInfo(AtomOffset, cnt) - TruncateWidth Then
'            AtomInfo(AtomWidths, cnt) = Length - AtomInfo(AtomOffset, cnt) - TruncateWidth
'        End If
'        If AtomInfo(AtomWidths, cnt) < 0 Then AtomInfo(AtomWidths, cnt) = 0
'    Next
'End Sub
'
'Private Sub DeleteAllAtoms()
'    Do While AtomInfo(AtomUnique, 1) <> 0
'        DelOffsetAtom AtomInfo(AtomUnique, 1)
'    Loop
'End Sub

Public Sub Reset(Optional ByRef Scopes As MemoryScopes)
    IStream.Reset Scopes
End Sub

