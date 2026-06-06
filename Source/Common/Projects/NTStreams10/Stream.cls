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

Implements ntnodes10.Stream

Private MyStream As ntnodes10.Stream

Public Property Get Length() As Long
    Length = MyStream.Length
End Property

Public Property Let Length(ByVal RHS As Long)
    If RHS >= 0 Then
        MyStream.Length = RHS
    Else
        Err.Raise 380, App.Title, "Invalid property value"
    End If
End Property

Public Function charAt(ByVal index As Long) As Stream
    Set charAt = New Stream
    charAt.Post MyStream.Peek(index)
End Function

Public Function charCodeAt(ByVal index As Long) As Byte
    charCodeAt = Asc(MyStream.Peek(index))
End Function

Public Function at(ByVal index As Long) As Stream
    Set at = New Stream
    If (index >= 0 And index < MyStream.Length) Or (index < 0 And index >= -MyStream.Length) Then
        If (index > 0) Then
            at.Post MyStream.Peek(index)
        Else
            at.Post MyStream.Peek((MyStream.Length + index))
        End If
    Else
        Err.Raise 9, App.Title, "Subscript out of range"
    End If
End Function

Public Function slice(ByVal startIndex As Long, ByVal endIndex As Long) As Stream
    Set slice = New Stream
    If (startIndex >= 0 And startIndex < MyStream.Length) Or (startIndex < 0 And startIndex >= -MyStream.Length) Then
        If (startIndex > 0) Then
            slice.ConcatBytes MyStream.Partial(startIndex, (endIndex - startIndex))
        Else
            slice.ConcatBytes MyStream.Partial((MyStream.Length + startIndex), (((MyStream.Length + endIndex) - 1) - (MyStream.Length + startIndex)))
        End If
    Else
        Err.Raise 9, App.Title, "Subscript out of range"
    End If
End Function

Public Function substring(ByVal startIndex As Long, ByVal endIndex As Long) As Stream
    Set substring = New Stream
    If (startIndex >= 0 And startIndex < MyStream.Length) And (endIndex < MyStream.Length) Then
        substring.ConcatBytes MyStream.Partial(startIndex, (endIndex - startIndex))
    Else
        Err.Raise 9, App.Title, "Subscript out of range"
    End If
End Function

Public Function substr(ByVal startIndex As Long, ByVal charCount As Long) As Stream
    Set substr = New Stream
    If (startIndex >= 0 And startIndex < MyStream.Length) And ((startIndex + charCount) < MyStream.Length) Then
        substr.ConcatBytes MyStream.Partial(startIndex, charCount)
    Else
        Err.Raise 9, App.Title, "Subscript out of range"
    End If
End Function

Public Function toUpperCase() As Stream
    'Set toUpperCase = New Stream
    Set toUpperCase = Strings.UCase(MyStream)
End Function

Public Function toLowerCase() As Stream
    Set toLowerCase = Strings.LCase(MyStream)
End Function

Public Function concat(ByRef str0, Optional ByRef str1, Optional ByRef str2, Optional ByRef str3, Optional ByRef str4, _
            Optional ByRef str5, Optional ByRef str6, Optional ByRef str7, Optional ByRef str8) As Stream
    Set concat = New Stream
    concat.Clone Me
    concat = Strings.ConcatenateEx(concat, str0, str1, str2, str3, str4, str5, str6, str7, str8)
End Function

Public Function trim() As Stream
    Set trim = Strings.trim(Stream)
End Function

Public Function trimStart() As Stream
    Set trimStart = Strings.LTrim(Stream)
End Function

Public Function trimEnd() As Stream
    Set trimEnd = Strings.RTrim(Stream)
End Function

Public Function padStart(ByVal charCount As Long, ByRef padString As Stream) As Stream
    Set padStart = New Stream
    Do While padStart.Length < charCount
        padStart.ConcatBytes padString.Partial
    Loop
    padStart.Length = charCount
    padStart.ConcatBytes MyStream.Partial
End Function

Public Function padEnd(ByVal charCount As Long, ByRef padString As Stream) As Stream
    Set padEnd = New Stream
    Do While padEnd.Length < charCount
        padEnd.ConcatBytes padString.Partial
    Loop
    padEnd.Length = charCount
    padEnd.Prepend MyStream.Partial
End Function

Public Function repeat(ByVal repeatCount As Long) As Stream
    Set repeat = New Stream
    Do While repeatCount > 0
        repeat.ConcatBytes MyStream.Partial
        repeatCount = repeatCount - 1
    Loop
End Function

Public Function replace(ByRef findString As Stream, ByRef replaceWith As Stream) As Stream
    Dim pos As Long
    Set replace = New Stream
    replace.ConcatBytes MyStream.Partial
    If findString.Length > 0 And replace.Length > 0 Then
        pos = InStr(1, replace, findString)
        If pos > 0 And pos <= replace.Length Then
            If replaceWith.Length > 0 Then
                replace.Pyramid replaceWith, (pos - 1), findString.Length
                pos = (pos + replaceWith.Length)
            Else
                replace.Pinch (pos - 1), findString.Length
                pos = (pos + replaceWith.Length)
            End If
        End If
    End If
End Function

Public Function replaceAll(ByRef findString As Stream, ByRef replaceWith As Stream) As Stream
    Dim pos As Long
    Set replaceAll = New Stream
    replaceAll.ConcatBytes MyStream.Partial
    If findString.Length > 0 And replaceAll.Length > 0 Then
        pos = 1
        Do
            pos = InStr(pos, replaceAll, findString)
            If pos > 0 And pos <= replaceAll.Length Then
                If replaceWith.Length > 0 Then
                    replaceAll.Pyramid replaceWith, (pos - 1), findString.Length
                    pos = (pos + replaceWith.Length)
                Else
                    replaceAll.Pinch (pos - 1), findString.Length
                    pos = (pos + replaceWith.Length)
                End If
            End If
        Loop While pos > 0 And pos <= replaceAll.Length
    End If
End Function

Public Function split(ByRef delimiter As Stream) As Stream()
    Dim ary() As Stream
    Dim str As New Stream
    str.ConcatBytes MyStream.Partial
    ReDim ary(0) As Stream
    Set ary(0) = str
    Set str = Nothing
    Do While Strings.InStrP(1, ary(UBound(ary)), delimiter) > 0
        ReDim Preserve ary(LBound(ary) To UBound(ary) + 1) As Stream
        Set ary(UBound(ary)) = Strings.RemoveArgP(ary(UBound(ary) - 1), delimiter)
        Set ary(UBound(ary) - 1) = Strings.NextArgP(ary(UBound(ary) - 1), delimiter)
    Loop
    split = ary
End Function

Friend Property Get Address(Optional ByVal Offset As Long = 0&) As Long
    Address = MyStream.Address(Offset)
End Property

Friend Sub Clone(Origin As ntnodes10.Stream, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&)
    MyStream.Clone Origin, Offset, Width
End Sub

Friend Sub ConcatBytes(ByRef Data() As Byte)
    MyStream.concat Data
End Sub

Friend Function Partial(Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Byte()
    Partial = MyStream.Partial(Offset, Width)
End Function

Friend Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Long
    Pass = MyStream.Pass(Data, Offset, Width)
End Function

Friend Function Peek(Optional ByVal Offset As Long = 0&) As Byte
    Peek = MyStream.Peek(Offset)
End Function

Friend Sub Pinch(ByVal Offset As Long, ByVal Width As Long)
    MyStream.Pinch Offset, Width
End Sub

Friend Sub PlaceAt(Data() As Byte, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&)
    MyStream.PlaceAt Data, Offset, Width
End Sub

Friend Sub Plot(ByVal Offset As Long, ByVal Data As Byte)
    MyStream.Plot Offset, Data
End Sub

Friend Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Long
    Poll = MyStream.Poll(Data, Ballot, Offset, Width)
End Function

Friend Sub Post(Data As Byte)
    MyStream.Post Data
End Sub

Friend Sub Prepend(Data() As Byte)
    MyStream.Prepend Data
End Sub

Friend Sub Pull(ByVal Offset As Long)
    MyStream.Pull Offset
End Sub

Friend Sub Push(ByVal Offset As Long)
    MyStream.Push Offset
End Sub

Friend Sub Pyramid(SubStream As Stream, ByVal Offset As Long, ByVal Width As Long)
    MyStream.Pyramid SubStream, Offset, Width
End Sub

Friend Sub Reset(Optional ScopeAs As ntnodes10.MemoryScopes)
    MyStream.Reset ScopeAs
End Sub

Friend Property Get Scope() As ntnodes10.MemoryScopes
    Scope = MyStream.Scope
End Property

Public Property Get from()
Attribute from.VB_UserMemId = 0
Attribute from.VB_MemberFlags = "40"
    from = StrConv(MyStream.Partial, vbUnicode)
End Property

Public Property Let from(ByVal RHS)
    If VBA.TypeName(RHS) = "String" Then
        Me.Reset
        Me.ConcatBytes StrConv(RHS, vbFromUnicode)
    ElseIf VBA.TypeName(RHS) = "Stream" Then
        Dim sp As New ntnodes10.Stream
        Set sp = RHS
        Me.Clone sp
        Set sp = Nothing
    End If
End Property

Private Sub Class_Initialize()
    Set MyStream = New ntnodes10.Stream
    vbaObjSetAddref Stream, ObjPtr(MyStream)
End Sub

Private Property Get Stream_Address(Optional ByVal Offset As Long = 0&) As Long
    Stream_Address = MyStream.Address(Offset)
End Property

Private Sub Stream_Clone(Origin As ntnodes10.Stream, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&)
    MyStream.Clone Origin, Offset, Width
End Sub

Private Sub Stream_Concat(Data() As Byte)
    MyStream.concat Data
End Sub

Private Property Let Stream_Length(ByVal RHS As Long)
    MyStream.Length = RHS
End Property

Private Property Get Stream_Length() As Long
    Stream_Length = MyStream.Length
End Property

Private Function Stream_Partial(Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Byte()
    Stream_Partial = MyStream.Partial(Offset, Width)
End Function

Private Function Stream_Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Long
    Stream_Pass = MyStream.Pass(Data, Offset, Width)
End Function

Private Function Stream_Peek(Optional ByVal Offset As Long = 0&) As Byte
    Stream_Peek = MyStream.Peek(Offset)
End Function

Private Sub Stream_Pinch(ByVal Offset As Long, ByVal Width As Long)
    MyStream.Pinch Offset, Width
End Sub

Private Sub Stream_PlaceAt(Data() As Byte, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&)
    MyStream.PlaceAt Data, Offset, Width
End Sub

Private Sub Stream_Plot(ByVal Offset As Long, ByVal Data As Byte)
    MyStream.Plot Offset, Data
End Sub

Private Function Stream_Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0&, Optional Width As Long = -1&) As Long
    Stream_Poll = MyStream.Poll(Data, Ballot, Offset, Width)
End Function

Private Sub Stream_Post(Data As Byte)
    MyStream.Post Data
End Sub

Private Sub Stream_Prepend(Data() As Byte)
    MyStream.Prepend Data
End Sub

Private Sub Stream_Pull(ByVal Offset As Long)
    MyStream.Pull Offset
End Sub

Private Sub Stream_Push(ByVal Offset As Long)
    MyStream.Push Offset
End Sub

Private Sub Stream_Pyramid(SubStream As ntnodes10.Stream, ByVal Offset As Long, ByVal Width As Long)
    MyStream.Pyramid SubStream, Offset, Width
End Sub

Private Sub Stream_Reset(Optional ScopeAs As ntnodes10.MemoryScopes)
    MyStream.Reset ScopeAs
End Sub

Private Property Get Stream_Scope() As ntnodes10.MemoryScopes
    Stream_Scope = MyStream.Scope
End Property
