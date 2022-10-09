VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function xGetcpy Lib "kernel32" Alias "xGetcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function xGetlen Lib "kernel32" Alias "xGetlenA" (ByVal lpString As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GPTR = &H40
Private Const GHND = &H42

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapCompact Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long) As Long
Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapValidate Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private lPtr As Long

Friend Function Pointer() As Long
    Pointer = lPtr
End Function

Private Sub Class_Initialize()
    Dim lSize As Long: lSize = 0
    lPtr = HeapAlloc(GetProcessHeap, 0, 4)
    CopyMemory ByVal lPtr, lSize, 4
End Sub

Private Sub Class_Terminate()
    HeapFree GetProcessHeap, 0, lPtr
End Sub

Public Sub Reset()
    Class_Terminate
    Class_Initialize
End Sub

Public Function GetString(Optional ByVal Upto As Variant = "", Optional ByVal Commit As Boolean = False) As String
    Dim lSize As Long
    Dim Ret As String
    CopyMemory lSize, ByVal lPtr, 4
    Ret = Space(lSize \ 2)
    CopyMemory ByVal StrPtr(Ret), ByVal lPtr + 4, lSize

    Dim strs() As String
    If TypeName(Upto) = "Strand" Then
        strs = Split(Ret, Upto.GetString(, False), , vbTextCompare)
        If UBound(strs) > 0 Then
            GetString = Left(Ret, Len(strs(0)))
            If Commit Then
                Ret = Mid(Ret, Len(strs(0)) + Upto.Size + 1)
                Concat Ret, True
            End If
        ElseIf UBound(strs) = 0 Then
            Ret = strs(0)
            GetString = Ret
            If Commit Then Reset
        End If
    ElseIf TypeName(Upto) = "String" Then
        If Not (Upto = "") Then
            strs = Split(Ret, Upto, , vbTextCompare)
            If UBound(strs) > 0 Then
                GetString = Left(Ret, Len(strs(0)))
                If Commit Then
                    Ret = Mid(Ret, Len(strs(0)) + Len(Upto) + 1)
                    Concat Ret, True
                End If
            ElseIf UBound(strs) = 0 Then
                Ret = strs(0)
                GetString = Ret
                If Commit Then Reset
            End If
        Else
            GetString = Ret
            If Commit Then Reset
        End If
    ElseIf IsNumeric(Upto) Then
        GetString = Left(Ret, Upto)
        If Commit Then
            Ret = Mid(Ret, Upto + 1)
            Concat Ret, True
        End If
    Else
        Err.Raise 4
    End If
End Function

Public Function Replace(ByVal Find As String, ByVal Place As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = -1, Optional ByVal Compare As Variant = vbBinaryCompare, Optional ByVal Commit As Boolean = False) As String
    Dim lSize As Long
    Dim Ret As String
    CopyMemory lSize, ByVal lPtr, 4
    Ret = Space(lSize \ 2)
    CopyMemory ByVal StrPtr(Ret), ByVal lPtr + 4, lSize
    Ret = VBA.Replace(Ret, Find, Place, Start, Count, Compare)
    If Commit Then Concat Ret, True
    Replace = Ret
End Function

Public Function InStr(Optional ByVal Start As Variant = Empty, Optional ByVal Find As Variant = Empty, Optional ByVal Compare As Variant = vbBinaryCompare) As Long
    Dim lSize As Long
    Dim Ret As String
    CopyMemory lSize, ByVal lPtr, 4
    Ret = Space(lSize \ 2)
    CopyMemory ByVal StrPtr(Ret), ByVal lPtr + 4, lSize
    If IsNumeric(Start) And ((Find = Empty) Or (Find = vbTextCompare) Or (Find = vbBinaryCompare) Or (Find = vbDatabaseCompare)) Then
        If (Compare = vbBinaryCompare) And ((Find = vbTextCompare) Or (Find = vbDatabaseCompare)) Then Compare = Find
        If (Find = Empty) Then
            InStr = VBA.InStr(1, Ret, Start, Compare)
    CopyMemory lSize, ByVal lPtr, 4
    Ret = Space(lSize \ 2)
    CopyMemory ByVal StrPtr(Ret), ByVal lPtr + 4, lSize
        Else
            InStr = VBA.InStr(Start, Ret, Find, Compare)
        End If
    ElseIf (Find = Empty) Or ((Find = Empty) Or (Find = vbTextCompare) Or (Find = vbBinaryCompare) Or (Find = vbDatabaseCompare)) Then
        If (Compare = vbBinaryCompare) And ((Find = vbTextCompare) Or (Find = vbDatabaseCompare)) Then Compare = Find
        InStr = VBA.InStr(1, Ret, Start, Compare)
    ElseIf (Not (Start = Empty)) And (Find = Empty) Then
        InStr = VBA.InStr(1, Ret, Start, Compare)
    ElseIf (Not (Start = Empty)) And (Not (Find = Empty)) Then
        InStr = VBA.InStr(Start, Ret, Find, Compare)
    ElseIf (Not IsNumeric(Start)) And (Not (Find = Empty)) Then
        InStr = VBA.InStr(1, Ret, Find, Compare)
    Else
        Err.Raise 13
    End If
End Function

Public Function Size(Optional ByVal InChar As Boolean = True) As Long
    Dim lSize As Long
    CopyMemory lSize, ByVal lPtr, 4
    Size = lSize \ (Abs(InChar) + 1)
End Function

Public Sub Concat(ByRef Value As String, Optional ByVal Reset As Boolean = False)
    If Reset Then Me.Reset
    Dim lSize As Long
    Dim nSize As Long
    nSize = (Len(Value) * 2)
    If nSize > 0 Then
        CopyMemory lSize, ByVal lPtr, 4
        lPtr = HeapReAlloc(GetProcessHeap, 0, lPtr, (nSize + lSize) + 4)
        CopyMemory ByVal lPtr, (nSize + lSize), 4
        CopyMemory ByVal (lPtr + 4) + lSize, ByVal StrPtr(Value), nSize
    End If
End Sub

Public Function Partial(ByVal Start As Long, Optional ByVal Length As Long = -1, Optional ByVal Commit As Boolean = False) As String
    Dim lSize As Long
    Dim Ret As String
    CopyMemory lSize, ByVal lPtr, 4
    Ret = Space(lSize \ 2)
    CopyMemory ByVal StrPtr(Ret), ByVal lPtr + 4, lSize
    If Length = -1 Then
        Ret = Mid(Ret, Start)
    ElseIf Start > 0 And Length > 0 Then
        Ret = Mid(Ret, Start, Length)
    End If
    If Commit Then Concat Ret, True
    Partial = Ret
End Function




Attribute 