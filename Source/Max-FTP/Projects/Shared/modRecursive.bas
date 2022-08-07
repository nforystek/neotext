Attribute VB_Name = "modRecursive"
#Const modRecursive = -1
Option Explicit
'TOP DOWN

Option Compare Binary

' Local Memory Flags
Const LMEM_FIXED = &H0
Const LMEM_MOVEABLE = &H2
Const LMEM_NOCOMPACT = &H10
Const LMEM_NODISCARD = &H20
Const LMEM_ZEROINIT = &H40
Const LMEM_MODIFY = &H80
Const LMEM_DISCARDABLE = &HF00
Const LMEM_VALID_FLAGS = &HF72
Const LMEM_INVALID_HANDLE = &H8000

Const LHND = (LMEM_MOVEABLE + LMEM_ZEROINIT)
Const lPtr = (LMEM_FIXED + LMEM_ZEROINIT)

Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long

Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long

Declare Function InitAtomTable Lib "kernel32" (ByVal nSize As Integer) As Boolean

Declare Function AddAtom Lib "kernel32" Alias "AddAtomA" (ByVal lpString As String) As Integer
Declare Function DeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function FindAtom Lib "kernel32" Alias "FindAtomA" (ByVal lpString As String) As Integer
Declare Function GetAtomName Lib "kernel32" Alias "GetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "MSVBVM60.DLL" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long
    
Public Declare Sub RtlMoveObject Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Object, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub RtlMoveVariant Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Variant, ByVal Source As Long, ByVal Length As Long)

Public typs() As Long
Public refs() As Long
Public vals() As Long
Public stack As Long

Private Atoms As Collection
Public Sub AsyncRecursiveReset()
    ReDim Preserve typs(-1 To -1) As Long
    ReDim Preserve refs(-1 To -1) As Long
    ReDim Preserve vals(-1 To -1) As Long
End Sub
Public Function AsyncRecursiveMakeAtom(ByVal Identity As String, Optional ByVal RefCount As Long, Optional ByVal ValCount As Long, Optional ByVal ReturnVal As Boolean = False) As Integer
    If Atoms Is Nothing Then Set Atoms = New Collection
    Dim args() As Long
    ReDim args(1 To ((RefCount + ValCount) + 1), 1 To 2) As Long
    args(1, 1) = RefCount
    args(1, 2) = ValCount
    Dim hMem As Long
    hMem = LocalAlloc(LMEM_MOVEABLE And VarPtr(hMem), ((((RefCount + ValCount) + 1) * 2) * 4) + 4)
    If hMem <> LocalLock(hMem) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
    RtlMoveMemory ByVal hMem, ByVal VarPtr(ReturnVal), 4
    RtlMoveMemory ByVal hMem + 4, ByVal VarPtr(args(LBound(args, 1), LBound(args, 2))), ((((RefCount + ValCount) + 1) * 2) * 4)
    Atoms.Add hMem, "A" & AddAtom(Identity)
End Function


Public Sub AsyncRecursivePushArgs(ByVal Identity As String)
    'Debug.Print "PUSH " & Identity
    Dim ReturnVal As Long
    Dim RefCount As Long
    Dim ValCount As Long
    Dim lSize As Long
    Dim hMem As Long
    hMem = Atoms.Item("A" & FindAtom(Identity))
    RtlMoveMemory ReturnVal, ByVal hMem, 4
    Dim args() As Long
    lSize = LocalSize(hMem)
    ReDim args(1 To (((lSize - 4) \ 4) \ 2), 1 To 2) As Long
    RtlMoveMemory ByVal VarPtr(args(LBound(args, 1), LBound(args, 2))), ByVal hMem + 4, lSize - 4
    Dim cnt As Long
    Dim idx As Long
    If args(1, 1) > 0 Then
        cnt = 2
        For idx = LBound(refs) + 1 To UBound(refs)
            args(cnt, 1) = refs(idx)
            args(cnt, 2) = typs(cnt - 2)
            'Debug.Print "args(" & cnt & ", 1)=" & refs(idx) & " typs(" & cnt - 2 & ")=" & typs(cnt - 2)
            cnt = cnt + 1
            If cnt - 1 > args(1, 1) Then Exit For
        Next
        cnt = 0
        For idx = LBound(vals) + 1 To UBound(vals)
            args(args(1, 1) + 2 + cnt, 1) = vals(idx)
            args(args(1, 1) + 2 + cnt, 2) = typs(args(1, 1) + cnt)
            'Debug.Print "args(" & args(1, 1) + 2 + cnt & ", 1)=" & vals(idx) & " typs(" & args(1, 1) + cnt & ")=" & typs(args(1, 1) + cnt)
            cnt = cnt + 1
            If cnt + 1 > args(1, 2) Then Exit For
        Next
    End If
    RtlMoveMemory ByVal hMem + 4, ByVal VarPtr(args(LBound(args, 1), LBound(args, 2))), lSize - 4
End Sub

Public Sub AsyncRecursivePokeArgs(ByVal Identity As String)
    'Debug.Print "POKE " & Identity
    Dim ReturnVal As Long
    Dim RefCount As Long
    Dim ValCount As Long
    Dim lSize As Long
    Dim hMem As Long
    hMem = Atoms.Item("A" & FindAtom(Identity))
    RtlMoveMemory ReturnVal, ByVal hMem, 4
    Dim args() As Long
    lSize = LocalSize(hMem)
    ReDim args(1 To (((lSize - 4) \ 4) \ 2), 1 To 2) As Long
    RtlMoveMemory ByVal VarPtr(args(LBound(args, 1), LBound(args, 2))), ByVal hMem + 4, lSize - 4
    Dim cnt As Long
    If args(1, 1) > 0 Then
        For cnt = 1 To args(1, 1)
            ReDim Preserve refs(LBound(refs) To UBound(refs) + 1) As Long
            ReDim Preserve typs(LBound(typs) To UBound(typs) + 1) As Long
            refs(UBound(refs)) = args(cnt + 1, 1)
            typs(UBound(typs)) = args(cnt + 1, 2)
            'Debug.Print refs(UBound(refs)) & "=args(" & cnt + 1 & ", 1) typs(" & UBound(typs) & ")=" & typs(UBound(typs))
        Next
    End If
    refs(-1) = UBound(refs)
    If args(1, 2) > 0 Then
        For cnt = (args(1, 1) + 1) + 1 To args(1, 1) + args(1, 2) + 1
            ReDim Preserve vals(LBound(vals) To UBound(vals) + 1) As Long
            ReDim Preserve typs(LBound(typs) To UBound(typs) + 1) As Long
            vals(UBound(vals)) = args(cnt, 1)
            typs(UBound(typs)) = args(cnt, 2)
            'Debug.Print vals(UBound(vals)) & "=args(" & cnt & ", 1) typs(" & UBound(typs) & ")=" & typs(UBound(typs))
        Next
    End If
    vals(-1) = UBound(vals)
    typs(-1) = UBound(typs)
End Sub

Public Sub AsyncRecursiveDoneAtom(ByVal Identity As String)
    Dim atom As Long
    Dim hMem As Long
    atom = FindAtom(Identity)
    hMem = Atoms.Item("A" & atom)
    LocalUnlock hMem
    LocalFree hMem
    Atoms.Remove "A" & atom
    If DeleteAtom(atom) <> 0 Then Err.Raise 8, App.Title, "Atom delete failed."
    If Atoms.Count = 0 Then Set Atoms = Nothing
End Sub

Public Sub SetByRefArg(Optional ByRef arg)
    ReDim Preserve refs(LBound(refs) To (UBound(refs) + 1)) As Long
    ReDim Preserve typs(LBound(typs) To (UBound(typs) + 1)) As Long

    If IsObject(arg) Then
        refs(UBound(refs)) = ObjPtr(arg)
    Else
        Select Case TypeName(arg)
            Case "String"
                Dim hMem As Long
                hMem = lCreateANSI(arg)
                RtlMoveMemory ByVal VarPtr(refs(UBound(refs))), hMem, 4
                typs(UBound(typs)) = 1
            Case Else
                If IsNumeric(arg) Then
                    refs(UBound(refs)) = arg
                    typs(UBound(typs)) = 2
                End If
        End Select
    End If
End Sub

Public Sub SetByValArg(Optional ByRef arg)
    ReDim Preserve vals(LBound(vals) To UBound(vals) + 1) As Long
    ReDim Preserve typs(LBound(typs) To UBound(typs) + 1) As Long
    If IsObject(arg) Then
        vals(UBound(vals)) = ObjPtr(arg)
    Else
        Select Case TypeName(arg)
            Case "String"
                Dim hMem As Long
                hMem = lCreateANSI(arg)
                RtlMoveMemory ByVal VarPtr(vals(UBound(vals))), hMem, 4
                typs(UBound(typs)) = 1
            Case Else
                If IsNumeric(arg) Then
                    vals(UBound(vals)) = arg
                    typs(UBound(typs)) = 2
                End If
        End Select
    End If
End Sub

Public Sub GetByRefArg(Optional ByRef arg)
    If IsObject(arg) Then
        If refs(refs(-1)) <> 0 Then
            RtlMoveMemory ObjPtr(arg), ByVal refs(refs(-1)), 4
        End If
    Else
        If refs(refs(-1)) <> 0 Then
            Select Case typs(typs(-1))
                Case 1
                    Dim hMem As Long
                    'RtlMoveMemory ByVal VarPtr(hMem), ByVal refs(refs(-1)), 4
                    arg = lStringANSI(refs(refs(-1)))
                    lDestroyANSI hMem
                Case 2
                    arg = refs(refs(-1))
            End Select
        End If
    End If
    refs(-1) = refs(-1) - 1
    typs(-1) = typs(-1) - 1
    ReDim Preserve refs(LBound(refs) To UBound(refs) - 1) As Long
    ReDim Preserve typs(LBound(typs) To UBound(typs) - 1) As Long
End Sub
Public Sub GetByValArg(Optional ByRef arg)
    If IsObject(arg) Then
        If vals(vals(-1)) <> 0 Then
            RtlMoveMemory ObjPtr(arg), ByVal vals(vals(-1)), 4
        End If
    Else
        If vals(vals(-1)) <> 0 Then
            Select Case typs(typs(-1))
                Case 1
                    Dim hMem As Long
                    'RtlMoveMemory ByVal VarPtr(hMem), ByVal vals(vals(-1)), 4
                    arg = lStringANSI(vals(vals(-1)))
                    lDestroyANSI hMem
                Case 2
                    arg = vals(vals(-1))
            End Select
        End If
    End If
    vals(-1) = vals(-1) - 1
    typs(-1) = typs(-1) - 1
    ReDim Preserve vals(LBound(vals) To UBound(vals) - 1) As Long
    ReDim Preserve typs(LBound(typs) To UBound(typs) - 1) As Long
End Sub

Private Sub part1(ByRef arg1 As NTAdvFTP61.url, ByRef arg2 As NTAdvFTP61.Client, ByVal arg3 As String, ByVal arg4 As Long)

    stack = stack + 1
    AsyncRecursiveMakeAtom "testID" & stack, 2, 2, False

    
    Select Case stack
        Case 2
            'push new values to the stacking procedure calls
            arg1.SetFolder "ftp://ftp.neotext.org/newfolder/", "/somethingelse"
            arg2.Account = "differentvalue"
            arg3 = "checking the integrity"
            arg4 = 937
        
        Case 3
            'push new values to the stacking procedure calls
            arg1.SetFolder "ftp://ftp.neotext.org/newfolder/", "/somethingelse"
            arg2.Account = "differentvalue"
            arg3 = "futher check integrity"
            arg4 = 9345
            

    End Select
    
    arg4 = arg4 + 30


    
    'do initial partial of the recursion procedure
    Debug.Print "Part1: " & ObjPtr(arg1) & " " & ObjPtr(arg2) & " " & arg3 & " " & arg4

    AsyncRecursiveReset
    SetByRefArg arg1
    SetByRefArg arg2
    SetByValArg arg3
    SetByValArg arg4
    AsyncRecursivePushArgs "testID" & stack
    
End Sub

Private Sub part2(ByRef arg1 As NTAdvFTP61.url, ByRef arg2 As NTAdvFTP61.Client, ByRef arg3 As String, ByRef arg4 As Long)

    AsyncRecursiveReset
    AsyncRecursivePokeArgs "testID" & stack
    GetByValArg arg4
    GetByValArg arg3
    GetByRefArg arg2
    GetByRefArg arg1
    


        
    'do return call portion the recursion procedure
    Debug.Print "Part2: " & ObjPtr(arg1) & " " & ObjPtr(arg2) & " " & arg3 & " " & arg4



    
    AsyncRecursiveDoneAtom "testID" & stack
    stack = stack - 1


End Sub

'    Dim arg1 As New NTAdvFTP61.url
'    Dim arg2 As New NTAdvFTP61.Client
'    Dim arg3 As String
'    Dim arg4 As Long
'
'    arg1.SetFolder "ftp://ftp.neotext.org/upload/", "/HtDocs"
'    arg2.Account = "RefTest"
'    arg3 = "this is a argument too!"
'    arg4 = 80
'
'    part1 arg1, arg2, arg3, arg4
'    part1 arg1, arg2, arg3, arg4
'    part1 arg1, arg2, arg3, arg4
'
'
'    part2 arg1, arg2, arg3, arg4
'    part2 arg1, arg2, arg3, arg4
'    part2 arg1, arg2, arg3, arg4
'
'End

