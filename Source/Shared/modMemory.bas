Attribute VB_Name = "modMemory"
#Const [True] = -1
#Const [False] = 0

Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public Type SAFEARRAYBOUND ' 8 bytes
    cElements As Long
    lLbound   As Long
End Type

Public Type SAFEARRAYHEADER ' 20 bytes (for one dimensional arrays
    Dimensions    As Integer
    fFeatures     As Integer
    DataSize      As Long
    cLocks        As Long
    dataPointer   As Long
    sab(0 To 0)        As SAFEARRAYBOUND
End Type

Public Const FADF_AUTO = &H1&        '// Array is allocated on the stack.
Public Const FADF_STATIC = &H2&      '// Array is statically allocated.
Public Const FADF_EMBEDDED = &H4&    '// Array is embedded in a structure.
Public Const FADF_FIXEDSIZE = &H10&  '// Array may not be resized or reallocated.
Public Const FADF_BSTR = &H100&      '// An array of BSTRs.
Public Const FADF_UNKNOWN = &H200&   '// An array of IUnknown*.
Public Const FADF_DISPATCH = &H400&  '// An array of IDispatch*.
Public Const FADF_VARIANT = &H800&   '// An array of VARIANTs.
Public Const FADF_RESERVED = &HF0E8& '// Bits reserved for future use.

Public Enum eDATASIZE
    byteArray = 1
    integerArray = 2  ' or Boolean Data Type
    longArray = 4
    singleArray = 4
    doubleArray = 8
End Enum

Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Var() As Any) As Long
Public Declare Function StrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function StrCpyReverse Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias _
    "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
    
Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" _
    (dstObject As Any, ByVal srcObjPtr As Long) As Long
    
Public Type MethodInfo
    MethodPointer As Long
    MethodAddress As Long
End Type

Public Type MemInfo
    MemH As Long
    MemP As Long
End Type

Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_FIXED = &H0
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = &H40

Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GPTR = &H40
Public Const GHND = &H42

Public Const LNOTIFY_OUTOFMEM = 0
Public Const LNOTIFY_MOVE = 1
Public Const LNOTIFY_DISCARD = 2

Public Type MEMORY_BASIC_INFORMATION
     BaseAddress As Long
     AllocationBase As Long
     AllocationProtect As Long
     RegionSize As Long
     State As Long
     Protect As Long
     lType As Long
End Type

Public Type Memory
    Pointer As Long
    Address As Long
End Type

Public Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Public Declare Function VirtualAlloc Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFree Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long

Public Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long

Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleHandle2 Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As Any) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

Public Declare Sub CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByRef Dest As Any, ByVal src As Any)
Public Declare Sub GetMem8 Lib "msvbvm60.dll" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub GetMem4 Lib "msvbvm60.dll" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub PutMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByVal newVal As Long)


Private Function GetPtr(VarVal As Variant) As Long
    GetMem4 ByVal (VarPtr(VarVal) + 8&), GetPtr  ' this gets the pointer held within VarVal
    GetMem4 ByVal GetPtr, GetPtr              ' that the pointer in VarVal points to
    GetMem4 ByVal (GetPtr + 12&), GetPtr
End Function

Public Sub Push(ByRef S() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Length As Long = 0)
    FillMemory ByVal VarPtr(S(Offset + 1)), IIf(Length, Length, (UBound(S) - LBound(S)) + 1), IIf(S(Offset + 1) + 1 > 255, 255, S(Offset + 1) + 1)
End Sub
Public Sub Pop(ByRef S() As Byte, Optional ByVal Offset As Long = 0, Optional ByVal Length As Long = -1)
    FillMemory ByVal VarPtr(S(Offset + 1)), IIf(Length, Length, (UBound(S) - LBound(S)) + 1), IIf(S(Offset + 1) - 1 < 0, 0, S(Offset + 1) - 1)
End Sub

Public Function ObjectLongProperties(ObjectClass) As Memory()
    Dim FPS() As Memory
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(ObjectClass)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4

    Dim SectionOf As Long
    Dim SimpleCount As Long
    Dim ComplexCount As Long
    Dim MethodCount As Long
    
    
    SectionOf = 1
    
    Dim PTX As Long
    Dim cnt As Long
    Select Case SectionOf
        Case 1 'public simple data types
            Do

              '  Debug.Print HeapSize(GetProcessHeap(), GlobalFlags(ByVal (VTable + 27 + (SimpleCount * 2 * 4))), ByVal (VTable + 27 + (SimpleCount * 2 * 4)))
                SimpleCount = SimpleCount + 1
                
                ReDim Preserve FPS(SimpleCount - 1)
             '   For cnt = 0 To SimpleCount - 1
                    PTX = VTable + 23 + ((SimpleCount - 1) * 2 * 4)
                    RtlMoveMemory FPS((SimpleCount - 1)).Pointer, PTX, 4
                    RtlMoveMemory FPS((SimpleCount - 1)).Address, ByVal PTX, 4
                    
                    
                Debug.Print HeapSize(GetProcessHeap(), GlobalFlags(FPS((SimpleCount - 1)).Pointer), FPS((SimpleCount - 1)).Pointer)
       
                    
             '   Next
             '   SimpleCount = SimpleCount + 1
             
             
             
             Loop Until SimpleCount = 20
           ' End If
        Case 2 'public objects and variants
            If ComplexCount > 0 Then
                ReDim FPS(ComplexCount - 1)
                For cnt = 0 To ComplexCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (cnt * 3 * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next

            End If
        Case 3 'public Functions and Subs
            If MethodCount > 0 Then
                ReDim FPS(MethodCount - 1)
                For cnt = 0 To MethodCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (ComplexCount * 3 * 4) + (cnt * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next

            End If

    End Select
    

End Function

Public Function ObjectPointers(ObjectClass, ByVal SectionOf As Long, Optional ByVal SimpleCount As Long, Optional ByVal ComplexCount As Long, Optional ByVal MethodCount As Long) As Memory()
    Dim FPS() As Memory
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(ObjectClass)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4

    Dim PTX As Long
    Dim cnt As Long
    Select Case SectionOf
        Case 1 'public simple data types
            If SimpleCount > 0 Then
                ReDim FPS(SimpleCount - 1)
                For cnt = 0 To SimpleCount - 1
                    PTX = VTable + 28 + (cnt * 2 * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next

            End If
        Case 2 'public objects and variants
            If ComplexCount > 0 Then
                ReDim FPS(ComplexCount - 1)
                For cnt = 0 To ComplexCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (cnt * 3 * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next

            End If
        Case 3 'public Functions and Subs
            If MethodCount > 0 Then
                ReDim FPS(MethodCount - 1)
                For cnt = 0 To MethodCount - 1
                    PTX = VTable + 28 + (SimpleCount * 2 * 4) + (ComplexCount * 3 * 4) + (cnt * 4)
                    RtlMoveMemory FPS(cnt).Pointer, PTX, 4
                    RtlMoveMemory FPS(cnt).Address, ByVal PTX, 4
                Next

            End If

    End Select
    
    ObjectPointers = FPS
End Function

Public Function ArraySize(InArray, Optional ByVal InBytes As Boolean = False) As Long
On Error GoTo 0
On Error GoTo -1
On Error Resume Next
On Error GoTo dimerror
On Local Error Resume Next
On Local Error GoTo dimerror:
    Static dimcheck As Long
    dimcheck = 0
    If UBound(InArray) = -1 Or LBound(InArray) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray) + -CInt(Not CBool(-LBound(InArray)))) * IIf(InBytes, LenB(InArray(LBound(InArray))), 1)
    End If
    Exit Function
startover:
'On Error GoTo 0
'On Error Resume Next
'On Local Error Resume Next
On Error GoTo 0
On Error GoTo -1
On Error Resume Next
On Error GoTo dimerror
On Local Error Resume Next
On Local Error GoTo dimerror:
'On Local Error GoTo dimerror:

    If UBound(InArray, dimcheck) = -1 Or LBound(InArray, dimcheck) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray, dimcheck) + -CInt(Not CBool(-LBound(InArray, dimcheck)))) * IIf(InBytes, LenB(InArray(LBound(InArray, dimcheck), LBound(InArray, dimcheck - 1))), 1)
    End If
    
    Exit Function
dimerror:
    If dimcheck = 0 Then
        dimcheck = 1
        Err.Clear
        GoTo startover
    End If
    ArraySize = 0
End Function
Public Function Convert(Info)
    'slow method of converting byte
    'array to string nd vice versa
    Dim N As Long
    Dim out() As Byte
    Dim ret As String
    Select Case TypeName(Info)
        Case "String"
            If Len(Info) > 0 Then
                ReDim out(0 To Len(Info) - 1) As Byte
                For N = 0 To Len(Info) - 1
                    out(N) = Asc(Mid(Info, N + 1, 1))
                Next
            Else
                ReDim out(-1 To -1) As Byte
            End If
            Convert = out
        Case "Byte()"
            If (ArraySize(Info) > 0) Then
                For N = LBound(Info) To UBound(Info)
                    ret = ret & Chr(Info(N))
                Next
            End If
            Convert = ret
    End Select
End Function

' + ArrayPtr ++++++++++++++++++++++++Rd+
' This function returns a pointer to the
' SAFEARRAY header of any Visual Basic
' array, including a Visual Basic string
' array.
' Substitutes both ArrPtr and StrArrPtr.
' This function will work with vb5 or
' vb6 without modification.
Public Function ArrayPtr(Arr) As Long
    ' Thanks to Francesco Balena and Monte Hansen
    Dim iDataType As Integer
    On Error GoTo UnInit
    RtlMoveMemory iDataType, Arr, 2& ' get the real VarType of the argument, this is similar to VarType(), but returns also the VT_BYREF bit
    If (iDataType And vbArray) = vbArray Then ' if a valid array was passed
        RtlMoveMemory ArrayPtr, ByVal VarPtr(Arr) + 8&, 4& ' get the address of the SAFEARRAY descriptor stored in the second half of the Variant parameter that has received the array. Thanks to Francesco Balena.
    End If
    Exit Function
UnInit:
    If Err Then Err.Clear
End Function
' ++++++++++++++++++++++++++++++++++++++


Public Function CreateANSI(Optional ByVal StringX As String = "") As Long 'global
    'creates the ansi memory for a stringx
    Dim out As Long
    out = GlobalAlloc(GMEM_MOVEABLE And VarPtr(out), LenB(StringX))
    If out <> GlobalLock(out) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
    RtlMoveMemory ByVal out, StringX, LenB(StringX)
    CreateANSI = out
End Function

Public Function StringANSI(ByRef ansi As Long) As String 'global
    'returns the string value of the ansi
    StringANSI = String(StrLen(ansi), Chr(0))
    StrCpy StringANSI, ansi
End Function

Public Sub AppendANSI(ByRef ansi As Long, ByVal StringX As String) 'global
    'concatenates the stringx to the end of the ansi
    GlobalUnlock ansi
    ansi = GlobalReAlloc(ansi, StrLen(ansi) + LenB(StringX), GMEM_MOVEABLE Or GPTR)
    If ansi <> GlobalLock(ansi) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
    StrCpyReverse ansi, StringANSI(ansi) & StringX
End Sub

Public Function LengthANSI(ByRef ansi As Long) As Long
    'returns the string length of the ansi
    LengthANSI = StrLen(ansi)
End Function

Public Sub DestroyANSI(ByRef ansi As Long) 'global
    'disposes of the ansi memory
    GlobalUnlock ansi
    GlobalFree ansi
End Sub

Public Function DeployANSI(ByRef ansi As Long) As String 'global
    'returns the string of the ansi disposing of the memory
    DeployANSI = String(StrLen(ansi), Chr(0))
    StrCpy DeployANSI, ansi
    GlobalUnlock ansi
    GlobalFree ansi
End Function

Public Function ArrayOfANSI(ByRef ansi As Long) As Byte() 'global
    'converts the ansi to array making it able native array use
    Dim out() As Byte
    Dim llen As Long
    llen = LengthANSI(ansi)
    If llen = 0 Then
        ReDim out(-1 To -1) As Byte
    Else
        ReDim out(1 To llen) As Byte
        RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi, llen
    End If
    ArrayOfANSI = out
End Function

Public Function ArrayToANSI(ByRef ansi() As Byte) As Long 'global
    'converts the array to ansi making it able the local ansi use
    Dim out As Long
    out = GlobalAlloc(GMEM_MOVEABLE Or VarPtr(out), ArraySize(ansi))
    If out <> GlobalLock(out) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
    RtlMoveMemory ByVal out, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
    ArrayToANSI = out
End Function

Public Function lCreateANSI(Optional ByVal StringX As String = "") As Long 'local
    'creates the ansi memory for a stringx
    Dim out As Long
    out = LocalAlloc(LMEM_MOVEABLE And VarPtr(out), LenB(StringX))
    If out <> LocalLock(out) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
    RtlMoveMemory ByVal out, StringX, LenB(StringX)
    lCreateANSI = out
End Function

Public Function lStringANSI(ByRef ansi As Long) As String 'local
    'returns the string value of the ansi
    lStringANSI = String(StrLen(ansi), Chr(0))
    StrCpy lStringANSI, ansi
End Function

Public Sub lAppendANSI(ByRef ansi As Long, ByVal StringX As String) 'local
    'concatenates the stringx to the end of the ansi
    LocalUnlock ansi
    ansi = LocalReAlloc(ansi, StrLen(ansi) + LenB(StringX), LMEM_MOVEABLE Or lPtr)
    If ansi <> LocalLock(ansi) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
    StrCpyReverse ansi, lStringANSI(ansi) & StringX
End Sub

Public Function lLengthANSI(ByRef ansi As Long) As Long 'local
    'returns the string length of the ansi
    lLengthANSI = StrLen(ansi)
End Function

Public Sub lDestroyANSI(ByRef ansi As Long) 'local
    'disposes of the ansi memory
    LocalUnlock ansi
    LocalFree ansi
End Sub

Public Function lDeployANSI(ByRef ansi As Long) As String 'local
    'returns the string of the ansi disposing of the memory
    lDeployANSI = String(StrLen(ansi), Chr(0))
    StrCpy lDeployANSI, ansi
    LocalUnlock ansi
    LocalFree ansi
End Function

Public Function lArrayOfANSI(ByRef ansi As Long) As Byte() 'local
    'converts the ansi to array making it able native array use
    Dim out() As Byte
    Dim llen As Long
    llen = lLengthANSI(ansi)
    If llen = 0 Then
        ReDim out(-1 To -1) As Byte
    Else
        ReDim out(1 To llen) As Byte
        RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi, llen
    End If
    lArrayOfANSI = out
End Function

Public Function lArrayToANSI(ByRef ansi() As Byte) As Long 'local
    'converts the array to ansi making it able the local ansi use
    Dim out As Long
    out = LocalAlloc(LMEM_MOVEABLE And VarPtr(out), ArraySize(ansi))
    If out <> LocalLock(out) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
    RtlMoveMemory ByVal out, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
    lArrayToANSI = out
End Function

Public Function nArrayOfString(ByVal str As String) As Byte()
    'transfer of string to array using ansi as the vessle
    Dim out() As Byte
    If str = "" Then
        ReDim out(-1 To -1) As Byte
    Else
        Dim ansi As Long
        ansi = lCreateANSI(str)
        ReDim out(1 To lLengthANSI(ansi)) As Byte
        RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi, lLengthANSI(ansi)
        lDestroyANSI ansi
    End If
    nArrayOfString = out
End Function

Public Function nArrayToString(ByRef ary() As Byte) As String
    'transfer of array to string using ansi as the vessle
    Dim ansi As Long
   ' ansi = lCreateANSI(String(ArraySize(ary), Chr(0)))
    ansi = LocalAlloc(LMEM_MOVEABLE And VarPtr(ansi), ArraySize(ary))
    If ansi <> LocalLock(ansi) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
    RtlMoveMemory ByVal ansi, ByVal VarPtr(ary(LBound(ary))), ArraySize(ary)
    nArrayToString = lStringANSI(ansi)
    lDestroyANSI ansi
End Function

Public Function RedimArray(ByVal DataSize As Long, ByVal lNumElements As Long, ByRef sa As SAFEARRAYHEADER, ByVal lDataPointer As Long, ByVal lArrayPointer As Long, Optional LoBound As Long = 0) As Long
  If lNumElements > 0 And lDataPointer <> 0 And lArrayPointer <> 0 Then
    With sa
      .DataSize = DataSize                              ' byte = 1 byte, integer = 2 bytes etc
      .Dimensions = 1 '2                                ' one dimensional
      .dataPointer = lDataPointer                       ' to unicode string data (or other?)
      .sab(0).lLbound = LoBound                         ' lower bound
      .sab(0).cElements = lNumElements                  ' number of elements
      '.sab(1).cElements = lNumElements
      '.sab(1).lLbound = LoBound
      RtlMoveMemory ByVal lArrayPointer, VarPtr(sa), 4& ' fake VB out
      RedimArray = True
    End With
  End If
End Function

Public Sub DestroyArray(ByVal lArrayPointer As Long)
  Dim lZero As Long
  RtlMoveMemory ByVal lArrayPointer, lZero, 4         ' put the array back to its original state
End Sub

Public Function GetObjectFunctionsPointers(obj As Object, ByVal NumberOfMethods As Long, Optional ByVal PublicVarNumber As Long, Optional ByVal PublicObjVariantNumber As Long) As MethodInfo()
    Dim FPS() As MethodInfo
    ReDim FPS(NumberOfMethods - 1)
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(obj)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4
    Dim PTX As Long
    Dim cnt As Long
    For cnt = 0 To NumberOfMethods - 1
        PTX = VTable + 28 + (PublicVarNumber * 2 * 4) + (PublicObjVariantNumber * 3 * 4) + cnt * 4
        RtlMoveMemory FPS(cnt).MethodPointer, PTX, 4
        RtlMoveMemory FPS(cnt).MethodAddress, ByVal PTX, 4
    Next
    GetObjectFunctionsPointers = FPS
End Function

Public Function AddObjectFunctionsPointers(ByRef FPS() As MethodInfo, obj As Object, ByVal NumberOfMethods As Long, Optional ByVal PublicVarNumber As Long, Optional ByVal PublicObjVariantNumber As Long) As Long
    Dim ubnds As Long
    ubnds = UBound(FPS)
    AddObjectFunctionsPointers = ubnds + 1
    ReDim FPS(ubnds + NumberOfMethods)
    Dim OBJ1 As Long
    OBJ1 = ObjPtr(obj)
    Dim VTable As Long
    RtlMoveMemory VTable, ByVal OBJ1, 4
    Dim PTX As Long
    Dim cnt As Long
    For cnt = 1 To NumberOfMethods
        PTX = VTable + 28 + (PublicVarNumber * 2 * 4) + (PublicObjVariantNumber * 3 * 4) + cnt * 4
        RtlMoveMemory FPS(ubnds + cnt).MethodPointer, PTX, 4
        RtlMoveMemory FPS(ubnds + cnt).MethodAddress, ByVal PTX, 4
    Next
End Function



