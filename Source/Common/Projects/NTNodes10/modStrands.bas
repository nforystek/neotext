Attribute VB_Name = "modStrands"
#Const modStrands = -1
Option Explicit
Option Compare Binary
Option Private Module

Public Type MemInfo
    MemP As Long
    MemH As Long
End Type

Public Const MirrorPath_Temp = "(Temp)"
Public Const MirrorPath_None = "(None)"

'Public Const StateAccess = 0 'Default behavior of ram construct conductly
'Public Const StateMirror = -2 'Coupled behavior of direct to file access
'
'Public Const RemmitSioux = 0  'disposes of accessed information as goes
'Public Const CommitFirms = 1 'commits the creating so future access keeps
'
'Public Const ScopeNormal = 0 'normal virtual memory heap
'Public Const ScopeLocale = 4 'locale virtual memory heap
'Public Const ScopeGlobal = -4 'global virtual disk drive
'
'Public Const Memory_Remmit = StateAccess Or RemmitSioux '0 'virtualmemory and heap Sioux to scope
'Public Const Mirror_Remmit = StateMirror Or RemmitSioux '-2 'direct file and heap Sioux to scope
'
'Public Const Memory_Commit = StateAccess Or CommitFirms ' 1 'virtualmemory and heap Firms to scope
'Public Const Mirror_Commit = StateMirror Or CommitFirms '-1 'direct file and heap Firms to scope
'
'Public Const Normal_Memory_Remmit = ScopeNormal Or StateAccess Or RemmitSioux     ' 0 'virtualmemory and heap Sioux to scope
'Public Const Locale_Memory_Remmit = ScopeLocale Or StateAccess Or RemmitSioux     ' 4 'virtualmemory and heap Sioux to scope
'Public Const Global_Memory_Remmit = ScopeGlobal Or StateAccess Or RemmitSioux     ' -4 'virtualmemory and heap Sioux to scope
'
'Public Const Normal_Mirror_Remmit = ScopeNormal Or StateMirror Or RemmitSioux     '-2 'direct file and heap Sioux to scope
'Public Const Locale_Mirror_Remmit = ScopeLocale Or StateMirror Or RemmitSioux     ' 2 'direct file and heap Sioux to scope
'Public Const Global_Mirror_Remmit = ScopeGlobal Or StateMirror Or RemmitSioux     '-6 'direct file and heap Sioux to scope
'
'Public Const Normal_Memory_Commit = ScopeNormal Or StateAccess Or CommitFirms     ' 1 'virtualmemory and heap Firms to scope
'Public Const Locale_Memory_Commit = ScopeLocale Or StateAccess Or CommitFirms     ' 5 'virtualmemory and heap Firms to scope
'Public Const Global_Memory_Commit = ScopeGlobal Or StateAccess Or CommitFirms     '-3 'virtualmemory and heap Firms to scope
'
'Public Const Normal_Mirror_Commit = ScopeNormal Or StateMirror Or CommitFirms     '-1  'direct file and heap Firms to scope
'Public Const Locale_Mirror_Commit = ScopeLocale Or StateMirror Or CommitFirms     '3 'direct file and heap Firms to scope
'Public Const Global_Mirror_Commit = ScopeGlobal Or StateMirror Or CommitFirms     '-5  'direct file and

Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40

Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_FIXED = &H0

Public Const HEAP_NO_SERIALIZE = &H1

Public Const LMEM_MOVEABLE = &H2
Public Const lPtr = &H40

Public Const GMEM_MOVEABLE = &H2
Public Const GPTR = &H40

Public Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (ary() As Any) As Long
 
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

'&H11000000
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Public Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32" () As Long

Public Const READ_CONTROL = &H20000

Public Const SYNCHRONIZE = &H100000
Public Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Public Const FILE_READ_DATA = (&H1)                     '  file pipe
Public Const FILE_READ_EA = (&H8)                       '  file directory
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const FILE_APPEND_DATA = (&H4)                   '  file
Public Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all
Public Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Public Const FILE_WRITE_EA = (&H10)                     '  file directory
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)
Public Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)

Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000

Public Const INVALID_SET_FILE_POINTER = -1

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const CREATE_ALWAYS = 2
Public Const OPEN_ALWAYS = 4
Public Const CREATE_NEW = 1

Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Public Declare Function VarPtrArray Lib "MSVBVM60.DLL" Alias "VarPtr" (Var() As Any) As Long
Public Declare Function StrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function StrCpyReverse Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Public Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Public Declare Function hWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function hReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Public Const MAX_PATH = 260

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

'Public Const Num_128 = &H80
'Public Const Num_255 = &HFF&
'Public Const Num_256 = &H100&
'Public Const Num_32767 = &H7FFF
'Public Const Num_32768 = &H8000&
'Public Const Num_Neg_32768 = &H8000
'Public Const Num_65280 = &HFF00&
'Public Const Num_65535 = &HFFFF&
'Public Const Num_65536 = &H10000
'Public Const Num_Neg_65536 = &HFFFF0000
'Public Const Num_142606336 = &H8800000
'Public Const Num_285212671 = &H10FFFFFF
'Public Const Num_285212672 = &H11000000
'Public Const Num_2147418112 = &H7FFF0000
'Public Const Num_Neg_2147483648 = &H80000000
'Public Const Num_Neg_285212672 = &HEF000000

'Public Function GetByteArray(ByRef lPtr As Long) As Byte()
'    Dim lZero As Long
'    Dim NewObj() As Byte
'    RtlMoveMemory AryPtr(NewObj), ByVal lPtr, 4&
'    GetObject = NewObj
'    RtlMoveMemory AryPtr(NewObj), ByVal lZero, 4&
'End Function

Public Function AppEXE(Optional ByVal TitleOnly As Boolean = True, Optional ByVal RootEXEOf As Boolean = False) As String
    Dim lpTemp As String
    Dim nLen As Long
    lpTemp = Space(256)
    If RootEXEOf Then
        nLen = GetModuleFileName(0&, lpTemp, Len(lpTemp))
        lpTemp = Left(lpTemp, nLen)
    Else
        nLen = GetModuleFileName(GetModuleHandle(App.EXEName), lpTemp, Len(lpTemp))
        lpTemp = Left(lpTemp, nLen)
    End If
    If TitleOnly And InStrRev(lpTemp, "\") > 0 Then
        lpTemp = Mid(lpTemp, InStrRev(lpTemp, "\") + 1)
    End If
    If TitleOnly And InStrRev(lpTemp, ".") = Len(lpTemp) - 3 Then
        lpTemp = Left(lpTemp, InStrRev(lpTemp, ".") - 1)
    End If
    AppEXE = Trim(lpTemp)
End Function

'Public Property Get LoWord(ByRef lThis As Long) As Integer
'    LoWord = (lThis And Num_32767) Or (Num_Neg_32768 And ((lThis And Num_Neg_32768) = Num_Neg_32768))
'End Property
'Public Property Get HiWord(ByRef lThis As Long) As Integer
'    HiWord = ((lThis And Num_2147418112) \ Num_65536) Or (Num_Neg_2147483648 And (lThis < 0))
'End Property
'Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Integer)
'    lThis = lThis And Not Num_65535 Or lLoWord
'End Property
'Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Integer)
'    If (lHiWord And Num_32768) = Num_32768 Then
'       lThis = lThis And Not Num_Neg_65536 Or ((lHiWord And Num_32767) * Num_65536) Or Num_Neg_2147483648
'    Else
'       lThis = lThis And Not Num_Neg_65536 Or (lHiWord * Num_65536)
'    End If
'End Property

Public Property Get LoWord(ByRef lThis As Long) As Long
   LoWord = (lThis And &HFFFF&)
End Property

Public Property Let LoWord(ByRef lThis As Long, ByVal lLoWord As Long)
   lThis = lThis And Not &HFFFF& Or lLoWord
End Property

Public Property Get HiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      HiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      HiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Property

Public Property Let HiWord(ByRef lThis As Long, ByVal lHiWord As Long)
   If (lHiWord And &H8000&) = &H8000& Then
      lThis = lThis And Not &HFFFF0000 Or ((lHiWord And &H7FFF&) * &H10000) Or &H80000000
   Else
      lThis = lThis And Not &HFFFF0000 Or (lHiWord * &H10000)
   End If
End Property

Public Function GetWinDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(MAX_PATH, Chr(0))
    ret = GetWindowsDirectory(winDir, MAX_PATH)
    winDir = Trim(Replace(winDir, Chr(0), ""))
    If Trim(Dir(winDir, vbDirectory)) = "" Then winDir = App.path
    If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    GetWinDir = winDir
End Function

Public Function GetWinTempDir() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetTempPath(255, winDir)
    If (ret <> 16) And (ret <> 34) Then
        winDir = GetWinDir()
        If LCase(Dir(winDir & "TEMP", vbDirectory)) = "" Then
            MkDir winDir + "TEMP"
        End If
        winDir = winDir + "TEMP\"
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetWinTempDir = winDir
End Function

Public Function GetTemporaryFile() As String
    Dim winDir As String
    Dim ret As Long
    winDir = String(255, Chr(0))
    ret = GetTempFileName(GetWinTempDir, App.Title, 0, winDir)
    If ret = 0 Then
        winDir = GetWinTempDir & "\" & Left(Left(App.Title, 3) & Hex(CLng(Mid(CStr(Rnd), 3))), 14) & ".tmp"
        ret = FreeFile
        Open winDir For Output As #ret
        Close #ret
    Else
        winDir = Trim(Replace(winDir, Chr(0), ""))
    End If
    GetTemporaryFile = winDir
End Function

Public Function Char(Optional ByVal Value As Byte = 10) As Byte()
    Dim tmp(0 To 0) As Byte
    tmp(0) = Value
    Char = tmp
End Function

Public Function ArraySize(InArray, Optional ByVal InBytes As Boolean = False) As Long
On Error GoTo dimerror

    Static dimcheck As Long

    If UBound(InArray) = -1 Or LBound(InArray) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray) + -CInt(Not CBool(-LBound(InArray)))) * IIf(InBytes, LenB(InArray(LBound(InArray))), 1)
    End If
    Exit Function
startover:
    If UBound(InArray, dimcheck) = -1 Or LBound(InArray, dimcheck) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray, dimcheck) + -CInt(Not CBool(-LBound(InArray, dimcheck)))) * IIf(InBytes, LenB(InArray(LBound(InArray, dimcheck), LBound(InArray, dimcheck - 1))), 1)
    End If

    Exit Function
dimerror:
    If dimcheck = 0 Then
        dimcheck = 2
        Err.Clear
        GoTo startover
    End If
    ArraySize = 0
End Function

Public Function Convert(Info)
    Dim N As Long
    Dim out() As Byte
    Dim ret As String
    Select Case VBA.TypeName(Info)
        Case "String"
            If Len(Info) > 0 Then
                ReDim out(0 To Len(Info) - 1) As Byte
                For N = 0 To Len(Info) - 1
                    out(N) = Asc(Mid(Info, N + 1, 1))
                Next
            End If
            Convert = out
        Case "Byte()"
            If (ArraySize(Info) > 0) Then
                On Error GoTo dimcheck
                For N = LBound(Info) To UBound(Info)
                    ret = ret & Chr(Info(N))
                Next
            End If
            Convert = ret
    End Select
    Exit Function
dimcheck:
    If Err Then Err.Clear
    For N = LBound(Info, 2) To UBound(Info, 2)
        ret = ret & Chr(Info(0, N))
    Next
    Convert = ret
End Function


'Public Function StringANSI(ByRef ansi As MemInfo) As String 'global
'    'returns the string value of the ansi
'    StringANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy StringANSI, ansi.MemP
'End Function
'
'Public Sub AppendANSI(ByRef ansi As MemInfo, ByVal StringX As String) 'global
'    'concatenates the stringx to the end of the ansi
'    GlobalUnlock ansi.MemP
'    ansi.MemH = GlobalReAlloc(ansi.MemH, StrLen(ansi.MemP) + LenB(StringX), GMEM_MOVEABLE Or GPTR)
'    ansi.MemP = GlobalLock(ansi.MemH)
'    StrCpyReverse ansi.MemP, StringANSI(ansi) & StringX
'End Sub
'
'Public Function LengthANSI(ByRef ansi As MemInfo) As Long
'    'returns the string length of the ansi
'    LengthANSI = StrLen(ansi.MemP)
'End Function
'
'Public Sub DestroyANSI(ByRef ansi As MemInfo) 'global
'    'disposes of the ansi memory
'    GlobalUnlock ansi.MemP
'    GlobalFree ansi.MemH
'End Sub
'
'Public Function DeployANSI(ByRef ansi As MemInfo) As String 'global
'    'returns the string of the ansi disposing of the memory
'    DeployANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy DeployANSI, ansi.MemP
'    GlobalUnlock ansi.MemP
'    GlobalFree ansi.MemH
'End Function
'
'Public Function ArrayOfANSI(ByRef ansi As MemInfo) 'global
'    'converts the ansi to array making it able native array use
'    Dim out() As Byte
'    ReDim out(0 To LengthANSI(ansi) - 1) As Byte
'    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, LengthANSI(ansi)
'    ArrayOfANSI = out
'End Function
'
'Public Function ArrayToANSI(ByRef ansi() As Byte) As MemInfo 'global
'    'converts the array to ansi making it able the local ansi use
'    ArrayToANSI.MemH = GlobalAlloc(GMEM_MOVEABLE Or GPTR, ArraySize(ansi))
'    ArrayToANSI.MemP = GlobalLock(ArrayToANSI.MemH)
'    RtlMoveMemory ByVal ArrayToANSI.MemP, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
'End Function
'
'Public Function CreateANSI(Optional ByVal StringX As String = "") As MemInfo 'global
'    'creates the ansi memory for a stringx
'    CreateANSI.MemH = GlobalAlloc(GMEM_MOVEABLE Or GPTR, LenB(StringX))
'    CreateANSI.MemP = GlobalLock(CreateANSI.MemH)
'    RtlMoveMemory ByVal CreateANSI.MemP, ByVal StringX, LenB(StringX)
'End Function
'
Public Function lCreateANSI(Optional ByVal StringX As String = "") As MemInfo 'local
    'creates the ansi memory for a stringx
    lCreateANSI.MemH = LocalAlloc((LMEM_FIXED Or lPtr), LenB(StringX))
    lCreateANSI.MemP = LocalLock(lCreateANSI.MemH)
    RtlMoveMemory ByVal lCreateANSI.MemP, ByVal StringX, LenB(StringX)
End Function
'
Public Function lStringANSI(ByRef ansi As MemInfo) As String 'local
    'returns the string value of the ansi
    lStringANSI = String(StrLen(ansi.MemP), Chr(0))
    StrCpy lStringANSI, ansi.MemP
End Function

'Public Sub lAppendANSI(ByRef ansi As MemInfo, ByVal StringX As String) 'local
'    'concatenates the stringx to the end of the ansi
'    LocalUnlock ansi.MemP
'    ansi.MemH = LocalReAlloc(ansi.MemH, StrLen(ansi.MemP) + LenB(StringX), (&H0 Or &H40))
'    ansi.MemP = LocalLock(ansi.MemH)
'    StrCpyReverse ansi.MemP, lStringANSI(ansi) & StringX
'End Sub

Public Function lLengthANSI(ByRef ansi As MemInfo) As Long 'local
    'returns the string length of the ansi
    lLengthANSI = StrLen(ansi.MemP)
End Function

Public Sub lDestroyANSI(ByRef ansi As MemInfo) 'local
    'disposes of the ansi memory
    LocalUnlock ansi.MemP
    LocalFree ansi.MemH
End Sub

'Public Function lDeployANSI(ByRef ansi As MemInfo) As String 'local
'    'returns the string of the ansi disposing of the memory
'    lDeployANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy lDeployANSI, ansi.MemP
'    LocalUnlock ansi.MemP
'    LocalFree ansi.MemH
'End Function

Public Function lArrayOfANSI(ByRef ansi As MemInfo) As Byte() 'local
    'converts the ansi to array making it able native array use
    On Error GoTo doresume
    Dim out() As Byte
    Dim llen As Long
    If llen = 0 Then llen = lLengthANSI(ansi)
    ReDim out(0 To llen - 1) As Byte
    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, llen
    lArrayOfANSI = out
finalnot:
    Exit Function
doresume:
    On Error GoTo finalnot
    Err.Clear
    Resume
End Function

Public Function lArrayToANSI(ByRef ansi() As Byte) As MemInfo 'local
    'converts the array to ansi making it able the local ansi use
    lArrayToANSI.MemH = LocalAlloc(LMEM_FIXED Or lPtr, ArraySize(ansi))
    lArrayToANSI.MemP = LocalLock(lArrayToANSI.MemH)
    RtlMoveMemory ByVal lArrayToANSI.MemP, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
End Function
'
Public Function nArrayOfString(ByVal str As String) As Byte()
    'transfer of string to array using ansi as the vessle
    Dim out() As Byte
    Dim ansi As MemInfo
    Dim llen As Long
    llen = StrLen(ansi.MemP)
    ansi = lCreateANSI(str)
    ReDim out(0 To llen - 1) As Byte
    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, llen
    nArrayOfString = out
    lDestroyANSI ansi
End Function

Public Function nArrayToString(ByRef ary() As Byte) As String
    'transfer of array to string using ansi as the vessle
    Dim ansi As MemInfo
    ansi.MemH = LocalAlloc(GMEM_MOVEABLE Or lPtr, ArraySize(ary))
    ansi.MemP = LocalLock(ansi.MemH)
    RtlMoveMemory ByVal ansi.MemP, ByVal VarPtr(ary(LBound(ary))), ArraySize(ary)
    nArrayToString = lStringANSI(ansi)
    lDestroyANSI ansi
End Function
'Public Function nArrayToString(ByRef ary() As Byte) As String
'    'transfer of array to string using ansi as the vessle
'    Dim ansi As MemInfo
'    ansi.MemH = LocalAlloc((&H0 Or &H40), ArraySize(ary))
'    ansi.MemP = LocalLock(ansi.MemH)
'    RtlMoveMemory ByVal ansi.MemP, ByVal VarPtr(ary(LBound(ary))), ArraySize(ary)
'    nArrayToString = lStringANSI(ansi)
'    lDestroyANSI ansi
'End Function
'

'
'Public Function CreateANSI(Optional ByVal StringX As String = "") As MemInfo 'global
'    'creates the ansi memory for a stringx
'    CreateANSI.MemH = GlobalAlloc(GMEM_MOVEABLE Or GPTR, LenB(StringX))
'    CreateANSI.MemP = GlobalLock(CreateANSI.MemH)
'    RtlMoveMemory ByVal CreateANSI.MemP, ByVal StringX, LenB(StringX)
'End Function
'
'Public Function StringANSI(ByRef ansi As MemInfo) As String 'global
'    'returns the string value of the ansi
'    StringANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy StringANSI, ansi.MemP
'End Function
'
'Public Sub AppendANSI(ByRef ansi As MemInfo, ByVal StringX As String) 'global
'    'concatenates the stringx to the end of the ansi
'    GlobalUnlock ansi.MemP
'    ansi.MemH = GlobalReAlloc(ansi.MemH, StrLen(ansi.MemP) + LenB(StringX), GMEM_MOVEABLE Or GPTR)
'    ansi.MemP = GlobalLock(ansi.MemH)
'    StrCpyReverse ansi.MemP, StringANSI(ansi) & StringX
'End Sub
'
'Public Function LengthANSI(ByRef ansi As MemInfo) As Long
'    'returns the string length of the ansi
'    LengthANSI = StrLen(ansi.MemP)
'End Function
'
'Public Sub DestroyANSI(ByRef ansi As MemInfo) 'global
'    'disposes of the ansi memory
'    GlobalUnlock ansi.MemP
'    GlobalFree ansi.MemH
'End Sub
'
'Public Function DeployANSI(ByRef ansi As MemInfo) As String 'global
'    'returns the string of the ansi disposing of the memory
'    DeployANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy DeployANSI, ansi.MemP
'    GlobalUnlock ansi.MemP
'    GlobalFree ansi.MemH
'End Function
'
'Public Function ArrayOfANSI(ByRef ansi As MemInfo) 'global
'    'converts the ansi to array making it able native array use
'    Dim out() As Byte
'    ReDim out(0 To LengthANSI(ansi)) As Byte
'    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, LengthANSI(ansi)
'    ArrayOfANSI = out
'End Function
'
'Public Function ArrayToANSI(ByRef ansi() As Byte) As MemInfo 'global
'    'converts the array to ansi making it able the local ansi use
'    ArrayToANSI.MemH = GlobalAlloc(GMEM_MOVEABLE Or GPTR, ArraySize(ansi))
'    ArrayToANSI.MemP = GlobalLock(ArrayToANSI.MemH)
'    RtlMoveMemory ByVal ArrayToANSI.MemP, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
'End Function
'
'Public Function lCreateANSI(Optional ByVal StringX As String = "") As MemInfo 'local
'    'creates the ansi memory for a stringx
'    lCreateANSI.MemH = LocalAlloc((&H0 Or &H40), LenB(StringX))
'    lCreateANSI.MemP = LocalLock(lCreateANSI.MemH)
'    RtlMoveMemory ByVal lCreateANSI.MemP, ByVal StringX, LenB(StringX)
'End Function
'
'Public Function lStringANSI(ByRef ansi As MemInfo) As String 'local
'    'returns the string value of the ansi
'    lStringANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy lStringANSI, ansi.MemP
'End Function
'
'Public Sub lAppendANSI(ByRef ansi As MemInfo, ByVal StringX As String) 'local
'    'concatenates the stringx to the end of the ansi
'    LocalUnlock ansi.MemP
'    ansi.MemH = LocalReAlloc(ansi.MemH, StrLen(ansi.MemP) + LenB(StringX), (&H0 Or &H40))
'    ansi.MemP = LocalLock(ansi.MemH)
'    StrCpyReverse ansi.MemP, lStringANSI(ansi) & StringX
'End Sub
'
'Public Function lLengthANSI(ByRef ansi As MemInfo) As Long 'local
'    'returns the string length of the ansi
'    lLengthANSI = StrLen(ansi.MemP)
'End Function
'
'Public Sub lDestroyANSI(ByRef ansi As MemInfo) 'local
'    'disposes of the ansi memory
'    LocalUnlock ansi.MemP
'    LocalFree ansi.MemH
'End Sub
'
'Public Function lDeployANSI(ByRef ansi As MemInfo) As String 'local
'    'returns the string of the ansi disposing of the memory
'    lDeployANSI = String(StrLen(ansi.MemP), Chr(0))
'    StrCpy lDeployANSI, ansi.MemP
'    LocalUnlock ansi.MemP
'    LocalFree ansi.MemH
'End Function
'
'Public Function lArrayOfANSI(ByRef ansi As MemInfo) As Byte() 'local
'    'converts the ansi to array making it able native array use
'    Dim out() As Byte
'    ReDim out(0 To lLengthANSI(ansi)) As Byte
'    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, lLengthANSI(ansi)
'    lArrayOfANSI = out
'End Function
'
'Public Function lArrayToANSI(ByRef ansi() As Byte) As MemInfo 'local
'    'converts the array to ansi making it able the local ansi use
'    lArrayToANSI.MemH = LocalAlloc((&H0 Or &H40), ArraySize(ansi))
'    lArrayToANSI.MemP = LocalLock(lArrayToANSI.MemH)
'    RtlMoveMemory ByVal lArrayToANSI.MemP, ByVal VarPtr(ansi(LBound(ansi))), ArraySize(ansi)
'End Function
'
'Public Function nArrayOfString(ByVal str As String) As Byte()
'    'transfer of string to array using ansi as the vessle
'    Dim out() As Byte
'    Dim ansi As MemInfo
'    ansi = lCreateANSI(str)
'    ReDim out(0 To lLengthANSI(ansi)) As Byte
'    RtlMoveMemory ByVal VarPtr(out(LBound(out))), ByVal ansi.MemP, lLengthANSI(ansi)
'    nArrayOfString = out
'    lDestroyANSI ansi
'End Function
'
'Public Function nArrayToString(ByRef ary() As Byte) As String
'    'transfer of array to string using ansi as the vessle
'    Dim ansi As MemInfo
'    ansi.MemH = LocalAlloc((&H0 Or &H40), ArraySize(ary))
'    ansi.MemP = LocalLock(ansi.MemH)
'    RtlMoveMemory ByVal ansi.MemP, ByVal VarPtr(ary(LBound(ary))), ArraySize(ary)
'    nArrayToString = lStringANSI(ansi)
'    lDestroyANSI ansi
'End Function
'
