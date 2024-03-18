Attribute VB_Name = "modDeclares"
Option Explicit
Option Compare Binary

Public Const HEAP_CREATE_ENABLE_EXECUTE = &H40000
Public Const HEAP_GENERATE_EXCEPTIONS = &H4
Public Const HEAP_NO_SERIALIZE = &H1

Public Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Public Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long

Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Public Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Public Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Public Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long

Public Const LMEM_DISCARDABLE = &HF00
Public Const LMEM_DISCARDED = &H4000
Public Const LMEM_FIXED = &H0
Public Const LMEM_INVALID_HANDLE = &H8000
Public Const LMEM_LOCKM = &HFF
Public Const LMEM_MODIFY = &H80
Public Const LMEM_MOVEABLE = &H2
Public Const LMEM_NOCOMPACT = &H10
Public Const LMEM_NODISCARD = &H20
Public Const LMEM_VALID_FLAGS = &HF72
Public Const LMEM_ZEROINIT = &H40
Public Const lPtr = (LMEM_FIXED + LMEM_ZEROINIT)

Public Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalHandle Lib "kernel32" (wMem As Any) As Long

Public Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function LocalShrink Lib "kernel32" (ByVal hMem As Long, ByVal cbNewSize As Long) As Long

Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_DISCARDED = &H4000
Public Const GMEM_FIXED = &H0
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GMEM_LOCKM = &HFF
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_MODIFY = &H80
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long

Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function vbaObjSetAddref Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (dstObject As Any, ByVal srcObjPtr As Long) As Long
Public Declare Function vbaObjSet Lib "msvbvm60.dll" Alias "__vbaObjSet" (dstObject As Any, ByVal srcObjPtr As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (Dest As Any, Source As Any, ByVal Length As Long)
Public Declare Sub RtlMoveLongs Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nbytes As Long)
Public Declare Sub ErlMoveObjRef Lib "kernel32" Alias "RtlMoveMemory" (xDest As Object, ByRef xSource As Long, ByVal nbytes As Long)

Public Declare Sub PutLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr As Long, num As Any, Optional ByVal Size As Long = 4&)
Public Declare Sub GetLong Lib "kernel32" Alias "RtlMoveMemory" (num As Any, ByVal ptr As Long, Optional ByVal Size As Long = 4&)

Public Declare Sub GetMem8 Lib "msvbvm60.dll" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub GetMem4 Lib "msvbvm60.dll" (ByRef pSrc As Any, ByRef pDest As Any)
Public Declare Sub PutMem4 Lib "msvbvm60.dll" (ByVal Addr As Long, ByVal newVal As Long)

Public Declare Sub CopyBytes Lib "msvbvm60.dll" Alias "__vbaCopyBytes" (ByVal ByteLen As Long, ByRef Dest As Any, ByVal src As Any)






