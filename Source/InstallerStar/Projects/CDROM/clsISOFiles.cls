VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsISOFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpLocalFileTime As FILETIME _
) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" ( _
    lpFileTime As FILETIME, _
    lpSystemTime As SYSTEMTIME _
) As Long

Private Declare Function GetFileTimeAPI Lib "kernel32" _
Alias "GetFileTime" ( _
    ByVal hFile As Long, _
    ByRef lpCreationTime As Any, _
    ByRef lpLastAccessTime As Any, _
    ByRef lpLastWriteTime As Any _
) As Long

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type SYSTEMTIME
  wYear                 As Integer
  wMonth                As Integer
  wDayOfWeek            As Integer
  wDay                  As Integer
  wHour                 As Integer
  wMinute               As Integer
  wSecond               As Integer
  wMilliseconds         As Integer
End Type

Private clsFiles()      As Long
Private lngFileCnt      As Long

Private lngBlocks       As Long
Private lngBlocksJoliet As Long
Private lngLBA          As Long
Private lngLBAJoliet    As Long


Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000
Private Const MEM_COMMIT = &H1000
Private Const MEM_RESERVE = &H2000
Private Const MEM_RESET = &H80000
Private Const MEM_TOP_DOWN = &H100000
Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE = &H10
Private Const PAGE_EXECUTE_READ = &H20
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const PAGE_GUARD = &H100
Private Const PAGE_NOACCESS = &H1
Private Const PAGE_NOCACHE = &H200
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Private Declare Function IsBadStringPtr Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpStringDest As String, ByVal lpStringSrc As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const LHND = &H42
Private Const LMEM_FIXED = &H0
Private Const LMEM_MOVEABLE = &H2
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = &H40

'
'Private Function PtrObj(ByRef lptr As Long) As Variant
'    Dim lZero As Long
'    Dim NewObj As Socket
'    RtlMoveMemory NewObj, lptr, 4&
'    Set PtrObj = NewObj
'    RtlMoveMemory NewObj, lZero, 4&
'End Function

Private Function NewFilePtr() As Long
    NewFilePtr = VirtualAlloc(ByVal 0&, 3112, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock NewFilePtr, 3112
End Function
Private Function DelFilePtr(ByVal PTR As Long)
    If LPTR <> 0 Then
        VirtualUnlock PTR, 3112
        VirtualFree PTR, 3112, MEM_DECOMMIT
        VirtualFree PTR, 0, MEM_RELEASE
        PTR = 0
    End If
End Function
Private Function GetFile(ByVal Addr As Long) As clsISOFile
    If Addr = 0 Then Exit Function 'exit if no addr
    Dim Node As clsISOFile 'get at alloc addr from memory
 '   rtlMovMem yVal VarPtr(Node), ByVal Addr, LenB(Node)
    Dim hptr As Long
    
    hptr = LocalAlloc(LPTR, 3112)
    LocalLock hptr

    CopyMemory ByVal hptr, ByVal Addr, 3112
    CopyMemory ByVal VarPtr(Node), ByVal hptr, 3112
    
    LocalUnlock hptr
    LocalFree hptr
    
    GetFile = Node
End Function

Private Sub SetFile(ByVal Addr As Long, ByRef Node As clsISOFile)
    If Addr = 0 Then Exit Sub 'exit if no addr
    'Debug.Print "v/" & Node.Prior & "@" & Addr

    Dim hptr As Long
    hptr = LocalAlloc(LPTR, 3112)
    LocalLock hptr
    
        
    CopyMemory ByVal hptr, ByVal VarPtr(Node), 3112
     CopyMemory ByVal Addr, ByVal hptr, 3112
     
    
    LocalUnlock hptr
    LocalFree hptr
   ' rtlMovMem ByVal Addr, ByVal NodePtr, LenB(Node)
End Sub

Public Function FileExists( _
    ByVal name As String _
) As Boolean
    
    Dim i   As Long

    For i = 0 To lngFileCnt - 1
        If StrComp(Me.name(i), name, vbTextCompare) = 0 Then
            FileExists = True
            Exit Function
        End If
    Next
End Function

Public Property Get Blocks( _
) As Long

    Blocks = lngBlocks
End Property

Public Property Let Blocks( _
    ByVal lngNewVal As Long _
)

    lngBlocks = lngNewVal
End Property

Public Property Get BlocksJoliet( _
) As Long

    BlocksJoliet = lngBlocksJoliet
End Property

Public Property Let BlocksJoliet( _
    ByVal lngNewVal As Long _
)


    lngBlocksJoliet = lngNewVal
End Property

Public Property Get LBAJoliet( _
) As Long

    LBAJoliet = lngLBAJoliet
End Property

Public Property Let LBAJoliet( _
    ByVal lngNewVal As Long _
)

    lngLBAJoliet = lngNewVal
End Property


Public Sub Clear( _
)
    Do Until lngFileCnt = 0
        DelFilePtr clsFiles(lngFileCnt - 1)
        lngFileCnt = lngFileCnt - 1
    Loop
    lngFileCnt = 0
    Erase clsFiles
End Sub

Public Function Count( _
) As Long

    Count = lngFileCnt
End Function

'Public Property Get File( _
'    ByVal Index As Long _
') As clsISOFile
'
'    File = clsFiles(Index)
'End Property

Public Sub Remove( _
    ByVal Index As Long _
)

    Dim i   As Long

    DelFilePtr clsFiles(Index)
    
    For i = Index + 1 To lngFileCnt - 1
        clsFiles(i - 1) = clsFiles(i)
    Next

    lngFileCnt = lngFileCnt - 1
End Sub

Public Sub Add( _
    ByVal localfile As String, _
    Optional ByVal name As String _
)

    If localfile = "" Then Exit Sub
    If InStr(name, "\") > 0 Then Exit Sub

    ReDim Preserve clsFiles(lngFileCnt) As Long
    clsFiles(lngFileCnt) = NewFilePtr()
    
    If SpecialDisc = 0 Then
        If FileExists(IIf(name = "", GetFilename(localfile), name)) Then
            Exit Sub
        End If

    End If

    Dim ftype As clsISOFile
    
    With ftype
        .LocalPath = localfile
        .name = IIf(name = "", GetFilename(localfile), name)
        .Size = FileLen(localfile)
        .CreationDate = GetFileTime(localfile)
    End With
    
    SetFile clsFiles(lngFileCnt), ftype

    lngFileCnt = lngFileCnt + 1
End Sub

Private Function GetFilename( _
    ByVal strPath As String _
) As String

    GetFilename = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function

Private Function GetFileTime( _
    ByVal strFile As String _
) As Date

    Dim FTCreationTime      As FILETIME
    Dim FTLastAccessTime    As FILETIME
    Dim FTLastWriteTime     As FILETIME
    Dim SysTime             As SYSTEMTIME
    Dim hF                  As hFile

    hF = FileOpen(strFile)
    If hF.handle = INVALID_HANDLE Then Exit Function
    GetFileTimeAPI hF.handle, FTCreationTime, FTLastAccessTime, FTLastWriteTime
    FileClose hF

    FileTimeToLocalFileTime FTCreationTime, FTCreationTime
    FileTimeToSystemTime FTCreationTime, SysTime

    With SysTime
        GetFileTime = DateSerial(.wYear, .wMonth, .wDay) + _
                      TimeSerial(.wHour, .wMinute, .wSecond)
    End With
End Function

Private Sub Class_Terminate()
    Clear
End Sub


Public Property Get CreationDate(ByVal Index As Long _
) As Date
    
    CreationDate = GetFile(clsFiles(Index)).CreationDate
End Property

Public Property Let CreationDate(ByVal Index As Long, _
    ByVal dtNewDate As Date _
)
    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.CreationDate = dtNewDate
     SetFile clsFiles(Index), ftype
End Property

Public Property Get SortNumber(ByVal Index As Long _
) As Long

    SortNumber = GetFile(clsFiles(Index)).SortNumber
End Property

Public Property Let SortNumber(ByVal Index As Long, _
    ByVal lngNewVal As Long _
)
    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.SortNumber = lngNewVal
     
     SetFile clsFiles(Index), ftype
End Property


Public Property Get LBA(Optional ByVal Index As Long = -1 _
) As Long
    If Index = -1 Then
    LBA = lngLBA
    Else
        LBA = GetFile(clsFiles(Index)).LBA
    End If
End Property

Public Property Let LBA(Optional ByVal Index As Long = -1, _
    ByVal lngNewVal As Long _
)

    If Index = -1 Then
    lngLBA = lngNewVal
    Else
    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
         ftype.LBA = lngNewVal
         SetFile clsFiles(Index), ftype
    End If
End Property

Public Property Get JolietName(ByVal Index As Long _
) As String

    JolietName = GetFile(clsFiles(Index)).JolietName
End Property

Public Property Let JolietName(ByVal Index As Long, _
    ByVal strVal As String _
)

    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.JolietName = strVal
     SetFile clsFiles(Index), ftype
End Property

Public Property Get DOSName(ByVal Index As Long _
) As String

    DOSName = GetFile(clsFiles(Index)).DOSName
End Property

Public Property Let DOSName(ByVal Index As Long, _
    ByVal strVal As String _
)

    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.DOSName = strVal
     SetFile clsFiles(Index), ftype
End Property

Public Property Get name(ByVal Index As Long _
) As String

    name = GetFile(clsFiles(Index)).name
End Property

Public Property Let name(ByVal Index As Long, _
    ByVal strVal As String _
)

    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))

     ftype.name = strVal
     SetFile clsFiles(Index), ftype
End Property

Public Property Get LocalPath(ByVal Index As Long _
) As String

    LocalPath = GetFile(clsFiles(Index)).LocalPath
End Property

Public Property Let LocalPath(ByVal Index As Long, _
    ByVal strVal As String _
)

    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.LocalPath = strVal
     SetFile clsFiles(Index), ftype
End Property

Public Property Get Size(ByVal Index As Long _
) As Long

    Size = GetFile(clsFiles(Index)).Size
End Property

Public Property Let Size(ByVal Index As Long, _
    ByVal lngVal As Long _
)
    Dim ftype As clsISOFile
    ftype = GetFile(clsFiles(Index))
    
     ftype.Size = lngVal
     SetFile clsFiles(Index), ftype
End Property


