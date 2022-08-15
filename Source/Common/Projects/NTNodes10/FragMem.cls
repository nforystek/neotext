VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FragMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private FileSafe As Boolean
Private FileDel As Boolean
Private FileLoc As String
Private FileCount As Long
Private FileIndex() As FileData

Public Property Get Ensured() As Boolean ' _
Gets whether the memory is ensured by saving it to the a file upon any use of the allocations.
Attribute Ensured.VB_Description = "Gets whether the memory is ensured by saving it to the a file upon any use of the allocations."
    Ensured = FileSafe
End Property
Public Property Let Ensured(ByVal RHS As Boolean) ' _
Sets whether the memory is ensured by saving it to the a file upon any use of the allocations.
Attribute Ensured.VB_Description = "Sets whether the memory is ensured by saving it to the a file upon any use of the allocations."
    FileSafe = RHS
End Property

Public Property Get FilePath() As String ' _
Gets the file path that the memory will be used when the Ensured property is set to true.
Attribute FilePath.VB_Description = "Gets the file path that the memory will be used when the Ensured property is set to true."
    FilePath = FileLoc
End Property
Public Property Let FilePath(ByVal RHS As String) ' _
Sets the file path that the memory will be used when the Ensured property is set to true.
Attribute FilePath.VB_Description = "Sets the file path that the memory will be used when the Ensured property is set to true."
    If RHS <> "" Then
        FileLoc = RHS
        FileDel = False
        modIndexer.LoadIndecies FileCount, FileIndex, FileLoc
    Else
        Class_Terminate
        FileLoc = GetTemporaryFile
        FileDel = True
    End If
End Property
Public Property Get Count() As Long ' _
Gets the number of ordinal handles that are allocated in memory.
Attribute Count.VB_Description = "Gets the number of ordinal handles that are allocated in memory."
    Count = FileCount
End Property

Public Function Handle(ByVal Ordinal As Long) As Long ' _
Returns the handle for the Ordinal where the Count property is one or greater, and the Ordinal not above Count.
Attribute Handle.VB_Description = "Returns the handle for the Ordinal where the Count property is one or greater, and the Ordinal not above Count."
    Handle = FileIndex(Ordinal).StartLoc
End Function

Public Property Get SizeOf(ByVal Handle As Long) As Long ' _
Retruns the size of the allocated memory held at the Handle supplied.
Attribute SizeOf.VB_Description = "Retruns the size of the allocated memory held at the Handle supplied."
    SizeOf = SizeOfAlloc(FileCount, FileIndex, Handle)
End Property

Public Property Get Data(ByVal Handle As Long) As String ' _
Gets the data stored at the allocated memory held at the Handle supplied, in string form.
Attribute Data.VB_Description = "Gets the data stored at the allocated memory held at the Handle supplied, in string form."
    Data = GetAlloc(FileCount, FileIndex, Handle, FileLoc)
End Property

Public Property Let Data(ByVal Handle As Long, ByVal RHS As String) ' _
Sets the data stored at the allocated memory held at the Handle supplied, in string form.
Attribute Data.VB_Description = "Sets the data stored at the allocated memory held at the Handle supplied, in string form."
    SetAlloc FileCount, FileIndex, Handle, RHS, FileLoc
End Property

Public Function Allocate(ByVal Size As Long) As Long ' _
Allocates a portion of memory returning the handle.
Attribute Allocate.VB_Description = "Allocates a portion of memory returning the handle."
    Allocate = modIndexer.Allocate(FileCount, FileIndex, Size)
    If FileSafe Then modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
End Function

Public Sub Dealloc(ByRef Handle As Long) ' _
Deallocates the entire a portion of memory held at the Handle supplied.
Attribute Dealloc.VB_Description = "Deallocates the entire a portion of memory held at the Handle supplied."
    modIndexer.Dealloc FileCount, FileIndex, Handle
    If FileSafe Then modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
End Sub

Public Sub Realloc(ByRef Handle As Long, ByVal Size As Long) ' _
Reallocates the Size of the portion of memory held at the Handle supplied.
Attribute Realloc.VB_Description = "Reallocates the Size of the portion of memory held at the Handle supplied."
    modIndexer.Realloc FileCount, FileIndex, Handle, Size
    If FileSafe Then modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
End Sub

Private Sub Class_Initialize()
    FileLoc = GetTemporaryFile
    FileDel = True
End Sub

Private Sub Class_Terminate()
    If FileDel Then
        If PathExists(FileLoc, True) Then Kill FileLoc
    Else
        modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
    End If
End Sub
