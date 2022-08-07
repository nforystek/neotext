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

Public Property Get Ensured() As Boolean
    Ensured = FileSafe
End Property
Public Property Let Ensured(ByVal RHS As Boolean)
    FileSafe = RHS
End Property

Public Property Get FilePath() As String
    FilePath = FileLoc
End Property
Public Property Let FilePath(ByVal RHS As String)
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
Public Property Get Count() As Long
    Count = FileCount
End Property

Public Function Handle(ByVal Ordinal As Long) As Long
    Handle = FileIndex(Ordinal).StartLoc
End Function

Public Property Get SizeOf(ByVal Handle As Long) As Long
    SizeOf = SizeOfAlloc(FileCount, FileIndex, Handle)
End Property

Public Property Get Data(ByVal Handle As Long) As String
    Data = GetAlloc(FileCount, FileIndex, Handle, FileLoc)
End Property

Public Property Let Data(ByVal Handle As Long, ByVal RHS As String)
    SetAlloc FileCount, FileIndex, Handle, RHS, FileLoc
End Property

Public Function Allocate(ByVal Size As Long) As Long
    Allocate = modIndexer.Allocate(FileCount, FileIndex, Size)
    If FileSafe Then modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
End Function

Public Sub Dealloc(ByRef Handle As Long)
    modIndexer.Dealloc FileCount, FileIndex, Handle
    If FileSafe Then modIndexer.SaveIndecies FileCount, FileIndex, FileLoc
End Sub

Public Sub Realloc(ByRef Handle As Long, ByVal Size As Long)
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
