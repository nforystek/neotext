VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'TOP DOWN
Option Compare Binary


Private pKey As String
Private pParent As Project
Private pFileName As String
Private pFileDate As String
Private pFileSize As Long
Private pFileText As String
Private pFileCondComp As String

Public Property Get Parent()
    Set Parent = pParent
End Property
Friend Property Set Parent(ByRef par)
    Set pParent = par
End Property
    
Public Property Get Key() As String
    Key = pKey
End Property
Friend Property Let Key(ByVal newVal As String)
    pKey = newVal
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal newVal As String)
    On Error GoTo nofail
    
    pFileName = newVal
    If PathExists(pFileName, True) Then
        pFileSize = FileLen(pFileName)
        pFileDate = FileDateTime(pFileName)
        pFileText = ReadFile(pFileName)
    End If
    Exit Property
nofail:
    Err.Clear
    Resume
End Property

Public Property Get FileDate() As String
    FileDate = pFileDate
End Property
Public Property Let FileDate(ByVal newVal As String)
    pFileDate = newVal
End Property

Public Property Get FileSize() As Long
    FileSize = pFileSize
End Property
Public Property Let FileSize(ByVal newVal As Long)
    pFileSize = newVal
End Property

Public Property Get FileText() As String
    'If pFileText = "" And PathExists(pFileName, True) Then
    '    pFileText = ReadFile(pFileName)
    'End If
    FileText = pFileText
End Property
Public Property Let FileText(ByVal newVal As String)
    'If Not GetFileExt(pFileName, True, True) = "vbg" Then
        pFileText = newVal
    'End If
End Property

Private Sub Class_Terminate()
    Set pParent = Nothing
End Sub
