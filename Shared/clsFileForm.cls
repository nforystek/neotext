VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public WithEvents Form As Form
Attribute Form.VB_VarHelpID = -1

Private nLoaded As Boolean
Private nChanged As Boolean

Private nFlags As Long
Private nFilter As String
Private nCaption As String
Private nFileTitle As New Collection
Private nFilePath As String
Private nFileExt As String

Public Event NewFile(ByRef Cancel As Boolean)
Public Event OpenFile()
Public Event SaveFile()
Public Event CloseFile(ByRef Cancel As Boolean)
Public Event Refresh()

Private Function PromptCancel() As Boolean
    Dim Cancel As Boolean
    
    If nChanged Then
        Select Case MsgBox("Do you want to save the current file?", vbYesNoCancel + vbQuestion)
            Case vbYes
            
                Cancel = Not SaveFile
            
            Case vbNo
                Cancel = False
            Case vbCancel
                Cancel = True
        End Select
    
    Else
        Cancel = False
    End If
    PromptCancel = Cancel
End Function
Public Property Get Changed() As Boolean
    Changed = nChanged
End Property
Public Property Let Changed(ByVal NewValue As Boolean)
    nChanged = NewValue
End Property

Public Property Get Exists() As Boolean
    Exists = PathExists(FileName, True)
End Property

Public Property Get Loaded() As Boolean
    Loaded = nLoaded
End Property
Friend Property Let Loaded(ByVal NewValue As Boolean)
    nLoaded = NewValue
End Property

Public Property Get flags() As Long
    flags = nFlags
End Property
Public Property Let flags(ByVal NewValue As Long)
    nFlags = NewValue
End Property

Public Property Get Filter() As String
    Filter = nFilter
End Property
Public Property Let Filter(ByVal NewValue As String)
    nFilter = NewValue
End Property

Public Property Get Caption() As String
    Caption = nCaption
End Property
Public Property Let Caption(ByVal NewValue As String)
    nCaption = NewValue
End Property

Public Property Get FileName(Optional ByVal Index As Variant = 1) As String
    FileName = nFilePath & nFileTitle(Index)
End Property

Public Property Get FileTitle(Optional ByVal Index As Variant = 1) As String
    FileTitle = nFileTitle(Index)
End Property
Public Property Let FileTitle(Optional ByVal Index As Variant = 1, ByVal NewValue As String)
    ClearCollection nFileTitle, False, False
    nFileTitle.Add NewValue
End Property

Public Property Get FilePath() As String
    FilePath = nFilePath
End Property
Public Property Let FilePath(ByVal NewValue As String)
    NewValue = Trim(Replace(NewValue, "/", "\"))
    nFilePath = IIf((NewValue = vbNullString), vbNullString, IIf((Right(NewValue, 1) = "\"), NewValue, NewValue & "\"))
End Property
Public Property Get FileExt() As String
    FileExt = nFileExt
End Property
Public Property Let FileExt(ByVal NewValue As String)
    nFileExt = NewValue
End Property

Private Sub Class_Initialize()
    
    flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
    Caption = ""
    Filter = "All Files (*.*)" & Chr(0) & "*.*"
    FileExt = ""
    
    FileTitle = "Untitled"
    
    If Ini.Setting("DialogDir") = "" Then
        FilePath = AppPath
    Else
        FilePath = Ini.Setting("DialogDir")
    End If
    
    Changed = False
    Loaded = False
End Sub
Public Sub NewFile()
    If Not PromptCancel Then
        Dim nCancel As Boolean
        RaiseEvent NewFile(nCancel)
        If Not nCancel Then
            Changed = False
            Loaded = True
            FileTitle = "Untitled" & nFileExt
            FilePath = ""
            RaiseEvent Refresh
        End If
    End If
End Sub
Public Sub CloseFile()
    If Not PromptCancel Then
        Dim nCancel As Boolean
        RaiseEvent CloseFile(nCancel)
        If Not nCancel Then
            Loaded = False
            Changed = False
            FileTitle = ""
            FilePath = ""
            RaiseEvent Refresh
        End If
    End If
End Sub
Public Sub OpenFile(Optional ByVal nOpenFile As String = vbNullString)
    If Not nOpenFile = vbNullString Then
        If Not PathExists(nOpenFile, True) Then
            MsgBox "File [" & nOpenFile & "] does not exist.", vbInformation
            Exit Sub
        End If
    End If
    
    If Not PromptCancel Then
        
        Dim sOpen As SelectedFile
    
        If nOpenFile = vbNullString Then
            FileDialog.sFilter = Filter
            FileDialog.flags = flags
            FileDialog.sDlgTitle = Trim("Open " & Caption)
            FileDialog.sInitDir = FilePath
            
            On Error Resume Next
            sOpen = ShowOpen(Form.hWnd)
            If (Not (Err.Number = 32755)) And (sOpen.bCanceled = False) Then
                FilePath = sOpen.sLastDirectory
                
                ClearCollection nFileTitle, False, False
                Dim cnt As Integer
                For cnt = 1 To sOpen.nFilesSelected
                    nFileTitle.Add sOpen.sFiles(cnt)
                Next
                
                Changed = False
                Loaded = True
                RaiseEvent OpenFile
                RaiseEvent Refresh
            End If
            
            If Err Then Err.Clear
            On Error GoTo 0
            
        Else
            FilePath = GetFilePath(nOpenFile)
            ClearCollection nFileTitle, False, False
            nFileTitle.Add GetFileName(nOpenFile)
            
            Changed = False
            Loaded = True
            RaiseEvent OpenFile
            RaiseEvent Refresh
        End If
        
    End If
End Sub

Public Function SaveFile(Optional ByVal ForcePrompt As Boolean = False) As Boolean
    Dim save As Boolean
    save = False
                
    If ForcePrompt Or (Not PathExists(FileName, True)) Then
        
        Dim sSave As SelectedFile
        
        FileDialog.sFilter = Filter
        FileDialog.flags = flags
        FileDialog.sDlgTitle = Trim("Save " & Caption)
        FileDialog.sInitDir = GetFilePath(FileName)
        FileDialog.sFile = GetFileName(FileName)
        
        On Error Resume Next
        sSave = ShowSave(Form.hWnd)
        If (Not (Err.Number = 32755)) And (sSave.bCanceled = False) Then
            
            ClearCollection nFileTitle, False, False
            nFileTitle.Add GetFileName(sSave.sFiles(1)) & IIf((GetFileExt(sSave.sFiles(1)) = vbNullString), TreeFileExt, vbNullString)
            FilePath = GetFilePath(sSave.sFiles(1))
        
            Ini.Setting("DefaultDir") = IIf(Right(FilePath, 1) = "\", Left(FilePath, Len(FilePath) - 1), FilePath)
            save = True
        End If
        
        If Err Then Err.Clear
        On Error GoTo 0
        
    Else
        save = True
    End If
    
    If save Then
        Changed = False
        RaiseEvent SaveFile
        RaiseEvent Refresh
        save = True
    End If
    
    SaveFile = save
    
End Function

Private Sub Class_Terminate()
    ClearCollection nFileTitle, False, True
    Set Form = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = PromptCancel
    End If
End Sub

