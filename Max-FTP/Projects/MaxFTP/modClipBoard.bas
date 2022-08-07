#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modClipBoard"



#Const modClipBoard = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private Declare Function DragQueryFile Lib "shell32" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Function DragQueryPoint Lib "shell32" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardOwner Lib "user32" () As Long

Private Const GHND = &H42
Private Const CF_HDROP = &HF
Private Const GET_DROP_COUNT = &HFFFFFFFF
Private Type POINTAPI
   x As Long
   Y As Long
End Type
Private Type DROPFILES
   pFiles As Long
   pt As POINTAPI
   fNC As Long
   fWide As Long
End Type

Dim DF As DROPFILES

Dim hGlobal As Long

Dim lpGlobal As Long
Dim hDrop As Long
Dim lFiles As Long
Dim strFile As String

Private Action As String

Private SourceForm As Form
Private SourceIndex As Integer

Public PasteFormID As String
Public PasteIndex As Integer
Public PasteFilenames As String

Private Function RootFolder(ByVal rFolder As String) As String
    Dim URL As New NTAdvFTP61.URL
    Dim DirChar As String
    
    DirChar = URL.GetDirChar(rFolder)
    If Right(rFolder, 1) = DirChar Then
        RootFolder = rFolder
    Else
        RootFolder = rFolder & DirChar
    End If
    
    Set URL = Nothing
End Function
Public Function CopyFiles(ByRef sForm As Form, ByVal sIndex As Integer, ByVal sAction As String) As String
    Action = sAction
    
    If Not OpenClipboard(sForm.hwnd) = 0 Then
        EmptyClipboard
        
        Dim FileNames As String
        Dim rFolder As String
        rFolder = RootFolder(sForm.setLocation(sIndex).Text)
        
        Dim lItem As ListItem
        For Each lItem In sForm.pView(sIndex).ListItems
            If lItem.Selected Then
                If Left(lItem.Text, 1) = "/" Or Left(lItem.Text, 1) = "\" Then
                    FileNames = FileNames & rFolder & Mid(lItem.Text, 2) & vbNullChar
                Else
                    FileNames = FileNames & rFolder & lItem.Text & vbNullChar
                End If
            End If
        Next
        CopyFiles = FileNames
        
        hGlobal = GlobalAlloc(GHND, Len(DF) + Len(FileNames))
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
            DF.pFiles = Len(DF)
      
            Call CopyMem(ByVal lpGlobal, DF, Len(DF))
            Call CopyMem(ByVal (lpGlobal + Len(DF)), ByVal FileNames, Len(FileNames))
            Call GlobalUnlock(hGlobal)
      
            SetClipboardData CF_HDROP, hGlobal
        End If
        CloseClipboard
        
        Set SourceForm = sForm
        SourceIndex = sIndex
    End If

End Function

Public Function PasteFromDragDrop(ByRef dForm As Form, ByVal dIndex As Integer, ByRef Data As MSComctlLib.DataObject)

    If Action = "" Then Action = "Copy"
    Set SourceForm = Nothing
    
    On Error GoTo catch
    
    If Data.Files.Count > 0 Then
        
        Dim cnt As Long
        Dim FileNames As String
        
        For cnt = 1 To Data.Files.Count

            If Data.GetFormat(vbCFFiles) Then
                FileNames = FileNames & CStr(Data.Files(cnt)) & vbCrLf
            End If

        Next
        
        Data.Files.Clear
        
        If Not FileNames = "" Then
            CallPreformPaste dForm.MyDescription, dIndex, FileNames
        End If

    End If
    
    Exit Function
catch:
    Err.Clear
End Function

Public Function PasteFromClipboard(ByRef dForm As Form, ByVal dIndex As Integer) As String

    If Not IsClipboardFormatAvailable(CF_HDROP) = 0 Then
        
        If Not OpenClipboard(dForm.hwnd) = 0 Then
            
            hDrop = GetClipboardData(CF_HDROP)
            lFiles = DragQueryFile(hDrop, -1&, "", 0)
            
            Dim cnt As Long
            strFile = Space(260)
            
            For cnt = 0 To lFiles - 1
                Call DragQueryFile(hDrop, cnt, strFile, Len(strFile))

                frmMain.CopyItems.AddItem strFile
            
            Next
            
            Dim FileNames As String
            If frmMain.CopyItems.ListCount > 0 Then
                For cnt = 0 To frmMain.CopyItems.ListCount - 1
                    FileNames = FileNames & frmMain.CopyItems.List(cnt) & vbCrLf
                Next
            End If
            frmMain.CopyItems.Clear
        
            CloseClipboard
            
            If Not FileNames = "" Then
                CallPreformPaste dForm.MyDescription, dIndex, FileNames
            End If
            PasteFromClipboard = FileNames
        End If
    
    End If

End Function

Public Sub CallPreformPaste(ByVal pFormID As String, ByVal pIndex As Integer, ByVal pFilenames As String)
'On Error GoTo redo
    
    PasteFormID = pFormID
    PasteIndex = pIndex
    PasteFilenames = pFilenames

    Dim dForm As Form
    Set dForm = GetFormByID(PasteFormID)
    
    Dim dIndex As Integer
    Dim FileNames As String

    dIndex = PasteIndex
    FileNames = PasteFilenames
    
    If Not GetClipboardOwner = dForm.hwnd Then
        If Not SourceForm Is Nothing Then
            If Not GetClipboardOwner = SourceForm.hwnd Then
                Set SourceForm = Nothing
                SourceIndex = 0
                If Action = "" Then Action = "Copy"
            End If
        Else
            Set SourceForm = Nothing
            SourceIndex = 0
            If Action = "" Then Action = "Copy"
        End If
    ElseIf Not (SourceForm Is Nothing) Then
        If (dForm.hwnd = SourceForm.hwnd) Then
            Set SourceForm = Nothing
            SourceIndex = 0
            If Action = "" Then Action = "Copy"
        End If
    End If
    
    Dim URL As New NTAdvFTP61.URL
    Dim firstFile As String
    
    If SourceForm Is Nothing Then
        
        Set SourceForm = dForm
        If dIndex = 0 Then
            SourceIndex = 1
        Else
            SourceIndex = 0
        End If
        
        If InStr(FileNames, vbCrLf) > 0 Then
            firstFile = Left(FileNames, InStr(FileNames, vbCrLf) - 1)
        Else
            firstFile = FileNames
        End If
        SourceForm.setLocation(SourceIndex).Text = URL.GetParentFolder(firstFile)
        SourceForm.FTPConnect SourceIndex
        
        Do Until Not SourceForm.GetState(SourceIndex) = st_Processing
            DoTasks
        Loop
        
    End If
    
    Dim dServer As String
    Dim sServer As String
    Dim dFolder As String
    Dim sFolder As String
    
    dServer = LCase(URL.GetServer(dForm.setLocation(dIndex).Text))
    sServer = LCase(URL.GetServer(SourceForm.setLocation(SourceIndex).Text))
    dFolder = LCase(URL.GetFolder(dForm.setLocation(dIndex).Text))
    sFolder = LCase(URL.GetFolder(SourceForm.setLocation(SourceIndex).Text))
        
    If ((dServer = sServer) And (dFolder = sFolder)) And (Not (dServer = "")) Then
        MsgBox "Can not copy files to themselves.", vbInformation, AppName
    Else
        
        Dim lIndex As Long
        Dim lItem As ListItem
        Dim fname As String
        Dim ClipData As New NTNodes10.Collection
        
        Do Until FileNames = ""
            fname = URL.GetFile(Left(FileNames, InStr(FileNames, vbCrLf) - 1))
            lIndex = IsPathOnItems(SourceForm.pView(SourceIndex), fname)
            If lIndex > 0 Then
                Set lItem = SourceForm.pView(SourceIndex).ListItems(lIndex)
                ClipData.Add lItem.Text & "|" & lItem.SubItems(1) & "|" & lItem.SubItems(2)
            End If
            FileNames = Mid(FileNames, InStr(FileNames, vbCrLf) + 2)
        Loop
                
        Dim destObj As NTAdvFTP61.Client
        Dim sourceObj As NTAdvFTP61.Client
        Set sourceObj = SourceForm.SetMyClient(SourceIndex)
        Set destObj = dForm.SetMyClient(dIndex)
    
       ' Debug.Print sourceObj.Folder & " " & destObj.Folder
        
        If sourceObj.ConnectedState() = False Then
            If SourceForm.Visible Then MsgBox "The source folder connection has been disconnected, unable to paste files.", vbInformation, AppName
        Else
            If destObj.ConnectedState() = False Then
                If SourceForm.Visible Then MsgBox "The destination folder connection is not connected, unable to paste files.", vbInformation, AppName
            Else

                SourceForm.FTPCopyMove Action, SourceIndex, dForm, dIndex, ClipData

            End If
        End If
        
        Set ClipData = Nothing
        Set sourceObj = Nothing
        Set destObj = Nothing
       
    End If

    Set dForm = Nothing
    Set URL = Nothing
'    Exit Sub
'redo:
'    Err.Clear
'    Resume
End Sub

Public Sub PasteAbort(ByVal clientForm As Form, ByVal ClientIndex As Integer)
    If clientForm.copyIsSource(ClientIndex) Then

        clientForm.FTPAbort ClientIndex
        If Not clientForm.copyToClientForm(ClientIndex) Is Nothing Then
            clientForm.copyToClientForm(ClientIndex).FTPAbort clientForm.copyToClientIndex(ClientIndex)
        End If
        
    Else
        If Not clientForm.copyToClientForm(ClientIndex) Is Nothing Then
            clientForm.copyToClientForm(ClientIndex).FTPAbort clientForm.copyToClientIndex(ClientIndex)
        End If
        clientForm.FTPAbort ClientIndex
    
    End If
End Sub
