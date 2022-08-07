Attribute VB_Name = "modRemove"
#Const modRemove = -1
Option Explicit
'TOP DOWN

Public Sub RemovePath(ByVal Path As String, Optional ByRef FolderList As String)
    Dim nxt As String
    On Error Resume Next
    nxt = Dir(Path & "\*.*", vbHidden)
    If Not Err Then
        Do Until nxt = ""
            If Not nxt = "." And Not nxt = ".." Then
                FolderList = FolderList & Path & "\" & nxt & vbCrLf
            End If
            nxt = Dir
        Loop
    End If
    Do Until FolderList = ""
        nxt = RemoveNextArg(FolderList, vbCrLf)
        SetAttr nxt, vbNormal
        Select Case GetFileExt(nxt, True, True)
            Case "madb", "mdb", "bak"
            Case Else
                Kill nxt
        End Select
    Loop
    nxt = Dir(Path & "\*", vbDirectory)
    If Err Then
        Err.Clear
        Select Case GetFileExt(Path, True, True)
            Case "madb", "mdb", "bak"
            Case Else
                Kill Path
        End Select
    Else
        Do Until nxt = ""
            If Not nxt = "." And Not nxt = ".." Then
                FolderList = FolderList & Path & "\" & nxt & vbCrLf
            End If
            nxt = Dir
        Loop
    End If
    Do Until FolderList = ""
        RemovePath RemoveNextArg(FolderList, vbCrLf), FolderList
    Loop
    RmDir Path
End Sub

Public Sub Main()

    If LCase(GetFileName(Command)) = "uninstall.exe" Then
        If GetFilePath(Command) <> "" Then
            If PathExists(GetFilePath(Command), False) Then

                On Error Resume Next
                Load frmRemove
                
                Do Until (Not PathExists(Command, True)) Or Forms.Count = 0
                    Kill Command
                    
                    If Err Then Err.Clear
                    DoEvents
                    Sleep 1
                Loop
                
                RemovePath GetFilePath(Command)
               
                End
            End If
        End If
    End If
End Sub
