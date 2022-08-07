VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

'DB Requirements:
'
'   Create an Access Database with the following two tables:
'
'       Table1 Name: "FileNames"
'           Column1 Name: "ID"              Type: AutoNumber (Primary) (Indexed: Yes (No Duplicates)) (Related to FileData.FileID)
'           Column2 Name: "FileName"        Type: Memo (Required: Yes)
'
'       Table2 Name: "FileData"
'           Column1 Name: "ID               Type: AutoNumber (Primary) (Indexed: Yes (No Duplicates))
'           Column1 Name: "FileID"          Type: Numeric (Required: No) (Indexed: Yes (Duplicates OK)) (Related to FileNames.ID)
'           Column2 Name: "FileChunk"       Type: Numeric (Required: No)
'           Column3 Name: "FileData"        Type: OLE Object (Required: No)
'
'

Private Const FileChunkSize = 4096


Public Function FilePut(ByVal FilePath As String, ByVal Name As String) As Long
    
    FileRemove Name
    
    Dim DBConn As New clsDBConnection
    Dim rsFile As New ADODB.Recordset

    Dim cnt As Long
    Dim fileID As Long
    Dim strData As String
    Dim intFile As Long
    Dim filePos As Long
    Dim nextSize As Long

    DBConn.rsQuery rsFile, "INSERT INTO FileNames (FileName) " & "VALUES ('" & Replace(Name, "'", "''") & "');"
    DBConn.rsQuery rsFile, "SELECT * FROM FileNames WHERE FileName='" & Replace(Name, "'", "''") & "';"
    
    fileID = rsFile("ID")
    
    DBConn.rsQuery rsFile, "SELECT * FROM FileData;"
        
    intFile = FreeFile
    Open FilePath & "\" & Name For Binary As intFile
        
        If LOF(intFile) > 0 Then
            filePos = 0
            cnt = 1
            Do
                
                nextSize = LOF(intFile) - filePos
                If nextSize > 0 Then
                    If nextSize > FileChunkSize Then
                        strData = Space(FileChunkSize)
                    Else
                        strData = Space(nextSize)
                    End If
                Else
                    strData = ""
                End If
                
                If Len(strData) > 0 Then
                    Get #intFile, filePos + 1, strData
                
                    rsFile.AddNew
                    rsFile("FileID") = fileID
                    rsFile("FileChunk") = cnt
                    rsFile("FileData").AppendChunk strData
                    rsFile.Update
                    
                    cnt = cnt + 1
                    
                End If
                
                filePos = filePos + Len(strData)
                
            Loop Until Len(strData) = 0
        End If
    
    Close intFile
        
    rsClose rsFile

End Function

Public Function FileGet(ByVal FilePath As String, ByVal Name As String) As String
    
    Dim DBConn As New clsDBConnection
    Dim rsFile As New ADODB.Recordset
    
    Dim lngLen As Long
    Dim strData As String
    Dim intFile As Long

    DBConn.rsQuery rsFile, "SELECT * FROM FileNames WHERE FileName='" & Replace(Name, "'", "''") & "';"
    
    If Not rsEnd(rsFile) Then
        DBConn.rsQuery rsFile, "SELECT * FROM FileData WHERE FileID=" & rsFile("ID") & " ORDER BY FileChunk;"
        
        intFile = FreeFile
        Open FilePath & "\" & Name For Output As intFile
        Close intFile
    
        intFile = FreeFile
        Open FilePath & "\" & Name For Binary As #intFile
            
            Do Until rsEnd(rsFile)
        
                lngLen = rsFile.Fields("FileData").ActualSize
                strData = rsFile("FileData").GetChunk(lngLen)
        
                If lngLen > 0 Then
                    Put intFile, , strData
                End If
        
                rsFile.MoveNext
            Loop
    
    
        Close intFile
        
        
        rsClose rsFile
        
        FileGet = FilePath
        
    End If


End Function

Public Function FileRemove(ByVal Name As String)
    Dim DBConn As New clsDBConnection
    Dim rsFile As New ADODB.Recordset
    
    Dim fileID As Long
    
    DBConn.rsQuery rsFile, "SELECT * FROM FileNames WHERE FileName='" & Replace(Name, "'", "''") & "';"
    
    If Not rsEnd(rsFile) Then
        
        fileID = rsFile("ID")
        
        DBConn.rsQuery rsFile, "DELETE FROM FileData WHERE FileID=" & fileID & ";"

        DBConn.rsQuery rsFile, "DELETE FROM FileNames WHERE ID=" & fileID & ";"

    End If
    
    rsClose rsFile

End Function

