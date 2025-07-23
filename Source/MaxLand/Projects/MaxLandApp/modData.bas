Attribute VB_Name = "modData"
#Const modData = -1
Option Explicit
'TOP DoWN
Option Compare Binary
Option Private Module



Public Function DatabaseFilePath(Optional ByVal UserLoc As Boolean = False) As String
    Dim retVal As String
  '  If retVal = "" Then
        retVal = AppPath & Replace(App.EXEName, ".exe", "") & ".mdb"
        If PathExists(Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare), True) And (Not PathExists(retVal, True)) Then
            retVal = Replace(retVal, GetProgramFilesFolder, GetAllUsersAppDataFolder, , , vbTextCompare)
        End If
   ' End If
    If UserLoc Then
        retVal = Replace(Replace(retVal, GetProgramFilesFolder, GetCurrentAppDataFolder, , , vbTextCompare), GetAllUsersAppDataFolder, GetCurrentAppDataFolder, , , vbTextCompare) & ".sql"
        MakeFolder GetFilePath(retVal)
    End If
    DatabaseFilePath = retVal
End Function

Public Sub BackupDB()
    
    Dim txt As String
    
    Set db = New Database
    db.rsQuery rs, "SELECT * FROM Settings;"
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do
            txt = txt & "INSERT INTO Settings (Username, Resolution, Windowed) VALUES ('" & Replace(rs("Username"), "'", "''") & "', '" & rs("Resolution") & "', " & rs("Windowed") & ");"
            rs.MoveNext
        Loop Until db.rsEnd(rs)
    End If
    
    db.rsQuery rs, "SELECT * FROM Serials;"
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do
            txt = txt & "INSERT INTO Serials (Username, PXFile, Script) VALUES ('" & Replace(rs("Username"), "'", "''") & "', '" & Replace(rs("PXFile"), "'", "''") & "', '" & Replace(rs("Script"), "'", "''") & "');"
            rs.MoveNext
        Loop Until db.rsEnd(rs)
    End If
    
    WriteFile DatabaseFilePath(True), txt

    Set db = Nothing
    
End Sub
Public Sub RestoreDB()
    
    Dim txt As String
    
    Set db = New Database
    txt = ReadFile(DatabaseFilePath(True))
    Do Until txt = ""
        db.dbQuery RemoveNextArg(txt, ";")
    Loop
    Kill DatabaseFilePath(True)

    Set db = Nothing

End Sub
Public Sub ResetDB()
    Set db = New Database
    db.dbQuery "DELETE * FROM Serials;"
    db.dbQuery "DELETE * FROM Settings;"
    Set db = Nothing
End Sub

Public Sub CompactDB()
On Error GoTo catch
    
    Dim JRO As New JRO.JetEngine
    Dim sPath As String
    Dim dPath As String

    sPath = DatabaseFilePath()
    dPath = DatabaseFilePath(True) & ".bak"
    
    If PathExists(dPath) Then Kill dPath
    JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=" & Replace(App.EXEName, ".exe", "", , , vbTextCompare) & ";", _
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dPath & ";Jet OLEDB:Database Password=" & Replace(App.EXEName, ".exe", "", , , vbTextCompare) & ";Jet OLEDB:Engine Type=5"

    Kill sPath
    FileCopy dPath, sPath
    Kill dPath

    Set JRO = Nothing
    
catch:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub


