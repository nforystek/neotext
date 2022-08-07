Attribute VB_Name = "modUser"
#Const modUser = -1
Option Explicit
'TOP DOWN
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

    Set db = New clsDatabase
    Dim rs As ADODB.Recordset
    Set rs = CreateObject("ADODB.Recordset")
    db.rsQuery rs, "SELECT * FROM Settings;"
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do
            txt = txt & "INSERT INTO Settings (Username, Resolution, Windowed, WireFrame) VALUES ('" & Replace(rs("Username"), "'", "''") & "', '" & Replace(rs("Resolution"), "'", "''") & "', " & rs("Windowed") & ", " & rs("WireFrame") & ");" & vbCrLf
            rs.MoveNext
        Loop Until db.rsEnd(rs)
        
    End If
    
    db.rsQuery rs, "SELECT * FROM Scores;"
    If Not db.rsEnd(rs) Then
        rs.MoveFirst
        Do
            txt = txt & "INSERT INTO Scores (Username, LevelNum, BestTime) VALUES ('" & Replace(rs("Username"), "'", "''") & "', " & rs("LevelNum") & ", " & rs("BestTime") & ");" & vbCrLf
            rs.MoveNext
        Loop Until db.rsEnd(rs)
        
    End If
   
    db.rsClose rs
    WriteFile DatabaseFilePath(True), txt
    Set db = Nothing
    

End Sub

Public Function RestoreDB(Optional ByVal RemoveINI As Boolean = True)
    Set db = New clsDatabase
    db.dbQuery "DELETE * FROM Settings;"
    db.dbQuery "DELETE * FROM Scores;"
    Dim txt As String
    Dim Data As String
    Dim Username As String
    Data = ReadFile(DatabaseFilePath(True))
    Do While Not (Data = "")
        txt = RemoveNextArg(Data, vbCrLf)
        If Not (txt = "") Then
            db.dbQuery txt
        End If
        
    Loop
    If RemoveINI Then Kill DatabaseFilePath(True)
    Set db = Nothing
End Function

Public Sub ResetDB()
    Set db = New clsDatabase
    db.dbQuery "DELETE * FROM Settings;"
    db.dbQuery "DELETE * FROM Scores;"
    Set db = Nothing
End Sub

Public Sub ClearUserData()
    Dim User As String
    User = GetUserLoginName
    db.dbQuery "DELETE * FROM Settings WHERE Username='" & Replace(User, "'", "''") & "';"
    db.dbQuery "DELETE * FROM Scores WHERE Username='" & Replace(User, "'", "''") & "';"
End Sub

Public Sub LoadUserData()
    Dim User As String
    User = GetUserLoginName
    db.rsQuery rs, "SELECT * FROM Settings WHERE Username = '" & Replace(User, "'", "''") & "';"
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


