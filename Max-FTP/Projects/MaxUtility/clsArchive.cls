VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private Const EncodeBackups = False
Private dbConn As clsDBConnection

Public Enum BackupOptions
    bo_None = 0
    bo_PublicSettings = 1
    bo_FileAssociations = 2
    bo_Favorites = 4
    bo_Projects = 8
    bo_Graphics = 16
    
    bo_UserSettings = 32
        
    bo_ActiveApp = 64
    bo_SessionDrives = 128
    bo_VisitedHistory = 256
    'bo_GlobalEvents = 512
    bo_Schedules = 1024
    
    'bo_AutoSetup = 1955
    bo_AllOptions = 2047
End Enum

Private Enum FieldTypes
    ft_Boolean = 1
    ft_Long = 2
    ft_String = 3
    ft_Date = 4
End Enum

Private outText As New Strand

Private tabNum As Integer

Private msgSup As Boolean

Private X As MSXML.DOMDocument

Private Function GetFieldType(ByVal FieldTypeName As String) As FieldTypes
    Select Case FieldTypeName
        Case "Boolean"
            GetFieldType = ft_Boolean
        Case "Long"
            GetFieldType = ft_Long
        Case "String"
            GetFieldType = ft_String
        Case "Date"
            GetFieldType = ft_Date
    End Select
End Function

Private Function TagBegin(ByVal TagName As String)
    outText.Concat "<" & TagName & ">"
    tabNum = tabNum + 1
End Function
Private Function TagEnd(ByVal TagName As String)
    tabNum = tabNum - 1
    outText.Concat "</" & TagName & ">"
End Function
Private Function TagFull(ByVal TagName As String, ByVal TagValue As String)
    outText.Concat "<" & TagName & ">" & URLEncode(TagValue) & "</" & TagName & ">"
End Function
Private Function ExportRow(ByRef rs As ADODB.Recordset)
    Dim col As ADODB.Field
    For Each col In rs.Fields
        If Not (LCase(col.Name) = "id" Or LCase(col.Name) = "parentid") Then
            TagFull col.Name & "_" & GetFieldType(TypeName(col.Value)), col.Value & ""
        End If
    Next
End Function
Private Function ExportTable(ByVal TableName As String, Optional ByVal ParentID As Long = -1)
    Dim rs As New ADODB.Recordset
    If ParentID = -1 Then
        dbConn.rsQuery rs, "SELECT * FROM " & TableName & ";"
    Else
        dbConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE ParentID=" & ParentID & ";"
    End If
    
    TagBegin "T"
    TagBegin TableName
    Do Until rsEnd(rs)
    
        TagBegin "R"
        ExportRow rs
        TagEnd "R"
        
        rs.MoveNext
    Loop
    TagEnd TableName
    TagEnd "T"
    
    rsClose rs
End Function

Private Function ExportScheduleTable(ByVal UserID As Long)
    Dim rs As New ADODB.Recordset
    dbConn.rsQuery rs, "SELECT * FROM Schedules WHERE ParentID=" & UserID & ";"
    
    TagBegin "T"
    TagBegin "Schedules"
    Do Until rsEnd(rs)
    
        TagBegin "R"
        ExportRow rs
        ExportTable "Operations", rs("ID")
        TagEnd "R"
        
        rs.MoveNext
    Loop
    TagEnd "Schedules"
    TagEnd "T"
    
    rsClose rs
End Function
Private Function ExportUserTable(ByVal Options As Integer)
    Dim rs As New ADODB.Recordset
    dbConn.rsQuery rs, "SELECT * FROM Users;"
    
    TagBegin "T"
    TagBegin "Users"
    Do Until rsEnd(rs)
    
        TagBegin "R"
        ExportRow rs
                
        ExportTable "ClientSettings", rs("ID")
        ExportTable "ProfileSettings", rs("ID")
        ExportTable "ScheduleSettings", rs("ID")
        ExportTable "ScriptingSettings", rs("ID")

        If CheckOption(Options, "ActiveApp") Then
            ExportTable "ActiveApp", rs("ID")
        End If
        If CheckOption(Options, "SessionDrives") Then
            ExportTable "SessionDrives", rs("ID")
        End If
        If CheckOption(Options, "History") Then
            ExportTable "History", rs("ID")
            ExportTable "SiteCache", rs("ID")
        End If
        If CheckOption(Options, "Schedules") Then
            ExportScheduleTable rs("ID")
        End If
        
        TagEnd "R"
        
        rs.MoveNext
    Loop
    TagEnd "Users"
    TagEnd "T"
    
    rsClose rs
End Function

Private Function ExportFolder(ByVal Path As String, ByVal f1 As Folder, Optional ByVal Silent As Boolean = False)
    Dim f2 As File
    Dim f3 As Folder
    
    TagBegin "F"
    
    TagFull "N", f1.Name
    
    If (f1.Files.count <= 50) And (f1.SubFolders.count <= 50) Then
    
        For Each f2 In f1.Files
            If (InStr(LCase(Path), MaxDBBackupExt) = 0) Then
                TagBegin "I"
                TagFull "N", f2.Name
                If (GetFileSize(Path & f1.Name & "\" & f2.Name) < 300000) Then
                    outText.Concat "<B>" & URLEncode(HesEncodeData(ReadFile(Path & f1.Name & "\" & f2.Name))) & "</B>"
                Else
                    If Not Silent And Not msgSup Then
                        msgSup = (MsgBox("WARNING: The file below is larger then 300kb and will not be included in the backup." & vbCrLf & _
                                                "NOTE: If this backup was initiated by the installation these files will not be impacted." & vbCrLf & vbCrLf & _
                                                Path & f1.Name & "\" & f2.Name & vbCrLf & vbCrLf & _
                                                "Do you want to supress further 'too large warning' messages similar to this message?", vbCritical + vbYesNo, AppName) = vbYes)
                    End If
                End If
                TagEnd "I"
            End If
        Next
        
        For Each f3 In f1.SubFolders
            ExportFolder Path & f1.Name & "\", f3, Silent
        Next
        
    Else
        If Not Silent And Not msgSup Then
            msgSup = (MsgBox("WARNING: The folder below has more then 50 files or folders and will not be included in the backup." & vbCrLf & _
                                    "NOTE: If this backup was initiated by the installation only original installation files will be impacted." & vbCrLf & vbCrLf & _
                                    Path & f1.Name & vbCrLf & vbCrLf & _
                                    "Do you want to supress further 'too large warning' messages similar to this message?", vbCritical + vbYesNo, AppName) = vbYes)
        End If
    End If
    TagEnd "F"

End Function
Public Function ExportToBackup(ByVal Options As Integer, ByVal FileName As String, Optional ByVal Silent As Boolean = False) As String
    On Error GoTo catch
    
    Set dbConn = New clsDBConnection
    
    outText.Reset
    
    tabNum = 0
    
    outText.Concat "<?xml version=""1.0""?>"
    TagBegin "O" & Trim(CStr(Options))
    
    If CheckOption(Options, "PublicSettings") Then
        ExportTable "PublicSettings"
    End If
    
    If CheckOption(Options, "FileAssociations") Then
        ExportTable "FileAssociations"
    End If
   
    If CheckOption(Options, "Users") Then
        ExportUserTable Options
    End If
    
    Dim fso As New FileSystemObject
    Dim f1 As Folder
    
    If CheckOption(Options, "Graphics") And (Not Silent) Then
        If PathExists(AppPath & GraphicsFolder) Then
            Set f1 = fso.GetFolder(AppPath & GraphicsFolder)
            ExportFolder AppPath, f1, Silent
        End If
    End If
    
    If CheckOption(Options, "Favorites") And (Not Silent) Then
        If PathExists(AppPath & FavoritesFolder) Then
            Set f1 = fso.GetFolder(AppPath & FavoritesFolder)
            ExportFolder AppPath, f1, Silent
        End If
    End If
    If CheckOption(Options, "ActiveApp") And (Not Silent) Then
        If PathExists(AppPath & ActiveAppFolder) Then
            Set f1 = fso.GetFolder(AppPath & ActiveAppFolder)
            ExportFolder AppPath, f1, Silent
        End If
    End If
    
    If CheckOption(Options, "Projects") And (Not Silent) Then
        If PathExists(AppPath & ProjectFolder) Then
            Set f1 = fso.GetFolder(AppPath & ProjectFolder)
            ExportFolder AppPath, f1, Silent
        End If
    End If
    
    TagEnd "O" & Trim(CStr(Options))
    
    Set f1 = Nothing
    Set fso = Nothing
    
    Dim tmpText As String
    If EncodeBackups Then
        WriteFile FileName, HesEncodeData(outText.GetString)
    Else
        WriteFile FileName, outText.GetString
    End If

    Set dbConn = Nothing
    
catch:
    ExportToBackup = Err.Description
    
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Private Function InsertRow(ByVal TableName As String, Optional ByVal ParentID As String = "-1") As Long
    Dim rs As New ADODB.Recordset
    Dim p As Integer

    Dim s As String
    Dim f As String

    s = Left(Replace(GUID, "-", ""), 50)
    Select Case TableName
        Case "PublicSettings"
            f = "ServiceInstallGUID"
            p = 0
        Case "FileAssociations"
            f = "DisplayName"
            p = 0
        Case "Users"
            f = "UserName"
            dbConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE " & f & "='" & ParentID & "';"
            p = IIf(Not rsEnd(rs), 3, 0)
            
        Case "ProfileSettings"
            p = 1
        Case "ClientSettings"
            p = 1
        Case "ScheduleSettings"
            p = 1
        Case "ScriptingSettings"
            p = 1
            
        Case "ActiveApp"
            f = "FileName"
            p = 2
        Case "SessionDrives"
            f = "ShareName"
            p = 2
        Case "History"
            f = "URL"
            p = 2
        Case "SiteCache"
            f = "HostURL"
            p = 2
        Case "Schedules"
            f = "ScheduleName"
            p = 2
        Case "Operations"
            f = "OperationName"
            p = 2
    End Select
    
    Select Case p
        Case 0
            dbConn.dbQuery "INSERT INTO " & TableName & " (" & f & ") VALUES ('" & s & "');"
            dbConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE " & f & "='" & s & "';"
        Case 1
            dbConn.dbQuery "INSERT INTO " & TableName & " (ParentID) VALUES (" & ParentID & ");"
            dbConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE ParentID=" & ParentID & ";"
        Case 2
            dbConn.dbQuery "INSERT INTO " & TableName & " (" & f & ",ParentID) VALUES ('" & s & "'," & ParentID & ");"
            dbConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE " & f & "='" & s & "' AND ParentID=" & ParentID & ";"
    End Select
    
    InsertRow = CLng(rs("ID"))
    rsClose rs
End Function
Private Function InsertField(ByVal TableName As String, ByVal ID As Long, ByVal FieldType As Integer, ByVal FieldName As String, ByVal FieldValue As String) As Long
    
    If dbConn.isField(TableName, FieldName) Then
        Dim rs As New ADODB.Recordset
    
        Select Case FieldType
            Case ft_Boolean, ft_Long
                dbConn.dbQuery "UPDATE " & TableName & " SET " & FieldName & "=" & URLDecode(FieldValue) & " WHERE ID=" & ID & ";"
            Case ft_String, ft_Date
                dbConn.dbQuery "UPDATE " & TableName & " SET " & FieldName & "='" & Replace(URLDecode(FieldValue), "'", "''") & "' WHERE ID=" & ID & ";"
        End Select
        
        rsClose rs
    End If
End Function

Private Function ImportRow(ByVal Options As Long, ByVal TableName As String, ByVal c As IXMLDOMNode, Optional ByVal ParentID As Long = -1)
    
    Dim NewID As Long
    NewID = InsertRow(TableName, IIf(TableName = "Users", c.childNodes(0).Text, ParentID))
    
    Dim n As Long
    For n = 0 To c.childNodes.Length - 1
        Select Case c.childNodes(n).baseName
            Case "T"
                ImportTable Options, c.childNodes(n), NewID
            Case Else
                InsertField TableName, NewID, CInt(Right(c.childNodes(n).baseName, 1)), Left(c.childNodes(n).baseName, InStrRev(c.childNodes(n).baseName, "_") - 1), c.childNodes(n).Text
        End Select
    Next
    
End Function
Private Function ImportTable(ByVal Options As Long, ByVal c As IXMLDOMNode, Optional ByVal ParentID As Long = -1)

    If c.childNodes.Length > 0 Then
        
        Dim TableName As String
        Dim n As Long
        
        TableName = c.childNodes(0).baseName

        If CheckOption(Options, TableName) Then
        
            For n = 0 To c.childNodes(0).childNodes.Length - 1
                Select Case c.childNodes(0).childNodes(n).baseName
                    Case "R"
                        If dbConn.isTable(TableName) Then
                            ImportRow Options, TableName, c.childNodes(0).childNodes(n), ParentID
                        End If
                End Select
            Next
        End If

    End If

End Function
Private Function ImportFile(ByVal RootPath As String, ByVal c As IXMLDOMNode)
    If c.childNodes.Length > 0 Then
        
        Dim FileName As String
        Dim FileData As String
        
        FileName = URLDecode(c.childNodes(0).Text)
        FileData = URLDecode(c.childNodes(1).Text)
        
        If PathExists(RootPath & FileName) Then
            Kill RootPath & FileName
        End If
        
        WriteFile RootPath & FileName, HesDecodeData(FileData)
        
    End If

End Function
Private Function CheckOption(ByVal Options As Integer, ByVal LabelName As String) As Boolean
    Dim retVal As Boolean
    Select Case LabelName
        Case "PublicSettings"
            retVal = BitWord(Options, bo_PublicSettings)
        Case "FileAssociations"
            retVal = BitWord(Options, bo_FileAssociations)
        Case "Graphics"
            retVal = BitWord(Options, bo_Graphics)
        Case "Favorites"
            retVal = BitWord(Options, bo_Favorites)
        Case "Projects"
            retVal = BitWord(Options, bo_Projects)
        Case "Users", "ProfileSettings", "ClientSettings", "ScheduleSettings", "ScriptingSettings"
            retVal = BitWord(Options, bo_UserSettings)
        Case "ActiveApp"
            retVal = BitWord(Options, bo_ActiveApp)
        Case "SessionDrives"
            retVal = BitWord(Options, bo_SessionDrives)
        Case "History", "SiteCache"
            retVal = BitWord(Options, bo_VisitedHistory)
        Case "Schedules", "Operations"
            retVal = BitWord(Options, bo_Schedules)
        Case Else
            retVal = True
    End Select
    CheckOption = retVal
End Function
Private Function ImportFolder(ByVal Options As Integer, ByVal RootPath As String, ByVal c As IXMLDOMNode)
    
    If c.childNodes.Length > 0 Then
        
        Dim FolderName As String
        Dim n As Long
        
        FolderName = URLDecode(c.childNodes(0).Text)
        If CheckOption(Options, FolderName) Then
        
            If Not PathExists(RootPath & FolderName) Then MkDir RootPath & FolderName
            
            If c.childNodes.Length > 1 Then
                For n = 1 To c.childNodes.Length - 1
                    Select Case c.childNodes(n).baseName
                        Case "I"
                            ImportFile RootPath & FolderName & "\", c.childNodes(n)
                        Case "F"
                            ImportFolder Options, RootPath & FolderName & "\", c.childNodes(n)
                    End Select
                Next
            End If
        End If
    End If
    
End Function

Public Function ImportFromBackup(ByVal Options As Long, ByVal FileName As String, Optional ByVal Silent As Boolean = False) As String
    On Error GoTo catch
   
    If Options = bo_None Then Options = GetOptionsFromBackup(FileName)
   
    ResetDatabase Options, Silent, False
   
    Set dbConn = New clsDBConnection
    Set X = New MSXML.DOMDocument
   
    Dim n As Long
    Dim inText As String
    
    inText = ReadFile(FileName)
    If Nor((InStr(inText, ">") > 0), (InStr(inText, "</") > 0)) Then
        inText = HesDecodeData(inText)
    End If
    
    inText = Replace(inText, "<?xml version=""1.0""?>", "")
    inText = Replace(inText, "<?xml version=""1.0"" ?>", "")
    inText = Replace(inText, "<? xml version=""1.0""?>", "")
    inText = Replace(inText, "<? xml version=""1.0"" ?>", "")
    
    X.async = "false"
    X.loadXML inText
    
    If X.parseError.errorCode = 0 Then
        If X.childNodes(0).childNodes.Length > 0 Then
            For n = 0 To X.childNodes(0).childNodes.Length - 1
                Select Case X.childNodes(0).childNodes(n).baseName
                    Case "T"
                        ImportTable Options, X.childNodes(0).childNodes(n)
                    Case "F"
                        If (Not Silent) Then ImportFolder Options, AppPath, X.childNodes(0).childNodes(n)
                End Select
            Next
        End If
    End If
    
catch:
    
    If (Not (X Is Nothing)) And (Not Err) Then
        If Not X.parseError.errorCode = 0 Then
            ImportFromBackup = X.parseError.reason
        End If
        Set X = Nothing
    Else
        ImportFromBackup = Err.Description
    End If
    
    Set dbConn = Nothing
    
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function GetOptionsFromBackup(ByVal FileName As String) As Long
On Error GoTo catch
   
    Set X = New MSXML.DOMDocument
   
    Dim inText As String
    
    inText = ReadFile(FileName)
    If Nor((InStr(inText, ">") > 0), (InStr(inText, "</") > 0)) Then
        inText = HesDecodeData(inText)
    End If
    
    inText = Replace(inText, "<?xml version=""1.0""?>", "")
    inText = Replace(inText, "<?xml version=""1.0"" ?>", "")
    inText = Replace(inText, "<? xml version=""1.0""?>", "")
    inText = Replace(inText, "<? xml version=""1.0"" ?>", "")
    
    X.async = "false"
    X.loadXML inText
    
    GetOptionsFromBackup = CLng(Mid(X.childNodes(0).baseName, 2))

catch:
    Set X = Nothing
    
    Err.Clear
    On Error GoTo 0
End Function

Public Function ResetDatabase(ByVal Options As Integer, Optional ByVal Silent As Boolean = False, Optional SetOriginal As Boolean = True) As String
    
On Error GoTo catch
    
    Set dbConn = New clsDBConnection
    
    Dim rs As New ADODB.Recordset

    If CheckOption(Options, "PublicSettings") Then
        dbConn.rsQuery rs, "DELETE * FROM PublicSettings;"
        If SetOriginal Then
            dbConn.rsQuery rs, "INSERT INTO PublicSettings (ServiceInstallGUID) VALUES ('" & Replace(GUID, "-", "") & "');"
        End If
    End If

    If CheckOption(Options, "FileAssociations") Then
        dbConn.rsQuery rs, "DELETE * FROM FileAssociations;"
        If SetOriginal Then
        
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (No, 'Archive Files', '*.ZIP, *.ARJ, *.RAR, *.LBR, *.TAR, *.GZ, *.BKF',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (Yes, 'Executables', '*.COM, *.BAT, *.EXE, *.DLL, *.OCX, *.MSI',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (No, 'Image File', '*.GIF, *.JPG, *.JPEG, *.BMP, *.PCX, *.TIF, *.ICO, *.PNG, *.CUR, *.PSD, *.PSP',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (Yes, 'Max-FTP Files', '*.MFTP, *.MPRJ, *.MSCR, *.MADB',  1, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                                    "VALUES (No, 'NSIS Scripts', '*.NSI, *.NSH',  1, No, '%windowsfolder%\notepad.exe', Yes);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (No, 'Office Files', '*.DOC, *.XLS, *.MDB, *.PPT, *.EML',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (No, 'Sound Files', '*.WAV, *.WAVE, *.MID, *.MIDI, *.MP3, *.WMA, *.AIFF, *.AU',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                                    "VALUES (No, 'Text Document', '*.TXT, *.TEXT, *.NFO, *.INFO, *.INI, *.INIT',  1, No, '%windowsfolder%\notepad.exe', Yes);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, AssumeLineFeed) " & _
                                    "VALUES (No, 'Video Files', '*.AVI, *.MPG, *.MPEG, *.WMV, *.ASF, *.MOV, *.MOD, *.SWF',  0, Yes, No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                                    "VALUES (No, 'Visual Basic Files', '*.VBG, *.VBP, *.BAS, *.CLS, *.CTL, *.FRM, *.DSR, *.PAG, *.DOB, *.VBZ',  1, No, '%windowsfolder%\notepad.exe', No);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                                    "VALUES (No, 'Web Documents', '*.HTM, *.HTML, *.SHTML, *.ASA, *.ASP, *.JS, *.VBS, *.CSS, *.XML, * .XSL, *.PAD',  1, No, '%windowsfolder%\notepad.exe', Yes);"
            dbConn.rsQuery rs, "INSERT INTO FileAssociations (Locked, DisplayName, Extentions, TransferType, WindowsApp, ApplicationExe, AssumeLineFeed) " & _
                                    "VALUES (No, '.NET Documents', '*.CS, *.ASPX, *.ASAX, *.CONFIG, *.COMPILED',  1, No, '%windowsfolder%\notepad.exe', No);"

        End If
    End If
    
    If CheckOption(Options, "Users") Then
        Dim delUser As Boolean
        delUser = True
        
        If CheckOption(Options, "ActiveApp") Then
            dbConn.rsQuery rs, "DELETE * FROM ActiveApp;"
        Else
            delUser = False
        End If
        
        If CheckOption(Options, "History") Then
            dbConn.rsQuery rs, "DELETE * FROM SiteCache;"
            dbConn.rsQuery rs, "DELETE * FROM History;"
        Else
            delUser = False
        End If
        
        If CheckOption(Options, "SessionDrives") Then
            dbConn.rsQuery rs, "DELETE * FROM SessionDrives;"
        Else
            delUser = False
        End If
        
        If CheckOption(Options, "Schedules") Then
            dbConn.rsQuery rs, "DELETE * FROM Operations;"
            dbConn.rsQuery rs, "DELETE * FROM Schedules;"
        Else
            delUser = False
        End If
        
        dbConn.rsQuery rs, "DELETE * FROM ScriptingSettings;"
        dbConn.rsQuery rs, "DELETE * FROM ScheduleSettings;"
        dbConn.rsQuery rs, "DELETE * FROM ProfileSettings;"
        dbConn.rsQuery rs, "DELETE * FROM ClientSettings;"
        
        If delUser Then
            dbConn.rsQuery rs, "DELETE * FROM Users;"
        End If
        
    End If
   
    If SetOriginal Then
    
        dbConn.rsQuery rs, "UPDATE TipOfDay SET Viewed = 0;"
    
        dbConn.rsQuery rs, "DELETE * FROM MessageQueue;"
    
    End If
   
    rsClose rs
    Set dbConn = Nothing
    
catch:
    ResetDatabase = Err.Description
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function CompactDatabase(Optional ByVal Silent As Boolean = False) As String
    On Error Resume Next
    
    Dim JRO As New JRO.JetEngine
    Dim sPath As String
    Dim dPath As String

    sPath = DatabaseFilePath
    dPath = Replace(DatabaseFilePath, ".mdb", ".bak")
    
    If PathExists(dPath) Then Kill dPath
    JRO.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=" & LCase(GetFileTitle(DatabaseFilePath)) & ";", _
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dPath & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=" & LCase(GetFileTitle(DatabaseFilePath)) & ";"
    
    If PathExists(dPath) Then
        Kill sPath
        FileCopy dPath, sPath
        Kill dPath
    End If

    Set JRO = Nothing

    CompactDatabase = Err.Description
    If (Err.Number <> 0) Then
        If (Not Silent) Then MsgBox "Error: " & Err.Number & ", " & Err.Description, vbCritical, AppName
        Err.Clear
    End If
    On Error GoTo 0
End Function


