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

Private Const EncodeBackups = True
Private DBConn As clsDBConnection

Public Enum BackupOptions
    bo_None = 0
    bo_Settings = 1
    bo_Operations = 2
   
    bo_AllOptions = 3
End Enum

Private Enum FieldTypes
    ft_Boolean = 1
    ft_Long = 2
    ft_String = 3
    ft_Date = 4
End Enum

Private outText As New Strand
Private tabNum As Integer
   
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
    outText.Concat String(tabNum, Chr(9)) & "<" & TagName & ">" & vbCrLf
    tabNum = tabNum + 1
End Function
Private Function TagEnd(ByVal TagName As String)
    tabNum = tabNum - 1
    outText.Concat String(tabNum, Chr(9)) & "</" & TagName & ">" & vbCrLf
End Function
Private Function TagFull(ByVal TagName As String, ByVal TagValue As String)
    outText.Concat String(tabNum, Chr(9)) & "<" & TagName & ">" & URLEncode(TagValue) & "</" & TagName & ">" & vbCrLf
End Function
Private Function ExportRow(ByRef rs As ADODB.Recordset)
    Dim col As ADODB.Field
    For Each col In rs.Fields
        If Not (LCase(col.name) = "id" Or LCase(col.name) = "parentid") Then
            TagFull col.name & "_" & GetFieldType(TypeName(col.Value)), col.Value & ""
        End If
    Next
End Function
Private Function ExportTable(ByVal TableName As String, Optional ByVal ParentID As Long = -1)
    Dim rs As New ADODB.Recordset
    If ParentID = -1 Then
        DBConn.rsQuery rs, "SELECT * FROM " & TableName & ";"
    Else
        DBConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE ParentID=" & ParentID & ";"
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

Private Function ExportOperationsTable()
    Dim rs As New ADODB.Recordset
    DBConn.rsQuery rs, "SELECT * FROM Operations;"
    
    TagBegin "T"
    TagBegin "Operations"
    Do Until rsEnd(rs)
    
        TagBegin "R"
        ExportRow rs
        ExportTable "OperationParams", rs("ID")
        TagEnd "R"
                
        rs.MoveNext
    Loop
    TagEnd "Operations"
    TagEnd "T"
    
    rsClose rs
End Function

Public Function ExportToBackup(ByVal Options As Integer, ByVal FileName As String, Optional ByVal Silent As Boolean = False) As String
    On Error GoTo catch
    
    Set DBConn = New clsDBConnection
    outText.Reset
    tabNum = 0
    
    outText.Concat "<?xml version=""2.6""?>" & vbCrLf
    
    TagBegin "O" & Trim(CStr(Options))
    
    If CheckOption(Options, "Settings") Then
        ExportTable "Settings"
    End If
    
    If CheckOption(Options, "Operations") Then
        ExportOperationsTable
    End If
    
    TagEnd "O" & Trim(CStr(Options))
    
    If EncodeBackups Then
        WriteFile FileName, HesEncodeData(outText.GetString)
    Else
        WriteFile FileName, outText.GetString
    End If

    Set DBConn = Nothing

catch:
    ExportToBackup = Err.Description
    If Err Then Err.Clear
    On Error GoTo 0
End Function
Private Function InsertRow(ByVal TableName As String, Optional ByVal ParentID As Long = -1) As Long
    Dim rs As New ADODB.Recordset
    Dim s As String
    Dim f As String
    Dim p As Integer
    s = Left(Replace(GUID, "-", ""), 50)
    Select Case TableName
        Case "Settings"
            f = "wTop"
            p = 1
        Case "Operations"
            f = "Name"
            p = 0
        Case "OperationParams"
            f = "ParamValue"
            p = 2

    End Select
    
    Select Case p
        Case 0
            DBConn.dbQuery "INSERT INTO " & TableName & " (" & f & ") VALUES ('" & s & "');"
            DBConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE " & f & "='" & s & "';"
        Case 1
            DBConn.dbQuery "INSERT INTO " & TableName & " (ID) VALUES (1);"
            DBConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE ID=1;"
        Case 2
            DBConn.dbQuery "INSERT INTO " & TableName & " (" & f & ",ParentID) VALUES ('" & s & "'," & ParentID & ");"
            DBConn.rsQuery rs, "SELECT * FROM " & TableName & " WHERE " & f & "='" & s & "' AND ParentID=" & ParentID & ";"
    End Select
    
    InsertRow = CLng(rs("ID"))
    rsClose rs
End Function
Private Function InsertField(ByVal TableName As String, ByVal ID As Long, ByVal FieldType As Integer, ByVal FieldName As String, ByVal FieldValue As String) As Long
    If DBConn.isField(TableName, FieldName) Then
        Dim rs As New ADODB.Recordset
    
        Select Case FieldType
            Case ft_Boolean, ft_Long
                DBConn.dbQuery "UPDATE " & TableName & " SET " & FieldName & "=" & URLDecode(FieldValue) & " WHERE ID=" & ID & ";"
            Case ft_String, ft_Date
                DBConn.dbQuery "UPDATE " & TableName & " SET " & FieldName & "='" & Replace(URLDecode(FieldValue), "'", "''") & "' WHERE ID=" & ID & ";"
        End Select
        
        rsClose rs
    End If
End Function

Private Function ImportRow(ByVal Options As Integer, ByVal TableName As String, ByVal c As IXMLDOMNode, Optional ByVal ParentID As Long = -1)
    
    Dim NewID As Long
    NewID = InsertRow(TableName, ParentID)
    
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
Private Function ImportTable(ByVal Options As Integer, ByVal c As IXMLDOMNode, Optional ByVal ParentID As Long = -1)

    If c.childNodes.Length > 0 Then
        
        Dim TableName As String
        Dim n As Long
        
        TableName = c.childNodes(0).baseName
        If CheckOption(Options, TableName) Then
        
            For n = 0 To c.childNodes(0).childNodes.Length - 1
                Select Case c.childNodes(0).childNodes(n).baseName
                    Case "R"
                        If DBConn.isTable(TableName) Then
                            ImportRow Options, TableName, c.childNodes(0).childNodes(n), ParentID
                        End If
                End Select
            Next
        End If
        
    End If

End Function

Private Function CheckOption(ByVal Options As Integer, ByVal LabelName As String) As Boolean
    Dim retVal As Boolean
    Select Case LabelName
        Case "Settings"
            retVal = BitWord(Options, bo_Settings)
        Case "Operations"
            retVal = BitWord(Options, bo_Operations)

        Case Else
            retVal = True
    End Select
    CheckOption = retVal
End Function

Public Function ImportFromBackup(ByVal Options As Integer, ByVal FileName As String, Optional ByVal Silent As Boolean = False) As String
    On Error GoTo catch
   
    If Options = bo_None Then Options = GetOptionsFromBackup(FileName)
   
    ResetDatabase Options, False
   
    Set DBConn = New clsDBConnection
    Set X = New MSXML.DOMDocument
       
    Dim n As Long
    Dim inText As String
    
    inText = ReadFile(FileName)
    If Nor((InStr(inText, ">") > 0), (InStr(inText, "</") > 0)) Then
        inText = modHexStr.HesDecodeData(inText)
    End If
    inText = Replace(inText, "<?xml version=""2.6""?>", "")
    
    X.async = "false"
    X.loadXML inText
    
    If X.parseError.errorCode = 0 Then
        If X.childNodes(0).childNodes.Length > 0 Then
            For n = 0 To X.childNodes(0).childNodes.Length - 1
                Select Case X.childNodes(0).childNodes(n).baseName
                    Case "T"
                        ImportTable Options, X.childNodes(0).childNodes(n)
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
    
    Set DBConn = Nothing
    
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function GetOptionsFromBackup(ByVal FileName As String) As Long
On Error GoTo catch
   
    Set X = New MSXML.DOMDocument
   
    Dim inText As String
        
    inText = ReadFile(FileName)
    If Nor((InStr(inText, ">") > 0), (InStr(inText, "</") > 0)) Then
        inText = modHexStr.HesDecodeData(inText)
    End If
    inText = Replace(inText, "<?xml version=""2.6""?>", "")
    
    X.async = "false"
    X.loadXML inText
    
    GetOptionsFromBackup = CLng(Mid(X.childNodes(0).baseName, 2))

catch:
    Set X = Nothing
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function ResetDatabase(ByVal Options As Integer, Optional ByVal Silent As Boolean = False, Optional SetOriginal As Boolean = True) As String
    
On Error GoTo catch
    
    Set DBConn = New clsDBConnection
    
    Dim rs As New ADODB.Recordset

    If CheckOption(Options, "Settings") Then
        DBConn.rsQuery rs, "DELETE * FROM Settings;"
        If SetOriginal Then
            Load frmMain
            DBConn.rsQuery rs, "INSERT INTO Settings (ID, VBScriptText, JScriptText) VALUES (1, '" & Replace(frmMain.DefaultVBScript.Text, "'", "''") & "', '" & Replace(frmMain.DefaultJScript.Text, "'", "''") & "');"
            Unload frmMain
        End If
    End If

    If CheckOption(Options, "Operations") Then
    
        DBConn.rsQuery rs, "DELETE * FROM OperationParams;"
        DBConn.rsQuery rs, "DELETE * FROM Operations;"
                
    End If
      
    If SetOriginal Then
    
        DBConn.rsQuery rs, "DELETE * FROM MessageQueue;"
    
    End If
       
    rsClose rs
    Set DBConn = Nothing

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
    dPath = Replace(sPath, ".mdb", ".bak")
    
    If PathExists(dPath) Then Kill dPath
    JRO.CompactDatabase _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPath & ";Jet OLEDB:Database Password=" & LCase(GetFileTitle(DBFileName)) & ";", _
            "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dPath & ";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=" & LCase(GetFileTitle(DBFileName)) & ";"

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


