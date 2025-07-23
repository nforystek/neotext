VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dbConnection As ADODB.Connection

Private Function OpenConnection() As Boolean
    
    Dim Retry As Boolean
    On Error Resume Next
    Do
        If Err Then Err.Clear
        
        If dbConnection.State <> adStateOpen Then
            If dbConnection.State <> 0 Then dbConnection.Close
            dbConnection.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & DatabaseFilePath & ";DefaultDir=" & GetFilePath(DatabaseFilePath) & ";pwd=" & Replace(App.EXEName, ".exe", "")
        End If
        OpenConnection = (dbConnection.State = adStateOpen)
        
        If Err.number <> 0 And Retry = False Then
            Debug.Print "DATABASE ERROR: " & Err.number & " " & Err.description
            Retry = True
        ElseIf Retry Then
            Err.Clear
        End If
        
    Loop Until (Err.number = 0)

    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function rsQuery(ByRef rs As ADODB.Recordset, ByVal sqlStr As String) As Boolean
    If rs Is Nothing Then
        Set rs = New ADODB.Recordset
    End If
    Dim Retry As Boolean
    On Error Resume Next
    Do
        Retry = False
            
        If rs.State <> 0 Then rs.Close
        rs.Open sqlStr, dbConnection, adOpenDynamic, adLockOptimistic
        rsQuery = (rs.State = adStateOpen)
        If Not rs.State = 0 Then
            If Not rsEnd(rs) Then rs.MoveFirst
        End If

        If Err Then
            Debug.Print "DATABASE ERROR: " & Err.number & " " & Err.description
            Err.Clear
        End If
    
    Loop Until Retry = False
    
    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function dbQuery(ByVal sqlStr As String) As Boolean

    Dim Retry As Boolean
    Dim rs As ADODB.Recordset
    
    On Error Resume Next
    
    Set rs = New ADODB.Recordset

    Do
        Retry = False

        If rs.State <> 0 Then rs.Close
        rs.Open sqlStr, dbConnection, adOpenDynamic, adLockOptimistic
        dbQuery = (rs.State = adStateOpen)
        
        If Err Then
            Debug.Print "DATABASE ERROR: " & Err.number & " " & Err.description
            Err.Clear
        End If
    
    Loop Until Retry = False
    
    rsClose rs
    
    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function rsEnd(ByRef rs As ADODB.Recordset) As Boolean
    rsEnd = (rs.EOF Or rs.BOF)
End Function

Public Sub rsClose(ByRef rs As ADODB.Recordset, Optional ByVal SetNothing As Boolean = True)
    If Not (rs Is Nothing) Then
        If Not rs.State = 0 Then rs.Close
        If SetNothing Then Set rs = Nothing
    End If
End Sub

Public Function isField(ByVal TableName As String, ByVal FieldName As String) As Boolean
    Dim rs As New ADODB.Recordset

    isField = False
    
    rsQuery rs, "SELECT * FROM " & TableName & ";"
    
    Dim fld As Field
    For Each fld In rs.Fields
        If LCase(Trim(fld.Name)) = LCase(Trim(FieldName)) Then
            isField = True
        End If
    Next fld
    
    rsClose rs

End Function

Public Function isTable(ByVal TableName As String) As Boolean
    Dim rs As ADODB.Recordset
    isTable = False

    Set rs = dbConnection.OpenSchema(adSchemaTables)
    
    rs.MoveFirst
    Do Until rsEnd(rs)
        
        Select Case rs.Fields("TABLE_TYPE")
            Case "TABLE", "SYSTEM TABLE"
                If LCase(Trim(rs.Fields("TABLE_NAME"))) = LCase(Trim(TableName)) Then isTable = True
        End Select
            
        rs.MoveNext
    Loop
    
    rsClose rs
    
End Function
Private Sub Class_Initialize()
    Stats_Database_Count = Stats_Database_Count + 1
    
    Set dbConnection = New ADODB.Connection
    OpenConnection
    
End Sub

Private Sub Class_Terminate()
    
    If dbConnection.State <> 0 Then dbConnection.Close
    Set dbConnection = Nothing
    Stats_Database_Count = Stats_Database_Count - 1

End Sub

