VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private DBConn As ADODB.Connection

Public Event DBError(ByVal Number As Long, ByVal Description As String, ByRef Retry As Boolean)

Private Function OpenConnection() As Boolean
    On Error Resume Next
    Dim Retry As Boolean
    Do
        If DBConn.State <> adStateOpen Then
            If DBConn.State <> 0 Then DBConn.Close
            DBConn.Open "Driver={Microsoft Access Driver (*.mdb)};dbq=" & DatabaseFilePath & ";DefaultDir=" & GetFilePath(DatabaseFilePath) & ";Pwd=" & LCase(GetFileTitle(DBFileName)) & ";"
        End If
        OpenConnection = (DBConn.State = adStateOpen)
        
        If Err.Number <> 0 Then
            RaiseEvent DBError(Err.Number, Err.Description, Retry)
            Err.Clear
        End If
        
    Loop Until Retry = False

    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function rsQuery(ByVal rs As ADODB.Recordset, ByVal sqlStr As String) As Boolean
    On Error Resume Next
    
    Dim Retry As Boolean
    Do
        Retry = False
            
        If rs.State <> 0 Then rs.Close
        rs.Open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
        rsQuery = (rs.State = adStateOpen)
        If Not rs.State = 0 Then
            If Not rsEnd(rs) Then rs.MoveFirst
        End If
        
        If Err.Number <> 0 Then
            RaiseEvent DBError(Err.Number, Err.Description, Retry)
            Err.Clear
        End If
    
    Loop Until Retry = False
    
    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function dbQuery(ByVal sqlStr As String) As Boolean
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset
    
    Dim Retry As Boolean
    Do
        Retry = False

        If rs.State <> 0 Then rs.Close
        rs.Open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
        dbQuery = (rs.State = adStateOpen)
                
        If Err.Number <> 0 Then
            RaiseEvent DBError(Err.Number, Err.Description, Retry)
            Err.Clear
        End If
    
    Loop Until Retry = False
    
    rsClose rs
    
    If Err Then Err.Clear
    On Error GoTo 0
    
End Function

Public Function isField(ByVal TableName As String, ByVal FieldName As String) As Boolean
    isField = False
    
    Dim rs As New ADODB.Recordset
    rsQuery rs, "SELECT * FROM " & TableName & ";"
    
    Dim fld As Field
    For Each fld In rs.Fields
        If LCase(Trim(fld.name)) = LCase(Trim(FieldName)) Then
            isField = True
        End If
    Next fld
    
    rsClose rs

End Function

Public Function isTable(ByVal TableName As String) As Boolean
    isTable = False
    
    Dim rs As ADODB.Recordset
    Set rs = DBConn.OpenSchema(adSchemaTables)
    
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

Public Function rsClose(ByRef rs As ADODB.Recordset)
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End Function

Public Function dbClose()
    If DBConn.State <> 0 Then DBConn.Close
    Set DBConn = Nothing
End Function

Private Sub Class_Initialize()

    Set DBConn = New ADODB.Connection
    OpenConnection
    
End Sub

Private Sub Class_Terminate()
    
    dbClose

End Sub

