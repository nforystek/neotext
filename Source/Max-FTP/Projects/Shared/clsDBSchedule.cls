VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dbConn As clsDBConnection

Public Property Get ScheduleIDs() As VBA.Collection
    Dim rs As New ADODB.Recordset
    
    Dim ids As New VBA.Collection
    Dim ID As Long
    
    dbConn.rsQuery rs, "SELECT * FROM Schedules;"
    
    Do Until rsEnd(rs)
        ID = rs("ID")
        ids.Add ID

        rs.MoveNext
    Loop

    Set ScheduleIDs = ids
    
    rsClose rs
End Property

Public Function ScheduleNameExists(ByVal UserID As Long, ByVal ScheduleName As String) As Boolean
    Dim rs As New ADODB.Recordset
    
    dbConn.rsQuery rs, "SELECT * FROM Schedules WHERE ParentID=" & UserID & " AND ScheduleName='" & Replace(ScheduleName, "'", "''") & "';"
    ScheduleNameExists = Not rsEnd(rs)

    rsClose rs
End Function

Public Function GetScheduleOwner(ByVal ScheduleID As Long) As String
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT ParentID FROM Schedules WHERE ID = " & ScheduleID & ";"
    If Not rsEnd(rs) Then
        dbConn.rsQuery rs, "SELECT UserName FROM Users WHERE ID = " & rs("ParentID") & ";"
        If Not rsEnd(rs) Then
            GetScheduleOwner = rs("UserName")
        Else
            GetScheduleOwner = ""
        End If
    Else
        GetScheduleOwner = ""
    End If

    rsClose rs
End Function

Public Function GetOperationOwner(ByVal OperationID As Long) As String
    Dim rs As New ADODB.Recordset

    dbConn.rsQuery rs, "SELECT ParentID FROM Operations WHERE ID = " & OperationID & ";"
    If Not rsEnd(rs) Then
        dbConn.rsQuery rs, "SELECT ParentID FROM Schedules WHERE ID = " & rs("ParentID") & ";"
        If Not rsEnd(rs) Then
            dbConn.rsQuery rs, "SELECT UserName FROM Users WHERE ID = " & rs("ParentID") & ";"
            If Not rsEnd(rs) Then
                GetOperationOwner = rs("UserName")
            Else
                GetOperationOwner = ""
            End If
        Else
            GetOperationOwner = ""
        End If
    Else
        GetOperationOwner = ""
    End If

    rsClose rs
End Function

Public Function GetScheduleValue(ByVal ScheduleID As Long, ByVal Column As String) As Variant
    
    Dim sValue As Variant
    
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM Schedules WHERE ID=" & ScheduleID & ";") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If

    rsClose rs
    
    GetScheduleValue = sValue

End Function

Public Sub SetScheduleValue(ByVal ScheduleID As Long, ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE Schedules SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ID=" & ScheduleID & ";"
    Else
        dbConn.dbQuery "UPDATE Schedules SET " & Column & "=" & Value & " WHERE ID=" & ScheduleID & ";"
    End If

End Sub

Public Function GetOperationValue(ByVal OperationID As Long, ByVal Column As String) As Variant
    
    Dim sValue As Variant
    
    Dim rs As New ADODB.Recordset
    
    If dbConn.rsQuery(rs, "SELECT * FROM Operations WHERE ID=" & OperationID & ";") Then
        sValue = rs(Column)
    Else
        sValue = Null
    End If

    rsClose rs
    
    GetOperationValue = sValue

End Function

Public Sub SetOperationValue(ByVal OperationID As Long, ByVal Column As String, ByVal Value As Variant)
   
    If TypeName(Value) = "String" Then
        dbConn.dbQuery "UPDATE Operations SET " & Column & "='" & Replace(Value, "'", "''") & "' WHERE ID=" & OperationID & ";"
    Else
        dbConn.dbQuery "UPDATE Operations SET " & Column & "=" & Value & " WHERE ID=" & OperationID & ";"
    End If

End Sub

Private Sub Class_Initialize()
    Set dbConn = New clsDBConnection
End Sub

Private Sub Class_Terminate()
    Set dbConn = Nothing
End Sub
