Attribute VB_Name = "modDBSchedule"





#Const modDBSchedule = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module

Public Type ScheduleType
    
    ParentID As Long
    ScheduleName As String
    IsPublic As Boolean
    Disabled As Boolean
    
    ScheduleType As Integer
    IncrementType As Integer
    IncrementInterval As Integer
    ExecuteTime As String
    ExecuteDate As String

End Type

Public Type ScheduleOperationType
    
    ScheduleID As Long
    OperationID As Long
    Disabled As Boolean
    
    OperationName As String
    OperationOrder As Long
    
    Action As String
    Overwrite As Boolean
    SubFolders As Boolean

    LastRun As String
    WildCard As String
    RenameNew As String
    
    SLogin As String
    SPass As String
    SURL As String
    sPort As Integer
    SAnon As Boolean
    
    DLogin As String
    DPass As String
    DURL As String
    DPort As Integer
    DAnon As Boolean
    
End Type

Public Function AddSchedule(ScheduleInfo As ScheduleType) As Long

    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset

    With ScheduleInfo
        If Not (.ScheduleName = "") Then
            dbConn.rsQuery rs, "SELECT * FROM Schedules WHERE ParentID=" & .ParentID & " AND ScheduleName='" & Replace(.ScheduleName, "'", "''") & "';"
            If rsEnd(rs) Then .ScheduleName = ""
        End If
        
        If (.ScheduleName = "") Then
        
            .ScheduleName = Replace(GUID, "-", "")
            
            dbConn.rsQuery rs, "INSERT INTO Schedules (ParentID, ScheduleName, IsPublic," & _
                                                            "ScheduleType, IncrementType, IncrementInterval, ExecuteDate, ExecuteTime" & _
                                                            ") VALUES (" & .ParentID & ", '" & Replace(.ScheduleName, "'", "''") & "', " & CheckToBool(.IsPublic) & ", " & _
                                                            .ScheduleType & ", " & .IncrementType & ", " & .IncrementInterval & ", '" & Replace(.ExecuteDate, "'", "''") & "', '" & Replace(.ExecuteTime, "'", "''") & "');"
        
            dbConn.rsQuery rs, "SELECT * FROM Schedules WHERE ScheduleName='" & Replace(.ScheduleName, "'", "''") & "';"
        
            AddSchedule = rs("ID")
                        
            dbConn.rsQuery rs, "UPDATE Schedules SET ScheduleName='' WHERE ScheduleName='" & Replace(.ScheduleName, "'", "''") & "';"
            .ScheduleName = ""
            
        Else
        
            AddSchedule = rs("ID")
        End If
        
    End With

    rsClose rs
    Set dbConn = Nothing
End Function

Public Function RemoveSchedule(ByVal ScheduleID As Long)
    Dim dbConn As New clsDBConnection
    
    dbConn.dbQuery "DELETE FROM Operations WHERE ParentID=" & ScheduleID & ";"
    dbConn.dbQuery "DELETE FROM Schedules WHERE ID=" & ScheduleID & ";"
    
    Set dbConn = Nothing
End Function

Public Function AddOperation(OperationInfo As ScheduleOperationType) As Long

    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    With OperationInfo
        
        dbConn.rsQuery rs, "SELECT * FROM Operations WHERE ParentID=" & .ScheduleID & " AND ID=" & .OperationID & ";"
        
        If rsEnd(rs) Then
            Dim tmpG As String
            tmpG = Replace(modGuid.GUID, "-", "")
            dbConn.rsQuery rs, "INSERT INTO Operations (" & _
                                            "ParentID, OperationName" & _
                                            ") VALUES (" & _
                                            "" & .ScheduleID & ", '" & tmpG & "');"
        
            dbConn.rsQuery rs, "SELECT * FROM Operations WHERE OperationName='" & tmpG & "';"
        
        End If
        
        AddOperation = rs("ID")
        
    End With

    rsClose rs
    Set dbConn = Nothing
End Function

Public Function RemoveOperation(ByVal OperationID As Long)
    Dim dbConn As New clsDBConnection
    
    dbConn.dbQuery "DELETE FROM Operations WHERE ID=" & OperationID & ";"
    
    Set dbConn = Nothing
End Function

Public Function CountOperation(ByVal ScheduleID As Long) As Long
    Dim dbConn As New clsDBConnection
    Dim rs As New ADODB.Recordset
    
    dbConn.rsQuery rs, "SELECT Count(OperationOrder) as Cnt FROM Operations WHERE ParentID=" & ScheduleID & ";"
    CountOperation = CLng(rs("Cnt"))
    
    rsClose rs
    Set dbConn = Nothing
End Function

