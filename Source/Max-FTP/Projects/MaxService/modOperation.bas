Attribute VB_Name = "modOperation"
Option Explicit
'TOP DOWN
Option Private Module
Public AllSchedules As MaxService.clsSchedules

Public Sub LoadSchedule(Optional ByVal ScheduleID As Long = 0)
On Error GoTo errorcatch:
    
    Dim dbConn As New clsDBConnection
    Dim rsSchedule As New ADODB.Recordset
    Dim rsOperation As New ADODB.Recordset
    
    Dim newSchedule As MaxService.clsSchedule
    Dim newOperation As MaxService.clsOperation
    
    If ScheduleID = 0 Then
        UnloadSchedules
        dbConn.rsQuery rsSchedule, "SELECT * FROM Schedules;"
    Else
        dbConn.rsQuery rsSchedule, "SELECT * FROM Schedules WHERE ID=" & ScheduleID & ";"
    End If
    
    Do Until rsEnd(rsSchedule)
    
        Set newSchedule = AllSchedules.Add(rsSchedule("ID"), rsSchedule("ScheduleName"))
        newSchedule.ThreadID = rsSchedule("ThreadID")
        
        With newSchedule.RemindMe
            .ScheduleType = rsSchedule("ScheduleType")
            .IncrementType = rsSchedule("IncrementType")
            If Not rsSchedule("IncrementInterval") = 0 Then .IncrementInterval = rsSchedule("IncrementInterval")
            If Not rsSchedule("ExecuteDate") = "" Then .ExecuteDate = rsSchedule("ExecuteDate")
            If Not rsSchedule("ExecuteTime") = "" Then .ExecuteTime = rsSchedule("ExecuteTime")

            dbConn.rsQuery rsOperation, "SELECT * FROM Operations WHERE ParentID=" & rsSchedule("ID") & " ORDER BY OperationOrder;"
            
            Do Until rsEnd(rsOperation)
                
                Set newOperation = newSchedule.Operations.Add(rsOperation("ID"))
                
                With newOperation
                
                    .UserID = rsSchedule("ParentID")
                    .ScheduleID = rsSchedule("ID")
                    .ThreadID = newSchedule.ThreadID
                    
                    .Action = rsOperation("Action")
                    .Overwrite = rsOperation("Overwrite")
                    .SubFolders = rsOperation("SubFolders")
                    .WildCard = rsOperation("WildCard")
                    .RenameNew = rsOperation("RenameNew")
                    
                    .SURL = rsOperation("SURL")
                    .SLogin = rsOperation("SLogin")
                    .SPass = rsOperation("SPass")
                    .SPort = rsOperation("SPort")
                    .SPasv = rsOperation("SPasv")
                    .SData = rsOperation("SData")
                    
                    .DURL = rsOperation("DURL")
                    .DLogin = rsOperation("DLogin")
                    .DPass = rsOperation("DPass")
                    .DPort = rsOperation("DPort")
                    .DPasv = rsOperation("DPasv")
                    .DData = rsOperation("DData")
                    
                End With
                
                rsOperation.MoveNext
                
                Set newOperation = Nothing
                
            Loop
            
        
            .Enabled = True
                    
        End With
        
        rsSchedule.MoveNext
        
        Set newSchedule = Nothing
    Loop
    
    
    rsClose rsOperation
    rsClose rsSchedule
    
    Set dbConn = Nothing


errorcatch:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub UnloadSchedules()
    Do Until AllSchedules.Count = 0
        AllSchedules.Item(1).StopSchedule
        AllSchedules.Remove 1
    Loop
End Sub

Public Sub UnloadSchedule(ByVal ScheduleID As Long)
    Dim op As MaxService.clsOperation
    Dim sch As MaxService.clsSchedule
    
    Dim found As Integer
    Dim cnt As Integer
    found = 0
    For cnt = 1 To AllSchedules.Count
        Set sch = AllSchedules.Item(cnt)
        If ScheduleID = sch.id Then
            found = cnt
            Exit For
        End If
        Set sch = Nothing
    Next
    
    If Not (sch Is Nothing) Then
        sch.Operations.Clear
        AllSchedules.Remove found
        Set sch = Nothing
    End If
End Sub

Public Sub ReloadSchedule(ByVal ScheduleID As Long)
    UnloadSchedule ScheduleID
    LoadSchedule ScheduleID
End Sub



