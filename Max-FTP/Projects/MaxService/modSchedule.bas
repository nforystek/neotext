Attribute VB_Name = "modSchedule"
#Const modSchedule = -1
Option Explicit
'TOP DOWN
Option Compare Binary
Option Private Module
Private AllSchedules As NTNodes10.Collection

Public Sub PauseSchedules(ByVal Pause As Boolean)
    Dim sch As Schedule
    Dim opr As Operation
    Dim cnt As Long
    
    For Each sch In AllSchedules
        If sch.OperationCount > 0 Then
            For cnt = 1 To sch.OperationCount
                Set opr = sch.GetOperation(cnt)
                opr.PauseTransfers = True
                Set opr = Nothing
            Next
        End If
    Next
End Sub

Public Function AddSchedule() As Schedule
    Dim obj As New Schedule
    AllSchedules.Add obj
    Set obj = Nothing
    Set AddSchedule = AllSchedules(AllSchedules.Count)
End Function

Public Function GetSchedule(ByVal Sid As Long) As Schedule
    If AllSchedules.Count > 0 Then
        Dim cnt As Long
        For cnt = 1 To AllSchedules.Count
            If AllSchedules(cnt).ID = Sid Then
                Set GetSchedule = AllSchedules(cnt)
                Exit For
            End If
        Next
    End If
End Function

Public Function ScheduleExists(ByVal Sid As Long) As Boolean
    ScheduleExists = False
    If Not AllSchedules Is Nothing Then
        If AllSchedules.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To AllSchedules.Count
                If AllSchedules(cnt).ID = Sid Then
                    ScheduleExists = True
                    Exit Function
                End If
            Next
        End If
    End If
End Function

Public Sub LoadCmdLine(ByVal InCOmmand As String, ByVal InParams As String)
    Dim cmdSchedule As Schedule
    Dim cmdOperation As Operation
    Dim url As New NTAdvFTP61.url
    
    Dim tmpValue As String
    
    If AllSchedules Is Nothing Then
        Set AllSchedules = New NTNodes10.Collection
        Set ThreadManager = New clsThreadManager
    End If
    
    If Not ScheduleExists(-1) Then
        Set cmdSchedule = AddSchedule
    Else
        Set cmdSchedule = GetSchedule(-1)
    End If

    cmdSchedule.ID = -1
    cmdSchedule.name = dbSettings.GetUserLoginName

    Set cmdOperation = cmdSchedule.AddOperation
    cmdOperation.ID = cmdSchedule.OperationCount
    
    With cmdOperation
    
        .UserID = dbSettings.CurrentUserID
        .ScheduleID = cmdSchedule.ID
        
        .Action = UCase(Left(InCOmmand, 1)) & Mid(InCOmmand, 2)
        .WildCard = RemoveQuotedArg(InParams, """", """")
        
        If .Action = "Copy" Or .Action = "Move" Or .Action = "Delete" Then
            If UCase(NextArg(InParams, " ")) = "S" Then
                .SubFolders = True
                RemoveNextArg InParams, " "
            Else
                .SubFolders = False
            End If
        End If

        If .Action = "Copy" Or .Action = "Move" Then
            If UCase(NextArg(InParams, " ")) = "O" Then
                .Overwrite = True
                RemoveNextArg InParams, " "
            Else
                .Overwrite = False
            End If
        End If

        If .Action = "Copy" Or .Action = "Move" Then
            If UCase(NextArg(InParams, " ")) = "N" Then
                .OnlyNewer = True
                RemoveNextArg InParams, " "
            Else
                .OnlyNewer = False
            End If
        End If


        If .Action = "Script" Then
            If UCase(NextArg(InParams, " ")) = "W" Then
                .WaitForScript = True
                RemoveNextArg InParams, " "
            Else
                .WaitForScript = False
            End If
        End If

        If .Action = "Script" Then
            If IsNumeric(NextArg(InParams, " ")) Then
                .Seconds = CLng(NextArg(InParams, " "))
                RemoveNextArg InParams, " "
            Else
                .Seconds = IIf(.WaitForScript, -1, 0)
            End If
        End If
        
        
        If .Action = "Script" Then
            If UCase(NextArg(InParams, " ")) = "T" Then
                .ForceTerminate = True
                RemoveNextArg InParams, " "
            Else
                .ForceTerminate = False
            End If
        End If
        
        
        
'            chkOverwrite.Value = BoolToCheck(.GetOperationValue(OpId, "Overwrite"))
'            chkWaitForScript.Value = chkOverwrite.Value
'            chkOnlyNewer.Value = BoolToCheck(.GetOperationValue(OpId, "OnlyNewer"))
'            chkForceTerminate.Value = chkOnlyNewer.Value
'            chkSubFolders.Value = BoolToCheck(.GetOperationValue(OpId, "SubFolders"))
'            chkNoInterface.Value = chkSubFolders.Value
'            txtWildCard.Tag = False
'            txtWildCard.Text = .GetOperationValue(OpId, "WildCard")
'            If cmbOperation.Text = "Script" Then
'                If IsNumeric(.GetOperationValue(OpId, "RenameNew")) Then
'                    Check1.Value = Abs((CLng(.GetOperationValue(OpId, "RenameNew")) > 0))
'                End If
'                txtSeconds.Text = .GetOperationValue(OpId, "RenameNew")
'
'            Else
'                txtRename(1).Text = .GetOperationValue(OpId, "RenameNew")
'            End If
'            If IsNumeric(txtRename(1).Text) Then txtSeconds.Text = txtRename(1).Text
            
            
            
        If .Action = "Rename" Then
            .RenameNew = RemoveQuotedArg(InParams, """", """")
        End If
        
        tmpValue = RemoveQuotedArg(InParams, """", """")
        If Not (tmpValue = "") Then .SURL = tmpValue
        .sport = url.GetPort(tmpValue)
        If Not (url.GetUserName(tmpValue) = "") Then .SLogin = url.GetUserName(tmpValue)
        If Not (url.GetPassword(tmpValue) = "") Then .SPass = url.GetPassword(tmpValue)
        If UCase(NextArg(Left(InParams, 1), " ")) = "A" Then
            .SPasv = False
            tmpValue = Mid(RemoveNextArg(InParams, " "), 2)
            If InStr(tmpValue, ":") > 0 Then
                .DData = NextArg(tmpValue, ":")
                .SAdap = RemoveArg(tmpValue, ":")
            End If
        Else
            .SPasv = True
        End If

        If .Action = "Copy" Or .Action = "Move" Then
        
            tmpValue = RemoveQuotedArg(InParams, """", """")
            If Not (tmpValue = "") Then .DURL = tmpValue
            .DPort = url.GetPort(tmpValue)
            If Not (url.GetUserName(tmpValue) = "") Then .DLogin = url.GetUserName(tmpValue)
            If Not (url.GetPassword(tmpValue) = "") Then .DPass = url.GetPassword(tmpValue)
            If UCase(NextArg(InParams, " ")) = "A" Then
                .DPasv = False
                tmpValue = Mid(RemoveNextArg(InParams, " "), 2)
                If InStr(tmpValue, ":") > 0 Then
                    .DData = NextArg(tmpValue, ":")
                    .SAdap = RemoveArg(tmpValue, ":")
                End If
            Else
                .DPasv = True
            End If
        End If
        
    End With

    Set cmdOperation = Nothing

    Set cmdSchedule = Nothing

    Set dbSettings = Nothing
End Sub
Public Sub LoadSchedule(Optional ByVal ScheduleID As Long = 0)
On Error GoTo errorcatch:

    If AllSchedules Is Nothing Then
        Set AllSchedules = New NTNodes10.Collection
        Set ThreadManager = New clsThreadManager
    End If
    
    Dim dbConn As New clsDBConnection
    Dim rsSchedule As New ADODB.Recordset
    Dim rsOperation As New ADODB.Recordset
    
    Dim newSchedule As Schedule
    Dim newOperation As Operation
    
    If ScheduleID = 0 Then
        UnloadSchedules True
        
        dbConn.rsQuery rsSchedule, "SELECT * FROM Schedules;"
    Else
        If ScheduleExists(ScheduleID) Then
            UnloadSchedule ScheduleID, True
        End If
        dbConn.rsQuery rsSchedule, "SELECT * FROM Schedules WHERE ID=" & ScheduleID & ";"
    End If
    
    Do Until rsEnd(rsSchedule)
    
        If Not ScheduleExists(rsSchedule("ID")) Then
            Set newSchedule = AddSchedule
        Else
            Set newSchedule = GetSchedule(rsSchedule("ID"))
            newSchedule.ClearOperations
        End If

        newSchedule.ID = rsSchedule("ID")
        newSchedule.name = rsSchedule("ScheduleName")
        
        With newSchedule.RemindMe
            
            .Enabled = False
            .ScheduleType = rsSchedule("ScheduleType")
            .IncrementType = rsSchedule("IncrementType")
            If Not rsSchedule("IncrementInterval") = 0 Then .IncrementInterval = rsSchedule("IncrementInterval")
            If Not rsSchedule("ExecuteDate") = "" Then .ExecuteDate = rsSchedule("ExecuteDate")
            If Not rsSchedule("ExecuteTime") = "" Then .ExecuteTime = rsSchedule("ExecuteTime")
            
            dbConn.rsQuery rsOperation, "SELECT * FROM Operations WHERE ParentID=" & rsSchedule("ID") & " ORDER BY OperationOrder;"
            
            Do Until rsEnd(rsOperation)
                
                Set newOperation = newSchedule.AddOperation
                                
                newOperation.ID = rsOperation("ID")
                
                With newOperation
                
                    .UserID = rsSchedule("ParentID")
                    .ScheduleID = rsSchedule("ID")
                    
                    .Action = rsOperation("Action")
                    .Overwrite = rsOperation("Overwrite")
                    .OnlyNewer = rsOperation("OnlyNewer")
                    .SubFolders = rsOperation("SubFolders")
                    .WildCard = rsOperation("WildCard")
                    .RenameNew = rsOperation("RenameNew")
                    
                    .SURL = rsOperation("SURL")
                    .SLogin = rsOperation("SLogin")
                    .SPass = rsOperation("SPass")
                    .sport = rsOperation("SPort")
                    .SPasv = rsOperation("SPasv")
                    .SData = rsOperation("SData")
                    .SAdap = rsOperation("SAdap")
                    .SSSL = rsOperation("SSSL")
                    
                    .DURL = rsOperation("DURL")
                    .DLogin = rsOperation("DLogin")
                    .DPass = rsOperation("DPass")
                    .DPort = rsOperation("DPort")
                    .DPasv = rsOperation("DPasv")
                    .DData = rsOperation("DData")
                    .DAdap = rsOperation("DAdap")
                    .DSSL = rsOperation("DSSL")
                    
                End With

                Set newOperation = Nothing
                
                rsOperation.MoveNext
                
            Loop
        
            .Enabled = IsServiceFormStarted
                    
        End With
        
        Set newSchedule = Nothing
        
        rsSchedule.MoveNext
        
    Loop
    
    rsClose rsOperation
    rsClose rsSchedule
    
    Set dbConn = Nothing

errorcatch:
    If Err Then Err.Clear
    On Error GoTo 0
End Sub

Public Sub UnloadSchedules(Optional ByVal NoUnload As Boolean = False)
    If Not (AllSchedules Is Nothing) Then
        If AllSchedules.Count > 0 Then
            Dim cnt As Long
            For cnt = 1 To AllSchedules.Count
            
                AllSchedules(cnt).ClearOperations
            Next
            Do Until AllSchedules.Count = 0
                AllSchedules.Remove 1
            Loop
            
        End If
        If Not NoUnload Then
            Set AllSchedules = Nothing
            Set ThreadManager = Nothing
        End If
    End If
End Sub

Public Sub UnloadSchedule(ByVal ScheduleID As Long, Optional ByVal NoUnload As Boolean = False)
    
    If (AllSchedules.Count > 0) Then

        Dim cnt As Long
        Dim ScheduleIndex As Long
        For cnt = 1 To AllSchedules.Count
            If AllSchedules(cnt).ID = ScheduleID Then
                AllSchedules(cnt).ClearOperations
                ScheduleIndex = cnt
                Exit For
            End If
        Next
        
        If ScheduleIndex > 0 Then
            AllSchedules.Remove ScheduleIndex
        End If
        If Not NoUnload Then
            If AllSchedules.Count = 0 Then
                Set AllSchedules = Nothing
                Set ThreadManager = Nothing
            End If
        End If
    End If
    
End Sub






