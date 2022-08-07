VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private mCol As Collection

Public Sub Clear()
    StopOperations
    Dim obj As clsOperation
    Do Until mCol.Count = 0
        Set obj = mCol.Item(1)
        obj.Enabled = False
        mCol.Remove 1
        Set obj = Nothing
    Loop
End Sub
Public Sub Remove(ByVal ID As Long)
    If Exists(ID) Then
        StopOperation ID
        mCol.Remove "K" & ID
    End If
End Sub

Public Function Add(ByVal ID As Long) As clsOperation
    Dim obj As New clsOperation
    mCol.Add obj, "K" & ID
    Set Add = obj
    Set obj = Nothing
End Function

Public Property Get Exists(ByVal ID As Long) As Boolean
    On Error Resume Next
    Dim Item As clsOperation
    Set Item = mCol.Item("K" & ID)
    If Err Then
        Err.Clear
        Exists = False
    Else
        Exists = True
    End If
    Set Item = Nothing
    On Error GoTo 0
End Property

Public Property Get Item(ByVal ID As Long) As clsOperation
    On Error Resume Next
    Set Item = mCol.Item("K" & ID)
    If Err Then Err.Clear
    On Error GoTo 0
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Clear
    Set mCol = Nothing
End Sub

Private Function GetParameters(ByVal ParentID As String) As String
    Dim dbRS As ADODB.Recordset
    Set dbRS = New ADODB.Recordset
    
    Dim tmpVal As String
    Dim retVal As String
    
    DBConn.rsQuery dbRS, "SELECT * FROM OperationParams WHERE ParentID=" & ParentID & " ORDER BY ParamNum;"
    While Not dbRS.EOF
        Select Case dbRS("ParamType")
            Case 1
                Select Case LCase(Trim(dbRS("ParamValue")))
                    Case "-1", "1", "yes", "true", "on"
                        retVal = retVal & "true, "
                    Case "0", "no", "false", "off"
                        retVal = retVal & "false, "
                End Select
            Case 2, 4
                retVal = retVal & dbRS("ParamValue") & ", "
            Case 3
                Select Case dbSettings.GetSetting("Language")
                    Case "VBScript"
                        tmpVal = Replace(dbRS("ParamValue"), vbCrLf, """ & vbCrLf & """)
                        tmpVal = Replace(Replace(tmpVal, vbCr, """ & vbCr & """), vbLf, """ & vbLf & """)
                        tmpVal = Replace(tmpVal, """", """""")
                        retVal = retVal & """" & tmpVal & """, "
                    Case "JScript"
                        tmpVal = Replace(Replace(Replace(dbRS("ParamValue"), vbCrLf, "\r\n"), vbCr, "\r"), vbLf, "\n")
                        tmpVal = Replace(Replace(Replace(tmpVal, "\", "\\"), "'", "\'"), """", "\""")
                        retVal = retVal & """" & tmpVal & """, "
                End Select
        End Select
        dbRS.MoveNext
    Wend
    
    If Not Trim(retVal) = "" Then
        retVal = Left(retVal, Len(retVal) - 2)
    End If

    GetParameters = retVal
    
    If Not dbRS.State = 0 Then dbRS.Close
    Set dbRS = Nothing
End Function

Friend Sub Load(Optional ByVal ID As Long = -1)
    
    Dim newItem As clsOperation
    Dim dbRS As ADODB.Recordset
    Set dbRS = New ADODB.Recordset
                
    If ID = -1 Then
        Clear
        DBConn.rsQuery dbRS, "SELECT * FROM Operations;"
    Else
        DBConn.rsQuery dbRS, "SELECT * FROM Operations WHERE ID=" & ID & ";"
    End If
    If Not rsEnd(dbRS) Then
        dbRS.MoveFirst
        Do
            If Exists(dbRS("ID")) Then
                Set newItem = Item(dbRS("ID"))
            Else
                Set newItem = Add(dbRS("ID"))
            End If
            With newItem
                .Enabled = False
                
                .ID = dbRS("ID")
                .Name = dbRS("Name") & ""
                .Procedure = dbRS("ProcName")
                .Parameters = GetParameters(dbRS("ID"))
                
                .ScheduleType = dbRS("ScheduleType")
                .IncrementType = dbRS("IncrementType")
                
                If Not dbRS("IncrementInterval") = 0 Then .IncrementInterval = dbRS("IncrementInterval")
                If Not dbRS("ExecuteDate") = "" Then .ExecuteDate = dbRS("ExecuteDate")
                If Not dbRS("ExecuteTime") = "" Then .ExecuteTime = dbRS("ExecuteTime")
                
                .Enabled = CBool(dbRS("Enabled"))
            End With
            
            Set newItem = Nothing
            dbRS.MoveNext
        Loop Until rsEnd(dbRS)
    End If
    
    If Not dbRS.State = 0 Then dbRS.Close
    Set dbRS = Nothing
End Sub

Public Sub StopOperation(ByVal OperationID As Long)
    If Exists(OperationID) Then
        Item(OperationID).StopOperation
    End If
End Sub
Public Sub StartOperation(ByVal OperationID As Long)
    If Exists(OperationID) Then
        Item(OperationID).StartOperation
    End If
End Sub

Public Sub StopOperations()
    Dim Index As Long
    If mCol.Count > 0 Then
        For Index = 1 To mCol.Count
            mCol.Item(Index).StopOperation
        Next
    End If
End Sub

