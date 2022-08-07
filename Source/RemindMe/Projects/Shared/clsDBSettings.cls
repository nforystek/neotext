VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDBSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dbRecord As ADODB.Recordset

Private Sub Class_Initialize()
    
    Set dbRecord = New ADODB.Recordset

End Sub
Private Sub Class_Terminate()
    
    If dbRecord.State <> 0 Then dbRecord.Close
    Set dbRecord = Nothing

End Sub

Public Function GetSetting(ByVal Column As String) As Variant
    
    If DBConn.rsQuery(dbRecord, "SELECT * FROM Settings WHERE ID = 1;") Then
        GetSetting = dbRecord(Column)
    Else
        GetSetting = Null
    End If

End Function

Public Sub SetSetting(ByVal Column As String, ByVal Value As Variant)
    
    If TypeName(Value) = "String" Then
        DBConn.rsQuery dbRecord, "UPDATE Settings SET " & Column & "='" & Value & "' WHERE ID=1;"
    Else
        DBConn.rsQuery dbRecord, "UPDATE Settings SET " & Column & "=" & Value & " WHERE ID=1;"
    End If

End Sub

Public Function MessageWaiting(ByVal MsgTo As String) As Boolean
    DBConn.rsQuery dbRecord, "SELECT Count(*) as Cnt FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"
    MessageWaiting = (dbRecord("Cnt") > 0)
    If dbRecord.State <> 0 Then dbRecord.Close
End Function

Public Function MessageQueue(ByVal MsgTo As String, Optional ByVal Count As Long = -1) As Collection
    Dim mCol As New Collection
    Dim rsDel As New ADODB.Recordset
    
    If Not (Count = 0) Then
    
        DBConn.rsQuery dbRecord, "SELECT * FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"
    
        If (Not rsEnd(dbRecord)) Then
            Dim nextStr As String
            Dim sqlStr As String
            dbRecord.MoveFirst
            Do
                nextStr = dbRecord("MessageText")
                DBConn.rsQuery rsDel, "DELETE FROM MessageQueue WHERE ID=" & dbRecord("ID") & ";"
                
                mCol.Add nextStr
                
                dbRecord.MoveNext
                If Not (Count = -1) Then Count = Count - 1
            Loop Until rsEnd(dbRecord) Or (Count = 0)
        End If
        If dbRecord.State <> 0 Then dbRecord.Close
        
        If Not rsDel.State = 0 Then rsDel.Close
        Set rsDel = Nothing
        
    End If
    
    Set MessageQueue = mCol
End Function
Public Function Message(ByVal MsgTo As String, ByVal Msg As String)
    DBConn.rsQuery dbRecord, "INSERT INTO MessageQueue (MessageTo, MessageText) VALUES ('" & MsgTo & "','" & Replace(Msg, "'", "''") & "');"
End Function


