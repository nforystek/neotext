#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMsgQueue"



#Const modMsgQueue = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Private db As New clsDBConnection
Private rs As New ADODB.Recordset
    
Public Function MessageQueueLog(ByVal MsgTo As String) As Long

    db.rsQuery rs, "SELECT Count(*) as Cnt FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"
    
    Dim ret As Long
    ret = CLng(rs("Cnt"))
    rsClose rs
    
    MessageQueueLog = ret
        
End Function

Public Function MessageQueueGet(ByVal MsgTo As String) As VBA.Collection
    
    Dim mCol As New VBA.Collection

    db.rsQuery rs, "SELECT * FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"

    MsgTo = "DELETE FROM MessageQueue WHERE "
    Do Until rsEnd(rs)

        mCol.Add CStr(rs("MessageText"))
        MsgTo = MsgTo & "ID=" & rs("ID") & " OR "

        rs.MoveNext
    Loop
    
    MsgTo = Left(MsgTo, Len(MsgTo) - 4) & ";"
    
    rsClose rs
    
    db.dbQuery MsgTo
    
    Set MessageQueueGet = mCol

End Function

Public Sub MessageQueueAdd(ByVal MsgTo As String, ByVal Msg As String)
    
    db.dbQuery "INSERT INTO MessageQueue (MessageTo, MessageText) VALUES ('" & MsgTo & "','" & Replace(Msg, "'", "''") & "');"

End Sub




