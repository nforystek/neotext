#Const [True] = -1
#Const [False] = 0
Attribute VB_Name = "modMsgQueue"
Option Explicit
'TOP DOWN

Option Compare Binary

Option Private Module

Private rs As New ADODB.Recordset
    
Public Function MessageQueueLog(ByVal MsgTo As String) As Long

    Access.rsQuery rs, "SELECT Count(*) as Cnt FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"
    
    Dim ret As Long
    ret = CLng(rs("Cnt"))
    Access.rsClose rs
    
    MessageQueueLog = ret
        
End Function

Public Function MessageQueueGet(ByVal MsgTo As String) As Collection
    
    Dim mCol As New Collection

    Access.rsQuery rs, "SELECT * FROM MessageQueue WHERE MessageTo='" & MsgTo & "';"

    MsgTo = "DELETE FROM MessageQueue WHERE "
    Do Until Access.rsEnd(rs)

        mCol.Add CStr(rs("MessageText"))
        MsgTo = MsgTo & "ID=" & rs("ID") & " OR "

        rs.MoveNext
    Loop
    
    MsgTo = Left(MsgTo, Len(MsgTo) - 4) & ";"
    
    Access.rsClose rs
    
    Access.dbQuery MsgTo
    
    Set MessageQueueGet = mCol

End Function

Public Sub MessageQueueAdd(ByVal MsgTo As String, ByVal msg As String)
    
    Access.dbQuery "INSERT INTO MessageQueue (MessageTo, MessageText) VALUES ('" & MsgTo & "','" & Replace(msg, "'", "''") & "');"

End Sub


