
Attribute VB_Name = "modDatabase"

#Const modDatabase = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module
Dim dbConnection As Object

Function OpenConnection()
    On Error Resume Next
    
    If dbConnection Is Nothing Then
        Set dbConnection = CreateObject("ADODB.Connection")
    End If
    
    If dbConnection.State <> 0 Then dbConnection.Close
    dbConnection.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & Settings.Location & ";"
    
    OpenConnection = (Not dbConnection.State = 0)

End Function

Function OpenRecordSet(ByRef OnRecord, ByVal SQLStr)
    
    If OnRecord Is Nothing Then
        Set OnRecord = CreateObject("ADODB.RecordSet")
    Else
        If OnRecord.State <> 0 Then OnRecord.Close
    End If
    OnRecord.Open SQLStr, dbConnection, , 3
    OpenRecordSet = (OnRecord.State = 1)

End Function

 Function CloseRecordSet(ByRef OnRecord)
    If Not OnRecord Is Nothing Then
        If OnRecord.State <> 0 Then OnRecord.Close
        Set OnRecord = Nothing
    End If
End Function

Function CloseConnection()
    If Not dbConnection Is Nothing Then
        If dbConnection.State <> 0 Then dbConnection.Close
        Set dbConnection = Nothing
    End If
End Function


Public Sub DBRefreshList(ByVal ListView, ByVal Filter As String)
    OpenConnection

    ListView.ListItems.Clear
    Dim newnode
    
    Dim rs As Object
            
    If Filter = "None" Then
        OpenRecordSet rs, "SELECT * FROM Changes ORDER BY cDateTime;"
    Else
        OpenRecordSet rs, "SELECT * FROM Changes WHERE NOT cStatus='" & Filter & "' ORDER BY cDateTime;"
    End If
    
    If Not rs.EOF And Not rs.BOF Then
        
        rs.MoveFirst
        Do
        
            Set newnode = ListView.ListItems.Add(, , DecryptString(rs("cDateTime") & ""))
            newnode.SubItems(1) = DecryptString(rs("cProduct") & "")
            newnode.SubItems(2) = DecryptString(rs("cType") & "")
            newnode.SubItems(3) = DecryptString(rs("cComments") & "")
            newnode.SubItems(4) = DecryptString(rs("cStatus") & "")
        
            newnode.Tag = rs("ID")
            
            rs.MoveNext
        Loop Until rs.EOF Or rs.BOF
        
    End If
    
    CloseRecordSet rs
    CloseConnection
     
End Sub

Public Function DBInsertRecord(ByVal cDateTime As String, ByVal cProduct As String, ByVal cType As String, ByVal cComments As String, ByVal cStatus As String) As Long
    OpenConnection
    Dim rs As New ADODB.Recordset
    
    Dim tmpID As String
    
    tmpID = -Int((10000 * Rnd) + 1) & " tmp " & -Int((10000 * Rnd) + 1)
    
    OpenRecordSet rs, "INSERT INTO Changes (cDateTime, cProduct, cType, cComments, cStatus) VALUES ('" & tmpID & "','" & EncryptString(cProduct) & "','" & EncryptString(cType) & "','" & EncryptString(cComments) & "','" & EncryptString(cStatus) & "');"
    
    OpenRecordSet rs, "SELECT * FROM Changes WHERE cDateTime='" & tmpID & "';"
    
    tmpID = rs("ID")
    
    OpenRecordSet rs, "UPDATE Changes SET cDateTime='" & EncryptString(cDateTime) & "' WHERE ID=" & rs("ID") & ";"

    DBInsertRecord = CLng(tmpID)

    CloseRecordSet rs
    CloseConnection
End Function
Public Function DBUpdateRecord(ByVal ID As Long, ByVal cDateTime As String, ByVal cProduct As String, ByVal cType As String, ByVal cComments As String, ByVal cStatus As String) As Long
    OpenConnection
    Dim rs As New ADODB.Recordset

    OpenRecordSet rs, "UPDATE Changes SET cDateTime='" & EncryptString(cDateTime) & "', cProduct='" & EncryptString(cProduct) & "', cType='" & EncryptString(cType) & "', cComments='" & EncryptString(cComments) & "', cStatus='" & EncryptString(cStatus) & "' WHERE ID=" & ID & ";"

    CloseRecordSet rs
    CloseConnection
End Function
Public Function DBDeleteRecord(ByVal ID As Long) As Long
    OpenConnection
    Dim rs As New ADODB.Recordset

    OpenRecordSet rs, "DELETE * FROM Changes WHERE ID=" & ID & ";"

    CloseRecordSet rs
    CloseConnection
End Function
 