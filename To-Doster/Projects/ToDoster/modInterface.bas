
Attribute VB_Name = "modInterface"
#Const modInterface = -1
Option Explicit
'TOP DOWN
Option Compare Text
Option Private Module

Public Function TestConnection() As Boolean
    On Error Resume Next
    If IsURL(Settings.Location) Then
        Debug.Print PostExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=TestConnection", GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))
        TestConnection = PostExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=TestConnection", GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))
    Else
        TestConnection = OpenConnection
        CloseConnection
    End If
End Function


Public Sub RefreshList(ByVal ListView, ByVal Filter As String)
    On Error GoTo failure
    If IsURL(Settings.Location) Then
        Dim xml
        
        Set xml = ServerExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=RefreshList&Filter=" & Filter, GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))
    
        ListView.ListItems.Clear
    
        Dim X, newnode
        
        For Each X In xml.documentelement.childnodes
            Set newnode = ListView.ListItems.Add(, , DecryptString(URLDecode(X.childnodes(1).Text)))
            newnode.SubItems(1) = DecryptString(URLDecode(X.childnodes(2).Text))
            newnode.SubItems(2) = DecryptString(URLDecode(X.childnodes(3).Text))
            newnode.SubItems(3) = DecryptString(URLDecode(X.childnodes(4).Text))
            newnode.SubItems(4) = DecryptString(URLDecode(X.childnodes(5).Text))
        
            newnode.Tag = X.childnodes(0).Text
            
        Next
    
    Else
        DBRefreshList ListView, Filter
    End If
    Exit Sub
failure:
    MsgBox "There was an error attempting to query the database." & vbCrLf & _
            "Often times this is because of permission accessing" & vbCrLf & _
            "the web server or database setup preventing ability.", vbInformation, vbOKOnly
    Err.Clear
End Sub

Public Function InsertRecord(ByVal cDateTime As String, ByVal cProduct As String, ByVal cType As String, ByVal cComments As String, ByVal cStatus As String) As Long
    On Error GoTo failure
    
    If IsURL(Settings.Location) Then
        
        InsertRecord = PostExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=InsertRecord&cDateTime=" & URLEncode(EncryptString(cDateTime)) & "&cProduct=" & URLEncode(EncryptString(cProduct)) & "&cType=" & URLEncode(EncryptString(cType)) & "&cComments=" & URLEncode(EncryptString(cComments)) & "&cStatus=" & URLEncode(EncryptString(cStatus)), GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))

    Else
        InsertRecord = DBInsertRecord(cDateTime, cProduct, cType, cComments, cStatus)
    End If
    Exit Function
failure:
    MsgBox "There was an error attempting to query the database." & vbCrLf & _
            "Often times this is because of permission accessing" & vbCrLf & _
            "the web server or database setup preventing ability.", vbInformation, vbOKOnly
    Err.Clear
End Function

Public Function UpdateRecord(ByVal ID As Long, ByVal cDateTime As String, ByVal cProduct As String, ByVal cType As String, ByVal cComments As String, ByVal cStatus As String) As String
    On Error GoTo failure
    If IsURL(Settings.Location) Then
        
        UpdateRecord = PostExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=UpdateRecord&ID=" & ID & "&cDateTime=" & URLEncode(EncryptString(cDateTime)) & "&cProduct=" & URLEncode(EncryptString(cProduct)) & "&cType=" & URLEncode(EncryptString(cType)) & "&cComments=" & URLEncode(EncryptString(cComments)) & "&cStatus=" & URLEncode(EncryptString(cStatus)), GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))

    Else
        UpdateRecord = DBUpdateRecord(ID, cDateTime, cProduct, cType, cComments, cStatus)
    End If
    Exit Function
failure:
    MsgBox "There was an error attempting to query the database." & vbCrLf & _
            "Often times this is because of permission accessing" & vbCrLf & _
            "the web server or database setup preventing ability.", vbInformation, vbOKOnly
    Err.Clear
End Function

Public Function DeleteRecord(ByVal ID As Long) As String
    On Error GoTo failure
    If IsURL(Settings.Location) Then
        
        DeleteRecord = PostExecute(GetServerName(Settings.Location), GetServerWebForm(Settings.Location), "Method=DeleteRecord&ID=" & ID, GetUsername(Settings.Location), GetPassword(Settings.Location), IsSSL(Settings.Location))

    Else
        DeleteRecord = DBDeleteRecord(ID)
    End If
    Exit Function
failure:
    MsgBox "There was an error attempting to query the database." & vbCrLf & _
            "Often times this is because of permission accessing" & vbCrLf & _
            "the web server or database setup preventing ability.", vbInformation, vbOKOnly
    Err.Clear
End Function
 
