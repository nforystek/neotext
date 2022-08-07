


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private dbConn As clsDBConnection

Private Sub Class_Initialize()
    
    Set dbConn = New clsDBConnection
End Sub
Private Sub Class_Terminate()

    Set dbConn = Nothing
End Sub

Private Function GetAssocRecord(ByRef rs As ADODB.Recordset, ByVal FileExt As String) As Integer
    dbConn.rsQuery rs, "SELECT * FROM FileAssociations WHERE Extentions LIKE '%" & FileExt & "%';"
End Function

Public Function GetTransferType(ByVal FileExt As String) As Integer
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetTransferType = rs("TransferType")
    
    Else
        
        GetTransferType = 0

    End If

    rsClose rs
End Function

Public Function GetAssumeLineFeed(ByVal FileExt As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetAssumeLineFeed = CBool(rs("AssumeLineFeed"))
    
    Else
        
        GetAssumeLineFeed = True

    End If

    rsClose rs
End Function

Public Function GetDisplayName(ByVal FileExt As String) As String
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetDisplayName = rs("DisplayName")
    
    Else
        
        GetDisplayName = "Unknown File Type"

    End If

    rsClose rs
End Function

Public Function GetExtentions(ByVal FileExt As String) As String
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetExtentions = rs("Extentions")
        
    Else
        
        GetExtentions = ""

    End If

    rsClose rs
End Function

Public Function GetWindowsApp(ByVal FileExt As String) As Boolean
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetWindowsApp = rs("WindowsApp")
        
    Else
    
        GetWindowsApp = True

    End If

    rsClose rs
End Function

Public Function GetApplicationExe(ByVal FileExt As String) As String
    
    Dim rs As New ADODB.Recordset
    GetAssocRecord rs, FileExt
    
    If Not rsEnd(rs) Then
    
        GetApplicationExe = rs("ApplicationExe")
        
    Else
    
        GetApplicationExe = ""

    End If

    rsClose rs
End Function