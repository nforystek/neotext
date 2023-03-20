Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
On Error GoTo finish
    Dim g As New NTAdvFTP61.Group

    g.Server = "neotext.org"
    g.Details = post
    g.PostAs = "nforystek@neotext.org"
    g.NewsGroup = "neotext.binaries.documents"
    
    Dim gid As String
    Dim col As Collection
    
    g.Connect
    
    Set col = g.PutBinary("C:\Documents and Settings\Nickels\Desktop\RANDALL.PNG", gid)
    
    
    g.GetBinary "C:\Documents and Settings\Nickels\Desktop\RANDALL.nws", gid, col
    
    
finish:
    If Err Then Debug.Print Err.Description
    
    g.Disconnect
    
    Set g = Nothing
End Sub
