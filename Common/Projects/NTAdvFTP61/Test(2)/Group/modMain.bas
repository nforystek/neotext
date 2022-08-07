Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
On Error GoTo finish
    Dim g As New NTAdvFTP61.Group
    g.AuthUser = "117953718345"
    g.AuthPass = "4fb3b5d74f"
    g.Server = "free.xsusenet.com"
    g.Details = post
    g.PostAs = "nforystek@neotext.org"
    g.NewsGroup = "alt.bumbling.idiots.the.fbi"
    
    Dim gid As String
    Dim col As Collection
    
    g.Connect
    
    Set col = g.PutBinary("C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.exe", gid)
    
    
    g.GetBinary "C:\Development\Neotext\Max-FTP\Deploy\Max-FTP v6.1.0.nws", gid, col
    
    
finish:
    If Err Then Debug.Print Err.Description
    
    g.Disconnect
    
    Set g = Nothing
End Sub
