Attribute VB_Name = "modWebForm"
Option Explicit
'TOP DOWN

Option Private Module
Public Function PostForm(ByVal HostName As String, ByVal HostFile As String, ByVal FormData As String, Optional ByVal ReturnXML As Boolean = False) As Variant
    Dim retVal As String
    
    retVal = PostToWebsite(HostName, HostFile, FormData)
    
    If ReturnXML Then
        If Right(retVal, 2) = "OK" Then
            Dim xml As New msxml2.DOMDocument
            xml.async = "false"
            xml.loadXML Left(retVal, Len(retVal) - 2)
            Set PostForm = xml
            Set xml = Nothing
        Else
            Set PostForm = Nothing
        End If
    Else
        PostForm = CBool((Right(retVal, 2) = "OK"))
    End If
End Function

