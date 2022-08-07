VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Public ItemName As String
Public ItemClass As String
Public ItemSource As String
Public ItemPath As String

Public Function GetXMLText() As String
    Dim e As New NTShell22.Internet
    Dim outText As String
    
    outText = "<Item>" & vbCrLf
    outText = outText & "<Name>" & e.URLEncode(ItemName) & "</Name>" & vbCrLf
    outText = outText & "<Class>" & e.URLEncode(ItemClass) & "</Class>" & vbCrLf
    outText = outText & "<Source>" & e.URLEncode(ItemSource) & "</Source>" & vbCrLf
    outText = outText & "</Item>" & vbCrLf
    
    Set e = Nothing
    GetXMLText = outText
End Function

Public Function LoadFromFile(Optional ByVal fName As String = "", Optional ByVal fScript As String = "")
    Dim txt As String
    
    Dim xml As New msxml.DOMDocument
    
    txt = Replace(ReadFile(fName), "<?xml version=""1.0""?>", "")
    txt = Replace(txt, "<?xml version=""1.0"" ?>", "")
    txt = Replace(txt, "<? xml version=""1.0""?>", "")
    txt = Replace(txt, "<? xml version=""1.0"" ?>", "")
    
    xml.async = "false"
    xml.loadXML txt
    
    LoadFromXML xml.childNodes(0), fScript
    
    Set xml = Nothing
    
    ItemPath = fName
End Function

Public Function LoadFromXML(ByRef xml As msxml.IXMLDOMNode, Optional ByVal fScript As String = "")
    ItemPath = ""
    
    If xml.baseName = "Item" Then
        Dim pst As New NTShell22.Internet
        Dim child As msxml.IXMLDOMNode
        For Each child In xml.childNodes
            Select Case child.baseName
                Case "Name"
                    ItemName = pst.URLDecode(child.Text)
                Case "Class"
                    ItemClass = pst.URLDecode(child.Text)
                Case "Source"
                    Dim nSource As String
                    nSource = pst.URLDecode(child.Text)
                    If ((InStr(nSource, "#JScript" & vbCrLf) > 0) Or (InStr(nSource, "#VBScript" & vbCrLf) > 0)) And (Not (fScript = "")) Then
                    
                        Dim nScript As String
                        Dim nLine As String
                        Do Until nSource = ""
                            nLine = RemoveNextArg(nSource, vbCrLf)
                            If Left(nLine, 1) = "#" Then
                                nScript = Mid(nLine, 2)
                            ElseIf nScript = fScript Then
                                ItemSource = ItemSource & nLine & vbCrLf
                            End If
                        Loop
                    Else
                        ItemSource = nSource
                    End If
                    
            End Select
        Next
        Set child = Nothing
        Set pst = Nothing
        
    End If
End Function


Attribute 