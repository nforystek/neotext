Attribute VB_Name = "modTree"
#Const modTree = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module
Public XMLFilePath As String
Public XMLPadding As String
Public Function MakeDir(ByVal Path As String)
    If Not PathExists(Path) Then MkDir Path
End Function
Public Function WriteHTMLText(ByRef nBase As clsItem, ByVal FilePath As String, ByVal IncludeCustom As Boolean, ByVal IncludeCode As Boolean) As String
    Dim html As String
    html = "<HTML>" & vbCrLf & _
            "<HEAD>" & vbCrLf & _
            "<META HTTP-EQUIV=""Content-Style-Type"" CONTENT=""text/css"">" & vbCrLf
    If PathExists(FilePath & "head.js", True) Then Kill FilePath & "head.js"
    If IncludeCode Then
        html = html & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf & _
                        "<!--" & vbCrLf & _
                        frmMain.txtHead.Text & vbCrLf & _
                        "// -->" & vbCrLf & _
                        "</SCRIPT>" & vbCrLf
    Else
        html = html & "<SCRIPT LANGUAGE=""JavaScript"" SRC=""head.js""></SCRIPT>" & vbCrLf
        WriteFile FilePath & "head.js", frmMain.txtHead.Text
    End If
    
    If PathExists(FilePath & "custom.js", True) Then Kill FilePath & "custom.js"
    If IncludeCustom Then
        html = html & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf & _
                        "<!--" & vbCrLf & _
                        GenerateTree(nBase) & vbCrLf & _
                        "// -->" & vbCrLf & _
                        "</SCRIPT>" & vbCrLf
    Else
        html = html & "<SCRIPT LANGUAGE=""JavaScript"" SRC=""custom.js""></SCRIPT>" & vbCrLf
        WriteFile FilePath & "custom.js", GenerateTree(nBase)
    End If
    html = html & "</HEAD>" & vbCrLf & _
                    "<BODY" & IIf((nBase.Value("BackColor") = ""), "", " BGCOLOR='" & nBase.Value("BackColor") & "'") & _
                                IIf((nBase.Value("BackImage") = ""), "", " BACKGROUND='Media\" & nBase.Value("BackImage") & "'") & _
                                IIf((nBase.Value("FontColor") = ""), "", " TEXT='" & nBase.Value("FontColor") & "'") & ">" & vbCrLf

    If PathExists(FilePath & "body.js", True) Then Kill FilePath & "body.js"
    If IncludeCode Then
        html = html & "<SCRIPT LANGUAGE=""JavaScript"">" & vbCrLf & _
                        "<!--" & vbCrLf & _
                        frmMain.txtBody.Text & vbCrLf & _
                        "// -->" & vbCrLf & _
                        "</SCRIPT>" & vbCrLf
    Else
        html = html & "<SCRIPT LANGUAGE=""JavaScript"" SRC=""body.js""></SCRIPT>" & vbCrLf
        WriteFile FilePath & "body.js", frmMain.txtBody.Text

    End If
    
    html = html & "</BODY>" & vbCrLf & _
                    "</HTML>" & vbCrLf
                    
    WriteFile FilePath & "tree.html", html
    
End Function
Public Function ExportHTMLTree(ByRef nBase As clsItem, ByVal FilePath As String, ByVal IncludeCustom As Boolean, ByVal IncludeCode As Boolean) As Boolean
On Error GoTo catch

    Screen.MousePointer = 11
    
    XMLFilePath = FilePath
    XMLPadding = Chr(9)

    MakeDir FilePath
    
    WriteHTMLText nBase, FilePath, IncludeCustom, IncludeCode
    
    MakeDir FilePath & "Media"
    FileCopy MediaFolder & "202.gif", FilePath & "Media\202.gif"
    
    If nBase.Value("UsePlusMinus") Then
        MakeDir FilePath & "Media\PlusMinus"
        MakeDir FilePath & "Media\PlusMinus\" & nBase.Value("PlusMinusColor")
        FileCopy MediaFolder & "PlusMinus\" & nBase.Value("PlusMinusColor") & "\plus.gif", _
                    FilePath & "Media\PlusMinus\" & nBase.Value("PlusMinusColor") & "\plus.gif"
        FileCopy MediaFolder & "PlusMinus\" & nBase.Value("PlusMinusColor") & "\minus.gif", _
                    FilePath & "Media\PlusMinus\" & nBase.Value("PlusMinusColor") & "\minus.gif"
                    
    End If
    
    If nBase.Value("UseTreelines") Then
        MakeDir FilePath & "Media\Treelines"
        MakeDir FilePath & "Media\Treelines\" & nBase.Value("TreelineColor")
        
        FileCopy MediaFolder & "Treelines\" & nBase.Value("TreelineColor") & "\vline.gif", _
                    FilePath & "Media\Treelines\" & nBase.Value("TreelineColor") & "\vline.gif"
        FileCopy MediaFolder & "Treelines\" & nBase.Value("TreelineColor") & "\hline.gif", _
                    FilePath & "Media\Treelines\" & nBase.Value("TreelineColor") & "\hline.gif"
                    
        FileCopy MediaFolder & "Treelines\" & nBase.Value("TreelineColor") & "\btm.gif", _
                    FilePath & "Media\Treelines\" & nBase.Value("TreelineColor") & "\btm.gif"
        FileCopy MediaFolder & "Treelines\" & nBase.Value("TreelineColor") & "\mid.gif", _
                    FilePath & "Media\Treelines\" & nBase.Value("TreelineColor") & "\mid.gif"
        FileCopy MediaFolder & "Treelines\" & nBase.Value("TreelineColor") & "\top.gif", _
                    FilePath & "Media\Treelines\" & nBase.Value("TreelineColor") & "\top.gif"
                    
    End If
    
    If Not (nBase.Value("BackImage") = "") Then
        MakeDir FilePath & "Media\Backgrounds"
        
        FileCopy MediaFolder & nBase.Value("BackImage"), _
                    FilePath & "Media\" & nBase.Value("BackImage")
        
    End If
    
    Dim nItem As clsItem
    For Each nItem In nBase.SubItems
        ExportHTMLImages nItem, FilePath
    Next

catch:
    Screen.MousePointer = 0
    If Err Then
        MsgBox "WARNING!! Unable to export tree." & vbCrLf & "Reason: " & Err.Description, vbCritical
        Err.Clear
        ExportHTMLTree = False
    Else
        ExportHTMLTree = True
    End If

On Error GoTo 0
End Function
Public Function ExportHTMLImages(ByRef nBase As clsItem, ByVal FilePath As String) As Boolean
    ExportImage nBase, FilePath, "Collapsed"
    ExportImage nBase, FilePath, "Expanded"
    ExportImage nBase, FilePath, "Mouseover"
    ExportImage nBase, FilePath, "Mouseout"

    Dim nItem As clsItem
    For Each nItem In nBase.SubItems
        ExportHTMLImages nItem, FilePath
    Next
End Function
Public Function ExportImage(ByRef nBase As clsItem, ByVal FilePath As String, ByVal ImageName As String)
    If nBase.Value("Use" & ImageName) Then
        Dim nPic As String
        nPic = GetFilePath(nBase.Value(ImageName))
        MakeDir FilePath & "Media\" & GetFilePath(nBase.Value(ImageName))
        If Not PathExists(FilePath & "Media\" & nBase.Value(ImageName), True) Then
            FileCopy MediaFolder & nBase.Value(ImageName), FilePath & "Media\" & nBase.Value(ImageName)
        End If
    End If
    
End Function

Public Function SaveXMLFile(ByRef nBase As clsItem, ByVal FileName As String) As Boolean
On Error GoTo catch

    Screen.MousePointer = 11
    XMLFilePath = GetFilePath(FileName)
    XMLPadding = Chr(9)
    
    WriteFile FileName, "<?xml version=""1.0""?>" & vbCrLf & "<Tree>" & vbCrLf & _
                            nBase.XMLText & _
                            GetXMLImages(nBase) & _
                            "<ExportFolder>" & URLEncode(ExportFolder) & "</ExportFolder>" & _
                            "<ExportCustom>" & Ini.Setting("IncludeCustom") & "</ExportCustom>" & _
                            "<ExportCode>" & Ini.Setting("IncludeCode") & "</ExportCode>" & _
                            "</Tree>"
catch:
    Screen.MousePointer = 0
    If Err Then
        MsgBox "WARNING!! Unable to save tree file - [" & GetFileName(FileName) & "]" & vbCrLf & "Reason: " & Err.Description, vbCritical
        Err.Clear
        SaveXMLFile = False
    Else
        SaveXMLFile = True
    End If
On Error GoTo 0
End Function

Public Function OpenXMLFile(ByRef nBase As clsItem, ByVal FileName As String) As Boolean
On Error GoTo catch
    Screen.MousePointer = 11
    
    XMLFilePath = GetFilePath(FileName)
    
    Dim nXML As String
    
    nXML = ReadFile(FileName)
    
    nXML = Replace(nXML, "<?xml version=""1.0""?>", "")
    nXML = Replace(nXML, "<?xml version=""1.0"" ?>", "")
    nXML = Replace(nXML, "<? xml version=""1.0""?>", "")
    nXML = Replace(nXML, "<? xml version=""1.0"" ?>", "")
    
    Dim xml As New MSXML.DOMDocument
    xml.async = "false"
    xml.loadXML nXML

    If xml.parseError.errorCode = 0 Then
        If SafeKey(xml.childNodes(0).baseName) = "tree" Then
                
            Dim cImages As New Collection
            Dim child As MSXML.IXMLDOMNode
            For Each child In xml.childNodes(0).childNodes
                Select Case SafeKey(child.baseName)
                    Case "item"
                        nBase.XMLText = child.xml
                    Case "image"
                        CheckXMLImage child
                    Case "exportfolder"
                        ExportFolder = URLDecode(child.Text)
                    Case "exportcustom"
                    
                    Case "exportcode"
                    
                End Select
            Next
            ClearCollection cImages, False, True
            
        End If
        
        GetXMLImages nBase

    End If
    Set xml = Nothing

catch:
    Screen.MousePointer = 0
    If Err Then
        MsgBox "WARNING!! Unable to open tree file - [" & GetFileName(FileName) & "]" & vbCrLf & "Reason: " & Err.Description, vbCritical
        Err.Clear
        OpenXMLFile = False
    ElseIf Not (xml.parseError.errorCode = 0) Then
        MsgBox "WARNING!! Unable to open tree file - [" & GetFileName(FileName) & "]" & vbCrLf & "Reason: " & xml.parseError.reason, vbCritical
        Err.Clear
        OpenXMLFile = False
    Else
        OpenXMLFile = True
    End If
On Error GoTo 0
End Function

Private Function CheckXMLImage(ByRef xml As MSXML.IXMLDOMNode) As Boolean
    Dim imgName As String
    Dim imgData As String
    
    Dim child As MSXML.IXMLDOMNode
    For Each child In xml.childNodes
        Select Case SafeKey(child.baseName)
            Case "name"
                imgName = Trim(URLDecode(child.Text))
            Case "data"
                If Not (imgName = vbNullString) Then
                
                    If PathExists(MediaFolder & imgName, True) Then

                    Else
                        WriteFile MediaFolder & imgName, HesDecodeData(child.Text)

                    End If
                Else

                End If
        End Select
    Next

End Function

Private Function GetXMLImages(ByRef nBase As clsItem) As String
    Dim cImages As New Collection
    
    RecurseImages nBase, cImages
    
    Dim img As Variant
    Dim Ret As String
    For Each img In cImages
        Ret = Ret & CStr(img)
    Next
    ClearCollection cImages, False, True
    
    GetXMLImages = Ret
    
End Function

Private Sub RecurseImages(ByRef nBase As clsItem, ByRef cImages As Collection)
    
    Dim img As Variant
    For Each img In nBase.Images
        If Not (nBase.Key(img, True) = "none") Then
            If Not NodeExists(cImages, nBase.Key(img, True)) Then
                cImages.Add XMLPadding & "<Image>" & vbCrLf & _
                                XMLPadding & Chr(9) & "<Name>" & URLEncode(nBase.Value(img)) & "</Name>" & vbCrLf & _
                                XMLPadding & Chr(9) & "<Data>" & HesEncodeData(ReadFile(MediaFolder & nBase.Value(img))) & "</Data>" & vbCrLf & _
                                XMLPadding & "</Image>" & vbCrLf, nBase.Key(img, True)
                
            End If
        End If
    Next
    
    Dim nItem As clsItem
    For Each nItem In nBase.SubItems
        RecurseImages nItem, cImages
    Next
    
End Sub
