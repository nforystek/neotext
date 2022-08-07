VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN

Option Compare Binary

Private pFileName As String
Private pChanged As Boolean
Private pLoaded As Boolean
Private pAllowUI As Boolean
Private pLanguage As String

Private pItems As New Collection
Private pTemp As clsItem

Public Event Started()

Private WithEvents pDebugger As Form
Attribute pDebugger.VB_VarHelpID = -1
Private WithEvents pCompiler As frmProjectCompiler
Attribute pCompiler.VB_VarHelpID = -1

Friend Property Get AllowUI() As Boolean
    AllowUI = pAllowUI
End Property
Friend Property Let AllowUI(ByVal newVal As Boolean)
    pAllowUI = newVal
End Property

Friend Property Get Compiler() As Object
    If Not (pCompiler Is Nothing) Then
        Set Compiler = pCompiler.ScriptControl1
    End If
End Property

Friend Property Get Debugger() As Object
    If Not (pDebugger Is Nothing) Then
        Set Debugger = pDebugger.ScriptControl1
    End If
End Property

Public Property Get Changed() As Boolean
    Changed = pChanged
End Property
Public Property Let Changed(ByVal newVal As Boolean)
    pChanged = newVal
End Property

Public Property Get FileName() As String
    FileName = pFileName
End Property
Public Property Let FileName(ByVal newVal As String)
    pFileName = newVal
End Property

Public Property Get Loaded() As Boolean
    Loaded = pLoaded
End Property

Public Property Get Template() As clsItem
    Set Template = pTemp
End Property
Public Property Get IsTemplate(Optional ByVal fName As String = "") As Boolean
    IsTemplate = (GetFileExt(IIf(fName = "", pFileName, fName), True, False) = MaxScriptExt)
End Property

Public Property Get Language() As String
    Language = pLanguage
End Property
Friend Property Let Language(ByVal newVal As String)
    pLanguage = newVal
End Property

Public Property Get Items() As Collection
    Set Items = pItems
End Property
Public Function Remove(ByVal ItemName As String)
    Dim nItem As clsItem
    Set nItem = pItems(ItemName)
    pItems.Remove ItemName
    Set nItem = Nothing
End Function

Public Function IsRunning() As Boolean
    IsRunning = Not (pCompiler Is Nothing)
End Function

Public Function RunProject() As String
    On Error Resume Next
    If IsTemplate Then
        MsgBox "Templates are not able to be run individually.", vbInformation, AppName
    Else
        Set pDebugger = frmDebug
        Set pCompiler = New frmProjectCompiler
        pCompiler.RunProject Me
    End If
    If Not Err.Number = 0 Then
        Debug.Print "Unhandled exception: " & Err.Number & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Function
Public Function StopProject() As String
    If Not pCompiler Is Nothing Then
        pCompiler.StopProject
    End If
End Function

Public Function LoadFromFile(ByVal fName As String, Optional ByVal fScript As String = "") As Boolean
    On Error GoTo catch:
    
    Dim xml As New MSXML.DOMDocument
    xml.async = "false"
    xml.loadXML Replace(Replace(Replace(Replace(ReadFile(fName), "<?xml version=""1.0""?>", ""), "<?xml version=""1.0"" ?>", ""), "<? xml version=""1.0""?>", ""), "<? xml version=""1.0"" ?>", "")
            
    If (xml.childNodes(0).baseName = "Project") Or (xml.childNodes(0).baseName = "Item") Then
    
        ResetProject
        pFileName = fName
        
        Dim Item As clsItem
        Dim child As MSXML.IXMLDOMNode
        Dim URL As New NTAdvFTP61.URL
        
        If IsTemplate(fName) Then
        
            Set pTemp = New clsItem
            pTemp.LoadFromFile fName
            
        Else
            If Not (fScript = "") Then
                pLanguage = fScript
            End If
            
            For Each child In xml.childNodes(0).childNodes
                Select Case child.baseName
                    Case "Language"
                        pLanguage = URL.URLDecode(child.Text)
                    Case "File", "Item"
                        Set Item = New clsItem
                        With Item
                            If child.baseName = "File" Then
                                .ItemName = Replace(URL.URLDecode(child.Text), MaxScriptExt, "")
                                If InStr(.ItemName, ".") > 0 Then
                                    .ItemName = Mid(.ItemName, InStr(.ItemName, ".") + 1)
                                End If
                                .LoadFromFile URL.GetParentFolder(fName) & "\" & URL.URLDecode(child.Text), pLanguage
                            Else
                                .LoadFromXML child, pLanguage
                            End If
                        End With
                        pItems.Add Item, Item.ItemName
                        Set Item = Nothing
                End Select
            Next
            
        End If
        
        Set child = Nothing
        Set URL = Nothing
            
        pLoaded = True

        pChanged = False
            
    End If
    
    Set xml = Nothing
    
    LoadFromFile = True
Exit Function
catch:
    Err.Clear
    LoadFromFile = False
End Function

Public Function SaveToFile(ByVal fName As String) As Boolean
    On Error GoTo catch:
    
    Dim URL As New NTAdvFTP61.URL
    Dim cItem As clsItem
    Dim txt As String
    
    If IsTemplate(fName) Then
    
        WriteFile fName, "<?xml version=""1.0""?>" & vbCrLf & pTemp.GetXMLText & vbCrLf
        
    Else
    
        For Each cItem In pItems
            txt = txt & cItem.GetXMLText()
        Next
    
        WriteFile fName, "<?xml version=""1.0""?>" & vbCrLf & "<Project>" & vbCrLf & _
                "<Language>" & URL.URLEncode(pLanguage) & "</Language>" & vbCrLf & _
                txt & _
                "</Project>" & vbCrLf
    End If
    
    pFileName = fName
    pChanged = False
    Set URL = Nothing
    
    SaveToFile = True
Exit Function
catch:
    Err.Clear
    SaveToFile = False
End Function

Public Function ResetProject()
    StopProject
    
    ClearCollection pItems, True
    Set pTemp = Nothing
    
    pLoaded = False
    pLanguage = ""
    pFileName = ""
    pChanged = False
    
End Function

Private Sub Class_Initialize()
    pAllowUI = True
End Sub

Private Sub pCompiler_Finished()
    If Not (pCompiler Is Nothing) Then
        Unload pCompiler
        Set pCompiler = Nothing
    End If
    If Not (pDebugger Is Nothing) Then
        Set pDebugger = Nothing
    End If
End Sub

Friend Sub HandleError()
    pCompiler.HandleNormalError
End Sub


Attribute 