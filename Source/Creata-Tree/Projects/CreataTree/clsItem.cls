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

Private nFilePath As String

Private nPropSet As clsPropSet
Private nSubItems As clsItemSet

Private nKey As String
Public Property Get Key(Optional ByVal cPropKey As Variant = vbNullString, Optional ByVal ImageKey As Boolean = False) As String
    If cPropKey = vbNullString Then
        Key = nKey
    Else
        Key = nPropSet.Item(cPropKey).Key(ImageKey)
    End If
End Property
Public Function Display() As String
    If IsBase Then
        Display = "(" & Value("Label") & ")"
    Else
        If Value("UseText") And (Not (Value("Text") = vbNullString)) Then
            If Value("Text") = Value("Label") Then
                Display = Value("text")
            Else
                Display = Value("Text") & " - [" & Value("Label") & "]"
            End If
        Else
            Display = "[" & Value("Label") & "]"
        End If
    End If
End Function

Public Function BulletUse(ByVal Expanded As Boolean) As Boolean
    BulletUse = IIf(Expanded, Value("UseExpanded"), Value("UseCollapsed"))
End Function

Public Function Bullet(ByVal Expanded As Boolean) As String
    Dim bVal As String
    bVal = IIf(Expanded, Value("Expanded"), Value("Collapsed"))
    If (bVal = vbNullString) Then
        Bullet = BlankImageKey
    ElseIf PathExists(MediaFolder & bVal) Then
        Bullet = bVal
    Else
        Bullet = ErrorImageKey
    End If
End Function
Public Function BulletKey(ByVal Expanded As Boolean) As String
    Dim ret As String
    Dim bVal As String
    bVal = Bullet(Expanded)
    If Expanded Then
        If BulletUse(Expanded) Then
            ret = SafeKey(Bullet(Expanded))
        Else
            Expanded = Not Expanded
            bVal = Bullet(Expanded)
        End If
    End If
    
    If ret = vbNullString Then
        ret = SafeKey(Bullet(Expanded))
    End If
    
    BulletKey = ret
End Function

Public Property Get IsBase(Optional ByVal cSubKey As Variant = vbNullString) As Boolean
    If cSubKey = vbNullString Then
        IsBase = (InStr(LCase(Value("Type")), "base") > 0)
    Else
        IsBase = nSubItems.Item(cSubKey).IsBase
    End If
End Property
Public Property Get IsTemplate(Optional ByVal cSubKey As Variant = vbNullString) As Boolean
    If cSubKey = vbNullString Then
        IsTemplate = (InStr(LCase(Value("Type")), "template") > 0)
    Else
        IsTemplate = nSubItems.Item(cSubKey).IsTemplate
    End If
End Property

Public Property Get IsBlank(Optional ByVal cSubKey As Variant = vbNullString) As Boolean
    If cSubKey = vbNullString Then
        IsBlank = (InStr(LCase(Value("Type")), "blank") > 0)
    Else
        IsBlank = nSubItems.Item(cSubKey).IsBlank
    End If
End Property

Public Property Get Value(ByVal cPropKey As Variant) As Variant
    Value = nPropSet.Item(cPropKey).Value
End Property

Public Property Let Value(ByVal cPropKey As Variant, ByVal NewValue As Variant)
    nPropSet.Item(cPropKey).Value = NewValue
End Property

Public Property Get Cast(ByVal cPropKey As Variant) As Variant
    Cast = nPropSet.Item(cPropKey).Cast
End Property

Public Property Get SubItem(ByVal cKey As Variant) As clsItem
    Set SubItem = nSubItems.Item(cKey)
End Property
Public Property Get SubItems() As Collection
    Set SubItems = nSubItems.Items
End Property

Public Function AddItem(ByVal nXML As String) As clsItem
    
    Dim nItem As clsItem
    
    Set nItem = nSubItems.Items(nSubItems.Add(nXML))
    
    nItem.Value("Type") = "Item"
    
    Set AddItem = nItem
    Set nItem = Nothing
End Function

Public Sub RemoveItem(ByVal cKey As Variant)
    nSubItems.Remove cKey
End Sub

Private Sub Class_Initialize()
    nKey = SafeKey(GUID)
    Set nPropSet = New clsPropSet
    Set nSubItems = New clsItemSet
End Sub

Private Sub Class_Terminate()
    Set nSubItems = Nothing
    Set nPropSet = Nothing
End Sub

Public Property Get Images() As Collection
    Dim nProp As clsProp
    Dim ret As New Collection
    For Each nProp In nPropSet.Items
        If nProp.Cast = "i" Then
            ret.Add nProp.Key
        End If
    Next
    Set Images = ret
    Set ret = Nothing
End Property

Public Property Let XMLText(ByVal NewValue As String)
    
    Dim xml As New MSXML.DOMDocument
    xml.async = "false"
    xml.loadXML Replace(NewValue, "<?xml version=""1.0""?>", "")
    
    If xml.parseError.errorCode = 0 Then
        If SafeKey(xml.childNodes(0).baseName) = "item" Then
            Set nSubItems = Nothing
            Set nPropSet = Nothing
            
            Set nPropSet = New clsPropSet
            Set nSubItems = New clsItemSet
            
            Dim child As MSXML.IXMLDOMNode
            For Each child In xml.childNodes(0).childNodes
                Select Case SafeKey(child.baseName)
                    Case "file"
                        If PathExists(XMLFilePath & "\" & URLDecode(child.Text), True) Then
                            nSubItems.Add Replace(ReadFile(XMLFilePath & "\" & URLDecode(child.Text)), "<?xml version=""1.0""?>", "")
                        End If
                    Case "item"
                        nSubItems.Add child.xml
                    Case Else
                        nPropSet.Add child.baseName, URLDecode(child.Text)
                        
                End Select
            Next
            
        End If
    End If
    Set xml = Nothing
End Property

Public Property Get XMLText() As String
    Dim ret As String
    
    ret = ret & XMLPadding & "<Item>" & vbCrLf
    XMLPadding = XMLPadding & Chr(9)
    
    Dim nProp As clsProp
    For Each nProp In nPropSet.Items
        ret = ret & XMLPadding & nProp.XMLText & vbCrLf
    Next
    
    Dim nItem As clsItem
    For Each nItem In nSubItems.Items
        ret = ret & nItem.XMLText
    Next
    XMLPadding = Left(XMLPadding, Len(XMLPadding) - 1)
    
    ret = ret & XMLPadding & "</Item>" & vbCrLf
    
    XMLText = ret
End Property



